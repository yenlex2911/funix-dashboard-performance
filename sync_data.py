# -*- coding: utf-8 -*-
"""
FUNiX Way - Google Sheets Realtime Synchronization Script
Author: Agent CCO FX
Syncs all 4 menus + P1 report + OLAP Fact/Dimension data bundle according to FUNiX Way rules:
- Menu 1: Lead (Col K == 'FUNiX Way'), Doanh thu (Col R == 'FUNiX Way'), Sales & Rooms (Category Col G & H), Dimension 'Phòng Tuyển Sinh'.
- Menu 2: Tổng lead (Col N), P0 (DA), P1 (BA), P2 (EA), P3 (BN+CA), P4 (EN+FA), Other (CN).
- Menu 3: Sales Performance normalized to Category Col G & H (13 sales + 'Khác').
- Menu 4: Doanh thu Report with Dòng SP according to Category Col B, Hannix (Col G) & Phòng (Col H).
"""

import os
import sys
import json
import time
import re
import urllib.request
import urllib.parse
from google.oauth2 import service_account

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')
if hasattr(sys.stderr, 'reconfigure'):
    sys.stderr.reconfigure(encoding='utf-8')

CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_DIR = os.path.abspath(os.path.join(CURRENT_DIR, ".."))
ROOT_DIR = os.path.abspath(os.path.join(CURRENT_DIR, "../../.."))

SERVICE_ACCOUNT_PATHS = [
    os.path.join(CURRENT_DIR, "service_account.json"),
    os.path.join(ROOT_DIR, "service_account.json"),
    "D:/FUNiX/service_account.json",
    "C:/Users/ADMIN/Downloads/service_account.json"
]

DASHBOARD_DEV_HTML = os.path.join(CURRENT_DIR, "dashboard_dev.html")
GITHUB_REPO_DIR = CURRENT_DIR if os.path.exists(os.path.join(CURRENT_DIR, "index.html")) else "D:/FUNiX/funix-dashboard-performance"

SPREADSHEET_DOANH_THU = "1vFoo_Oe3zOZMDGYGA0R-6rCIOvZ-n8Z2iKDhg2CJ6gs"
SPREADSHEET_P1_REPORT = "1wgD_sas1F4-zq-ukey5pqALecFAl8jtYZD-uKq_pVT4"

SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets.readonly',
    'https://www.googleapis.com/auth/drive.readonly'
]

def parse_num(val):
    if not val: return 0
    if isinstance(val, (int, float)): return val
    s = str(val).strip().replace(' ', '')
    if s.endswith('%'):
        s = s.replace(',', '.').replace('%', '')
        try: return float(s) / 100.0
        except ValueError: return 0
    s = s.replace('.', '').replace(',', '.')
    try: return float(s)
    except ValueError: return 0

def get_google_access_token():
    # 1. Check environment variable (for GitHub Actions or cloud execution)
    sa_env = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON") or os.environ.get("GCP_SA_KEY")
    if sa_env:
        try:
            info = json.loads(sa_env)
            creds = service_account.Credentials.from_service_account_info(info, scopes=SCOPES)
            assertion = creds._make_authorization_grant_assertion()
            data = urllib.parse.urlencode({
                'grant_type': 'urn:ietf:params:oauth:grant-type:jwt-bearer',
                'assertion': assertion
            }).encode('utf-8')
            req = urllib.request.Request('https://oauth2.googleapis.com/token', data=data)
            with urllib.request.urlopen(req, timeout=30) as resp:
                return json.loads(resp.read().decode('utf-8'))['access_token']
        except Exception as e:
            print(f"[Cảnh báo] Lỗi xác thực từ biến môi trường: {e}")

    # 2. Check local file paths
    sa_file = None
    for p in SERVICE_ACCOUNT_PATHS:
        if os.path.exists(p):
            sa_file = p
            break
    if not sa_file:
        raise FileNotFoundError("Không tìm thấy file service_account.json hoặc biến môi trường GOOGLE_SERVICE_ACCOUNT_JSON")

    for attempt in range(5):
        try:
            creds = service_account.Credentials.from_service_account_file(sa_file, scopes=SCOPES)
            assertion = creds._make_authorization_grant_assertion()
            data = urllib.parse.urlencode({
                'grant_type': 'urn:ietf:params:oauth:grant-type:jwt-bearer',
                'assertion': assertion
            }).encode('utf-8')
            req = urllib.request.Request('https://oauth2.googleapis.com/token', data=data)
            with urllib.request.urlopen(req, timeout=30) as resp:
                return json.loads(resp.read().decode('utf-8'))['access_token']
        except Exception as e:
            time.sleep(2)
    raise Exception("Lỗi xác thực Google Service Account")

def fetch_sheet_values(token, spreadsheet_id, sheet_range):
    for attempt in range(4):
        try:
            enc = urllib.parse.quote(sheet_range)
            url = f"https://sheets.googleapis.com/v4/spreadsheets/{spreadsheet_id}/values/{enc}"
            req = urllib.request.Request(url, headers={'Authorization': f'Bearer {token}'})
            with urllib.request.urlopen(req, timeout=60) as resp:
                if resp.status != 200:
                    raise Exception(f"HTTP {resp.status}")
                return json.loads(resp.read().decode('utf-8')).get('values', [])
        except Exception:
            time.sleep(3)
    raise Exception(f"Lỗi đọc sheet {sheet_range}")

def main():
    print("=" * 70)
    print("  FUNIX WAY - TIẾN TRÌNH ĐỒNG BỘ DỮ LIỆU GOOGLE SHEETS TOÀN DIỆN V2")
    print("=" * 70)
    print("[1/6] Xác thực Google Service Account...")
    token = get_google_access_token()
    print("      -> Xác thực thành công!")

    # 1. Fetch Category sheet
    print(f"\n[2/6] Đang tải dữ liệu mapping từ tab 'Category' (GID: 770210720)...")
    cat_rows = fetch_sheet_values(token, SPREADSHEET_DOANH_THU, "Category!A1:H100")
    category_map = {}
    sales_norm_map = {}
    sales_list = []
    room_map = {}

    for r in cat_rows[1:]:
        prod = (r[0] if len(r) > 0 else '').strip()
        cat = (r[1] if len(r) > 1 else '').strip()
        khts = (r[2] if len(r) > 2 else '').strip()
        sale = (r[6] if len(r) > 6 else '').strip()
        phong = (r[7] if len(r) > 7 else '').strip()

        if cat:
            if prod and prod != "Product": category_map[prod.lower()] = cat
            if khts and khts != "Loại KHTS NE": category_map[khts.lower()] = cat

        if sale:
            low_sale = sale.lower()
            if low_sale not in sales_norm_map:
                sales_norm_map[low_sale] = sale
                sales_list.append(sale)
                room_map[low_sale] = phong or 'Khác'

    if 'khác' not in sales_norm_map:
        sales_norm_map['khác'] = 'Khác'
        sales_list.append('Khác')
        room_map['khác'] = 'Khác'

    def get_dong_sp(course, khts):
        raw = (khts or course or '').strip()
        if not raw: return 'Other'
        if 'branding' in raw.lower(): return 'Branding IT'
        low = raw.lower()
        if low in category_map:
            c = category_map[low]
            if c in ['P0', 'P1', 'P2', 'P3', 'P4']: return c
            if c == 'Bán chéo': return 'Other'
            return c
        if 'robotics' in low or 'wings' in low: return 'P0'
        if 'se' in raw.upper() or 'software engineering' in low or raw.startswith('P1'): return 'P1'
        if 'xup' in low or 'p2' in raw.upper(): return 'P2'
        if 'p3' in raw.upper(): return 'P3'
        if 'p4' in raw.upper(): return 'P4'
        return 'Other'

    def norm_sale(s):
        if not s: return 'Khác'
        low = str(s).strip().lower()
        return sales_norm_map.get(low, 'Khác')

    def get_room(sale):
        return room_map.get(sale.lower(), 'Khác')

    print(f"      -> Nạp thành công {len(category_map)} quy tắc mapping Category.")
    print(f"      -> Danh sách {len(sales_list)} Sales chuẩn từ Cột G: {sales_list}")

    # 2. Fetch Lead upto date (Menu 2)
    print(f"\n[3/6] Đang tải dữ liệu sheet 'Lead upto date' (Menu 2)...")
    mkt_rows = fetch_sheet_values(token, SPREADSHEET_DOANH_THU, "Lead upto date!A1:FM1000")
    mkt_data_by_month = {}
    mkt_data_by_day = {}
    for r in mkt_rows:
        label = (r[0] if len(r) > 0 else '').strip()
        def gc(idx): return parse_num(r[idx] if idx < len(r) else 0)

        # 1. Parse monthly summary rows (TỔNG YYYY-MM)
        if label and label.startswith('TỔNG 202'):
            parts = label.split(' ')
            mKey = parts[1]
            tot_lead = gc(13)
            p1_lead = gc(52)
            p3_lead = gc(65) + gc(78)
            p0_lead = gc(104)
            p2_lead = gc(130)
            p4_lead = gc(143) + gc(156)
            other_lead = gc(91)

            mkt_data_by_month[mKey] = {
                "month": mKey,
                "totalLead": tot_lead,
                "seedingLead": gc(14),
                "qcLead": gc(15),
                "orgLead": gc(16),
                "leadCL": gc(17),
                "leadCLQC": gc(19),
                "cost": gc(21),
                "revenueDate": gc(24),
                "revenueAll": gc(25),
                "p0Lead": p0_lead, "p0LeadCL": gc(108), "p0Cost": gc(112), "p0Rev": gc(116),
                "p1Lead": p1_lead, "p1LeadCL": gc(56), "p1Cost": gc(60), "p1Rev": gc(64),
                "p2Lead": p2_lead, "p2LeadCL": gc(134), "p2Cost": gc(138), "p2Rev": gc(142),
                "p3Lead": p3_lead, "p3LeadCL": gc(69) + gc(82), "p3Cost": gc(73) + gc(86), "p3Rev": gc(77) + gc(90),
                "p4Lead": p4_lead, "p4LeadCL": gc(147) + gc(160), "p4Cost": gc(151) + gc(164), "p4Rev": gc(155) + gc(168),
                "otherLead": other_lead, "otherLeadCL": gc(95), "otherCost": gc(99), "otherRev": gc(103),
                "seLead": p1_lead, "seLeadCL": gc(56), "seCost": gc(60), "seRev": gc(64),
                "daLead": p3_lead, "daLeadCL": gc(69) + gc(82), "daCost": gc(73) + gc(86), "daRev": gc(77) + gc(90),
                "xupLead": p2_lead, "xupLeadCL": gc(134), "xupCost": gc(138), "xupRev": gc(142)
            }

        # 2. Parse daily rows (DD/MM/YYYY)
        elif label and '/' in label:
            pts = label.split('/')
            if len(pts) == 3:
                try:
                    d_iso = f"{pts[2].zfill(4)}-{pts[1].zfill(2)}-{pts[0].zfill(2)}"
                    m_key = f"{pts[2].zfill(4)}-{pts[1].zfill(2)}"
                    mkt_data_by_day[d_iso] = {
                        "date": d_iso,
                        "month": m_key,
                        "totalLead": gc(13),
                        "seedingLead": gc(14),
                        "qcLead": gc(15),
                        "orgLead": gc(16),
                        "leadCL": gc(17),
                        "leadCLQC": gc(19),
                        "cost": gc(21),
                        "revenueDate": gc(24),
                        "revenueAll": gc(25),
                        "p0Lead": gc(104), "p0LeadCL": gc(108), "p0Cost": gc(112), "p0Rev": gc(116),
                        "p1Lead": gc(52), "p1LeadCL": gc(56), "p1Cost": gc(60), "p1Rev": gc(64),
                        "p2Lead": gc(130), "p2LeadCL": gc(134), "p2Cost": gc(138), "p2Rev": gc(142),
                        "p3Lead": gc(65) + gc(78), "p3LeadCL": gc(69) + gc(82), "p3Cost": gc(73) + gc(86), "p3Rev": gc(77) + gc(90),
                        "p4Lead": gc(143) + gc(156), "p4LeadCL": gc(147) + gc(160), "p4Cost": gc(151) + gc(164), "p4Rev": gc(155) + gc(168),
                        "otherLead": gc(91), "otherLeadCL": gc(95), "otherCost": gc(99), "otherRev": gc(103)
                    }
                except Exception:
                    pass
    print(f"      -> Phân tích thành công {len(mkt_data_by_month)} tháng và {len(mkt_data_by_day)} ngày dữ liệu Marketing.")

    # 3. Fetch hannix upto date (Menu 3)
    print(f"\n[4/6] Đang tải dữ liệu sheet 'hannix upto date' (Menu 3)...")
    sales_rows = fetch_sheet_values(token, SPREADSHEET_DOANH_THU, "hannix upto date!A1:AZ12000")
    sales_data_by_month = {}
    currentDateStr = ''
    for r in sales_rows[1:]:
        d = (r[0] if len(r) > 0 else '').strip() or (r[38] if len(r) > 38 else '').strip()
        if d: currentDateStr = d
        raw_name = (r[1] if len(r) > 1 else '').strip()
        if not raw_name: continue

        mKey = ''
        if '/' in currentDateStr:
            pts = currentDateStr.split('/')
            if len(pts) >= 2:
                m = int(pts[1]) if pts[1].isdigit() else 0
                if 1 <= m <= 12:
                    y = pts[2].zfill(4) if len(pts) >= 3 and pts[2].isdigit() else ('2025' if m == 12 else '2026')
                    mKey = f"{y}-{str(m).zfill(2)}"
        if not mKey: mKey = '2026-10'

        name = norm_sale(raw_name)
        room = get_room(name)

        if mKey not in sales_data_by_month:
            sales_data_by_month[mKey] = {}
        target = sales_data_by_month[mKey]
        if name not in target:
            target[name] = {
                "name": name, "room": room,
                "leadTouch": 0, "clientTouch": 0, "leadMKT": 0, "leadSeeding": 0, "leadKhac": 0,
                "dtTong": 0, "dtFW": 0, "dtCross": 0, "teTotal": 0, "teFW": 0, "teMKT": 0, "teSeeding": 0,
                "crHotleadSum": 0, "crCount": 0
            }

        dt_tong_str = (r[15] if len(r) > 15 else '').strip()
        if dt_tong_str == '':
            dt_tong_str = (r[39] if len(r) > 39 else '0').strip()
        dt_fw_str = (r[16] if len(r) > 16 else '').strip()
        if dt_fw_str == '':
            dt_fw_str = (r[40] if len(r) > 40 else '0').strip()
        dt_cross_str = (r[19] if len(r) > 19 else '').strip()
        if dt_cross_str == '':
            dt_cross_str = (r[41] if len(r) > 41 else '0').strip()

        cr_raw = parse_num((r[26] if len(r) > 26 else (r[24] if len(r) > 24 else 0)))
        target[name]["leadTouch"] += parse_num(r[4] if len(r) > 4 else 0)
        target[name]["clientTouch"] += parse_num(r[5] if len(r) > 5 else 0)
        target[name]["leadMKT"] += parse_num(r[8] if len(r) > 8 else 0)
        target[name]["leadSeeding"] += parse_num(r[9] if len(r) > 9 else 0)
        target[name]["leadKhac"] += parse_num(r[10] if len(r) > 10 else 0)
        target[name]["dtTong"] += parse_num(dt_tong_str)
        target[name]["dtFW"] += parse_num(dt_fw_str)
        target[name]["dtCross"] += parse_num(dt_cross_str)
        target[name]["teTotal"] += parse_num(r[20] if len(r) > 20 else 0)
        target[name]["teFW"] += parse_num(r[21] if len(r) > 21 else 0)
        target[name]["teMKT"] += parse_num(r[22] if len(r) > 22 else 0)
        target[name]["teSeeding"] += parse_num(r[23] if len(r) > 23 else 0)
        if cr_raw > 0:
            target[name]["crHotleadSum"] += cr_raw
            target[name]["crCount"] += 1
    print(f"      -> Phân loại thành công {len(sales_data_by_month)} tháng dữ liệu Sales.")

    # 4. Fetch Doanh thu (Menu 4 & OLAP)
    print(f"\n[5/6] Đang tải dữ liệu sheet 'Doanh thu' (Menu 4 & OLAP)...")
    dt_rows = fetch_sheet_values(token, SPREADSHEET_DOANH_THU, "Doanh thu!A1:AC10000")
    dt_transactions_by_month = {}
    total_valid_tx = 0
    for i in range(1, len(dt_rows)):
        row = dt_rows[i]
        if len(row) < 6: continue
        ngay_khop = (row[5] if len(row) > 5 else '').strip()
        if not ngay_khop: continue
        parts = ngay_khop.split('/')
        if len(parts) < 3: continue
        month_key = f"{parts[2].zfill(4)}-{parts[1].zfill(2)}"

        course = (row[2] if len(row) > 2 else '').strip()
        khts = (row[15] if len(row) > 15 else '').strip()
        raw_money = (row[4] if len(row) > 4 else '0').replace(',', '').replace('.', '')
        try: so_tien = float(raw_money) if raw_money else 0
        except ValueError: so_tien = 0

        dong_sp_val = get_dong_sp(course, khts)
        dong_sp_formatted = f"Dòng SP {dong_sp_val}" if dong_sp_val.startswith('P') else dong_sp_val
        hannix = norm_sale(row[7] if len(row) > 7 else '')
        phong = get_room(hannix)
        
        raw_cycle = (row[28] if len(row) > 28 else '0').strip()
        try: chu_ky = int(raw_cycle)
        except ValueError: chu_ky = 0

        tx = {
            "maNhapHoc": (row[0] if len(row) > 0 else '').strip(),
            "maThuPhi": (row[1] if len(row) > 1 else '').strip(),
            "khoaHoc": course,
            "mssv": (row[3] if len(row) > 3 else '').strip(),
            "soTien": so_tien,
            "ngayKhop": ngay_khop,
            "loaiNhapHoc": (row[6] if len(row) > 6 else '').strip(),
            "hannix": hannix,
            "phongHannix": phong,
            "nguonLead": (row[12] if len(row) > 12 else '').strip(),
            "nguoiChayQC": (row[13] if len(row) > 13 else '').strip(),
            "loaiKHTSLead": (row[14] if len(row) > 14 else '').strip(),
            "dongSPLead": get_dong_sp('', (row[14] if len(row) > 14 else '').strip()),
            "loaiKHTSNE": khts,
            "dongSP": dong_sp_formatted,
            "khoiNganh": (row[20] if len(row) > 20 else '').strip(),
            "sanPham": (row[17] if len(row) > 17 else '').strip(),
            "chuKyBanHang": chu_ky
        }
        if month_key not in dt_transactions_by_month:
            dt_transactions_by_month[month_key] = []
        dt_transactions_by_month[month_key].append(tx)
        total_valid_tx += 1
    print(f"      -> Nạp và chuẩn hóa thành công {total_valid_tx} giao dịch.")

    # 5. Fetch Raw Lead & Build OLAP Bundle
    print(f"\n[6/6] Đang tải dữ liệu sheet 'Raw Lead' & xây dựng OLAP Bundle...")
    raw_lead_rows = fetch_sheet_values(token, SPREADSHEET_DOANH_THU, "Raw Lead!A:X")

    bundle_path = os.path.join(CURRENT_DIR, "olap_data_bundle.js")
    if not os.path.exists(bundle_path):
        bundle_path = os.path.join(GITHUB_REPO_DIR, "olap_data_bundle.js")
    with open(bundle_path, 'r', encoding='utf-8') as f:
        txt = f.read()
        old_bundle = json.loads(txt[txt.find('{'):txt.rfind('}')+1])

    lookups = old_bundle['lookups']
    lookups['h'] = sales_list
    lookups['dept'] = ['GiangLTT', 'LanNH1', 'Khác']

    def get_idx(lst, val):
        val_str = str(val).strip()
        if not val_str: return 0
        if val_str in lst: return lst.index(val_str)
        lst.append(val_str)
        return len(lst) - 1

    new_lead_records = []
    mkt_lead_types_by_day = {}
    mkt_lead_types_by_month = {}
    for r in raw_lead_rows[1:]:
        if len(r) <= 10: continue
        if r[10].strip() != 'FUNiX Way': continue  # Col K strictly == 'FUNiX Way'
        created_date = (r[23] if len(r) > 23 else '').strip()
        parts = created_date.split(' ')[0].split('/')
        if len(parts) < 3: continue
        m_key = f"{parts[2]}-{parts[1].zfill(2)}"
        d_key = f"{parts[2]}-{parts[1].zfill(2)}-{parts[0].zfill(2)}"

        ch = (r[19] if len(r) > 19 else '').strip() or 'Other'
        ls = (r[20] if len(r) > 20 else '').strip() or 'Other'
        course = (r[8] if len(r) > 8 else '').strip()
        dsp = get_dong_sp(course, '')
        mid = (r[17] if len(r) > 17 else '').strip() or 'Khác'
        h = norm_sale(r[3] if len(r) > 3 else '')
        tg = (r[11] if len(r) > 11 else '').strip() or '(Hannix chưa chọn Đối tượng)'
        st = (r[18] if len(r) > 18 else '').strip() or 'L1 - Lead mới'

        date_idx = get_idx(lookups.setdefault('date', []), d_key)

        new_lead_records.append([
            get_idx(lookups['m'], m_key),
            get_idx(lookups['ch'], ch),
            get_idx(lookups['ls'], ls),
            get_idx(lookups['dsp'], dsp),
            get_idx(lookups['c'], course),
            get_idx(lookups['mid'], mid),
            lookups['h'].index(h),
            get_idx(lookups['tg'], tg),
            get_idx(lookups['st'], st),
            date_idx
        ])

        if d_key not in mkt_lead_types_by_day:
            mkt_lead_types_by_day[d_key] = {}
        mkt_lead_types_by_day[d_key][ls] = mkt_lead_types_by_day[d_key].get(ls, 0) + 1

        if m_key not in mkt_lead_types_by_month:
            mkt_lead_types_by_month[m_key] = {}
        mkt_lead_types_by_month[m_key][ls] = mkt_lead_types_by_month[m_key].get(ls, 0) + 1

    print(f"      -> Tổng hợp cơ cấu Loại Lead: {len(mkt_lead_types_by_month)} tháng và {len(mkt_lead_types_by_day)} ngày.")

    new_dt_records = []
    for idx_dt, r in enumerate(old_bundle['dt']):
        r_copy = list(r)
        sheet_row_idx = idx_dt + 1
        if sheet_row_idx < len(dt_rows):
            dt_r = dt_rows[sheet_row_idx]
            # Cột H sheet Doanh thu: Hannix thực hiện
            if len(dt_r) > 7:
                norm_h = norm_sale(dt_r[7])
                if norm_h in lookups['h']:
                    r_copy[6] = lookups['h'].index(norm_h)
                else:
                    r_copy[6] = lookups['h'].index('Khác')
            # Cột AC sheet Doanh thu: Chu kỳ bán hàng
            if len(dt_r) > 28:
                try:
                    cycle_val = int(str(dt_r[28]).strip())
                    r_copy[10] = cycle_val
                    # Bỏ qua Loại KHTS là Branding-IT... và Loại nhập học chứa Refund
                    loai_nh = (dt_r[6] if len(dt_r) > 6 else '').strip().lower()
                    loai_khts = (dt_r[15] if len(dt_r) > 15 else '').strip().lower()
                    course_name = (dt_r[2] if len(dt_r) > 2 else '').strip().lower()
                    is_excluded = ('refund' in loai_nh) or ('branding' in loai_khts) or ('branding' in course_name)
                    is_eligible = 1 if (cycle_val > 0 and not is_excluded) else 0
                    if len(r_copy) > 13:
                        r_copy[13] = is_eligible
                    else:
                        r_copy.append(is_eligible)
                except (ValueError, TypeError):
                    pass
            # Cột Z sheet Doanh thu (index 25): ngày tạo lead -> Tháng tạo lead
            lead_m = None
            if len(dt_r) > 25 and dt_r[25].strip():
                p = dt_r[25].strip().split(' ')[0].split('/')
                if len(p) == 3:
                    lead_m = f"{p[2]}-{p[1].zfill(2)}"
            if not lead_m:
                dt_date_str = lookups['date'][r_copy[12]] if 'date' in lookups and r_copy[12] < len(lookups['date']) else '2026-01-01'
                days = r_copy[10] if len(r_copy) > 10 else 0
                try:
                    dt_d = datetime.date.fromisoformat(dt_date_str)
                    lead_d = dt_d - datetime.timedelta(days=days)
                    lead_m = lead_d.strftime('%Y-%m')
                except Exception:
                    lead_m = lookups['m'][r_copy[0]]
            if 'lm' not in lookups:
                lookups['lm'] = list(lookups['m'])
            if lead_m not in lookups['lm']:
                lookups['lm'].append(lead_m)
            lm_idx = lookups['lm'].index(lead_m)
            if len(r_copy) > 14:
                r_copy[14] = lm_idx
            else:
                r_copy.append(lm_idx)
            # Cột O sheet Doanh thu (index 14): Loại KHTS Lead & Dòng SP Lead
            lead_khts = (dt_r[14] if len(dt_r) > 14 else '').strip()
            dsp_lead = get_dong_sp('', lead_khts)
            khts_lead_idx = get_idx(lookups['c'], lead_khts)
            dsp_lead_idx = get_idx(lookups['dsp'], dsp_lead)
            if len(r_copy) > 15:
                r_copy[15] = khts_lead_idx
            else:
                r_copy.append(khts_lead_idx)
            if len(r_copy) > 16:
                r_copy[16] = dsp_lead_idx
            else:
                r_copy.append(dsp_lead_idx)
        new_dt_records.append(r_copy)

    olap_bundle_obj = {
        "lookups": lookups,
        "dt": new_dt_records,
        "lead": new_lead_records
    }
    olap_js_content = 'window.OLAP_BUNDLE = ' + json.dumps(olap_bundle_obj, ensure_ascii=False) + ';\n'
    olap_json_content = json.dumps(olap_bundle_obj, ensure_ascii=False)

    for target_dir in [CURRENT_DIR, GITHUB_REPO_DIR]:
        if os.path.exists(target_dir):
            with open(os.path.join(target_dir, "olap_data_bundle.js"), 'w', encoding='utf-8') as f:
                f.write(olap_js_content)
            with open(os.path.join(target_dir, "olap_data_bundle.json"), 'w', encoding='utf-8') as f:
                f.write(olap_json_content)

    print(f"      -> Ghi nhận thành công OLAP Bundle: {len(new_lead_records):,} leads và {len(new_dt_records):,} GD!")

    # 6. Fetch P1-SE report
    p1_rows = fetch_sheet_values(token, SPREADSHEET_P1_REPORT, "báo cáo P1!A1:AZ35")
    def parse_p1_report_rows(rows):
        monthly_data = []
        for i in range(4, len(rows)):
            r = rows[i]
            if len(r) < 2: continue
            month = (r[1] if len(r) > 1 else '').strip()
            if not month or '/' not in month: continue
            parts = month.split('/')
            m_num = int(parts[0]) if len(parts) > 0 and parts[0].isdigit() else 0
            y_num = int(parts[1]) if len(parts) > 1 and parts[1].isdigit() else 0
            def gc(idx): return parse_num(r[idx] if idx < len(r) else 0)
            item = {
                "month": month, "year": y_num, "monthNum": m_num,
                "tongQuan": {
                    "chiPhi": gc(2), "lead": gc(3), "doanhThu": gc(4), "teNew": gc(5), "cr": gc(6)
                },
                "leadTheoNguon": {
                    "facebook": gc(7), "google": gc(8), "other": gc(9), "seo": gc(10), "prGioiThieu": gc(11), "gioiThieu": gc(12), "qcKhac": gc(13), "organicKhac": gc(14), "facebook2": gc(15)
                },
                "doanhThuTheoKhoa": {
                    "p1": gc(16), "p3": gc(17), "brandingIt": gc(18), "xup": gc(19), "funixWings": gc(20), "coding": gc(21), "tutor": gc(22), "mentor": gc(23)
                },
                "doanhThuTheoLoaiLead": {
                    "p1": gc(24), "p3": gc(25), "brandingIt": gc(26), "xup": gc(27), "funixWings": gc(28), "coding": gc(29)
                },
                "doanhThuTheoNguon": {
                    "facebook": gc(30), "google": gc(31), "other": gc(32), "seo": gc(33), "prGioiThieu": gc(34), "gioiThieu": gc(35), "qcKhac": gc(36), "organicKhac": gc(37), "facebook2": gc(38)
                }
            }
            monthly_data.append(item)
        return monthly_data
    p1_data = parse_p1_report_rows(p1_rows)

    # 7. Update HTML files in both directories
    html_targets = [
        DASHBOARD_DEV_HTML,
        os.path.join(GITHUB_REPO_DIR, "index.html")
    ]
    for h_path in html_targets:
        if not os.path.exists(h_path): continue
        with open(h_path, 'r', encoding='utf-8') as f:
            content = f.read()

        app_data_start = content.find('const APP_DATA = {')
        if app_data_start != -1:
            app_data_end = content.find(';\n', app_data_start)
            old_json_str = content[app_data_start + len('const APP_DATA = '):app_data_end]
            app_data_obj = json.loads(old_json_str)
            app_data_obj['mktDataByMonth'] = mkt_data_by_month
            app_data_obj['mktDataByDay'] = mkt_data_by_day
            app_data_obj['mktLeadTypesByDay'] = mkt_lead_types_by_day
            app_data_obj['mktLeadTypesByMonth'] = mkt_lead_types_by_month
            app_data_obj['salesDataByMonth'] = sales_data_by_month
            app_data_obj['dtTransactionsByMonth'] = dt_transactions_by_month
            app_data_obj['p1ReportData'] = p1_data
            app_data_obj['standardSales'] = sales_list
            app_data_obj['roomMap'] = room_map
            new_json_str = json.dumps(app_data_obj, ensure_ascii=False)
            content = content[:app_data_start + len('const APP_DATA = ')] + new_json_str + content[app_data_end:]

        # Auto-update month dropdown text to latest active October day
        active_days = sorted([d for d, v in mkt_data_by_day.items() if d.startswith('2026-10') and ((v.get('totalLead') or 0) > 0 or (v.get('cost') or 0) > 0)])
        if active_days:
            latest_day_str = active_days[-1][8:10] + "/10"
            content = re.sub(
                r'Tháng 10/2026 \(Lũy kế đến \d{2}/\d{2} - Hiện tại\)',
                f'Tháng 10/2026 (Lũy kế đến {latest_day_str} - Hiện tại)',
                content
            )

        with open(h_path, 'w', encoding='utf-8') as f:
            f.write(content)
        print(f"      -> Cập nhật HTML: {h_path}")

    print("\n" + "=" * 70)
    print("  HOÀN TẤT ĐỒNG BỘ TOÀN DIỆN!")
    print("=" * 70)

if __name__ == '__main__':
    try:
        main()
    except Exception as e:
        print(f"\n[LỖI]: {e}", file=sys.stderr)
        sys.exit(1)
