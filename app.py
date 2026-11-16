import streamlit as st
import openpyxl
from openpyxl.utils import get_column_letter
from copy import copy
import io
import datetime
import re
import zipfile
import os

st.set_page_config(page_title="Customer Data Batch Auto-Filler", layout="wide")
st.title("📦 ระบบเติมข้อมูลส่งลูกค้าอัตโนมัติ (แยกตามโมเดล + กรอง WW และวันที่)")

# ==========================================
# ฟังก์ชันแปลงช่วงแถว และสกัด WW / วันที่ / Lab ID
# ==========================================
def expand_ranges(ranges_list):
    rows = []
    for item in ranges_list:
        if isinstance(item, tuple):
            rows.extend(range(item[0], item[1] + 1))
        else:
            rows.append(item)
    return sorted(list(set(rows)))

def extract_ww_num(text):
    if not text:
        return None
    match = re.search(r'WW\.?\s*0*(\d+)', str(text), re.IGNORECASE)
    return int(match.group(1)) if match else None

def extract_unique_lot_or_lab_id(text):
    """ดึงเฉพาะรหัส Lab No. หรือ Lot No. ที่ไม่ซ้ำกัน (เช่น LAB26001868, HNRJ1106A1)"""
    if not text:
        return None
    s = str(text).strip()
    # ค้นหารหัส LAB... หรือ HNRJ...
    m = re.search(r'(LAB\d{6,}|HNRJ\w+)', s, re.IGNORECASE)
    if m:
        return m.group(1).upper()
    return None

def extract_date_val(val):
    if val is None:
        return None
    if isinstance(val, datetime.datetime):
        return val.date()
    if isinstance(val, datetime.date):
        return val
    if isinstance(val, (int, float)):
        if 40000 < val < 55000:
            return (datetime.datetime(1899, 12, 30) + datetime.timedelta(days=int(val))).date()

    s = str(val).strip()
    if not s or len(s) < 4:
        return None

    # ค้นหาวันที่ในข้อความยาวๆ เช่น "washing date 15 Oct'26" หรือ "14-Oct-26"
    m = re.search(r'(\d{1,2})[\-\/\s]+([A-Za-z]{3})[\-\/\s\']*(\d{2,4})', s)
    if m:
        day, mon, yr = m.group(1), m.group(2), m.group(3)
        clean_d = f"{day}-{mon}-{yr}"
        for fmt in ["%d-%b-%y", "%d-%b-%Y"]:
            try:
                return datetime.datetime.strptime(clean_d, fmt).date()
            except ValueError:
                pass

    date_formats = [
        "%d-%b-%y", "%d-%b-%Y", "%d/%b/%y", "%d/%b/%Y", "%d %b %y", "%d %b %Y",
        "%d-%m-%y", "%d-%m-%Y", "%d/%m/%y", "%d/%m/%Y",
        "%Y-%m-%d", "%Y/%m/%d", "%b-%y", "%b %y", "%b-%Y", "%B-%y"
    ]
    for fmt in date_formats:
        try:
            return datetime.datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    return None

# ==========================================
# ตั้งค่าพิกัดแถวและ Spec แยกตาม 4 โปรไฟล์โมเดล
# ==========================================
MODEL_PROFILES = {
    "DORADO": {
        "LPC": {
            "step": 1, "header_rows": [29, 30, 31], "start_col": 2,
            "copy_rows": expand_ranges([(29, 34), 44]),
            "spec_checks": [(32, 41), (33, 41), (34, 41)]
        },
        "IC": {
            "step": 2, "header_rows": [19, 20], "start_col": 2,
            "copy_rows": expand_ranges([(19, 27)]),
            "spec_checks": [(28, 32)]
        },
        "FTIR": {
            "step": 2, "header_rows": [5, 6], "start_col": 2,
            "copy_rows": expand_ranges([5, 6, 7, 9]),
            "spec_checks": [(7, 8), (9, 10)]
        },
        "APA": {
            "step": 2, "header_rows": [38, 39, 40], "start_col": 4,
            "copy_rows": expand_ranges([(38, 58), (62, 64), (66, 71), 75, 77, (79, 94), (98, 116), 118, (120, 129), 131]),
            "spec_checks": [(58, 59), (62, 65), (63, 65), (64, 65), (71, 72), (94, 95), (118, 119)]
        },
        "DHS": {
            "step": 1, "header_rows": [51, 52, 53], "start_col": 4,
            "copy_rows": expand_ranges([(51, 54), 57, 60, 63, 66, 69, 72, 75, 78, 81, 84, 87, 90, 93]),
            "spec_checks": [(r, r + 1) for r in [54, 57, 60, 63, 66, 69, 72, 75, 78, 81, 84, 87, 90, 93]]
        },
        "GCMS": {
            "step": 2, "header_rows": [40, 41, 42], "start_col": 4,
            "copy_rows": expand_ranges([40, (43, 45), 49, 53]),
            "spec_checks": [(r, r + 1) for r in [43, 44, 45, 49, 53]]
        },
        "GCMS RHC": {
            "step": 2, "header_rows": [36, 37, 38], "start_col": 4,
            "copy_rows": expand_ranges([36, 37, (39, 41), 45, 49, (53, 55)]),
            "spec_checks": [(41, 42), (45, 46), (49, 50)]
        },
        "Ghost Test": {
            "step": 1, "header_rows": [4, 5, 6], "start_col": 2,
            "copy_rows": expand_ranges([(4, 9)]),
            "spec_checks": []
        }
    },

    "MARLIN_V4": {
        "LPC": {
            "step": 1, "header_rows": [34, 35, 36, 37], "start_col": 2,
            "copy_rows": expand_ranges([(34, 37), (39, 41), 49]),
            "spec_checks": [(39, 46), (40, 46), (41, 46)]
        },
        "IC": {
            "step": 2, "header_rows": [20, 21, 22, 23], "start_col": 2,
            "copy_rows": expand_ranges([(20, 31)]),
            "spec_checks": [(31, 35)]
        },
        "FTIR": {
            "step": 2, "header_rows": [3, 4, 5], "start_col": 2,
            "copy_rows": expand_ranges([(3, 6)]),
            "spec_checks": [(6, 7)]
        },
        "DHS": {
            "step": 1, "header_rows": [53, 54, 55], "start_col": 4,
            "copy_rows": expand_ranges([(53, 56), 59, 62, 65, 68, 71, 74, 77, 80, 83, 86, 89, 92, 95]),
            "spec_checks": [(r, r + 1) for r in [56, 59, 62, 65, 68, 71, 74, 77, 80, 83, 86, 89, 92, 95]]
        },
        "APA": {
            "step": 2, "header_rows": [39, 40, 41, 42], "start_col": 4,
            "copy_rows": expand_ranges([(39, 42), (44, 60), (66, 68), (70, 75), 81, 83, (85, 95), (101, 119), 121, (127, 135), 137]),
            "spec_checks": [(60, 61), (75, 76), (95, 96), (121, 122), (135, 136)]
        },
        "GCMS": {
            "step": 2, "header_rows": [41, 42, 43, 44], "start_col": 4,
            "copy_rows": expand_ranges([(41, 45), 49, 53]),
            "spec_checks": [(45, 46), (53, 54)]
        },
        "GCMS RHC": {
            "step": 2, "header_rows": [37, 38, 39, 40], "start_col": 4,
            "copy_rows": expand_ranges([(37, 41), 47, 53, 59, (64, 65)]),
            "spec_checks": [(41, 42), (47, 48), (53, 54)]
        },
        "Ghost Test": {
            "step": 1, "header_rows": [11, 12, 13, 14], "start_col": 2,
            "copy_rows": expand_ranges([(11, 17)]),
            "spec_checks": []
        }
    },

    "SUMMIT": {
        "LPC": {
            "step": 1, "header_rows": [33, 34, 35], "start_col": 2,
            "copy_rows": expand_ranges([(33, 38), 46]),
            "spec_checks": [(36, 43), (37, 43), (38, 43)]
        },
        "IC": {
            "step": 2, "header_rows": [19, 20], "start_col": 2,
            "copy_rows": expand_ranges([(19, 28)]),
            "spec_checks": [(28, 32)]
        },
        "FTIR": {
            "step": 2, "header_rows": [3, 4, 5], "start_col": 2,
            "copy_rows": expand_ranges([(3, 6), 8]),
            "spec_checks": [(6, 7), (8, 9)]
        },
        "DHS": {
            "step": 1, "header_rows": [51, 52, 53], "start_col": 4,
            "copy_rows": expand_ranges([(51, 54), 57, 60, 63, 66, 69, 72, 75, 78, 81, 84, 87, 90, 93]),
            "spec_checks": [(r, r + 1) for r in [54, 57, 60, 63, 66, 69, 72, 75, 78, 81, 84, 87, 90, 93]]
        },
        "APA": {
            "step": 2, "header_rows": [38, 39, 40], "start_col": 4,
            "copy_rows": expand_ranges([(38, 40), (42, 58), (64, 66), (68, 73), 79, 81, (83, 95), (101, 119), 121, (127, 135), 137]),
            "spec_checks": [(58, 59), (73, 74), (95, 96), (121, 122)]
        },
        "GCMS": {
            "step": 2, "header_rows": [40, 41, 42], "start_col": 4,
            "copy_rows": expand_ranges([(40, 43), 47, 51]),
            "spec_checks": [(43, 44), (51, 52)]
        },
        "GCMS RHC": {
            "step": 2, "header_rows": [36, 37, 38], "start_col": 4,
            "copy_rows": expand_ranges([(36, 39), 45, 51, 57, (62, 63)]),
            "spec_checks": [(39, 40), (45, 46), (51, 52)]
        },
        "Ghost Test": {
            "step": 1, "header_rows": [10, 11], "start_col": 2,
            "copy_rows": expand_ranges([(10, 13)]),
            "spec_checks": []
        }
    },

    "MARLIN_OTHER": {
        "LPC": {
            "step": 1, "header_rows": [34, 35, 36, 37, 38], "start_col": 2,
            "copy_rows": expand_ranges([(34, 41), 51]),
            "spec_checks": [(39, 48), (40, 48), (41, 48)]
        },
        "IC": {
            "step": 2, "header_rows": [18, 19, 20, 21, 22], "start_col": 2,
            "copy_rows": expand_ranges([(18, 30)]),
            "spec_checks": [(30, 34)]
        },
        "FTIR": {
            "step": 2, "header_rows": [3, 4, 5, 6], "start_col": 2,
            "copy_rows": expand_ranges([(3, 7), 9]),
            "spec_checks": [(7, 8), (9, 10)]
        },
        "DHS": {
            "step": 1, "header_rows": [52, 53, 54, 55], "start_col": 4,
            "copy_rows": expand_ranges([(52, 56), 59, 62, 65, 68, 71, 74, 77, 80, 83, 86, 89, 92, 95]),
            "spec_checks": [(r, r + 1) for r in [56, 59, 62, 65, 68, 71, 74, 77, 80, 83, 86, 89, 92, 95]]
        },
        "APA": {
            "step": 2, "header_rows": [39, 40, 41, 42], "start_col": 4,
            "copy_rows": expand_ranges([(39, 42), (44, 60), (64, 66), (68, 73), 77, 79, (81, 94), (98, 116), 118, (120, 128), 130]),
            "spec_checks": [(60, 61), (73, 74), (94, 95), (118, 119)]
        },
        "GCMS": {
            "step": 2, "header_rows": [41, 42, 43, 44], "start_col": 4,
            "copy_rows": expand_ranges([(41, 45), 49, 53]),
            "spec_checks": [(45, 46), (53, 54)]
        },
        "GCMS RHC": {
            "step": 2, "header_rows": [37, 38, 39, 40], "start_col": 4,
            "copy_rows": expand_ranges([(37, 41), 45, 49, (53, 55)]),
            "spec_checks": [(41, 42), (45, 46), (49, 50)]
        },
        "Ghost Test": {
            "step": 1, "header_rows": [10, 11, 12, 13], "start_col": 2,
            "copy_rows": expand_ranges([(10, 17)]),
            "spec_checks": []
        }
    }
}

def detect_model_profile(filename):
    fn = filename.upper()
    if "GZ8427V4" in fn:
        return "MARLIN_V4", MODEL_PROFILES["MARLIN_V4"]
    elif "MARLIN" in fn:
        return "MARLIN_OTHER", MODEL_PROFILES["MARLIN_OTHER"]
    elif "SUMMIT" in fn:
        return "SUMMIT", MODEL_PROFILES["SUMMIT"]
    else:
        return "DORADO", MODEL_PROFILES["DORADO"]

def find_sheet_in_wb(wb, target_name):
    for s in wb.sheetnames:
        if s.strip().lower() == target_name.strip().lower():
            return wb[s]
    return None

def normalize_filename(name):
    s = os.path.splitext(name)[0].lower()
    for word in ["ใหม่", "copy of", "updated_", "raw_", "ctm_", "customer_"]:
        s = s.replace(word, "")
    return re.sub(r'[\s_\-]+', '', s)

# ==========================================
# Helper Functions สำหรับอ่านและตรวจสอบข้อมูลจริง
# ==========================================
def format_val(val):
    if val is None:
        return ""
    if isinstance(val, (datetime.datetime, datetime.date)):
        return val.strftime("%d-%b-%y")
    if isinstance(val, (int, float)):
        if 40000 < val < 55000:
            d = (datetime.datetime(1899, 12, 30) + datetime.timedelta(days=int(val))).date()
            return d.strftime("%d-%b-%y")
    return re.sub(r'\s+', ' ', str(val).strip())

def to_float(val):
    if val is None:
        return None
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip().replace(",", "")
    if s == "" or s.upper() in ["NA", "N/A", "ND", "-", "PASS", "FAIL"]:
        return None
    if s.startswith("#"):
        return None
    try:
        return float(s)
    except:
        return None

def find_spec_value(ws_raw, ws_cust, spec_row, col):
    for ws in (ws_raw, ws_cust):
        v = to_float(ws.cell(row=spec_row, column=col).value) if ws else None
        if v is not None:
            return v
    for ws in (ws_cust, ws_raw):
        if not ws:
            continue
        for c in range(2, min(ws.max_column + 1, 60)):
            v = to_float(ws.cell(row=spec_row, column=c).value)
            if v is not None:
                return v
    return None

def evaluate_entry_spec(ws_raw, ws_cust, start_col, config):
    failures = []
    for val_row, spec_row in config.get("spec_checks", []):
        for offset in range(config["step"]):
            c = start_col + offset
            val_num = to_float(ws_raw.cell(row=val_row, column=c).value)
            if val_num is None:
                continue
            spec_num = find_spec_value(ws_raw, ws_cust, spec_row, c)
            if spec_num is not None and val_num > spec_num:
                failures.append(f"Row {val_row} ({get_column_letter(c)}{val_row}={val_num:g} > Spec {spec_num:g})")
    return failures

def is_valid_header_token(val):
    if val is None:
        return False
    s = str(val).strip()
    if s == "" or s.startswith("=") or s.startswith("#"):
        return False
    ignore_list = [
        "WW.", "WW", "DATE", "ION", "SPEC.", "SPEC", "NG/CM2", "NG/CM²",
        "SILICONE", "HYDROCARBON", "MOLD", "REMARK", "PASS", "FAIL", "ND",
        "PART NAME", "MATERIAL", "PART NO.", "PART NO", "OUTGASSED COMPOUNDS"
    ]
    return s.upper() not in ignore_list

def get_data_rows(config):
    h_set = set(config["header_rows"])
    d_rows = [r for r in config["copy_rows"] if r not in h_set]
    return d_rows if d_rows else config["copy_rows"]

def column_has_actual_measurement(ws, col, config):
    data_rows = get_data_rows(config)
    step = config["step"]

    for r in data_rows:
        for offset in range(step):
            val = ws.cell(row=r, column=col + offset).value
            if val is None:
                continue
            s = str(val).strip()
            if s == "" or s.startswith("=") or s.startswith("#"):
                continue
            if to_float(val) is not None:
                return True
            if s.upper() in ["ND", "PASS", "FAIL"] or "GHOST" in s.upper():
                return True
    return False

def scan_all_valid_entries(ws_target, ws_cust_ref, config, check_spec=False):
    entries = []
    header_rows = config["header_rows"]
    data_rows = get_data_rows(config)
    step = config["step"]
    col = config["start_col"]
    max_col = ws_target.max_column

    last_seen_ww = None
    last_seen_date = None

    while col <= max_col:
        col_ww = None
        col_date = None
        unique_id = None
        tokens = []

        # กวาดหาวันที่, WW, และรหัส LAB/Lot ในแถวหัวตาราง
        search_rows = sorted(list(set(header_rows + [min(header_rows) - 1, max(header_rows) + 1])))
        for r in search_rows:
            if r < 1:
                continue
            for offset in range(step):
                raw_v = ws_target.cell(row=r, column=col + offset).value
                if raw_v is not None:
                    s_raw = str(raw_v)
                    if col_ww is None:
                        col_ww = extract_ww_num(s_raw)
                    # ให้ความสำคัญกับวันที่จากแถวหัวตารางโดยตรง
                    d_extracted = extract_date_val(raw_v)
                    if d_extracted is not None:
                        col_date = d_extracted
                    if unique_id is None:
                        unique_id = extract_unique_lot_or_lab_id(s_raw)

                    if r in header_rows and offset == 0 and is_valid_header_token(raw_v):
                        s_val = format_val(raw_v)
                        if s_val and s_val not in tokens:
                            tokens.append(s_val)

        if column_has_actual_measurement(ws_target, col, config):
            if col_ww is not None:
                last_seen_ww = col_ww
            else:
                col_ww = last_seen_ww

            if col_date is not None:
                last_seen_date = col_date
            else:
                col_date = last_seen_date

            # สร้างลายเซ็นจากตัวเลขผลเทสทั้งหมดในคอลัมน์นั้น (ไม่จำกัดแค่ 6 แถวแรก เพื่อให้ APA/DHS แม่นยำ 100%)
            data_vals = []
            for r in data_rows:
                for offset in range(step):
                    v = ws_target.cell(row=r, column=col + offset).value
                    fv = to_float(v)
                    if fv is not None:
                        data_vals.append(f"{r}_{offset}:{fv:.4f}")
                    elif v is not None and not str(v).startswith("#") and str(v).strip() != "":
                        data_vals.append(f"{r}_{offset}:{str(v).strip()}")
            data_sig = "|".join(data_vals)

            full_sig = " | ".join(tokens) if tokens else f"Col {get_column_letter(col)}"
            failures = evaluate_entry_spec(ws_target, ws_cust_ref, col, config) if check_spec else []

            date_str = col_date.strftime("%d-%b-%y") if col_date else "ไม่ระบุวันที่"
            ww_str = f"WW{col_ww:02d} | " if col_ww is not None else ""
            base_label = f"[Col {get_column_letter(col)} | {ww_str}{date_str}] {full_sig}"

            entries.append({
                "col": col,
                "unique_id": unique_id,
                "tokens": tokens,
                "full_sig": full_sig,
                "data_sig": data_sig,
                "label": f"⚠️ {base_label} — หลุด Spec: {', '.join(failures)}" if failures else f"✅ {base_label}",
                "is_ng": len(failures) > 0,
                "failures": failures,
                "ww_num": col_ww,
                "date_val": col_date
            })
            col += step
        else:
            col += 1

    return entries

def filter_new_raw_entries(cust_entries, raw_entries):
    """
    หาชุดข้อมูลในไฟล์ดิบที่ยังไม่ได้ส่งให้ลูกค้า (แก้ไขบั๊กชื่อโมเดลซ้ำกันทุกคอลัมน์):
    1. เช็คจากรหัส Unique ID (เช่น LAB26001868 หรือ HNRJ1106A1) หรือ ค่าตัวเลขผลเทสทั้งหมด (data_sig)
    2. ดึงทุกคอลัมน์ในไฟล์ดิบที่อยู่ถัดจากตัวสุดท้ายของไฟล์ลูกค้าเป็นต้นไป
    """
    if not cust_entries:
        return raw_entries
    if not raw_entries:
        return []

    last_c = cust_entries[-1]

    # 1. เช็คหาตำแหน่งคอลัมน์ในไฟล์ดิบที่ตรงกับตัวสุดท้ายของไฟล์ลูกค้า
    match_indices = []
    for i, r_ent in enumerate(raw_entries):
        # ถ้ารหัส LAB / HNRJ ตรงกันเป๊ะ
        if last_c["unique_id"] and r_ent["unique_id"] and last_c["unique_id"] == r_ent["unique_id"]:
            match_indices.append(i)
        # หรือถ้าชุดตัวเลขผลเทสทั้งหมดในคอลัมน์ตรงกันเป๊ะ
        elif last_c["data_sig"] and r_ent["data_sig"] == last_c["data_sig"]:
            match_indices.append(i)

    if match_indices:
        cutoff_idx = match_indices[-1]
        return raw_entries[cutoff_idx + 1:]

    # 2. ถ้าหาตัวสุดท้ายไม่เจอในไฟล์ดิบ ให้ดึงทุกคอลัมน์ที่มีวันที่ใหม่กว่าไฟล์ลูกค้า
    if last_c["date_val"] is not None:
        newer_by_date = [r for r in raw_entries if r["date_val"] is not None and r["date_val"] > last_c["date_val"]]
        if newer_by_date:
            return newer_by_date

    # 3. หรือดึงตามจำนวนคอลัมน์ที่เกินมาจากไฟล์ลูกค้า
    if len(raw_entries) > len(cust_entries):
        return raw_entries[len(cust_entries):]

    return []

def get_merged_width(ws, row, col):
    for m in ws.merged_cells.ranges:
        if m.min_row <= row <= m.max_row and m.min_col <= col <= m.max_col:
            return m.max_col - m.min_col + 1
    return 1

def copy_cell_exact(src_cell, dst_cell, protect_formulas=True):
    if protect_formulas and isinstance(dst_cell.value, str) and dst_cell.value.startswith("="):
        return
    dst_cell.value = src_cell.value
    if src_cell.has_style:
        dst_cell.font, dst_cell.fill = copy(src_cell.font), copy(src_cell.fill)
        dst_cell.border, dst_cell.alignment = copy(src_cell.border), copy(src_cell.alignment)
        dst_cell.number_format = src_cell.number_format
    dst_cell.comment = copy(src_cell.comment) if src_cell.comment else None
    if src_cell.hyperlink:
        dst_cell.hyperlink = copy(src_cell.hyperlink)

def append_entry_to_customer(ws_raw, ws_cust, src_col, dst_col, config, protect_formulas=True):
    step, copy_rows = config["step"], config["copy_rows"]
    for r in copy_rows:
        merge_w = get_merged_width(ws_raw, r, src_col)
        for m_range in list(ws_cust.merged_cells.ranges):
            if m_range.min_row <= r <= m_range.max_row and not (m_range.max_col < dst_col or m_range.min_col >= dst_col + step):
                ws_cust.unmerge_cells(str(m_range))
        for offset in range(step):
            copy_cell_exact(
                ws_raw.cell(row=r, column=src_col + offset),
                ws_cust.cell(row=r, column=dst_col + offset),
                protect_formulas
            )
        if merge_w > 1 and merge_w <= step:
            ws_cust.merge_cells(start_row=r, start_column=dst_col, end_row=r, end_column=dst_col + merge_w - 1)

def load_uploaded_files(upload_input):
    files_dict = {}
    if not upload_input:
        return files_dict
    items = upload_input if isinstance(upload_input, list) else [upload_input]
    for item in items:
        item.seek(0)
        raw_bytes = item.read()
        if item.name.lower().endswith(".zip"):
            with zipfile.ZipFile(io.BytesIO(raw_bytes)) as z:
                for n in z.namelist():
                    base = os.path.basename(n)
                    if base and base.lower().endswith(('.xlsx', '.xlsm')) and not base.startswith('~') and not base.startswith('.'):
                        files_dict[base] = z.read(n)
        elif item.name.lower().endswith(('.xlsx', '.xlsm')):
            files_dict[item.name] = raw_bytes
    return files_dict

# ==========================================
# ส่วนแสดงผลหน้าเว็บ (Main UI)
# ==========================================
st.markdown("### 1️⃣ ตั้งค่าจุดตัดข้อมูล (WW / วันที่) และอัปโหลดไฟล์")

col_cfg1, col_cfg2, col_up1, col_up2 = st.columns([1, 1, 1.3, 1.3])

with col_cfg1:
    user_ww_input = st.text_input("🎯 WW ที่ต้องการส่งถึง", placeholder="เช่น 17 หรือ WW17 (ว่าง=ไม่จำกัด)")
    protect_formulas = st.checkbox("🛡️ ป้องกันเขียนทับสูตร", value=True)

with col_cfg2:
    user_date_input = st.date_input("📅 วันที่ต้องการส่งถึง (กรณีชีทไม่มี WW)", value=None)

with col_up1:
    cust_zip_upload = st.file_uploader(
        "📂 อัปโหลดไฟล์ลูกค้า (.zip หรือ .xlsx/.xlsm)",
        type=["zip", "xlsx", "xlsm"],
        accept_multiple_files=True
    )

with col_up2:
    raw_zip_upload = st.file_uploader(
        "📂 อัปโหลดไฟล์ข้อมูลดิบ (.zip หรือ .xlsx/.xlsm)",
        type=["zip", "xlsx", "xlsm"],
        accept_multiple_files=True
    )

target_ww_num = None
if user_ww_input.strip():
    match = re.search(r'\d+', user_ww_input)
    if match:
        target_ww_num = int(match.group())
target_date_val = user_date_input

if cust_zip_upload and raw_zip_upload:
    with st.spinner("กำลังอ่านไฟล์และจับคู่โมเดล..."):
        cust_files_map = load_uploaded_files(cust_zip_upload)
        raw_files_map = load_uploaded_files(raw_zip_upload)

        if not cust_files_map or not raw_files_map:
            st.error("❌ ไม่พบไฟล์ Excel (.xlsx / .xlsm) ข้างในไฟล์ที่อัปโหลด กรุณาตรวจสอบไฟล์อีกครั้ง")
            st.stop()

        paired_models = []
        raw_norm_map = {normalize_filename(rn): rn for rn in raw_files_map.keys()}

        for c_name in sorted(cust_files_map.keys()):
            if c_name in raw_files_map:
                paired_models.append((c_name, c_name))
            else:
                c_norm = normalize_filename(c_name)
                if c_norm in raw_norm_map:
                    paired_models.append((c_name, raw_norm_map[c_norm]))

        if not paired_models and len(cust_files_map) == 1 and len(raw_files_map) == 1:
            paired_models.append((list(cust_files_map.keys())[0], list(raw_files_map.keys())[0]))

        if not paired_models:
            st.error("❌ ไม่พบไฟล์ที่ชื่อตรงกันระหว่างไฟล์ลูกค้าและไฟล์ดิบ")
            st.stop()

    st.markdown("---")
    st.subheader(f"📋 จับคู่สำเร็จ {len(paired_models)} โมเดล (คลิกที่ชื่อโมเดลเพื่อตรวจสอบและติ๊กเลือกข้อมูล)")

    all_models_data = {}

    for file_idx, (c_name, r_name) in enumerate(paired_models):
        profile_name, model_cfg = detect_model_profile(c_name)
        pair_title = f"📦 โมเดล: {c_name}  —  [ใช้โปรไฟล์แถว: {profile_name}]"

        with st.expander(pair_title, expanded=(file_idx == 0)):
            wb_cust_scan = openpyxl.load_workbook(io.BytesIO(cust_files_map[c_name]), data_only=True)
            wb_raw_scan = openpyxl.load_workbook(io.BytesIO(raw_files_map[r_name]), data_only=True)

            model_selections = {}
            sheet_names_list = list(model_cfg.keys())
            tabs = st.tabs(sheet_names_list)

            for idx, sheet_key in enumerate(sheet_names_list):
                config = model_cfg[sheet_key]
                with tabs[idx]:
                    ws_c = find_sheet_in_wb(wb_cust_scan, sheet_key)
                    ws_r = find_sheet_in_wb(wb_raw_scan, sheet_key)

                    if ws_c is None or ws_r is None:
                        st.warning(f"⚠️ ไม่พบชีท **{sheet_key}** ในโมเดลนี้")
                        continue

                    cust_entries = scan_all_valid_entries(ws_c, ws_c, config, check_spec=False)
                    raw_entries = scan_all_valid_entries(ws_r, ws_c, config, check_spec=True)

                    last_cust = cust_entries[-1] if cust_entries else None
                    next_dest_col = last_cust["col"] + config["step"] if last_cust else config["start_col"]

                    if last_cust:
                        st.info(f"📌 **ข้อมูลล่าสุดในไฟล์ลูกค้า:** `{last_cust['label']}` (คอลัมน์ถัดไปที่จะกรอก: **{get_column_letter(next_dest_col)}**)")
                    else:
                        st.info(f"📌 **ยังไม่มีข้อมูลในไฟล์ลูกค้า** (จะเริ่มกรอกที่คอลัมน์ **{get_column_letter(next_dest_col)}**)")

                    available_to_select = filter_new_raw_entries(cust_entries, raw_entries)

                    if not available_to_select:
                        st.success("✅ ข้อมูลในชีทนี้ส่งครบถ้วนแล้ว")
                        model_selections[sheet_key] = []
                    else:
                        sheet_selections = []
                        for entry in available_to_select:
                            # เงื่อนไขตัดรอบ: ถ้ามี WW ให้เช็ค WW, ถ้าไม่มี WW ให้เช็ควันที่
                            within_cutoff = True
                            if entry["ww_num"] is not None and target_ww_num is not None:
                                within_cutoff = (entry["ww_num"] <= target_ww_num)
                            elif entry["date_val"] is not None and target_date_val is not None:
                                within_cutoff = (entry["date_val"] <= target_date_val)

                            default_checked = (not entry["is_ng"]) and within_cutoff

                            checked = st.checkbox(
                                entry["label"],
                                value=default_checked,
                                key=f"chk_{c_name}_{sheet_key}_{entry['col']}_{target_ww_num}_{target_date_val}"
                            )
                            if checked:
                                sheet_selections.append(entry)

                        model_selections[sheet_key] = {
                            "next_col": next_dest_col,
                            "entries": sheet_selections,
                            "config": config
                        }

            all_models_data[c_name] = {
                "raw_name": r_name,
                "selections": model_selections
            }

    st.markdown("---")
    if st.button("🚀 ประมวลผลและดาวน์โหลดไฟล์", type="primary", use_container_width=True):
        with st.spinner("กำลังคัดลอกข้อมูลและบันทึกไฟล์..."):
            total_added_all = 0
            processed_files = {}

            for c_name, data in all_models_data.items():
                r_name = data["raw_name"]
                is_xlsm = c_name.lower().endswith(".xlsm")
                wb_cust_out = openpyxl.load_workbook(io.BytesIO(cust_files_map[c_name]), keep_vba=is_xlsm)
                wb_raw_src = openpyxl.load_workbook(io.BytesIO(raw_files_map[r_name]), data_only=True)

                for sheet_key, sheet_data in data["selections"].items():
                    if not sheet_data or not sheet_data["entries"]:
                        continue

                    ws_c_out = find_sheet_in_wb(wb_cust_out, sheet_key)
                    ws_r_src = find_sheet_in_wb(wb_raw_src, sheet_key)
                    if ws_c_out is None or ws_r_src is None:
                        continue

                    cfg = sheet_data["config"]
                    curr_dst_col = sheet_data["next_col"]

                    for entry in sheet_data["entries"]:
                        append_entry_to_customer(ws_r_src, ws_c_out, entry["col"], curr_dst_col, cfg, protect_formulas)
                        curr_dst_col += cfg["step"]
                        total_added_all += 1

                file_out_buffer = io.BytesIO()
                wb_cust_out.save(file_out_buffer)
                processed_files[c_name] = file_out_buffer.getvalue()

            st.success(f"🎉 เติมข้อมูลสำเร็จทั้งหมด **{total_added_all}** ชุดข้อมูล จาก {len(processed_files)} โมเดล!")

            if len(processed_files) == 1:
                single_name = list(processed_files.keys())[0]
                is_xlsm = single_name.lower().endswith(".xlsm")
                st.download_button(
                    label=f"📥 ดาวน์โหลดไฟล์ {single_name}",
                    data=processed_files[single_name],
                    file_name=single_name,
                    mime="application/vnd.ms-excel.sheet.macroEnabled.12" if is_xlsm else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            else:
                out_zip_buffer = io.BytesIO()
                with zipfile.ZipFile(out_zip_buffer, "w", zipfile.ZIP_DEFLATED) as out_zip:
                    for f_name, f_bytes in processed_files.items():
                        out_zip.writestr(f_name, f_bytes)
                st.download_button(
                    label="📥 ดาวน์โหลดไฟล์ ZIP ลูกค้าทั้งหมด (ชื่อไฟล์ข้างในเหมือนเดิมทุกโมเดล)",
                    data=out_zip_buffer.getvalue(),
                    file_name="Customer_Files.zip",
                    mime="application/zip",
                    use_container_width=True
                )
