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
# ฟังก์ชันแปลงช่วงแถว และสกัด WW / วันที่
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
    """สกัดตัวเลข WW เช่น WW15, WW.15, WW 15"""
    if not text:
        return None
    match = re.search(r'WW\.?\s*(\d+)', str(text), re.IGNORECASE)
    return int(match.group(1)) if match else None

def extract_date_val(val):
    """สกัดค่าวันที่จาก Cell (รองรับทั้ง datetime object และข้อความวันที่ เช่น 16-Oct-26, 2026-10-16)"""
    if val is None:
        return None
    if isinstance(val, datetime.datetime):
        return val.date()
    if isinstance(val, datetime.date):
        return val
    s = str(val).strip()
    if not s or len(s) < 5:
        return None
    # ลองแปลงข้อความรูปแบบต่างๆ เป็น date
    date_formats = [
        "%d-%b-%y", "%d-%b-%Y", "%d/%b/%y", "%d/%b/%Y",
        "%d-%m-%y", "%d-%m-%Y", "%d/%m/%y", "%d/%m/%Y",
        "%Y-%m-%d", "%Y/%m/%d", "%b-%y", "%b %y"
    ]
    for fmt in date_formats:
        try:
            return datetime.datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    # ค้นหารูปแบบ 16-Oct-26 ที่อาจปนอยู่ในข้อความยาวๆ
    m = re.search(r'(\d{1,2}[-/][A-Za-z]{3}[-/]\d{2,4})', s)
    if m:
        for fmt in ["%d-%b-%y", "%d-%b-%Y", "%d/%b/%y", "%d/%b/%Y"]:
            try:
                return datetime.datetime.strptime(m.group(1), fmt).date()
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
            "step": 1, "header_rows": [10, 11, 12], "start_col": 2,
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
    """แยกโปรไฟล์ตามชื่อไฟล์"""
    fn = filename.upper()
    if "GZ8427V4" in fn: # ครอบคลุมทั้ง GZ8427V4 และ GZ8427V40
        return "MARLIN_V4", MODEL_PROFILES["MARLIN_V4"]
    elif "MARLIN" in fn:
        return "MARLIN_OTHER", MODEL_PROFILES["MARLIN_OTHER"]
    elif "SUMMIT" in fn:
        return "SUMMIT", MODEL_PROFILES["SUMMIT"]
    else:
        return "DORADO", MODEL_PROFILES["DORADO"]

def find_sheet_in_wb(wb, target_name):
    """ค้นหาชีทแบบไม่สนตัวพิมพ์เล็ก-ใหญ่ (เช่น Ghost Test vs Ghost test)"""
    for s in wb.sheetnames:
        if s.strip().lower() == target_name.strip().lower():
            return wb[s]
    return None

# ==========================================
# Helper Functions สำหรับอ่านและคัดลอกข้อมูล
# ==========================================
def format_val(val):
    if val is None:
        return ""
    if isinstance(val, (datetime.datetime, datetime.date)):
        return val.strftime("%d-%b-%y")
    return str(val).strip()

def to_float(val):
    if val is None:
        return None
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip().replace(",", "")
    if s == "" or s.upper() in ["NA", "N/A", "ND", "-", "PASS", "FAIL"]:
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
    if s == "" or s.startswith("="):
        return False
    ignore_list = [
        "WW.", "WW", "DATE", "ION", "SPEC.", "SPEC", "NG/CM2", "NG/CM²",
        "SILICONE", "HYDROCARBON", "MOLD", "REMARK", "PASS", "FAIL", "ND"
    ]
    return s.upper() not in ignore_list

def get_merged_width(ws, row, col):
    for m in ws.merged_cells.ranges:
        if m.min_row <= row <= m.max_row and m.min_col <= col <= m.max_col:
            return m.max_col - m.min_col + 1
    return 1

def scan_entries(ws_target, ws_cust_ref, config, check_spec=False):
    entries = []
    header_rows = config["header_rows"]
    step = config["step"]
    col = config["start_col"]
    max_col = ws_target.max_column

    while col <= max_col:
        tokens = []
        ww_num = None
        entry_date = None

        # กวาดอ่านข้อมูลในกลุ่มแถวหัวตารางของคอลัมน์นั้น
        for r in header_rows:
            raw_v = ws_target.cell(row=r, column=col).value
            if is_valid_header_token(raw_v):
                s_val = format_val(raw_v)
                if s_val and s_val not in tokens:
                    tokens.append(s_val)
                if ww_num is None:
                    ww_num = extract_ww_num(s_val)
                if entry_date is None:
                    entry_date = extract_date_val(raw_v)

        if tokens:
            # ใช้ค่าตัวสุดท้ายในหัวตาราง (มักเป็น Lot No. หรือ Lab No.) เป็น ID หลักในการเทียบจุดตัด
            # และเก็บ signature รวมทุกแถวเผื่อกรณีชื่อซ้ำ
            primary_id = tokens[-1]
            full_sig = " | ".join(tokens)

            failures = evaluate_entry_spec(ws_target, ws_cust_ref, col, config) if check_spec else []
            base_label = f"[Col {get_column_letter(col)}] {full_sig}"

            entries.append({
                "col": col,
                "id": primary_id,
                "full_sig": full_sig,
                "label": f"⚠️ {base_label} — หลุด Spec: {', '.join(failures)}" if failures else f"✅ {base_label}",
                "is_ng": len(failures) > 0,
                "failures": failures,
                "ww_num": ww_num,
                "date_val": entry_date
            })
            col += step
        else:
            col += 1
    return entries

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

# ==========================================
# ส่วนแสดงผลหน้าเว็บ (Main UI)
# ==========================================
st.markdown("### 1️⃣ ตั้งค่าจุดตัดข้อมูล (WW / วันที่) และอัปโหลดไฟล์")

col_cfg1, col_cfg2, col_up1, col_up2 = st.columns([1, 1, 1.3, 1.3])

with col_cfg1:
    user_ww_input = st.text_input("🎯 WW ที่ต้องการส่งถึง", placeholder="เช่น 15 หรือ WW15 (ว่าง=ไม่จำกัด)")
    protect_formulas = st.checkbox("🛡️ ป้องกันเขียนทับสูตร", value=True)

with col_cfg2:
    user_date_input = st.date_input("📅 วันที่ต้องการส่งถึง (กรณีชีทไม่มี WW)", value=None)

with col_up1:
    cust_zip_upload = st.file_uploader("📂 อัปโหลดไฟล์ลูกค้า (.zip หรือ .xlsx/.xlsm)", type=["zip", "xlsx", "xlsm"])

with col_up2:
    raw_zip_upload = st.file_uploader("📂 อัปโหลดไฟล์ข้อมูลดิบ (.zip หรือ .xlsx/.xlsm)", type=["zip", "xlsx", "xlsm"])

# แปลงค่า WW เป้าหมาย
target_ww_num = None
if user_ww_input.strip():
    match = re.search(r'\d+', user_ww_input)
    if match:
        target_ww_num = int(match.group())

target_date_val = user_date_input # เป็น datetime.date หรือ None

def load_uploaded_files(upload_obj):
    """รองรับทั้งการอัปโหลดเป็นไฟล์ .zip และไฟล์ Excel เดี่ยวๆ"""
