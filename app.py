import streamlit as st
import openpyxl
from openpyxl.utils import get_column_letter, column_index_from_string
from lxml import etree
import pandas as pd
import io
import datetime
import re
import zipfile
import os
import posixpath

st.set_page_config(page_title="Customer Data Batch Auto-Filler", layout="wide")
st.title("📦 ระบบเติมข้อมูลส่งลูกค้าอัตโนมัติ (6 โปรไฟล์โมเดล + ลบ NG เก่า + กรอง WW/วันที่)")

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
XML_NS = "http://www.w3.org/XML/1998/namespace"

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
    if not text:
        return None
    s = str(text).strip()
    m = re.search(r'(LAB\d{6,}|HNRJ\w+)', s, re.IGNORECASE)
    if m:
        return m.group(1).upper()
    return None

def extract_date_val(val):
    """สกัดวันที่จากเซลล์ (ไม่แปลงตัวเลขดิบเป็นวันที่ เพื่อป้องกันค่าฝุ่น 40,000-55,000 ถูกอ่านเป็นปี 2016)"""
    if val is None:
        return None
    if isinstance(val, datetime.datetime):
        return val.date()
    if isinstance(val, datetime.date):
        return val
    if isinstance(val, (int, float)):
        return None

    s = str(val).strip()
    if not s or len(s) < 4:
        return None

    m1 = re.search(r'(\d{1,2})[\-\/\s]+([A-Za-z]{3})[\-\/\s\']*(\d{2,4})', s)
    if m1:
        day, mon, yr = m1.group(1), m1.group(2), m1.group(3)
        clean_d = f"{day}-{mon}-{yr}"
        for fmt in ["%d-%b-%y", "%d-%b-%Y"]:
            try:
                return datetime.datetime.strptime(clean_d, fmt).date()
            except ValueError:
                pass

    m2 = re.search(r'(\d{1,2})/(\d{1,2})/(\d{2,4})', s)
    if m2:
        for fmt in ["%m/%d/%Y", "%d/%m/%Y", "%m/%d/%y", "%d/%m/%y"]:
            try:
                return datetime.datetime.strptime(m2.group(0), fmt).date()
            except ValueError:
                pass

    date_formats = [
        "%d-%b-%y", "%d-%b-%Y", "%d/%b/%y", "%d/%b/%Y", "%d %b %y", "%d %b %Y",
        "%d-%m-%y", "%d-%m-%Y", "%d/%m/%y", "%d/%m/%Y",
        "%Y-%m-%d", "%Y/%m/%d", "%b-%y", "%b %y", "%b-%Y", "%B-%y", "%b'%y"
    ]
    for fmt in date_formats:
        try:
            return datetime.datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    return None

# ==========================================
# ตั้งค่าพิกัดแถวและ Spec แยกตาม 6 โปรไฟล์โมเดล
# ==========================================
MODEL_PROFILES = {
    "DORADO_OTHER": {
        "LPC": {
            "step": 1, "header_rows": [29, 30, 31], "start_col": 2,
            "copy_rows": expand_ranges([(29, 34), 44]),
            "spec_checks": [(32, 41), (33, 41), (34, 41)]
        },
        "IC": {
            "step": 2, "header_rows": [19, 20], "start_col": 2,
            "copy_rows": expand_ranges([(19, 28)]),
            "spec_checks": [(28, 32)]
        },
        "FTIR": {
            "step": 2, "header_rows": [5, 6], "start_col": 2,
            "copy_rows": expand_ranges([5, 6, 7, 9]),
            "spec_checks": [(7, 8), (9, 10)]
        },
        "DHS": {
            "step": 1, "header_rows": [51, 52, 53], "start_col": 4,
            "copy_rows": expand_ranges([(51, 54), 57, 60, 63, 66, 69, 72, 75, 78, 81, 84, 87, 90, 93]),
            "spec_checks": [(r, r + 1) for r in [54, 57, 60, 63, 66, 69, 72, 75, 78, 81, 84, 87, 90, 93]]
        },
        "APA": {
            "step": 2, "header_rows": [38, 39, 40], "start_col": 4,
            "copy_rows": expand_ranges([(38, 58), (62, 64), (66, 71), 75, (79, 94), (98, 116), 118, (120, 129), 131]),
            "spec_checks": [(58, 59), (62, 65), (63, 65), (64, 65), (71, 72), (94, 95), (118, 119)]
        },
        "GCMS": {
            "step": 2, "header_rows": [40, 41, 42], "start_col": 4,
            "copy_rows": expand_ranges([(40, 44), 49, 53]),
            "spec_checks": [(43, 46), (53, 54)]
        },
        "GCMS RHC": {
            "step": 2, "header_rows": [36, 37, 38, 39, 40], "start_col": 4,
            "copy_rows": expand_ranges([(36, 41), 45, 49, (53, 55)]),
            "spec_checks": [(41, 42), (45, 46), (49, 50)]
        },
        "Ghost Test": {
            "step": 1, "header_rows": [4, 5, 6], "start_col": 2,
            "copy_rows": expand_ranges([(4, 9)]),
            "spec_checks": []
        }
    },

    "DORADO_V4_V6": {
        "LPC": {
            "step": 1, "header_rows": [29, 30, 31], "start_col": 2,
            "copy_rows": expand_ranges([(29, 34), 44]),
            "spec_checks": [(32, 41), (33, 41), (34, 41)]
        },
        "IC": {
            "step": 2, "header_rows": [19, 20], "start_col": 2,
            "copy_rows": expand_ranges([(19, 28)]),
            "spec_checks": [(28, 32)]
        },
        "FTIR": {
            "step": 2, "header_rows": [5, 6], "start_col": 2,
            "copy_rows": expand_ranges([5, 6, 7, 9]),
            "spec_checks": [(7, 8), (9, 10)]
        },
        "DHS": {
            "step": 1, "header_rows": [51, 52, 53], "start_col": 4,
            "copy_rows": expand_ranges([(51, 54), 57, 60, 63, 66, 69, 72, 75, 78, 81, 84, 87, 90, 93]),
            "spec_checks": [(r, r + 1) for r in [54, 57, 60, 63, 66, 69, 72, 75, 78, 81, 84, 87, 90, 93]]
        },
        "APA": {
            "step": 2, "header_rows": [38, 39, 40], "start_col": 4,
            "copy_rows": expand_ranges([(38, 58), (64, 66), (68, 73), 79, 81, (83, 95), (101, 119), 121, (127, 135), 137]),
            "spec_checks": [(58, 59), (64, 67), (65, 67), (66, 67), (73, 74), (95, 96), (121, 122)]
        },
        "GCMS": {
            "step": 2, "header_rows": [40, 41, 42], "start_col": 4,
            "copy_rows": expand_ranges([(40, 43), 47, 51]),
            "spec_checks": [(43, 44), (47, 48), (51, 52)]
        },
        "GCMS RHC": {
            "step": 2, "header_rows": [36, 37, 38], "start_col": 4,
            "copy_rows": expand_ranges([(36, 39), 45, 51, 57, (62, 63)]),
            "spec_checks": [(39, 40), (45, 46), (51, 52), (57, 60)]
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
            "step": 2, "header_rows": [20, 21, 22], "start_col": 2,
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

    "MARLIN_V0": {
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
            "step": 2, "header_rows": [3, 4, 5], "start_col": 2,
            "copy_rows": expand_ranges([(3, 6), 8]),
            "spec_checks": [(6, 7), (8, 9)]
        },
        "DHS": {
            "step": 1, "header_rows": [52, 53, 54, 55], "start_col": 4,
            "copy_rows": expand_ranges([(52, 56), 59, 62, 65, 68, 71, 74, 77, 80, 83, 86, 89, 92, 95]),
            "spec_checks": [(r, r + 1) for r in [56, 59, 62, 65, 68, 71, 74, 77, 80, 83, 86, 89, 92, 95]]
        },
        "APA": {
            "step": 2, "header_rows": [39, 40, 41, 42], "start_col": 4,
            "copy_rows": expand_ranges([(39, 42), (44, 60), (66, 68), (70, 75), 81, 83, (85, 95), (101, 119), 121, (127, 135), 137]),
            "spec_checks": [(60, 61), (75, 76), (95, 96), (121, 122)]
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
    elif "GZ8427V0" in fn:
        return "MARLIN_V0", MODEL_PROFILES["MARLIN_V0"]
    elif "MARLIN" in fn:
        return "MARLIN_OTHER", MODEL_PROFILES["MARLIN_OTHER"]
    elif "SUMMIT" in fn:
        return "SUMMIT", MODEL_PROFILES["SUMMIT"]
    elif any(tag in fn for tag in ["GZ8431V4", "GZ8431V5", "GZ8431V6"]):
        return "DORADO_V4_V6", MODEL_PROFILES["DORADO_V4_V6"]
    else:
        return "DORADO_OTHER", MODEL_PROFILES["DORADO_OTHER"]

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
# Helper Functions สำหรับอ่านและสแกนข้อมูล
# ==========================================
def get_merge_bounds(ws, row, col):
    for m in ws.merged_cells.ranges:
        if m.min_row <= row <= m.max_row and m.min_col <= col <= m.max_col:
            return m.min_col, m.max_col
    return col, col

def get_cell_effective_val(ws, row, col):
    min_c, _ = get_merge_bounds(ws, row, col)
    return ws.cell(row=row, column=min_c).value

def format_val(val):
    if val is None:
        return ""
    if isinstance(val, (datetime.datetime, datetime.date)):
        return val.strftime("%d-%b-%y")
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
    except ValueError:
        return None

def parse_spec_float(val):
    f = to_float(val)
    if f is not None:
        return f
    if val is None:
        return None
    s = str(val).strip().replace(",", "")
    if s.upper() in ["NA", "N/A", "ND", "-", ""]:
        return None
    m = re.search(r'(\d+(?:\.\d+)?)', s)
    if m:
        try:
            return float(m.group(1))
        except ValueError:
            return None
    return None

def find_spec_value(ws_target, ws_ref, spec_row, col):
    for ws in (ws_target, ws_ref):
        if ws is None:
            continue
        v = parse_spec_float(ws.cell(row=spec_row, column=col).value)
        if v is not None:
            return v
    for ws in (ws_target, ws_ref):
        if ws is None:
            continue
        for c in range(max(2, col - 1), max(1, col - 15), -1):
            v = parse_spec_float(ws.cell(row=spec_row, column=c).value)
            if v is not None:
                return v
        for c in range(2, min(ws.max_column + 1, 60)):
            v = parse_spec_float(ws.cell(row=spec_row, column=c).value)
            if v is not None:
                return v
    return None

def evaluate_entry_spec(ws_target, ws_ref, start_col, span, sheet_key, config):
    failures = []
    is_ic = (sheet_key.strip().lower() == "ic")
    data_rows = get_data_rows(config)

    for val_row, spec_row in config.get("spec_checks", []):
        for offset in range(span):
            c = start_col + offset
            val_num = to_float(ws_target.cell(row=val_row, column=c).value)

            if val_num is None and is_ic:
                ion_vals = [to_float(ws_target.cell(row=r, column=c).value) for r in data_rows if r < val_row]
                ion_vals = [v for v in ion_vals if v is not None]
                if ion_vals:
                    val_num = sum(ion_vals)

            if val_num is None:
                continue
            spec_num = find_spec_value(ws_target, ws_ref, spec_row, c)
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
        "SILICONE", "HYDROCARBON", "HYDROCARBONS", "MOLD", "REMARK", "PASS", "FAIL", "ND",
        "PART NAME", "MATERIAL", "PART NO.", "PART NO", "OUTGASSED COMPOUNDS"
    ]
    return s.upper() not in ignore_list

def get_data_rows(config):
    h_set = set(config["header_rows"])
    d_rows = [r for r in config["copy_rows"] if r not in h_set]
    return d_rows if d_rows else config["copy_rows"]

def single_col_has_measurement(ws, col, config):
    for r in get_data_rows(config):
        val = ws.cell(row=r, column=col).value
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

def determine_entry_span(ws, col, config):
    default_step = config["step"]
    if default_step == 1:
        return 1
    for r in config["header_rows"]:
        min_c, max_c = get_merge_bounds(ws, r, col)
        if min_c == col and max_c > col:
            return min(max_c - min_c + 1, default_step)
        elif min_c < col <= max_c:
            return 1
    for r in config["header_rows"]:
        min_c_next, max_c_next = get_merge_bounds(ws, r, col + 1)
        if min_c_next == col + 1 and max_c_next > col + 1:
            return 1
    return default_step

def scan_all_valid_entries(ws_target, ws_ref, sheet_key, config, check_spec=False):
    entries = []
    header_rows = config["header_rows"]
    data_rows = get_data_rows(config)
    col = config["start_col"]
    max_col = ws_target.max_column

    while col <= max_col:
        if not single_col_has_measurement(ws_target, col, config):
            col += 1
            continue

        span = determine_entry_span(ws_target, col, config)
        if span > 1 and not single_col_has_measurement(ws_target, col + 1, config):
            span = 1

        col_ww = None
        col_date = None
        unique_id = None
        tokens = []

        for r in header_rows:
            if r < 1:
                continue
            for offset in range(span):
                raw_v = get_cell_effective_val(ws_target, r, col + offset)
                if raw_v is not None:
                    s_raw = str(raw_v)
                    if col_ww is None:
                        col_ww = extract_ww_num(s_raw)
                    if col_date is None:
                        d_extracted = extract_date_val(raw_v)
                        if d_extracted is not None:
                            col_date = d_extracted
                    if unique_id is None:
                        unique_id = extract_unique_lot_or_lab_id(s_raw)
                    if is_valid_header_token(raw_v):
                        s_val = format_val(raw_v)
                        if s_val and s_val not in tokens:
                            tokens.append(s_val)

        data_vals = []
        for r in data_rows:
            for offset in range(span):
                v = ws_target.cell(row=r, column=col + offset).value
                fv = to_float(v)
                if fv is not None:
                    data_vals.append(f"{r}_{offset}:{fv:.4f}")
                elif v is not None and not str(v).startswith("#") and str(v).strip() != "":
                    data_vals.append(f"{r}_{offset}:{str(v).strip()}")
        data_sig = "|".join(data_vals)

        full_sig = " | ".join(tokens) if tokens else f"Col {get_column_letter(col)}"
        failures = evaluate_entry_spec(ws_target, ws_ref, col, span, sheet_key, config) if check_spec else []

        date_str = col_date.strftime("%d-%b-%y") if col_date else "ไม่ระบุวันที่"
        ww_str = f"WW{col_ww:02d} | " if col_ww is not None else ""
        col_range_str = get_column_letter(col) if span == 1 else f"{get_column_letter(col)}-{get_column_letter(col + span - 1)}"
        base_label = f"[Col {col_range_str} | {ww_str}{date_str}] {full_sig}"

        entries.append({
            "col": col,
            "span": span,
            "end_col": col + span - 1,
            "col_range_str": col_range_str,
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
        col += span

    return entries

def filter_new_raw_entries(cust_entries, raw_entries):
    if not cust_entries:
        return raw_entries
    if not raw_entries:
        return []

    last_c = cust_entries[-1]
    match_indices = []
    for i, r_ent in enumerate(raw_entries):
        if last_c["unique_id"] and r_ent["unique_id"] and last_c["unique_id"] == r_ent["unique_id"]:
            match_indices.append(i)
        elif last_c["data_sig"] and r_ent["data_sig"] == last_c["data_sig"]:
            match_indices.append(i)

    if match_indices:
        candidates = raw_entries[match_indices[-1] + 1:]
    elif last_c["date_val"] is not None:
        candidates = [r for r in raw_entries if r["date_val"] is not None and r["date_val"] > last_c["date_val"]]
    elif last_c["ww_num"] is not None:
        candidates = [r for r in raw_entries if r["ww_num"] is not None and r["ww_num"] > last_c["ww_num"]]
    elif len(raw_entries) > len(cust_entries):
        candidates = raw_entries[len(cust_entries):]
    else:
        candidates = []

    if last_c["date_val"] is not None:
        candidates = [r for r in candidates if r["date_val"] is None or r["date_val"] >= last_c["date_val"]]

    return candidates

def matches_user_filter(entry, target_ww_num, target_date_val, filter_mode):
    if target_ww_num is None and target_date_val is None:
        return True

    if entry["ww_num"] is not None and target_ww_num is not None:
        if filter_mode == "gte":
            return entry["ww_num"] >= target_ww_num
        else:
            return entry["ww_num"] <= target_ww_num
    elif entry["date_val"] is not None and target_date_val is not None:
        if filter_mode == "gte":
            return entry["date_val"] >= target_date_val
        else:
            return entry["date_val"] <= target_date_val

    return True

# ==========================================
# ระบบ XML Patching (ลบคอลัมน์ NG แบบขยับชิด + เติมข้อมูลใหม่ รักษา Shape/กราฟ 100%)
# ==========================================
def clean_xml_str(s):
    return re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f]', '', str(s))

def parse_cell_ref(ref):
    m = re.match(r'^([A-Z]+)(\d+)$', ref)
    if not m:
        return 0, 0
    return int(m.group(2)), column_index_from_string(m.group(1))

def parse_range_ref(ref):
    parts = ref.split(":")
    if len(parts) == 1:
        r1, c1 = parse_cell_ref(parts[0])
        return r1, c1, r1, c1
    r1, c1 = parse_cell_ref(parts[0])
    r2, c2 = parse_cell_ref(parts[1])
    return r1, c1, r2, c2

def get_sheet_zip_paths(z):
    wb_xml = etree.fromstring(z.read("xl/workbook.xml"))
    rels_xml = etree.fromstring(z.read("xl/_rels/workbook.xml.rels"))

    rid_to_target = {}
    for rel in rels_xml.findall(f"{{{PKG_REL_NS}}}Relationship"):
        rid_to_target[rel.get("Id")] = rel.get("Target")

    sheet_map = {}
    sheets_el = wb_xml.find(f"{{{MAIN_NS}}}sheets")
    if sheets_el is not None:
        for sh in sheets_el.findall(f"{{{MAIN_NS}}}sheet"):
            name = sh.get("name")
            rid = sh.get(f"{{{REL_NS}}}id")
            target = rid_to_target.get(rid, "")
            if target:
                if target.startswith("/"):
                    full_path = target.lstrip("/")
                else:
                    full_path = posixpath.normpath(posixpath.join("xl", target))
                sheet_map[name.strip().lower()] = full_path
    return sheet_map

def get_or_create_row(sheet_data, row_num):
    for idx, row_el in enumerate(sheet_data):
        r_attr = row_el.get("r")
        if r_attr:
            r_int = int(r_attr)
            if r_int == row_num:
                return row_el
            elif r_int > row_num:
                new_row = etree.Element(f"{{{MAIN_NS}}}row", r=str(row_num))
                sheet_data.insert(idx, new_row)
                return new_row
    new_row = etree.SubElement(sheet_data, f"{{{MAIN_NS}}}row", r=str(row_num))
    return new_row

def get_or_create_cell(row_el, row_num, col_num):
    ref = f"{get_column_letter(col_num)}{row_num}"
    for idx, c_el in enumerate(row_el):
        c_ref = c_el.get("r")
        if c_ref:
            _, c_int = parse_cell_ref(c_ref)
            if c_int == col_num:
                return c_el
            elif c_int > col_num:
                new_c = etree.Element(f"{{{MAIN_NS}}}c", r=ref)
                row_el.insert(idx, new_c)
                return new_c
    new_c = etree.SubElement(row_el, f"{{{MAIN_NS}}}c", r=ref)
    return new_c

def find_inherited_style(row_el, target_col):
    best_style = None
    best_col = -1
    for c_el in row_el:
        c_ref = c_el.get("r")
        s_attr = c_el.get("s")
        if c_ref and s_attr is not None:
            _, c_int = parse_cell_ref(c_ref)
            if c_int < target_col and c_int > best_col:
                best_col = c_int
                best_style = s_attr
    return best_style

def write_value_to_cell_xml(c_el, val, row_el, col_num, protect_formulas=True, force_style=False):
    has_formula = c_el.find(f"{{{MAIN_NS}}}f") is not None
    if protect_formulas and has_formula:
        return

    if force_style or c_el.get("s") is None:
        inh_s = find_inherited_style(row_el, col_num)
        if inh_s is not None:
            c_el.set("s", inh_s)

    for child in list(c_el):
        c_el.remove(child)

    if val is None or (isinstance(val, str) and val == ""):
        c_el.attrib.pop("t", None)
        return

    if isinstance(val, bool):
        c_el.set("t", "b")
        v_el = etree.SubElement(c_el, f"{{{MAIN_NS}}}v")
        v_el.text = "1" if val else "0"
    elif isinstance(val, (datetime.datetime, datetime.date)):
        c_el.attrib.pop("t", None)
        if isinstance(val, datetime.datetime):
            serial = (val - datetime.datetime(1899, 12, 30)).total_seconds() / 86400.0
        else:
            serial = (val - datetime.date(1899, 12, 30)).days
        v_el = etree.SubElement(c_el, f"{{{MAIN_NS}}}v")
        v_el.text = str(serial)
    elif isinstance(val, (int, float)):
        c_el.attrib.pop("t", None)
        v_el = etree.SubElement(c_el, f"{{{MAIN_NS}}}v")
        v_el.text = str(val)
    else:
        c_el.set("t", "inlineStr")
        is_el = etree.SubElement(c_el, f"{{{MAIN_NS}}}is")
        t_el = etree.SubElement(is_el, f"{{{MAIN_NS}}}t")
        t_el.set(f"{{{XML_NS}}}space", "preserve")
        t_el.text = clean_xml_str(val)

def get_or_create_merge_cells(root):
    mc = root.find(f"{{{MAIN_NS}}}mergeCells")
    if mc is not None:
        return mc

    after_tags = {
        f"{{{MAIN_NS}}}{tag}" for tag in [
            "phoneticPr", "conditionalFormatting", "dataValidations", "hyperlinks",
            "printOptions", "pageMargins", "pageSetup", "headerFooter", "rowBreaks",
            "colBreaks", "customProperties", "cellWatches", "ignoredErrors",
            "smartTags", "drawing", "legacyDrawing", "legacyDrawingHF", "picture", "extLst"
        ]
    }
    mc = etree.Element(f"{{{MAIN_NS}}}mergeCells", count="0")
    for idx, child in enumerate(root):
        if child.tag in after_tags:
            root.insert(idx, mc)
            return mc
    root.append(mc)
    return mc

def patch_sheet_xml(sheet_xml_bytes, sheet_key, ws_cust_src, ws_raw_src, sheet_data_plan, protect_formulas=True):
    all_cust_entries = sheet_data_plan.get("all_cust_entries", [])
    delete_entries = sheet_data_plan.get("delete_entries", [])
    new_entries = sheet_data_plan.get("entries", [])
    config = sheet_data_plan["config"]

    if not delete_entries and not new_entries:
        return sheet_xml_bytes

    parser = etree.XMLParser(remove_blank_text=False)
    root = etree.fromstring(sheet_xml_bytes, parser=parser)
    sheet_data = root.find(f"{{{MAIN_NS}}}sheetData")
    if sheet_data is None:
        return sheet_xml_bytes

    copy_rows = config["copy_rows"]
    header_rows_set = set(config["header_rows"])
    is_lpc_sheet = (sheet_key.strip().lower() == "lpc")
    mold_row = max(copy_rows) if is_lpc_sheet else -1

    write_queue = []
    if delete_entries:
        first_del_col = min(e["col"] for e in delete_entries)
        old_end_col = max(e["end_col"] for e in all_cust_entries)
        del_col_set = {e["col"] for e in delete_entries}

        for c_ent in all_cust_entries:
            if c_ent["col"] > first_del_col and c_ent["col"] not in del_col_set:
                write_queue.append((ws_cust_src, c_ent))
        for r_ent in new_entries:
            write_queue.append((ws_raw_src, r_ent))

        curr_dst_col = first_del_col
        clear_up_to_col = old_end_col
    else:
        for r_ent in new_entries:
            write_queue.append((ws_raw_src, r_ent))
        curr_dst_col = sheet_data_plan["next_col"]
        clear_up_to_col = -1

    modified_cols = set()
    merges_to_add = []
    merges_to_clear = []

    for src_ws, entry in write_queue:
        src_col = entry["col"]
        span = entry.get("span", config["step"])

        for offset in range(span):
            modified_cols.add(curr_dst_col + offset)

        for r in copy_rows:
            row_el = get_or_create_row(sheet_data, r)
            min_c_src, max_c_src = get_merge_bounds(src_ws, r, src_col)
            src_merge_width = max_c_src - min_c_src + 1

            merges_to_clear.append((r, curr_dst_col, curr_dst_col + span - 1))

            for offset in range(span):
                s_val = src_ws.cell(row=r, column=src_col + offset).value
                if offset == 0 and s_val is None and min_c_src < src_col:
                    s_val = src_ws.cell(row=r, column=min_c_src).value

                c_el = get_or_create_cell(row_el, r, curr_dst_col + offset)
                is_header_or_mold = (r in header_rows_set) or (r == mold_row)
                write_value_to_cell_xml(
                    c_el,
                    s_val,
                    row_el,
                    curr_dst_col + offset,
                    protect_formulas=(False if is_header_or_mold else protect_formulas),
                    force_style=is_header_or_mold
                )

            if span > 1 and (src_merge_width >= span or (r in header_rows_set and src_ws.cell(row=r, column=src_col + 1).value is None)):
                m_ref = f"{get_column_letter(curr_dst_col)}{r}:{get_column_letter(curr_dst_col + span - 1)}{r}"
                merges_to_add.append(m_ref)

        curr_dst_col += span

    while curr_dst_col <= clear_up_to_col:
        modified_cols.add(curr_dst_col)
        for r in copy_rows:
            row_el = get_or_create_row(sheet_data, r)
            merges_to_clear.append((r, curr_dst_col, curr_dst_col))
            c_el = get_or_create_cell(row_el, r, curr_dst_col)
            is_header_or_mold = (r in header_rows_set) or (r == mold_row)
            write_value_to_cell_xml(
                c_el,
                None,
                row_el,
                curr_dst_col,
                protect_formulas=(False if is_header_or_mold else protect_formulas),
                force_style=False
            )
        curr_dst_col += 1

    for row_el in sheet_data:
        for c_el in row_el:
            c_ref = c_el.get("r")
            if not c_ref:
                continue
            _, c_int = parse_cell_ref(c_ref)
            if c_int in modified_cols and c_el.find(f"{{{MAIN_NS}}}f") is not None:
                v_el = c_el.find(f"{{{MAIN_NS}}}v")
                if v_el is not None:
                    c_el.remove(v_el)
                if c_el.get("t") == "e":
                    c_el.attrib.pop("t", None)

    if merges_to_clear or merges_to_add:
        mc = root.find(f"{{{MAIN_NS}}}mergeCells")
        if mc is not None:
            for m_el in list(mc):
                ref = m_el.get("ref")
                if not ref:
                    continue
                r1, c1, r2, c2 = parse_range_ref(ref)
                should_remove = False
                for cr, cc1, cc2 in merges_to_clear:
                    if r1 <= cr <= r2 and not (c2 < cc1 or c1 > cc2):
                        should_remove = True
                        break
                if should_remove:
                    mc.remove(m_el)

        if merges_to_add:
            mc = get_or_create_merge_cells(root)
            existing_refs = {m.get("ref") for m in mc}
            for new_ref in merges_to_add:
                if new_ref not in existing_refs:
                    etree.SubElement(mc, f"{{{MAIN_NS}}}mergeCell", ref=new_ref)
                    existing_refs.add(new_ref)

        if mc is not None:
            if len(mc) == 0:
                root.remove(mc)
            else:
                mc.set("count", str(len(mc)))

    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)

def patch_workbook_for_recalc(wb_xml_bytes):
    parser = etree.XMLParser(remove_blank_text=False)
    root = etree.fromstring(wb_xml_bytes, parser=parser)
    calc_pr = root.find(f"{{{MAIN_NS}}}calcPr")
    if calc_pr is None:
        calc_pr = etree.SubElement(root, f"{{{MAIN_NS}}}calcPr")
    calc_pr.set("fullCalcOnLoad", "1")
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)

def process_single_model_xml_safe(cust_bytes, raw_bytes, selections, protect_formulas=True):
    wb_cust_src = openpyxl.load_workbook(io.BytesIO(cust_bytes), data_only=True)
    wb_raw_src = openpyxl.load_workbook(io.BytesIO(raw_bytes), data_only=True)
    added_count = 0
    deleted_count = 0

    with zipfile.ZipFile(io.BytesIO(cust_bytes), "r") as zin:
        sheet_zip_map = get_sheet_zip_paths(zin)
        modified_parts = {}

        for sheet_key, sheet_data_plan in selections.items():
            if not sheet_data_plan:
                continue
            has_del = len(sheet_data_plan.get("delete_entries", [])) > 0
            has_add = len(sheet_data_plan.get("entries", [])) > 0
            if not has_del and not has_add:
                continue

            ws_c_src = find_sheet_in_wb(wb_cust_src, sheet_key)
            ws_r_src = find_sheet_in_wb(wb_raw_src, sheet_key)
            target_xml_path = sheet_zip_map.get(sheet_key.strip().lower())
            if ws_c_src is None or ws_r_src is None or not target_xml_path or target_xml_path not in zin.namelist():
                continue

            orig_sheet_xml = zin.read(target_xml_path)
            patched_xml = patch_sheet_xml(
                orig_sheet_xml,
                sheet_key,
                ws_c_src,
                ws_r_src,
                sheet_data_plan,
                protect_formulas=protect_formulas
            )
            modified_parts[target_xml_path] = patched_xml
            added_count += len(sheet_data_plan.get("entries", []))
            deleted_count += len(sheet_data_plan.get("delete_entries", []))

        if (added_count > 0 or deleted_count > 0) and "xl/workbook.xml" in zin.namelist():
            modified_parts["xl/workbook.xml"] = patch_workbook_for_recalc(zin.read("xl/workbook.xml"))

        out_buf = io.BytesIO()
        with zipfile.ZipFile(out_buf, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename in modified_parts:
                    zout.writestr(item, modified_parts[item.filename])
                else:
                    zout.writestr(item, zin.read(item.filename))

        return out_buf.getvalue(), added_count, deleted_count

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

col_cfg1, col_cfg2, col_up1, col_up2 = st.columns([1.1, 1.1, 1.2, 1.2])

with col_cfg1:
    user_ww_input = st.text_input("🎯 WW ที่ต้องการ", placeholder="เช่น 17 หรือ WW17 (ว่าง=ทั้งหมด)")
    filter_mode_label = st.selectbox(
        "🔍 โหมดการแสดงรายการให้เลือก",
        options=[
            "แสดงเฉพาะ <= WW/วันที่ที่เลือก (ซ่อนตัวที่เกิน)",
            "แสดงเฉพาะ >= WW/วันที่ที่เลือก (ซ่อนตัวที่ก่อนหน้า)",
            "แสดงทั้งหมด (ติ๊กเฉพาะ <= WW/วันที่ที่เลือก)"
        ],
        index=0
    )

with col_cfg2:
    user_date_input = st.date_input("📅 วันที่ที่ต้องการ (กรณีชีทไม่มี WW)", value=None)
    protect_formulas = st.checkbox("🛡️ ป้องกันเขียนทับสูตร", value=True)

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

if filter_mode_label.startswith("แสดงเฉพาะ <="):
    filter_mode = "lte"
    hide_out_of_range = True
elif filter_mode_label.startswith("แสดงเฉพาะ >="):
    filter_mode = "gte"
    hide_out_of_range = True
else:
    filter_mode = "lte"
    hide_out_of_range = False

if cust_zip_upload and raw_zip_upload:
    with st.spinner("กำลังสแกนไฟล์ ตรวจสอบ Spec ข้อมูลเดิม และจับคู่โมเดล..."):
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

        scanned_models = []
        global_ng_rows = []

        for c_name, r_name in paired_models:
            profile_name, model_cfg = detect_model_profile(c_name)
            wb_cust_scan = openpyxl.load_workbook(io.BytesIO(cust_files_map[c_name]), data_only=True)
            wb_raw_scan = openpyxl.load_workbook(io.BytesIO(raw_files_map[r_name]), data_only=True)

            sheets_info = {}
            model_ng_count = 0

            for sheet_key, config in model_cfg.items():
                ws_c = find_sheet_in_wb(wb_cust_scan, sheet_key)
                ws_r = find_sheet_in_wb(wb_raw_scan, sheet_key)
                if ws_c is None or ws_r is None:
                    sheets_info[sheet_key] = None
                    continue

                cust_entries = scan_all_valid_entries(ws_c, ws_c, sheet_key, config, check_spec=True)
                raw_entries = scan_all_valid_entries(ws_r, ws_c, sheet_key, config, check_spec=True)

                cust_ng_entries = [e for e in cust_entries if e["is_ng"]]
                model_ng_count += len(cust_ng_entries)
                for ng_e in cust_ng_entries:
                    global_ng_rows.append({
                        "โมเดล": c_name,
                        "ชีท": sheet_key,
                        "คอลัมน์": ng_e["col_range_str"],
                        "WW / วันที่": f"WW{ng_e['ww_num']:02d}" if ng_e["ww_num"] is not None else (ng_e["date_val"].strftime("%d-%b-%y") if ng_e["date_val"] else "-"),
                        "หัวตาราง": ng_e["full_sig"],
                        "จุดที่เกิน Spec": ", ".join(ng_e["failures"])
                    })

                last_cust = cust_entries[-1] if cust_entries else None
                next_dest_col = last_cust["end_col"] + 1 if last_cust else config["start_col"]
                available_to_select = filter_new_raw_entries(cust_entries, raw_entries)

                sheets_info[sheet_key] = {
                    "config": config,
                    "cust_entries": cust_entries,
                    "cust_ng_entries": cust_ng_entries,
                    "last_cust": last_cust,
                    "next_dest_col": next_dest_col,
                    "available": available_to_select
                }

            scanned_models.append({
                "c_name": c_name,
                "r_name": r_name,
                "profile_name": profile_name,
                "model_cfg": model_cfg,
                "sheets_info": sheets_info,
                "model_ng_count": model_ng_count
            })

    st.markdown("---")
    st.subheader(f"📋 จับคู่สำเร็จ {len(scanned_models)} โมเดล")

    if global_ng_rows:
        st.error(f"🚨 **ตรวจพบคอลัมน์เดิมในไฟล์ลูกค้าที่เกิน Spec ทั้งหมด {len(global_ng_rows)} คอลัมน์ (ระบบติ๊กเตรียมลบทิ้งทั้งคอลัมน์ให้อัตโนมัติแล้ว):**")
        st.dataframe(pd.DataFrame(global_ng_rows), use_container_width=True, hide_index=True)
    else:
        st.success("✅ ตรวจสอบคอลัมน์เดิมในไฟล์ลูกค้าทุกโมเดลแล้ว ไม่พบคอลัมน์ที่เกิน Spec")

    all_models_data = {}

    for file_idx, m_data in enumerate(scanned_models):
        c_name = m_data["c_name"]
        r_name = m_data["r_name"]
        profile_name = m_data["profile_name"]
        model_cfg = m_data["model_cfg"]
        sheets_info = m_data["sheets_info"]
        ng_badge = f" — 🚨 พบ NG เดิม {m_data['model_ng_count']} คอลัมน์" if m_data["model_ng_count"] > 0 else ""

        pair_title = f"📦 โมเดล: {c_name}  —  [โปรไฟล์: {profile_name}]{ng_badge}"

        with st.expander(pair_title, expanded=(file_idx == 0 or m_data["model_ng_count"] > 0)):
            model_selections = {}
            sheet_names_list = list(model_cfg.keys())
            tab_labels = []
            for sk in sheet_names_list:
                si = sheets_info.get(sk)
                if si and len(si["cust_ng_entries"]) > 0:
                    tab_labels.append(f"{sk} (🚨{len(si['cust_ng_entries'])})")
                else:
                    tab_labels.append(sk)

            tabs = st.tabs(tab_labels)

            for idx, sheet_key in enumerate(sheet_names_list):
                s_info = sheets_info.get(sheet_key)
                with tabs[idx]:
                    if s_info is None:
                        st.warning(f"⚠️ ไม่พบชีท **{sheet_key}** ในโมเดลนี้")
                        continue

                    config = s_info["config"]
                    cust_entries = s_info["cust_entries"]
                    cust_ng_entries = s_info["cust_ng_entries"]
                    last_cust = s_info["last_cust"]
                    next_dest_col = s_info["next_dest_col"]
                    available_to_select = s_info["available"]

                    delete_selections = []
                    if cust_ng_entries:
                        st.error(f"🗑️ **พบคอลัมน์เดิมที่เกิน Spec จำนวน {len(cust_ng_entries)} คอลัมน์ (ติ๊กเพื่อลบทิ้งทั้งคอลัมน์และขยับชิด):**")
                        for ng_e in cust_ng_entries:
                            del_chk = st.checkbox(
                                f"ลบทิ้ง: {ng_e['label']}",
                                value=True,
                                key=f"del_{c_name}_{sheet_key}_{ng_e['col']}"
                            )
                            if del_chk:
                                delete_selections.append(ng_e)
                        st.markdown("---")

                    if last_cust:
                        st.info(f"📌 **ข้อมูลล่าสุดในไฟล์ลูกค้า:** `{last_cust['label']}` (คอลัมน์ถัดไปที่จะกรอก: **{get_column_letter(next_dest_col)}**)")
                    else:
                        st.info(f"📌 **ยังไม่มีข้อมูลในไฟล์ลูกค้า** (จะเริ่มกรอกที่คอลัมน์ **{get_column_letter(next_dest_col)}**)")

                    sheet_selections = []
                    if not available_to_select:
                        st.success("✅ ไม่มีข้อมูลใหม่เพิ่มเติมในไฟล์ดิบ")
                    else:
                        shown_entries = []
                        hidden_count = 0
                        for entry in available_to_select:
                            in_range = matches_user_filter(entry, target_ww_num, target_date_val, filter_mode)
                            if hide_out_of_range and not in_range:
                                hidden_count += 1
                                continue
                            shown_entries.append((entry, in_range))

                        if hidden_count > 0:
                            st.caption(f"🙈 ซ่อนรายการที่อยู่นอกเงื่อนไข WW/วันที่ ไปแล้ว {hidden_count} รายการ")

                        if not shown_entries:
                            st.info("ไม่มีรายการใหม่ที่ตรงกับเงื่อนไข WW / วันที่ ที่เลือก")
                        else:
                            for entry, in_range in shown_entries:
                                default_checked = (not entry["is_ng"]) and in_range
                                checked = st.checkbox(
                                    entry["label"],
                                    value=default_checked,
                                    key=f"chk_{c_name}_{sheet_key}_{entry['col']}_{target_ww_num}_{target_date_val}_{filter_mode}"
                                )
                                if checked:
                                    sheet_selections.append(entry)

                    model_selections[sheet_key] = {
                        "all_cust_entries": cust_entries,
                        "delete_entries": delete_selections,
                        "next_col": next_dest_col,
                        "entries": sheet_selections,
                        "config": config
                    }

            all_models_data[c_name] = {
                "raw_name": r_name,
                "selections": model_selections
            }

    st.markdown("---")
    if st.button("🚀 ประมวลผล (ลบ NG เก่า + เติมข้อมูลใหม่) และดาวน์โหลดไฟล์", type="primary", use_container_width=True):
        with st.spinner("กำลังลบคอลัมน์ที่เกิน Spec และเติมข้อมูลใหม่ลงไฟล์ต้นฉบับโดยตรง (XML Patching)..."):
            total_added_all = 0
            total_deleted_all = 0
            deleted_log = []
            processed_files = {}

            for c_name, data in all_models_data.items():
                r_name = data["raw_name"]
                for sk, splan in data["selections"].items():
                    for de in splan.get("delete_entries", []):
                        deleted_log.append(f"**{c_name}** | ชีท **{sk}** | {de['label']}")

                patched_bytes, added_cnt, deleted_cnt = process_single_model_xml_safe(
                    cust_files_map[c_name],
                    raw_files_map[r_name],
                    data["selections"],
                    protect_formulas=protect_formulas
                )
                processed_files[c_name] = patched_bytes
                total_added_all += added_cnt
                total_deleted_all += deleted_cnt

            st.success(f"🎉 ประมวลผลสำเร็จ! เติมข้อมูลใหม่ **{total_added_all}** ชุด | ลบข้อมูลเดิมที่เกิน Spec ออก **{total_deleted_all}** คอลัมน์ (จาก {len(processed_files)} โมเดล)")

            if deleted_log:
                with st.expander(f"🗑️ รายละเอียดคอลัมน์ที่ถูกลบออกทั้งหมด ({len(deleted_log)} คอลัมน์)", expanded=True):
                    for item in deleted_log:
                        st.markdown(f"- {item}")

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
