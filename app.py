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
st.title("📦 ระบบเติมข้อมูลส่งลูกค้าอัตโนมัติ (รองรับ ZIP หลายโมเดล)")

# ==========================================
# ฟังก์ชันสกัดตัวเลข WW
# ==========================================
def extract_ww_num(text):
    match = re.search(r'WW\.?\s*(\d+)', str(text), re.IGNORECASE)
    return int(match.group(1)) if match else None

# ==========================================
# ตั้งค่าพิกัดแถว รูปแบบคอลัมน์ และเงื่อนไข Spec
# ==========================================
def expand_ranges(ranges_list):
    rows = []
    for item in ranges_list:
        if isinstance(item, tuple):
            rows.extend(range(item[0], item[1] + 1))
        else:
            rows.append(item)
    return sorted(list(set(rows)))

SHEET_CONFIGS = {
    "LPC": {"step": 1, "id_row": 31, "sub_rows": [29, 30], "start_col": 2, "copy_rows": expand_ranges([(29, 34), 44]), "spec_checks": [(32, 41), (33, 41), (34, 41)]},
    "IC": {"step": 2, "id_row": 19, "sub_rows": [20], "start_col": 2, "copy_rows": expand_ranges([(19, 27)]), "spec_checks": [(28, 32)]},
    "FTIR": {"step": 2, "id_row": 5, "sub_rows": [6], "start_col": 2, "copy_rows": expand_ranges([5, 6, 7, 9]), "spec_checks": [(7, 8), (9, 10)]},
    "APA": {"step": 2, "id_row": 38, "sub_rows": [39, 40], "start_col": 4, "copy_rows": expand_ranges([(38, 58), (62, 64), (66, 71), 75, 77, (79, 94), (98, 116), 118, (120, 129), 131]), "spec_checks": [(58, 59), (62, 65), (63, 65), (64, 65), (71, 72), (94, 95), (118, 119)]},
    "DHS": {"step": 1, "id_row": 51, "sub_rows": [52, 53], "start_col": 4, "copy_rows": expand_ranges([(51, 54), 57, 60, 63, 66, 69, 72, 75, 78, 81, 84, 87, 90, 93]), "spec_checks": [(r, r + 1) for r in [54, 57, 60, 63, 66, 69, 72, 75, 78, 81, 84, 87, 90, 93]]},
    "GCMS": {"step": 2, "id_row": 40, "sub_rows": [41, 42], "start_col": 4, "copy_rows": expand_ranges([40, (43, 45), 49, 53]), "spec_checks": [(r, r + 1) for r in [43, 44, 45, 49, 53]]},
    "GCMS RHC": {"step": 2, "id_row": 37, "sub_rows": [36, 38], "start_col": 4, "copy_rows": expand_ranges([36, 37, (39, 41), 45, 49, (53, 55)]), "spec_checks": [(41, 42), (45, 46), (49, 50)]},
    "Ghost Test": {"step": 1, "id_row": 4, "sub_rows": [5, 6], "start_col": 2, "copy_rows": expand_ranges([(4, 9)]), "spec_checks": []}
}

# ==========================================
# Helper Functions
# ==========================================
def format_val(val):
    if val is None: return ""
    if isinstance(val, datetime.datetime): return val.strftime("%d-%b-%y")
    return str(val).strip()

def to_float(val):
    if val is None: return None
    if isinstance(val, (int, float)): return float(val)
    s = str(val).strip().replace(",", "")
    if s == "" or s.upper() in ["NA", "N/A", "ND", "-", "PASS", "FAIL"]: return None
    try: return float(s)
    except: return None

def find_spec_value(ws_raw, ws_cust, spec_row, col):
    for ws in (ws_raw, ws_cust):
        v = to_float(ws.cell(row=spec_row, column=col).value) if ws else None
        if v is not None: return v
    for ws in (ws_cust, ws_raw):
        if not ws: continue
        for c in range(2, min(ws.max_column + 1, 60)):
            v = to_float(ws.cell(row=spec_row, column=c).value)
            if v is not None: return v
    return None

def evaluate_entry_spec(ws_raw, ws_cust, start_col, config):
    failures = []
    for val_row, spec_row in config.get("spec_checks", []):
        for offset in range(config["step"]):
            c = start_col + offset
            val_num = to_float(ws_raw.cell(row=val_row, column=c).value)
            if val_num is None: continue
            spec_num = find_spec_value(ws_raw, ws_cust, spec_row, c)
            if spec_num is not None and val_num > spec_num:
                failures.append(f"Row {val_row} ({get_column_letter(c)}{val_row}={val_num:g} > Spec {spec_num:g})")
    return failures

def is_valid_id(val):
    if val is None: return False
    s = str(val).strip()
    return not (s == "" or s.upper() in ["WW.", "WW", "DATE", "ION", "SPEC.", "SPEC", "NG/CM2", "NG/CM²"] or s.startswith("="))

def get_merged_width(ws, row, col):
    for m in ws.merged_cells.ranges:
        if m.min_row <= row <= m.max_row and m.min_col <= col <= m.max_col:
            return m.max_col - m.min_col + 1
    return 1

def scan_entries(ws_target, ws_cust_ref, config, check_spec=False):
    entries = []
    id_row, step, col, max_col = config["id_row"], config["step"], config["start_col"], ws_target.max_column
    while col <= max_col:
        val = ws_target.cell(row=id_row, column=col).value
        if is_valid_id(val):
            id_str = format_val(val)
            sub_info = [format_val(ws_target.cell(row=r, column=col).value) for r in config.get("sub_rows", [])]
            sub_info = [s for s in sub_info if s and s != id_str]
            failures = evaluate_entry_spec(ws_target, ws_cust_ref, col, config) if check_spec else []
            base_label = f"[Col {get_column_letter(col)}] {id_str}" + (f" ({' | '.join(sub_info)})" if sub_info else "")
            
            entries.append({
                "col": col, "id": id_str,
                "label": f"⚠️ {base_label} — หลุด Spec: {', '.join(failures)}" if failures else f"✅ {base_label}",
                "is_ng": len(failures) > 0, "failures": failures,
                "ww_num": extract_ww_num(base_label)
            })
            col += step
        else: col += 1
    return entries

def copy_cell_exact(src_cell, dst_cell, protect_formulas=True):
    if protect_formulas and isinstance(dst_cell.value, str) and dst_cell.value.startswith("="): return
    dst_cell.value = src_cell.value
    if src_cell.has_style:
        dst_cell.font, dst_cell.fill = copy(src_cell.font), copy(src_cell.fill)
        dst_cell.border, dst_cell.alignment = copy(src_cell.border), copy(src_cell.alignment)
        dst_cell.number_format = src_cell.number_format
    dst_cell.comment = copy(src_cell.comment) if src_cell.comment else None
    if src_cell.hyperlink: dst_cell.hyperlink = copy(src_cell.hyperlink)

def append_entry_to_customer(ws_raw, ws_cust, src_col, dst_col, config, protect_formulas=True):
    step, copy_rows = config["step"], config["copy_rows"]
    for r in copy_rows:
        merge_w = get_merged_width(ws_raw, r, src_col)
        for m_range in list(ws_cust.merged_cells.ranges):
            if m_range.min_row <= r <= m_range.max_row and not (m_range.max_col < dst_col or m_range.min_col >= dst_col + step):
                ws_cust.unmerge_cells(str(m_range))
        for offset in range(step):
            copy_cell_exact(ws_raw.cell(row=r, column=src_col + offset), ws_cust.cell(row=r, column=dst_col + offset), protect_formulas)
        if merge_w > 1 and merge_w <= step:
            ws_cust.merge_cells(start_row=r, start_column=dst_col, end_row=r, end_column=dst_col + merge_w - 1)

# ==========================================
# Main UI
# ==========================================
st.markdown("### 1️⃣ ตั้งค่าและอัปโหลดไฟล์ (ZIP)")
col1, col2, col3 = st.columns([1, 1.2, 1.2])

with col1:
    user_ww_input = st.text_input("🎯 พิมพ์ WW ที่ต้องการส่งถึง", placeholder="เช่น 15, WW15 (ว่าง=ทั้งหมด)")
    protect_formulas = st.checkbox("🛡️ ป้องกันเขียนทับสูตร", value=True)
with col2:
    cust_zip_upload = st.file_uploader("📂 อัปโหลดไฟล์ ZIP ลูกค้า", type=["zip"])
with col3:
    raw_zip_upload = st.file_uploader("📂 อัปโหลดไฟล์ ZIP ข้อมูลดิบ", type=["zip"])

target_ww_num = 9999
if user_ww_input.strip():
    match = re.search(r'\d+', user_ww_input)
    if match: target_ww_num = int(match.group())

if cust_zip_upload and raw_zip_upload:
    with st.spinner("กำลังอ่านไฟล์ ZIP และจับคู่โมเดล..."):
        cz = zipfile.ZipFile(cust_zip_upload)
        rz = zipfile.ZipFile(raw_zip_upload)
        
        # กรองเฉพาะไฟล์ Excel
        cust_names = {os.path.basename(n): n for n in cz.namelist() if n.endswith(('.xlsx', '.xlsm')) and not os.path.basename(n).startswith('~')}
        raw_names = {os.path.basename(n): n for n in rz.namelist() if n.endswith(('.xlsx', '.xlsm')) and not os.path.basename(n).startswith('~')}
        
        matched_files = sorted(list(set(cust_names.keys()) & set(raw_names.keys())))
        
        if not matched_files:
            st.error("❌ ไม่พบไฟล์ที่ชื่อตรงกันเลยใน 2 ZIP นี้ กรุณาตรวจสอบให้แน่ใจว่าชื่อไฟล์ลูกค้าและไฟล์ดิบของแต่ละโมเดลตั้งชื่อเหมือนกันเป๊ะ")
            st.stop()

    st.markdown("---")
    st.subheader(f"📋 พบไฟล์ที่จับคู่ได้ {len(matched_files)} โมเดล (ตรวจสอบและแก้ไขก่อนสร้างไฟล์)")

    all_models_data = {}
    
    # วนลูปสร้าง UI สำหรับแต่ละไฟล์
    for file_name in matched_files:
        with st.expander(f"📦 โมเดล: {file_name}", expanded=False):
            # อ่านไฟล์ลง Memory
            cust_bytes = io.BytesIO(cz.read(cust_names[file_name]))
            raw_bytes = io.BytesIO(rz.read(raw_names[file_name]))
            
            wb_cust_scan = openpyxl.load_workbook(cust_bytes, data_only=True)
            wb_raw_scan = openpyxl.load_workbook(raw_bytes, data_only=True)
            
            model_selections = {}
            tabs = st.tabs(list(SHEET_CONFIGS.keys()))
            
            for idx, (sheet_name, config) in enumerate(SHEET_CONFIGS.items()):
                with tabs[idx]:
                    if sheet_name not in wb_cust_scan.sheetnames or sheet_name not in wb_raw_scan.sheetnames:
                        st.warning(f"⚠️ ไม่พบชีทนี้")
                        continue
                    
                    ws_c = wb_cust_scan[sheet_name]
                    ws_r = wb_raw_scan[sheet_name]
                    
                    cust_entries = scan_entries(ws_c, ws_c, config, check_spec=False)
                    raw_entries = scan_entries(ws_r, ws_c, config, check_spec=True)
                    
                    last_cust_id = cust_entries[-1]["id"] if cust_entries else None
                    next_dest_col = cust_entries[-1]["col"] + config["step"] if cust_entries else config["start_col"]
                    
                    if cust_entries: st.info(f"📌 ข้อมูลล่าสุด: `{cust_entries[-1]['label']}`")
                    
                    cust_id_set = {e["id"] for e in cust_entries}
                    available_to_select = []
                    if last_cust_id:
                        match_indices = [i for i, e in enumerate(raw_entries) if e["id"] == last_cust_id]
                        if match_indices: available_to_select = raw_entries[match_indices[-1] + 1:]
                        else: available_to_select = [e for e in raw_entries if e["id"] not in cust_id_set]
                    else: available_to_select = raw_entries
                    
                    if not available_to_select:
                        st.success("✅ ส่งครบถ้วนแล้ว")
                        model_selections[sheet_name] = []
                    else:
                        sheet_selections = []
                        for entry in available_to_select:
                            is_within_ww = (entry["ww_num"] <= target_ww_num) if entry["ww_num"] is not None else True
                            default_checked = (not entry["is_ng"]) and is_within_ww
                            
                            # ตั้งชื่อ key ของ checkbox ให้ไม่ซ้ำกันข้ามไฟล์
                            checked = st.checkbox(
                                entry["label"], value=default_checked, 
                                key=f"chk_{file_name}_{sheet_name}_{entry['col']}"
                            )
                            if checked: sheet_selections.append(entry)
                        
                        model_selections[sheet_name] = {"next_col": next_dest_col, "entries": sheet_selections}
            
            all_models_data[file_name] = {
                "cust_zip_path": cust_names[file_name],
                "raw_zip_path": raw_names[file_name],
                "selections": model_selections
            }

    st.markdown("---")
    if st.button("🚀 ประมวลผลและสร้าง ZIP ก้อนใหม่", type="primary", use_container_width=True):
        with st.spinner("กำลังคัดลอกข้อมูลและแพ็กเป็น ZIP ให้คุณ..."):
            out_zip_buffer = io.BytesIO()
            total_added_all = 0
            
            # เปิดไฟล์ ZIP สำหรับเขียน
            with zipfile.ZipFile(out_zip_buffer, "w", zipfile.ZIP_DEFLATED) as out_zip:
                for file_name, data in all_models_data.items():
                    cust_bytes = io.BytesIO(cz.read(data["cust_zip_path"]))
                    raw_bytes = io.BytesIO(rz.read(data["raw_zip_path"]))
                    
                    is_xlsm = file_name.lower().endswith(".xlsm")
                    wb_cust_out = openpyxl.load_workbook(cust_bytes, keep_vba=is_xlsm)
                    wb_raw_src = openpyxl.load_workbook(raw_bytes, data_only=True)
                    
                    file_has_updates = False
                    
                    for sheet_name, sheet_data in data["selections"].items():
                        if not sheet_data or not sheet_data["entries"]: continue
                        
                        ws_c_out = wb_cust_out[sheet_name]
                        ws_r_src = wb_raw_src[sheet_name]
                        cfg = SHEET_CONFIGS[sheet_name]
                        curr_dst_col = sheet_data["next_col"]
                        
                        for entry in sheet_data["entries"]:
                            append_entry_to_customer(ws_r_src, ws_c_out, entry["col"], curr_dst_col, cfg, protect_formulas)
                            curr_dst_col += cfg["step"]
                            total_added_all += 1
                            file_has_updates = True
                    
                    # บันทึกไฟล์ที่ถูกอัปเดตแต่ละตัวลงใน Memory
                    file_out_buffer = io.BytesIO()
                    wb_cust_out.save(file_out_buffer)
                    # เขียนไฟล์นั้นยัดกลับเข้าไปใน ZIP ปลายทาง
                    out_zip.writestr(file_name, file_out_buffer.getvalue())

            st.success(f"🎉 สร้าง ZIP สำเร็จ! เติมข้อมูลไปทั้งหมด **{total_added_all}** ชุดจาก {len(matched_files)} โมเดล")
            st.download_button(
                label="📥 ดาวน์โหลดไฟล์ ZIP ลูกค้าที่อัปเดตแล้ว",
                data=out_zip_buffer.getvalue(),
                file_name="Updated_Customer_Files.zip",
                mime="application/zip",
                use_container_width=True
            )
