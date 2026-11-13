import streamlit as st
import openpyxl
from openpyxl.utils import get_column_letter
from copy import copy
import io
import datetime

st.set_page_config(page_title="Customer Data Auto-Filler", layout="wide")
st.title("📊 ระบบคัดกรองและเติมข้อมูลส่งลูกค้าอัตโนมัติ (พร้อมเช็ค Spec)")

# ==========================================
# ตั้งค่าพิกัดแถว รูปแบบคอลัมน์ และเงื่อนไข Spec ของทั้ง 8 ชีท
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
    "LPC": {
        "step": 1,
        "id_row": 31,
        "sub_rows": [29, 30],
        "start_col": 2,
        "copy_rows": expand_ranges([(29, 34), 44]),
        "spec_checks": [(32, 41), (33, 41), (34, 41)]
    },
    "IC": {
        "step": 2,
        "id_row": 19,
        "sub_rows": [20],
        "start_col": 2,
        "copy_rows": expand_ranges([(19, 27)]),
        "spec_checks": [(28, 32)]
    },
    "FTIR": {
        "step": 2,
        "id_row": 5,
        "sub_rows": [6],
        "start_col": 2,
        "copy_rows": expand_ranges([5, 6, 7, 9]),
        "spec_checks": [(7, 8), (9, 10)]
    },
    "APA": {
        "step": 2,
        "id_row": 38,
        "sub_rows": [39, 40],
        "start_col": 4,
        "copy_rows": expand_ranges([
            (38, 58), (62, 64), (66, 71), 75, 77,
            (79, 94), (98, 116), 118, (120, 129), 131
        ]),
        "spec_checks": [
            (58, 59),
            (62, 65), (63, 65), (64, 65),
            (71, 72),
            (94, 95),
            (118, 119)
        ]
    },
    "DHS": {
        "step": 1,
        "id_row": 51,
        "sub_rows": [52, 53],
        "start_col": 4,
        "copy_rows": expand_ranges([
            (51, 54), 57, 60, 63, 66, 69, 72, 75, 78, 81, 84, 87, 90, 93
        ]),
        "spec_checks": [
            (r, r + 1) for r in [54, 57, 60, 63, 66, 69, 72, 75, 78, 81, 84, 87, 90, 93]
        ]
    },
    "GCMS": {
        "step": 2,
        "id_row": 40,
        "sub_rows": [41, 42],
        "start_col": 4,
        "copy_rows": expand_ranges([40, (43, 45), 49, 53]),
        "spec_checks": [(r, r + 1) for r in [43, 44, 45, 49, 53]]
    },
    "GCMS RHC": {
        "step": 2,
        "id_row": 37,
        "sub_rows": [36, 38],
        "start_col": 4,
        "copy_rows": expand_ranges([36, 37, (39, 41), 45, 49, (53, 55)]),
        "spec_checks": [(41, 42), (45, 46), (49, 50)]
    },
    "Ghost Test": {
        "step": 1,
        "id_row": 4,
        "sub_rows": [5, 6],
        "start_col": 2,
        "copy_rows": expand_ranges([(4, 9)]),
        "spec_checks": []
    }
}

# ==========================================
# ฟังก์ชันจัดการข้อมูลและตรวจสอบ Spec
# ==========================================
def format_val(val):
    if val is None:
        return ""
    if isinstance(val, datetime.datetime):
        return val.strftime("%d-%b-%y")
    return str(val).strip()

def to_float(val):
    """แปลงค่าในเซลล์เป็นตัวเลขเพื่อเทียบ Spec (ถ้าเป็น NA, ND หรือข้อความจะคืนค่า None)"""
    if val is None:
        return None
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip().replace(",", "")
    if s == "" or s.upper() in ["NA", "N/A", "ND", "-", "PASS", "FAIL"]:
        return None
    try:
        return float(s)
    except ValueError:
        return None

def find_spec_value(ws_raw, ws_cust, spec_row, col):
    """ดึงค่า Spec จากคอลัมน์ปัจจุบัน ถ้าช่องนั้นว่างให้วิ่งหาค่า Spec จากคอลัมน์ข้างเคียงในแถวเดียวกัน"""
    for ws in (ws_raw, ws_cust):
        if ws is None:
            continue
        v = to_float(ws.cell(row=spec_row, column=col).value)
        if v is not None:
            return v
    # กรณีคอลัมน์ใหม่ในไฟล์ดิบยังไม่ได้ก๊อปเลข Spec ยาวไปทางขวา ให้ดึงค่า Spec มาตรฐานจากคอลัมน์ที่มีในแถวเดียวกัน
    for ws in (ws_cust, ws_raw):
        if ws is None:
            continue
        for c in range(2, min(ws.max_column + 1, 60)):
            v = to_float(ws.cell(row=spec_row, column=c).value)
            if v is not None:
                return v
    return None

def evaluate_entry_spec(ws_raw, ws_cust, start_col, config):
    """ตรวจสอบว่าชุดข้อมูลนั้นมีแถวไหนหลุด Spec หรือไม่"""
    step = config["step"]
    spec_checks = config.get("spec_checks", [])
    failures = []

    for val_row, spec_row in spec_checks:
        for offset in range(step):
            c = start_col + offset
            val_num = to_float(ws_raw.cell(row=val_row, column=c).value)
            if val_num is None:
                continue
            spec_num = find_spec_value(ws_raw, ws_cust, spec_row, c)
            if spec_num is not None and val_num > spec_num:
                col_let = get_column_letter(c)
                failures.append(f"Row {val_row} ({col_let}{val_row}={val_num:g} > Spec {spec_num:g})")
    return failures

def is_valid_id(val):
    if val is None:
        return False
    s = str(val).strip()
    if s == "" or s.upper() in ["WW.", "WW", "DATE", "ION", "SPEC.", "SPEC", "NG/CM2", "NG/CM²"]:
        return False
    if s.startswith("="):
        return False
    return True

def get_merged_width(ws, row, col):
    for m_range in ws.merged_cells.ranges:
        if m_range.min_row <= row <= m_range.max_row and m_range.min_col <= col <= m_range.max_col:
            return m_range.max_col - m_range.min_col + 1
    return 1

def scan_entries(ws_target, ws_cust_ref, config, check_spec=False):
    entries = []
    id_row = config["id_row"]
    step = config["step"]
    col = config["start_col"]
    max_col = ws_target.max_column

    while col <= max_col:
        val = ws_target.cell(row=id_row, column=col).value
        if is_valid_id(val):
            id_str = format_val(val)
            sub_info = []
            for r in config.get("sub_rows", []):
                sv = format_val(ws_target.cell(row=r, column=col).value)
                if sv and sv != id_str:
                    sub_info.append(sv)

            failures = []
            if check_spec:
                failures = evaluate_entry_spec(ws_target, ws_cust_ref, col, config)

            base_label = f"[Col {get_column_letter(col)}] {id_str}" + (f" ({' | '.join(sub_info)})" if sub_info else "")
            if failures:
                display_label = f"⚠️ {base_label} — หลุด Spec: {', '.join(failures)}"
            else:
                display_label = f"✅ {base_label}"

            entries.append({
                "col": col,
                "id": id_str,
                "label": display_label,
                "is_ng": len(failures) > 0,
                "failures": failures
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
        dst_cell.font = copy(src_cell.font)
        dst_cell.fill = copy(src_cell.fill)
        dst_cell.border = copy(src_cell.border)
        dst_cell.alignment = copy(src_cell.alignment)
        dst_cell.number_format = src_cell.number_format

    if src_cell.comment:
        dst_cell.comment = copy(src_cell.comment)
    else:
        dst_cell.comment = None

    if src_cell.hyperlink:
        dst_cell.hyperlink = copy(src_cell.hyperlink)

def append_entry_to_customer(ws_raw, ws_cust, src_col, dst_col, config, protect_formulas=True):
    step = config["step"]
    copy_rows = config["copy_rows"]

    for r in copy_rows:
        merge_w = get_merged_width(ws_raw, r, src_col)

        for m_range in list(ws_cust.merged_cells.ranges):
            if m_range.min_row <= r <= m_range.max_row and not (m_range.max_col < dst_col or m_range.min_col >= dst_col + step):
                ws_cust.unmerge_cells(str(m_range))

        for offset in range(step):
            s_cell = ws_raw.cell(row=r, column=src_col + offset)
            d_cell = ws_cust.cell(row=r, column=dst_col + offset)
            copy_cell_exact(s_cell, d_cell, protect_formulas=protect_formulas)

        if merge_w > 1 and merge_w <= step:
            ws_cust.merge_cells(
                start_row=r, start_column=dst_col,
                end_row=r, end_column=dst_col + merge_w - 1
            )

# ==========================================
# ส่วนแสดงผลหน้าเว็บ (Streamlit UI)
# ==========================================
col_up1, col_up2 = st.columns(2)
with col_up1:
    cust_file = st.file_uploader("1️⃣ อัปโหลดไฟล์ส่งลูกค้า (Customer File)", type=["xlsx", "xlsm"])
with col_up2:
    raw_file = st.file_uploader("2️⃣ อัปโหลดไฟล์ข้อมูลดิบ (Raw Data File)", type=["xlsx", "xlsm"])

protect_formulas = st.checkbox(
    "🛡️ ป้องกันการเขียนทับช่องที่มีสูตร (เช่น Total / Average) ที่ผูกไว้แล้วในไฟล์ลูกค้า", 
    value=True
)

if cust_file and raw_file:
    wb_cust_scan = openpyxl.load_workbook(cust_file, data_only=True)
    wb_raw_scan = openpyxl.load_workbook(raw_file, data_only=True)

    st.markdown("---")
    st.subheader("📋 เลือกข้อมูลที่ต้องการเติมลงในไฟล์ส่งลูกค้า (ระบบติ๊กเลือกเฉพาะตัวที่ผ่าน Spec ให้อัตโนมัติ)")

    selected_items = {}
    tabs = st.tabs(list(SHEET_CONFIGS.keys()))

    for idx, (sheet_name, config) in enumerate(SHEET_CONFIGS.items()):
        with tabs[idx]:
            if sheet_name not in wb_cust_scan.sheetnames or sheet_name not in wb_raw_scan.sheetnames:
                st.warning(f"⚠️ ไม่พบชีทชื่อ **{sheet_name}** ในไฟล์ใดไฟล์หนึ่ง")
                continue

            ws_c = wb_cust_scan[sheet_name]
            ws_r = wb_raw_scan[sheet_name]

            cust_entries = scan_entries(ws_c, ws_c, config, check_spec=False)
            raw_entries = scan_entries(ws_r, ws_c, config, check_spec=True)

            if cust_entries:
                last_cust_entry = cust_entries[-1]
                last_cust_id = last_cust_entry["id"]
                next_dest_col = last_cust_entry["col"] + config["step"]
                st.info(f"📌 **ข้อมูลล่าสุดในไฟล์ลูกค้า:** `{last_cust_entry['label']}` (คอลัมน์ถัดไปที่จะกรอกคือ **{get_column_letter(next_dest_col)}**)")
            else:
                last_cust_id = None
                next_dest_col = config["start_col"]
                st.info("📌 **ยังไม่มีข้อมูลในไฟล์ลูกค้า** ระบบจะแสดงข้อมูลทั้งหมดจากไฟล์ดิบ")

            cust_id_set = {e["id"] for e in cust_entries}
            available_to_select = []

            if last_cust_id:
                match_indices = [i for i, e in enumerate(raw_entries) if e["id"] == last_cust_id]
                if match_indices:
                    cutoff_idx = match_indices[-1]
                    available_to_select = raw_entries[cutoff_idx + 1:]
                else:
                    available_to_select = [e for e in raw_entries if e["id"] not in cust_id_set]
            else:
                available_to_select = raw_entries

            if not available_to_select:
                st.success("✅ ข้อมูลในชีทนี้ส่งครบถ้วนแล้ว ไม่มีข้อมูลใหม่ให้เลือก")
                selected_items[sheet_name] = []
            else:
                pass_count = sum(1 for e in available_to_select if not e["is_ng"])
                ng_count = sum(1 for e in available_to_select if e["is_ng"])
                st.write(f"พบข้อมูลใหม่ **{len(available_to_select)}** ชุด (✅ ผ่าน Spec **{pass_count}** ชุด | ⚠️ หลุด Spec **{ng_count}** ชุด):")

                sheet_selections = []
                for entry in available_to_select:
                    # ค่าเริ่มต้น: ถ้าไม่หลุด Spec (is_ng == False) จะติ๊กเลือกให้อัตโนมัติ (True)
                    # ถ้าหลุด Spec (is_ng == True) จะไม่ติ๊ก (False) เพื่อให้ผู้ใช้ดูและตัดสินใจติ๊กเอง
                    default_checked = not entry["is_ng"]
                    checked = st.checkbox(
                        entry["label"],
                        value=default_checked,
                        key=f"chk_{sheet_name}_{entry['col']}"
                    )
                    if checked:
                        sheet_selections.append(entry)

                selected_items[sheet_name] = {
                    "next_col": next_dest_col,
                    "entries": sheet_selections
                }

    st.markdown("---")
    if st.button("🚀 ประมวลผลและสร้างไฟล์ส่งลูกค้า", type="primary", use_container_width=True):
        with st.spinner("กำลังคัดลอกข้อมูล รูปแบบตาราง และ Comment ลงในไฟล์ลูกค้า..."):
            cust_file.seek(0)
            raw_file.seek(0)
            wb_cust_out = openpyxl.load_workbook(cust_file)
            wb_raw_src = openpyxl.load_workbook(raw_file)

            total_added = 0
            for sheet_name, data in selected_items.items():
                if not data or not data["entries"]:
                    continue

                ws_c_out = wb_cust_out[sheet_name]
                ws_r_src = wb_raw_src[sheet_name]
                cfg = SHEET_CONFIGS[sheet_name]
                curr_dst_col = data["next_col"]

                for entry in data["entries"]:
                    src_col = entry["col"]
                    append_entry_to_customer(
                        ws_r_src, ws_c_out,
                        src_col=src_col,
                        dst_col=curr_dst_col,
                        config=cfg,
                        protect_formulas=protect_formulas
                    )
                    curr_dst_col += cfg["step"]
                    total_added += 1

            output_buffer = io.BytesIO()
            wb_cust_out.save(output_buffer)
            output_buffer.seek(0)

            st.success(f"🎉 เติมข้อมูลสำเร็จทั้งหมด **{total_added}** ชุดข้อมูล! กราฟ สูตร และ Comment อยู่ครบถ้วน")
            st.download_button(
                label="📥 ดาวน์โหลดไฟล์ส่งลูกค้า (Updated Customer File)",
                data=output_buffer,
                file_name=f"Updated_{cust_file.name}",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
