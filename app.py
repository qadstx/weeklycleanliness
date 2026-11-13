import streamlit as st
import openpyxl
from openpyxl.utils import get_column_letter
from copy import copy
import io
import datetime

st.set_page_config(page_title="Customer Data Auto-Filler", layout="wide")
st.title("📊 ระบบคัดกรองและเติมข้อมูลส่งลูกค้าอัตโนมัติ (8 Sheets)")

# ==========================================
# ตั้งค่าพิกัดแถวและรูปแบบคอลัมน์ของทั้ง 8 ชีท
# ==========================================
def expand_ranges(ranges_list):
    """แปลงรายการแถวและช่วงแถว (tuple) ให้เป็น list ของตัวเลขแถวทั้งหมด"""
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
        "id_row": 31,          # แถวชื่อ Lot (เช่น HNRJ0415B2)
        "sub_rows": [29, 30],  # WW และ Date
        "start_col": 2,
        "copy_rows": expand_ranges([(29, 34), 44])
    },
    "IC": {
        "step": 2,
        "id_row": 19,          # แถวชื่อ Lab No. / WW
        "sub_rows": [20],      # Date / Month
        "start_col": 2,
        "copy_rows": expand_ranges([(19, 27)]) # ข้าม Row 28 (*Total anion) เพราะไฟล์ลูกค้ามีสูตรผูกไว้แล้ว
    },
    "FTIR": {
        "step": 2,
        "id_row": 5,           # แถวชื่อ Lab No. / WW
        "sub_rows": [6],       # Date / Month
        "start_col": 2,
        "copy_rows": expand_ranges([5, 6, 7, 9])
    },
    "APA": {
        "step": 2,
        "id_row": 38,          # แถวหัวตาราง APA
        "sub_rows": [39, 40],
        "start_col": 4,        # เริ่มสแกนตั้งแต่คอลัมน์ D เป็นต้นไป
        # หมายเหตุ: หากแถว Total ของ APA ในไฟล์ลูกค้าผูกสูตรไว้แล้ว โค้ดจะมีตัวเลือกให้ข้ามการทับสูตรอัตโนมัติ
        "copy_rows": expand_ranges([
            (38, 50), (62, 64), (66, 71), 75, 77,
            (79, 94), (98, 116), 118, (120, 129), 131
        ])
    },
    "DHS": {
        "step": 1,
        "id_row": 51,          # แถวชื่อ Lab No.
        "sub_rows": [52, 53],  # Date และ WW
        "start_col": 4,
        "copy_rows": expand_ranges([
            (51, 54), 57, 60, 63, 66, 69, 72, 75, 78, 81, 84, 87, 90, 93
        ])
    },
    "GCMS": {
        "step": 2,
        "id_row": 40,          # แถวหัวตาราง GCMS
        "sub_rows": [41, 42],
        "start_col": 4,
        "copy_rows": expand_ranges([40, (43, 45), 49, 53])
    },
    "GCMS RHC": {
        "step": 2,
        "id_row": 37,          # แถวหัวตาราง GCMS RHC (รวม 36-37)
        "sub_rows": [36, 38],
        "start_col": 4,
        "copy_rows": expand_ranges([36, 37, (39, 41), 45, 49, (53, 55)])
    },
    "Ghost Test": {
        "step": 1,
        "id_row": 4,           # แถวชื่อ Lab No.
        "sub_rows": [5, 6],    # Month และ WW
        "start_col": 2,
        "copy_rows": expand_ranges([(4, 9)])
    }
}

# ==========================================
# ฟังก์ชันจัดการข้อมูลใน Excel
# ==========================================
def format_val(val):
    if val is None:
        return ""
    if isinstance(val, datetime.datetime):
        return val.strftime("%d-%b-%y")
    return str(val).strip()

def is_valid_id(val):
    """เช็คว่าช่องหัวตารางมีข้อมูลจริงหรือไม่ (ไม่เอาช่องว่าง หรือช่องที่มีแค่คำว่า WW. รอไว้)"""
    if val is None:
        return False
    s = str(val).strip()
    if s == "" or s.upper() in ["WW.", "WW", "DATE", "ION", "SPEC.", "SPEC", "NG/CM2", "NG/CM²"]:
        return False
    if s.startswith("="):
        return False
    return True

def get_merged_width(ws, row, col):
    """เช็คว่าเซลล์นั้นถูก Merge กี่คอลัมน์"""
    for m_range in ws.merged_cells.ranges:
        if m_range.min_row <= row <= m_range.max_row and m_range.min_col <= col <= m_range.max_col:
            return m_range.max_col - m_range.min_col + 1
    return 1

def scan_entries(ws, config):
    """สแกนหาชุดข้อมูลทั้งหมดในชีท"""
    entries = []
    id_row = config["id_row"]
    step = config["step"]
    col = config["start_col"]
    max_col = ws.max_column

    while col <= max_col:
        val = ws.cell(row=id_row, column=col).value
        if is_valid_id(val):
            id_str = format_val(val)
            # ดึงข้อมูลแถวรอง (เช่น WW, Date) มาโชว์คู่กันเพื่อให้เลือกง่ายขึ้น
            sub_info = []
            for r in config.get("sub_rows", []):
                sv = format_val(ws.cell(row=r, column=col).value)
                if sv and sv != id_str:
                    sub_info.append(sv)
            
            label = f"{id_str}" + (f" ({' | '.join(sub_info)})" if sub_info else "")
            entries.append({
                "col": col,
                "id": id_str,
                "label": f"[Col {get_column_letter(col)}] {label}"
            })
            col += step
        else:
            col += 1
    return entries

def copy_cell_exact(src_cell, dst_cell, protect_formulas=True):
    """คัดลอกค่า, สูตร, สี, ฟอนต์, เส้นขอบ, Format และ Comment เป๊ะๆ"""
    # ถ้าช่องปลายทางในไฟล์ลูกค้ามีสูตรผูกไว้แล้ว (เช่น =SUM, =AVERAGE) และเปิดโหมดป้องกันสูตรไว้ จะไม่เขียนทับสูตร
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
    """นำข้อมูล 1 ชุด (1 หรือ 2 คอลัมน์) จากไฟล์ดิบไปวางต่อท้ายไฟล์ลูกค้า"""
    step = config["step"]
    copy_rows = config["copy_rows"]

    for r in copy_rows:
        # ตรวจสอบว่าในไฟล์ดิบ แถวนี้มีการ Merge คอลัมน์หรือไม่ (เช่น หัวตาราง 2 คอลัมน์)
        merge_w = get_merged_width(ws_raw, r, src_col)

        # ยกเลิก Merge เดิมในพื้นที่ปลายทางก่อน (ถ้ามี) เพื่อป้องกัน Error
        for m_range in list(ws_cust.merged_cells.ranges):
            if m_range.min_row <= r <= m_range.max_row and not (m_range.max_col < dst_col or m_range.min_col >= dst_col + step):
                ws_cust.unmerge_cells(str(m_range))

        # คัดลอกข้อมูลทีละคอลัมน์ย่อย (1 หรือ 2 คอลัมน์)
        for offset in range(step):
            s_cell = ws_raw.cell(row=r, column=src_col + offset)
            d_cell = ws_cust.cell(row=r, column=dst_col + offset)
            copy_cell_exact(s_cell, d_cell, protect_formulas=protect_formulas)

        # ถ้าต้นทาง Merge ไว้ ให้ Merge ปลายทางให้เหมือนกันเป๊ะ
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
    # โหลดไฟล์สำหรับอ่านค่า (data_only=False เพื่อเช็ครูปแบบและสูตร)
    wb_cust_scan = openpyxl.load_workbook(cust_file, data_only=True)
    wb_raw_scan = openpyxl.load_workbook(raw_file, data_only=True)

    st.markdown("---")
    st.subheader("📋 เลือกข้อมูลที่ต้องการเติมลงในไฟล์ส่งลูกค้า")

    selected_items = {}
    tabs = st.tabs(list(SHEET_CONFIGS.keys()))

    for idx, (sheet_name, config) in enumerate(SHEET_CONFIGS.items()):
        with tabs[idx]:
            if sheet_name not in wb_cust_scan.sheetnames or sheet_name not in wb_raw_scan.sheetnames:
                st.warning(f"⚠️ ไม่พบชีทชื่อ **{sheet_name}** ในไฟล์ใดไฟล์หนึ่ง")
                continue

            ws_c = wb_cust_scan[sheet_name]
            ws_r = wb_raw_scan[sheet_name]

            cust_entries = scan_entries(ws_c, config)
            raw_entries = scan_entries(ws_r, config)

            # หาตัวสุดท้ายที่ส่งลูกค้าไปแล้ว
            if cust_entries:
                last_cust_entry = cust_entries[-1]
                last_cust_id = last_cust_entry["id"]
                next_dest_col = last_cust_entry["col"] + config["step"]
                st.info(f"📌 **ข้อมูลล่าสุดในไฟล์ลูกค้า:** `{last_cust_entry['label']}` (คอลัมน์ถัดไปที่จะกรอกคือ **{get_column_letter(next_dest_col)}**)")
            else:
                last_cust_id = None
                next_dest_col = config["start_col"]
                st.info("📌 **ยังไม่มีข้อมูลในไฟล์ลูกค้า** ระบบจะแสดงข้อมูลทั้งหมดจากไฟล์ดิบ")

            # กรองเฉพาะข้อมูลในไฟล์ดิบที่อยู่หลังตัวสุดท้ายของไฟล์ลูกค้า
            cust_id_set = {e["id"] for e in cust_entries}
            available_to_select = []

            if last_cust_id:
                # หาตำแหน่งของ last_cust_id ในไฟล์ดิบ (เอาตัวสุดท้ายที่ชื่อตรงกันเผื่อชื่อซ้ำ)
                match_indices = [i for i, e in enumerate(raw_entries) if e["id"] == last_cust_id]
                if match_indices:
                    cutoff_idx = match_indices[-1]
                    available_to_select = raw_entries[cutoff_idx + 1:]
                else:
                    # กรณีหาชื่อตัวสุดท้ายไม่เจอในไฟล์ดิบ ให้ดึงทุกตัวที่ชื่อยังไม่อยู่ในไฟล์ลูกค้า
                    available_to_select = [e for e in raw_entries if e["id"] not in cust_id_set]
            else:
                available_to_select = raw_entries

            if not available_to_select:
                st.success("✅ ข้อมูลในชีทนี้ส่งครบถ้วนแล้ว ไม่มีข้อมูลใหม่ให้เลือก")
                selected_items[sheet_name] = []
            else:
                st.write(f"พบข้อมูลใหม่ที่ยังไม่ได้ส่งจำนวน **{len(available_to_select)}** ชุด:")
                
                # ปุ่มเลือกทั้งหมดในชีทนั้น
                select_all = st.checkbox(f"เลือกทั้งหมดในชีท {sheet_name}", value=True, key=f"all_{sheet_name}")
                
                sheet_selections = []
                cols_ui = st.columns(3)
                for i, entry in enumerate(available_to_select):
                    with cols_ui[i % 3]:
                        checked = st.checkbox(
                            entry["label"],
                            value=select_all,
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
            # โหลด Workbook จริงสำหรับแก้ไขและรักษาสูตร/กราฟทั้งหมด
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
                    curr_dst_col += cfg["step"] # ขยับชิดต่อกันทันที
                    total_added += 1

            # บันทึกไฟล์ลง Memory Buffer ให้กดดาวน์โหลด
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
