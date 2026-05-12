import streamlit as st
import pandas as pd
import io
from datetime import datetime

st.set_page_config(page_title="Supplier & DO System V7.0", layout="wide")

# --- ส่วนหัวของโปรแกรม ---
st.title("📦 ระบบจัดการข้อมูลขนส่ง (Splitter & DO Generator)")
st.markdown("---")

if 'name_memory' not in st.session_state:
    st.session_state['name_memory'] = {}

uploaded_file = st.file_uploader("อัปโหลดไฟล์ Excel (Transport)", type=['xlsx'])

if uploaded_file:
    # อ่านข้อมูลจาก Sheet 'Transport' เป็นหลัก
    df_raw = pd.read_excel(uploaded_file, sheet_name='Transport')
    
    # สร้าง Tab เพื่อแยกการทำงาน
    tab1, tab2 = st.tabs(["✂️ แยกข้อมูลซัพพลายเออร์ (V6.7.1)", "📄 สร้างใบส่งสินค้า (GENPrint DO)"])

    # --- TAB 1: ระบบเดิม V6.7.1 ---
    with tab1:
        df_split = df_raw.dropna(subset=['ซัพพลายเออร์'])
        unique_suppliers = sorted(df_split['ซัพพลายเออร์'].unique())

        st.subheader("📝 กำหนดชื่อตัวย่อชีต")
        with st.form("sheet_name_form"):
            cols = st.columns(3)
            current_mapping = {}
            for i, supplier in enumerate(unique_suppliers):
                with cols[i % 3]:
                    remembered_name = st.session_state['name_memory'].get(supplier, str(supplier)[:10].strip())
                    current_mapping[supplier] = st.text_input(f"{supplier}:", value=remembered_name, key=f"input_{supplier}")
            submit_split = st.form_submit_button("สร้างไฟล์แยกซัพพลายเออร์")

        if submit_split:
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                workbook = writer.book
                total_format = workbook.add_format({'bold': True, 'bg_color': '#D9EAD3', 'border': 1, 'align': 'right'})
                
                for supplier, sheet_name in current_mapping.items():
                    # บันทึกชื่อจำไว้ใน Session
                    st.session_state['name_memory'][supplier] = sheet_name
                    
                    clean_name = "".join([c if c not in r'[]:*?/\ ' else ' ' for c in sheet_name.strip()[:31]])
                    df_sup = df_split[df_split['ซัพพลายเออร์'] == supplier].copy()
                    
                    # ฝั่งซ้าย (มีชื่อสาขา)
                    left_display = df_sup.rename(columns={'Store Name': 'ชื่อสาขา', 'จำนวน': 'Total'})
                    left_cols = ['รหัสสาขา', 'ชื่อสาขา', 'โซน', 'รหัสสินค้า', 'รายการสินค้า', 'ซัพพลายเออร์', 'หน่วย', 'Total']
                    left_final = left_display[left_cols].copy()
                    width_ref_left = left_final.copy()
                    
                    mask = left_final['รหัสสาขา'].duplicated()
                    left_final.loc[mask, ['รหัสสาขา', 'ชื่อสาขา', 'โซน']] = ""
                    
                    # ฝั่งขวา
                    right_summary = df_sup.groupby(['รหัสสินค้า', 'รายการสินค้า'], as_index=False)['จำนวน'].sum()
                    right_summary.rename(columns={'จำนวน': 'Total'}, inplace=True)
                    right_summary.insert(0, 'ซัพพลายเออร์', "")

                    left_final.to_excel(writer, sheet_name=clean_name, index=False, startcol=0)
                    right_summary.to_excel(writer, sheet_name=clean_name, index=False, startcol=10)

                    worksheet = writer.sheets[clean_name]
                    worksheet.set_zoom(80)
                    worksheet.write(len(left_final) + 1, 0, "Grand Total", total_format)
                    worksheet.write(len(left_final) + 1, 7, df_sup['จำนวน'].sum(), total_format)
                    worksheet.write(len(right_summary) + 1, 10, f"{supplier} Total", total_format)
                    worksheet.write(len(right_summary) + 1, 13, right_summary['Total'].sum(), total_format)

            st.success("✅ สร้างไฟล์แยกซัพพลายเออร์เรียบร้อย!")
            st.download_button(label="📥 ดาวน์โหลดไฟล์ Splitter", data=output.getvalue(), file_name="Supplier_Splitter_V7.xlsx")
# --- TAB 2: ระบบ GENPrint DO (ฉบับสมบูรณ์ - แม่นยำสูง) ---
    with tab2:
        st.subheader("📑 ออกใบส่งสินค้า (Single Sheet - Precision Layout)")
        
        # กรองเฉพาะแถวที่มีรหัสสาขาเพื่อป้องกันข้อมูลขยะ
        df_clean = df_raw.dropna(subset=['รหัสสาขา']).copy()
        
        if st.button("🚀 สร้างไฟล์ใบส่งสินค้า (ฉบับพิมพ์ต่อเนื่อง)"):
            output_do = io.BytesIO()
            # ใช้ xlsxwriter เพื่อควบคุม Page Break และ Layout ให้เป๊ะที่สุด
            with pd.ExcelWriter(output_do, engine='xlsxwriter') as writer:
                workbook = writer.book
                worksheet = workbook.add_worksheet("DO_Continuous")
                
                # --- [1. การกำหนด Styles] ---
                f_title = workbook.add_format({'bold': True, 'font_size': 18, 'align': 'center'})
                f_header_bold = workbook.add_format({'bold': True, 'font_size': 11})
                f_std = workbook.add_format({'font_size': 10})
                f_right = workbook.add_format({'font_size': 10, 'align': 'right'})
                
                f_table_head = workbook.add_format({'border': 1, 'align': 'center', 'bold': True, 'bg_color': '#F2F2F2', 'font_size': 10})
                f_border = workbook.add_format({'border': 1, 'font_size': 10})
                f_border_center = workbook.add_format({'border': 1, 'align': 'center', 'font_size': 10})
                f_wrap = workbook.add_format({'border': 1, 'text_wrap': True, 'font_size': 10})
                
                f_ship_to = workbook.add_format({'font_size': 10, 'text_wrap': True, 'valign': 'top', 'align': 'left'})
                f_delivery = workbook.add_format({'bold': True, 'font_size': 11, 'align': 'right'})
                f_date_val = workbook.add_format({'bold': True, 'font_size': 11, 'num_format': 'dd/mm/yyyy', 'align': 'right'})
                
                f_footer_box = workbook.add_format({'border': 1, 'font_size': 9, 'valign': 'top', 'text_wrap': True})

                # --- [2. ตั้งค่าหน้ากระดาษ A4 และความกว้างคอลัมน์] ---
                worksheet.set_paper(9) # A4
                worksheet.set_margins(0.3, 0.3, 0.3, 0.3)
                worksheet.set_column('A:A', 5)   # No.
                worksheet.set_column('B:B', 14)  # Product Code
                worksheet.set_column('C:C', 45)  # Product Name (แกนกลาง)
                worksheet.set_column('D:D', 12)  # Unit
                worksheet.set_column('E:E', 15)  # QTY & Right Info
                
                curr = 0 # ตัวนับแถวปัจจุบัน (Row Cursor)
                page_breaks = [] # เก็บตำแหน่งตัดหน้าพิมพ์

                for store_code in df_clean['รหัสสาขา'].unique():
                    df_store = df_clean[df_clean['รหัสสาขา'] == store_code].copy()
                    if df_store.empty: continue
                    
                    first_row = df_store.iloc[0]

                    # --- [3. ระบบตรวจสอบการขึ้นหน้าใหม่] ---
                    # ถ้าขึ้น DO ใหม่แล้วจะเกินบรรทัดที่ 40 ให้สั่งตัดหน้า (Page Break)
                    if curr > 38:
                        page_breaks.append(curr)
                        # หมายเหตุ: เราไม่ reset curr เป็น 0 เพราะต้องการให้ไหลต่อเนื่องใน Sheet เดียว
                    
                    # --- [4. เขียนส่วนหัว (Header)] ---
                    worksheet.write(curr, 0, 'บริษัท โมบาย โลจิสติกส์ จำกัด', f_header_bold)
                    worksheet.write(curr + 1, 0, '279 หมู่ที่ 9 ตำบลบางโฉลง อำเภอบางพลี', f_std)
                    worksheet.write(curr + 2, 0, 'จังหวัดสมุทรปราการ 10540', f_std)
                    worksheet.write(curr + 3, 0, 'ติดต่อ/สอบถาม : Tel : 099-157-3114', f_std)

                    # Title อยู่ที่ C2 (ของ DO นั้นๆ)
                    worksheet.write(curr + 1, 2, 'ใบส่งสินค้าชั่วคราว', f_title)

                    # ข้อมูลฝั่งขวา (คอลัมน์ E)
                    worksheet.write(curr, 4, f'Do. No. {first_row["เลขที่ DO."]}', f_right)
                    worksheet.write(curr + 1, 4, f'Ref. Po. {first_row["เลขที่ PO."]}', f_right)
                    worksheet.write(curr + 3, 4, f'Zone {first_row["โซน"]}', f_right)
                    
                    # Delivery Date (ตำแหน่ง D8:E8)
                    worksheet.write(curr + 7, 3, 'Delivery Date', f_delivery)
                    worksheet.write(curr + 7, 4, first_row['Delivery Date'], f_date_val)

                    # --- [5. ข้อมูลสาขา และ Ship To (B:C Merge)] ---
                    worksheet.write(curr + 5, 0, f'Store Code: {store_code}', f_header_bold)
                    worksheet.write(curr + 6, 0, f'Store Name: {first_row["Store Name"]}', f_header_bold)
                    worksheet.write(curr + 7, 0, 'Ship To:', f_std)
                    # Merge B และ C สำหรับที่อยู่ (ไม่มีเส้นขอบ Stroke)
                    worksheet.merge_range(curr + 7, 1, curr + 8, 2, first_row['ที่อยู่'], f_ship_to)

                    # --- [6. ส่วนตารางสินค้า (Table)] ---
                    header_row = curr + 10
                    headers = ['No.', 'Product Code', 'Product Name', 'Unit/UOM', 'QTY']
                    for col_idx, text in enumerate(headers):
                        worksheet.write(header_row, col_idx, text, f_table_head)

                    row_ptr = header_row + 1
                    for i, (_, r) in enumerate(df_store.iterrows(), 1):
                        worksheet.write(row_ptr, 0, i, f_border_center)
                        worksheet.write(row_ptr, 1, r['รหัสสินค้า'], f_border_center)
                        worksheet.write(row_ptr, 2, r['รายการสินค้า'], f_wrap)
                        worksheet.write(row_ptr, 3, r['หน่วย'], f_border_center)
                        worksheet.write(row_ptr, 4, r['จำนวน'], f_border_center)
                        row_ptr += 1

                    # บรรทัดสรุปผลรวม (Total)
                    worksheet.merge_range(row_ptr, 0, row_ptr, 3, 'Total', f_table_head)
                    worksheet.write(row_ptr, 4, df_store['จำนวน'].sum(), f_border_center)

                    # --- [7. ส่วนท้ายกระดาษ (Signatures)] ---
                    f_ptr = row_ptr + 1 # ใช้ f_ptr เป็นตำแหน่งเริ่ม Footer
                    worksheet.set_row(f_ptr + 1, 60) # ปรับความสูงช่องเซ็นชื่อให้กว้างขึ้น
                    
                    worksheet.merge_range(f_ptr, 0, f_ptr, 1, 'ผู้รับสินค้า', f_table_head)
                    worksheet.merge_range(f_ptr+1, 0, f_ptr+1, 1, 'ชื่อ (ตัวบรรจง):\nวันที่:', f_footer_box)
                    
                    worksheet.merge_range(f_ptr, 2, f_ptr, 3, 'ผู้ส่งสินค้า / ทะเบียนรถ', f_table_head)
                    worksheet.merge_range(f_ptr+1, 2, f_ptr+1, 3, 'ชื่อ (ตัวบรรจง):\nวันที่:', f_footer_box)
                    
                    worksheet.write(f_ptr, 4, 'คลังสินค้า', f_table_head)
                    worksheet.write(f_ptr+1, 4, 'ชื่อ:\nวันที่:', f_footer_box)

                    # เลื่อนตำแหน่ง Cursor ไปยัง DO ถัดไป (เว้นระยะ 4 แถว)
                    curr = f_ptr + 4

                # --- [8. สั่งตัดหน้ากระดาษพิมพ์] ---
                if page_breaks:
                    worksheet.set_h_pagebreaks(page_breaks)

            st.success("✅ สร้างไฟล์ใบส่งสินค้าแบบ Master Continuous เรียบร้อยแล้ว!")
            st.download_button(
                label="📥 ดาวน์โหลดไฟล์ DO Master",
                data=output_do.getvalue(),
                file_name=f"DO_Master_Final_{datetime.now().strftime('%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )