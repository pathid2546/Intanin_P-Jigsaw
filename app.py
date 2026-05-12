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
    # อ่านข้อมูลจาก Sheet 'Transport'
    df_raw = pd.read_excel(uploaded_file, sheet_name='Transport')
    
    # สร้าง Tab เพื่อแยกการทำงาน
    tab1, tab2 = st.tabs(["✂️ แยกข้อมูลซัพพลายเออร์ (V6.7.1)", "📄 สร้างใบส่งสินค้า (GENPrint DO)"])

    # --- TAB 1: ปรับปรุง Format ให้เหมือนต้นฉบับ image_0b3a7c.png ---
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
                
                # --- [เพิ่ม Styles สำหรับ Tab 1] ---
                header_fmt = workbook.add_format({'bold': True, 'bg_color': '#D9EAD3', 'border': 1, 'align': 'center'})
                cell_fmt = workbook.add_format({'border': 1})
                num_fmt = workbook.add_format({'border': 1, 'align': 'right'})
                total_fmt = workbook.add_format({'bold': True, 'bg_color': '#F3F3F3', 'border': 1, 'align': 'right', 'num_format': '#,##0'})
                
                for supplier, sheet_name in current_mapping.items():
                    st.session_state['name_memory'][supplier] = sheet_name
                    clean_name = "".join([c if c not in r'[]:*?/\ ' else ' ' for c in sheet_name.strip()[:31]])
                    df_sup = df_split[df_split['ซัพพลายเออร์'] == supplier].copy()
                    
                    worksheet = workbook.add_worksheet(clean_name)
                    worksheet.set_zoom(85)

                    # --- ฝั่งซ้าย: รายการแยกสาขา ---
                    left_headers = ['รหัสสาขา', 'ชื่อสาขา', 'โซน', 'รหัสสินค้า', 'รายการสินค้า', 'ซัพพลายเออร์', 'หน่วย', 'Total']
                    for col, h in enumerate(left_headers):
                        worksheet.write(0, col, h, header_fmt)

                    curr_row = 1
                    # จัดกลุ่มข้อมูลเพื่อให้รหัสสาขาแสดงเฉพาะบรรทัดแรก (ตามรูป image_0b3a7c.png)
                    for (branch, s_name, zone), b_group in df_sup.groupby(['รหัสสาขา', 'Store Name', 'โซน'], sort=False):
                        for i, (_, row) in enumerate(b_group.iterrows()):
                            if i == 0:
                                worksheet.write(curr_row, 0, branch, cell_fmt)
                                worksheet.write(curr_row, 1, s_name, cell_fmt)
                                worksheet.write(curr_row, 2, zone, cell_fmt)
                            else:
                                worksheet.write(curr_row, 0, "", cell_fmt)
                                worksheet.write(curr_row, 1, "", cell_fmt)
                                worksheet.write(curr_row, 2, "", cell_fmt)
                            
                            worksheet.write(curr_row, 3, row['รหัสสินค้า'], cell_fmt)
                            worksheet.write(curr_row, 4, row['รายการสินค้า'], cell_fmt)
                            worksheet.write(curr_row, 5, row['ซัพพลายเออร์'], cell_fmt)
                            worksheet.write(curr_row, 6, row['หน่วย'], cell_fmt)
                            worksheet.write(curr_row, 7, row['จำนวน'], num_fmt)
                            curr_row += 1
                    
                    # ยอดรวมฝั่งซ้าย
                    worksheet.write(curr_row, 0, "Grand Total", total_fmt)
                    worksheet.write(curr_row, 7, df_sup['จำนวน'].sum(), total_fmt)

                    # --- ฝั่งขวา: สรุปรายการสินค้า (Summary) ---
                    col_offset = 10 # เริ่มที่คอลัมน์ K
                    right_headers = ['ซัพพลายเออร์', 'รหัสสินค้า', 'รายการสินค้า', 'Total']
                    for col, h in enumerate(right_headers):
                        worksheet.write(0, col_offset + col, h, header_fmt)

                    right_summary = df_sup.groupby(['รหัสสินค้า', 'รายการสินค้า'], as_index=False)['จำนวน'].sum()
                    for i, row in right_summary.iterrows():
                        worksheet.write(i + 1, col_offset, "", cell_fmt) # คอลัมน์ซัพพลายเออร์ว่างไว้ตามต้นฉบับ
                        worksheet.write(i + 1, col_offset + 1, row['รหัสสินค้า'], cell_fmt)
                        worksheet.write(i + 1, col_offset + 2, row['รายการสินค้า'], cell_fmt)
                        worksheet.write(i + 1, col_offset + 3, row['จำนวน'], num_fmt)
                    
                    # ยอดรวมฝั่งขวา
                    sum_row = len(right_summary) + 1
                    worksheet.write(sum_row, col_offset, f"{supplier} Total", total_fmt)
                    worksheet.write(sum_row, col_offset + 3, right_summary['จำนวน'].sum(), total_fmt)

                    # ปรับความกว้างคอลัมน์
                    worksheet.set_column('A:B', 12); worksheet.set_column('C:C', 8)
                    worksheet.set_column('D:D', 12); worksheet.set_column('E:F', 30)
                    worksheet.set_column('K:K', 15); worksheet.set_column('L:L', 12)
                    worksheet.set_column('M:M', 40)

            st.success("✅ สร้างไฟล์แยกซัพพลายเออร์เรียบร้อย!")
            st.download_button(label="📥 ดาวน์โหลดไฟล์ Splitter", data=output.getvalue(), file_name="Supplier_Splitter_V7.xlsx")

    # --- TAB 2: ระบบ GENPrint DO (คงเดิมตามความต้องการ) ---
    with tab2:
        st.subheader("📑 ออกใบส่งสินค้า (Single Sheet - Precision Layout)")
        df_clean = df_raw.dropna(subset=['รหัสสาขา']).copy()
        if st.button("🚀 สร้างไฟล์ใบส่งสินค้า (ฉบับพิมพ์ต่อเนื่อง)"):
            output_do = io.BytesIO()
            with pd.ExcelWriter(output_do, engine='xlsxwriter') as writer:
                workbook = writer.book
                worksheet = workbook.add_worksheet("DO_Continuous")
                
                f_title = workbook.add_format({'bold': True, 'font_size': 18, 'align': 'center'})
                f_header_bold = workbook.add_format({'bold': True, 'font_size': 11})
                f_std = workbook.add_format({'font_size': 10})
                f_right = workbook.add_format({'font_size': 10, 'align': 'right'})
                f_table_head = workbook.add_format({'border': 1, 'align': 'center', 'bold': True, 'bg_color': '#F2F2F2', 'font_size': 10})
                f_border_center = workbook.add_format({'border': 1, 'align': 'center', 'font_size': 10})
                f_wrap = workbook.add_format({'border': 1, 'text_wrap': True, 'font_size': 10})
                f_ship_to = workbook.add_format({'font_size': 10, 'text_wrap': True, 'valign': 'top', 'align': 'left'})
                f_delivery = workbook.add_format({'bold': True, 'font_size': 11, 'align': 'right'})
                f_date_val = workbook.add_format({'bold': True, 'font_size': 11, 'num_format': 'dd/mm/yyyy', 'align': 'right'})
                f_footer_box = workbook.add_format({'border': 1, 'font_size': 9, 'valign': 'top', 'text_wrap': True})

                worksheet.set_paper(9) 
                worksheet.set_margins(0.3, 0.3, 0.3, 0.3)
                worksheet.set_column('A:A', 5); worksheet.set_column('B:B', 14); worksheet.set_column('C:C', 45)
                worksheet.set_column('D:D', 12); worksheet.set_column('E:E', 15)
                
                curr = 0; page_breaks = []; do_count = 0 
                for store_code in df_clean['รหัสสาขา'].unique():
                    df_store = df_clean[df_clean['รหัสสาขา'] == store_code].copy()
                    if do_count > 0: page_breaks.append(curr)
                    
                    first_row = df_store.iloc[0]
                    worksheet.write(curr, 0, 'บริษัท โมบาย โลจิสติกส์ จำกัด', f_header_bold)
                    worksheet.write(curr + 1, 0, '279 หมู่ที่ 9 ตำบลบางโฉลง อำเภอบางพลี', f_std)
                    worksheet.write(curr + 2, 0, 'จังหวัดสมุทรปราการ 10540', f_std)
                    worksheet.write(curr + 3, 0, 'ติดต่อ/สอบถาม : Tel : 099-157-3114', f_std)
                    worksheet.write(curr + 1, 2, 'ใบส่งสินค้าชั่วคราว', f_title)
                    worksheet.write(curr, 4, f'Do. No. {first_row["เลขที่ DO."]}', f_right)
                    worksheet.write(curr + 1, 4, f'Ref. Po. {first_row["เลขที่ PO."]}', f_right)
                    worksheet.write(curr + 3, 4, f'Zone {first_row["โซน"]}', f_right)
                    worksheet.write(curr + 7, 3, 'Delivery Date', f_delivery)
                    worksheet.write(curr + 7, 4, first_row['Delivery Date'], f_date_val)
                    worksheet.write(curr + 5, 0, f'Store Code: {store_code}', f_header_bold)
                    worksheet.write(curr + 6, 0, f'Store Name: {first_row["Store Name"]}', f_header_bold)
                    worksheet.write(curr + 7, 0, 'Ship To:', f_std)
                    worksheet.merge_range(curr + 7, 1, curr + 8, 2, first_row['ที่อยู่'], f_ship_to)

                    header_row = curr + 10
                    headers = ['No.', 'Product Code', 'Product Name', 'Unit/UOM', 'QTY']
                    for col_idx, text in enumerate(headers):
                        worksheet.write(header_row, col_idx, text, f_table_head)

                    row_ptr = header_row + 1
                    for i, (_, r) in enumerate(df_store.iterrows(), 1):
                        worksheet.write(row_ptr, 0, i, f_border_center); worksheet.write(row_ptr, 1, r['รหัสสินค้า'], f_border_center)
                        worksheet.write(row_ptr, 2, r['รายการสินค้า'], f_wrap); worksheet.write(row_ptr, 3, r['หน่วย'], f_border_center)
                        worksheet.write(row_ptr, 4, r['จำนวน'], f_border_center)
                        row_ptr += 1

                    worksheet.merge_range(row_ptr, 0, row_ptr, 3, 'Total', f_table_head)
                    worksheet.write(row_ptr, 4, df_store['จำนวน'].sum(), f_border_center)

                    f_ptr = row_ptr + 1 
                    worksheet.set_row(f_ptr + 1, 60) 
                    worksheet.merge_range(f_ptr, 0, f_ptr, 1, 'ผู้รับสินค้า', f_table_head)
                    worksheet.merge_range(f_ptr+1, 0, f_ptr+1, 1, 'ชื่อ (ตัวบรรจง):\nวันที่:', f_footer_box)
                    worksheet.merge_range(f_ptr, 2, f_ptr, 3, 'ผู้ส่งสินค้า / ทะเบียนรถ', f_table_head)
                    worksheet.merge_range(f_ptr+1, 2, f_ptr+1, 3, 'ชื่อ (ตัวบรรจง):\nวันที่:', f_footer_box)
                    worksheet.write(f_ptr, 4, 'คลังสินค้า', f_table_head); worksheet.write(f_ptr+1, 4, 'ชื่อ:\nวันที่:', f_footer_box)

                    curr = f_ptr + 4; do_count += 1

                if page_breaks:
                    worksheet.set_h_pagebreaks(page_breaks)

            st.success("✅ แก้ไขระบบตัดหน้าเรียบร้อยแล้ว!")
            st.download_button(label="📥 ดาวน์โหลดไฟล์ DO Master", data=output_do.getvalue(), file_name=f"DO_Master_Fixed_{datetime.now().strftime('%H%M')}.xlsx")