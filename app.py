import streamlit as st
import pandas as pd
import io
from datetime import datetime

st.set_page_config(page_title="Supplier & DO System V7.2", layout="wide")

st.title("📦 ระบบจัดการข้อมูลขนส่ง (Splitter & DO Generator)")
st.markdown("---")

if 'name_memory' not in st.session_state:
    st.session_state['name_memory'] = {}

uploaded_file = st.file_uploader("อัปโหลดไฟล์ Excel (Transport)", type=['xlsx'])

if uploaded_file:
    df_raw = pd.read_excel(uploaded_file, sheet_name='Transport')
    tab1, tab2 = st.tabs(["✂️ แยกข้อมูลซัพพลายเออร์ (V6.7.1)", "📄 สร้างใบส่งสินค้า (GENPrint DO)"])

    # --- TAB 1: ระบบแยกซัพพลายเออร์ (Format ตาม image_0b3a7c.png) ---
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
                header_fmt = workbook.add_format({'bold': True, 'bg_color': '#D9EAD3', 'border': 1, 'align': 'center'})
                cell_fmt = workbook.add_format({'border': 1})
                num_fmt = workbook.add_format({'border': 1, 'align': 'right'})
                total_fmt = workbook.add_format({'bold': True, 'bg_color': '#F3F3F3', 'border': 1, 'align': 'right', 'num_format': '#,##0'})
                
                for supplier, sheet_name in current_mapping.items():
                    st.session_state['name_memory'][supplier] = sheet_name
                    clean_name = "".join([c if c not in r'[]:*?/\ ' else ' ' for c in sheet_name.strip()[:31]])
                    df_sup = df_split[df_split['ซัพพลายเออร์'] == supplier].copy()
                    worksheet = workbook.add_worksheet(clean_name)
                    
                    left_headers = ['รหัสสาขา', 'โซน', 'รหัสสินค้า', 'รายการสินค้า', 'ซัพพลายเออร์', 'หน่วย', 'Total']
                    for col, h in enumerate(left_headers): worksheet.write(0, col, h, header_fmt)

                    curr_row = 1
                    for (branch, zone), b_group in df_sup.groupby(['รหัสสาขา', 'โซน'], sort=False):
                        for i, (_, row) in enumerate(b_group.iterrows()):
                            worksheet.write(curr_row, 0, branch if i == 0 else "", cell_fmt)
                            worksheet.write(curr_row, 1, zone if i == 0 else "", cell_fmt)
                            worksheet.write(curr_row, 2, row['รหัสสินค้า'], cell_fmt)
                            worksheet.write(curr_row, 3, row['รายการสินค้า'], cell_fmt)
                            worksheet.write(curr_row, 4, row['ซัพพลายเออร์'], cell_fmt)
                            worksheet.write(curr_row, 5, row['หน่วย'], cell_fmt)
                            worksheet.write(curr_row, 6, row['จำนวน'], num_fmt)
                            curr_row += 1
                    
                    worksheet.write(curr_row, 0, "Grand Total", total_fmt); worksheet.write(curr_row, 6, df_sup['จำนวน'].sum(), total_fmt)
                    
                    col_offset = 9
                    right_headers = ['ซัพพลายเออร์', 'รหัสสินค้า', 'รายการสินค้า', 'Total']
                    for col, h in enumerate(right_headers): worksheet.write(0, col_offset + col, h, header_fmt)
                    right_summary = df_sup.groupby(['รหัสสินค้า', 'รายการสินค้า'], as_index=False)['จำนวน'].sum()
                    for i, row in right_summary.iterrows():
                        worksheet.write(i + 1, col_offset, "", cell_fmt)
                        worksheet.write(i + 1, col_offset + 1, row['รหัสสินค้า'], cell_fmt)
                        worksheet.write(i + 1, col_offset + 2, row['รายการสินค้า'], cell_fmt)
                        worksheet.write(i + 1, col_offset + 3, row['จำนวน'], num_fmt)
                    sum_row = len(right_summary) + 1
                    worksheet.write(sum_row, col_offset, f"{supplier} Total", total_fmt)
                    worksheet.write(sum_row, col_offset + 3, right_summary['จำนวน'].sum(), total_fmt)
                    worksheet.set_column('A:G', 15); worksheet.set_column('D:D', 35); worksheet.set_column('L:L', 35)

            st.success("✅ สร้างไฟล์แยกซัพพลายเออร์เรียบร้อย!")
            st.download_button(label="📥 ดาวน์โหลดไฟล์ Splitter", data=output.getvalue(), file_name="Supplier_Splitter_V7.xlsx")

    # --- TAB 2: แก้ไขส่วนหัว (Header) ให้ครบถ้วนตามภาพ ---
    with tab2:
        st.subheader("📑 ออกใบส่งสินค้า (ปรับปรุงส่วนหัวและ Layout)")
        df_clean = df_raw.dropna(subset=['รหัสสาขา']).copy()
        
        if st.button("🚀 สร้างไฟล์ใบส่งสินค้า (ฉบับพิมพ์ต่อเนื่อง)"):
            output_do = io.BytesIO()
            with pd.ExcelWriter(output_do, engine='xlsxwriter') as writer:
                workbook = writer.book
                worksheet = workbook.add_worksheet("DO_Continuous")
                
                # Styles
                f_title = workbook.add_format({'bold': True, 'font_size': 20, 'align': 'center'})
                f_header_company = workbook.add_format({'bold': True, 'font_size': 12})
                f_std = workbook.add_format({'font_size': 10})
                f_right = workbook.add_format({'font_size': 10, 'align': 'right'})
                f_table_head = workbook.add_format({'border': 1, 'align': 'center', 'bold': True, 'bg_color': '#F2F2F2'})
                f_border_center = workbook.add_format({'border': 1, 'align': 'center'})
                f_wrap = workbook.add_format({'border': 1, 'text_wrap': True})
                f_footer_box = workbook.add_format({'border': 1, 'font_size': 9, 'valign': 'top', 'text_wrap': True})

                worksheet.set_paper(9); worksheet.set_margins(0.3, 0.3, 0.3, 0.3)
                worksheet.set_column('A:A', 5); worksheet.set_column('B:B', 15); worksheet.set_column('C:C', 40)
                worksheet.set_column('D:D', 15); worksheet.set_column('E:E', 12)
                
                curr = 0; page_breaks = []; do_count = 0 
                for store_code in df_clean['รหัสสาขา'].unique():
                    df_store = df_clean[df_clean['รหัสสาขา'] == store_code].copy()
                    if do_count > 0: page_breaks.append(curr)
                    first_row = df_store.iloc[0]
                    
                    # --- [แก้ไขส่วนหัว: ฝั่งซ้าย] ---
                    worksheet.write(curr, 0, 'บริษัท โมบาย โลจิสติกส์ จำกัด', f_header_company)
                    worksheet.write(curr + 1, 0, '279 หมู่ที่ 9 ตำบลบางโฉลง อำเภอบางพลี', f_std)
                    worksheet.write(curr + 2, 0, 'จังหวัดสมุทรปราการ 10540', f_std)
                    # เพิ่ม ID Line Official ตามภาพ
                    worksheet.write(curr + 3, 0, f'ติดต่อ/สอบถาม : ID Line Official : @505phsps (มี @ ), Tel : 099-157-3114', f_std)
                    worksheet.write(curr + 4, 0, 'Customer Name', f_std) # ช่องชื่อลูกค้า
                    worksheet.write(curr + 5, 0, f'Store Code: {store_code}', f_header_company)
                    worksheet.write(curr + 6, 0, f'Store Name: {first_row["Store Name"]}', f_header_company)
                    worksheet.write(curr + 7, 0, f'Ship To: {first_row["ที่อยู่"]}', f_std)

                    # --- [แก้ไขส่วนหัว: ฝั่งขวา] ---
                    worksheet.write(curr + 2, 2, 'ใบส่งสินค้าชั่วคราว', f_title) # หัวข้อกลาง
                    worksheet.write(curr, 4, f'Do. No. {first_row["เลขที่ DO."]}', f_right)
                    worksheet.write(curr + 1, 4, f'Ref. Po. {first_row["เลขที่ PO."]}', f_right)
                    worksheet.write(curr + 2, 4, 'Ref. Po. -', f_right)
                    worksheet.write(curr + 3, 4, 'Ref. Po. -', f_right)
                    worksheet.write(curr + 4, 4, f'Zone {first_row["โซน"]}', f_right)
                    worksheet.write(curr + 5, 4, f'Cutoff Date {first_row.get("Cutoff", "8/5/2026")}', f_right)
                    worksheet.write(curr + 7, 4, f'Delivery Date: {first_row["Delivery Date"]}', f_right)

                    # ตารางสินค้า
                    header_row = curr + 9
                    headers = ['No.', 'Product Code', 'Product Name', 'Unit/UOM', 'QTY']
                    for col_idx, text in enumerate(headers): worksheet.write(header_row, col_idx, text, f_table_head)

                    row_ptr = header_row + 1
                    for i, (_, r) in enumerate(df_store.iterrows(), 1):
                        worksheet.write(row_ptr, 0, i, f_border_center); worksheet.write(row_ptr, 1, r['รหัสสินค้า'], f_border_center)
                        worksheet.write(row_ptr, 2, r['รายการสินค้า'], f_wrap); worksheet.write(row_ptr, 3, r['หน่วย'], f_border_center)
                        worksheet.write(row_ptr, 4, r['จำนวน'], f_border_center)
                        row_ptr += 1

                    worksheet.merge_range(row_ptr, 0, row_ptr, 3, 'Total', f_table_head)
                    worksheet.write(row_ptr, 4, df_store['จำนวน'].sum(), f_border_center)

                    # ลายเซ็น
                    f_ptr = row_ptr + 1 
                    worksheet.set_row(f_ptr + 1, 50) 
                    titles = ['ผู้รับสินค้า', 'ผู้ส่งสินค้า / ทะเบียนรถ', 'คลังสินค้า']
                    for i, title in enumerate(titles):
                        col = i * 2 if i < 2 else 4
                        worksheet.write(f_ptr, col, title, f_table_head)
                        worksheet.write(f_ptr+1, col, 'ชื่อ:\nวันที่:', f_footer_box)

                    curr = f_ptr + 4; do_count += 1

                if page_breaks: worksheet.set_h_pagebreaks(page_breaks)

            st.success("✅ สร้างไฟล์เรียบร้อยแล้ว!")
            st.download_button(label="📥 ดาวน์โหลดไฟล์ DO Master", data=output_do.getvalue(), file_name=f"DO_Fixed_{datetime.now().strftime('%H%M')}.xlsx")