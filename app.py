import streamlit as st
import pandas as pd
import io
from datetime import datetime

st.set_page_config(page_title="DO & Supplier System V8.6", layout="wide")

st.title("📦 ระบบจัดการขนส่งครบวงจร (V8.6 Final Layout)")
st.markdown("---")

if 'name_memory' not in st.session_state:
    st.session_state['name_memory'] = {}

uploaded_file = st.file_uploader("อัปโหลดไฟล์ Excel (Transport)", type=['xlsx'])

if uploaded_file:
    # อ่านข้อมูล
    df_raw = pd.read_excel(uploaded_file, sheet_name='Transport')
    
    # สร้าง Tab กลับมาให้ครบ
    tab1, tab2 = st.tabs(["✂️ 1. แยกซัพพลายเออร์", "📄 2. ออกใบส่งสินค้า (A4 Fix)"])

    # --- TAB 1: ระบบแยกซัพพลายเออร์ (กู้คืนกลับมา) ---
    with tab1:
        df_split = df_raw.dropna(subset=['ซัพพลายเออร์'])
        unique_suppliers = sorted(df_split['ซัพพลายเออร์'].unique())
        st.subheader("📝 กำหนดชื่อชีตย่อ")
        
        with st.form("sheet_name_form"):
            cols = st.columns(3)
            current_mapping = {}
            for i, supplier in enumerate(unique_suppliers):
                with cols[i % 3]:
                    remembered_name = st.session_state['name_memory'].get(supplier, str(supplier)[:10].strip())
                    current_mapping[supplier] = st.text_input(f"{supplier}:", value=remembered_name, key=f"input_{supplier}")
            submit_split = st.form_submit_button("สร้างไฟล์แยกซัพพลายเออร์")

        if submit_split:
            output_split = io.BytesIO()
            with pd.ExcelWriter(output_split, engine='xlsxwriter') as writer:
                workbook = writer.book
                header_fmt = workbook.add_format({'bold': True, 'bg_color': '#D9EAD3', 'border': 1, 'align': 'center'})
                cell_fmt = workbook.add_format({'border': 1})
                num_fmt = workbook.add_format({'border': 1, 'align': 'right'})
                
                for supplier, sheet_name in current_mapping.items():
                    st.session_state['name_memory'][supplier] = sheet_name
                    clean_name = "".join([c if c not in r'[]:*?/\ ' else ' ' for c in sheet_name.strip()[:31]])
                    df_sup = df_split[df_split['ซัพพลายเออร์'] == supplier].copy()
                    worksheet = workbook.add_worksheet(clean_name)
                    # (การจัดวางในชีตซัพพลายเออร์เหมือนเดิมตามมาตรฐานที่เคยใช้)
                    headers = ['รหัสสาขา', 'โซน', 'รหัสสินค้า', 'รายการสินค้า', 'ซัพพลายเออร์', 'หน่วย', 'จำนวน']
                    for col, h in enumerate(headers): worksheet.write(0, col, h, header_fmt)
                    for r_idx, row in enumerate(df_sup.values, 1):
                        for c_idx, val in enumerate(row): worksheet.write(r_idx, c_idx, val, cell_fmt if c_idx != 6 else num_fmt)
                    worksheet.set_column('A:G', 15); worksheet.set_column('D:D', 35)

            st.success("✅ สร้างไฟล์แยกซัพพลายเออร์สำเร็จ!")
            st.download_button(label="📥 ดาวน์โหลดไฟล์ Splitter", data=output_split.getvalue(), file_name="Supplier_Splitter.xlsx")

    # --- TAB 2: ออกใบ DO (Layout ใหม่ตามสั่ง) ---
    with tab2:
        st.subheader("📑 ออกใบส่งสินค้า (โครงสร้าง 5 คอลัมน์ ตามแจ้ง)")
        df_clean = df_raw.dropna(subset=['รหัสสาขา']).copy()
        
        if st.button("🚀 สร้างไฟล์ใบส่งสินค้า V8.6 (A4 Perfect)"):
            output_do = io.BytesIO()
            with pd.ExcelWriter(output_do, engine='xlsxwriter') as writer:
                workbook = writer.book
                worksheet = workbook.add_worksheet("DO_Master")
                
                # Setup A4
                worksheet.set_paper(9) 
                worksheet.set_margins(0.3, 0.3, 0.3, 0.3)
                worksheet.fit_to_pages(1, 0)
                
                # ความกว้างคอลัมน์ 5 คอลัมน์ (A-E)
                worksheet.set_column('A:A', 6)   # No.
                worksheet.set_column('B:B', 15)  # Code
                worksheet.set_column('C:C', 38)  # Name / Title
                worksheet.set_column('D:D', 12)  # Unit / Do. No.
                worksheet.set_column('E:E', 12)  # QTY / Ref Po.
                
                # Styles
                f_comp = workbook.add_format({'bold': True, 'font_size': 14, 'align': 'left'})
                f_title = workbook.add_format({'bold': True, 'font_size': 20, 'align': 'center', 'valign': 'vcenter'})
                f_std = workbook.add_format({'font_size': 10})
                f_bold = workbook.add_format({'bold': True, 'font_size': 10})
                f_right = workbook.add_format({'font_size': 10, 'align': 'right'})
                f_table_h = workbook.add_format({'border': 1, 'align': 'center', 'bold': True, 'bg_color': '#F2F2F2'})
                f_border = workbook.add_format({'border': 1, 'align': 'center'})
                f_wrap = workbook.add_format({'border': 1, 'text_wrap': True})
                f_footer = workbook.add_format({'border': 1, 'font_size': 9, 'valign': 'top'})

                curr = 0; do_count = 0; page_breaks = []
                for store_code in df_clean['รหัสสาขา'].unique():
                    df_store = df_clean[df_clean['รหัสสาขา'] == store_code].copy()
                    if do_count > 0: page_breaks.append(curr)
                    first = df_store.iloc[0]

                    # --- ROW 1: Merge 5 คอลัมน์ ใส่ชื่อบริษัท ---
                    worksheet.merge_range(curr, 0, curr, 4, 'บริษัท โมบาย โลจิสติกส์ จำกัด', f_comp)
                    
                    # --- ROW 2: เพิ่มความสูง และจัดวาง 3 ส่วน ---
                    worksheet.set_row(curr + 1, 30) # เพิ่มความสูง Row 2
                    worksheet.merge_range(curr+1, 0, curr+1, 1, '279 หมู่ที่ 9 ตำบลบางโฉลง อำเภอบางพลี', f_std)
                    worksheet.write(curr+1, 2, 'ใบส่งสินค้าชั่วคราว', f_title)
                    worksheet.write(curr+1, 3, 'Do. No.', f_right)
                    worksheet.write(curr+1, 4, first["เลขที่ DO."], f_std)
                    
                    # --- ROW 3 เป็นต้นไป (ข้อมูลทั่วไป) ---
                    worksheet.merge_range(curr+2, 0, curr+2, 1, 'จังหวัดสมุทรปราการ 10540', f_std)
                    worksheet.write(curr+2, 3, 'Ref. Po.', f_right)
                    worksheet.write(curr+2, 4, first["เลขที่ PO."], f_std)
                    
                    worksheet.merge_range(curr+3, 0, curr+3, 2, 'ติดต่อ/สอบถาม : ID Line Official : @505phsps (มี @ ), Tel : 099-157-3114', f_std)
                    
                    # ข้อมูลวันที่ (ตัดเวลา)
                    cutoff = str(first.get('Cut Off Date', '-')).split(' ')[0]
                    delivery = str(first["Delivery Date"]).split(' ')[0]
                    worksheet.write(curr+4, 3, 'Cut off Date:', f_right)
                    worksheet.write(curr+4, 4, cutoff, f_std)
                    
                    worksheet.write(curr+5, 0, f'Store Code: {store_code}', f_bold)
                    worksheet.write(curr+5, 3, 'Zone:', f_right)
                    worksheet.write(curr+5, 4, first["โซน"], f_std)
                    
                    worksheet.merge_range(curr+6, 0, curr+6, 2, f'Store Name: {first["Store Name"]}', f_bold)
                    worksheet.write(curr+7, 3, 'Delivery Date:', f_right)
                    worksheet.write(curr+7, 4, delivery, f_std)
                    
                    worksheet.merge_range(curr+7, 0, curr+7, 2, f'Ship To: {first["ที่อยู่"]}', f_std)

                    # --- Table Section ---
                    t_h = curr + 9
                    table_cols = ['No.', 'Product Code', 'Product Name', 'Unit/UOM', 'QTY']
                    for i, col_name in enumerate(table_cols):
                        worksheet.write(t_h, i, col_name, f_table_h)

                    r_ptr = t_h + 1
                    for i, (_, r) in enumerate(df_store.iterrows(), 1):
                        worksheet.write(r_ptr, 0, i, f_border)
                        worksheet.write(r_ptr, 1, r['รหัสสินค้า'], f_border)
                        worksheet.write(r_ptr, 2, r['รายการสินค้า'], f_wrap)
                        worksheet.write(r_ptr, 3, r['หน่วย'], f_border)
                        worksheet.write(r_ptr, 4, r['จำนวน'], f_border)
                        r_ptr += 1
                    
                    # Total Row
                    worksheet.merge_range(r_ptr, 0, r_ptr, 3, 'Total', f_table_h)
                    worksheet.write(r_ptr, 4, df_store['จำนวน'].sum(), f_border)

                    # --- Footer Section (สมมาตร) ---
                    f_row = r_ptr + 1
                    worksheet.merge_range(f_row, 0, f_row, 1, 'ผู้รับสินค้า', f_table_h)
                    worksheet.merge_range(f_row, 2, f_row, 3, 'ผู้ส่งสินค้า / ทะเบียนรถ', f_table_h)
                    worksheet.write(f_row, 4, 'คลังสินค้า', f_table_h)
                    
                    labels = ['ชื่อ (ตัวบรรจง):', 'วันที่:', 'เวลา:', 'หมายเหตุ:']
                    for lbl in labels:
                        f_row += 1
                        worksheet.merge_range(f_row, 0, f_row, 1, lbl, f_footer)
                        worksheet.merge_range(f_row, 2, f_row, 3, lbl, f_footer)
                        worksheet.write(f_row, 4, lbl, f_footer)

                    curr = f_row + 2
                    do_count += 1

                if page_breaks: worksheet.set_h_pagebreaks(page_breaks)

            st.success("✅ แก้ไข Layout และดึง Tab 1 กลับมาเรียบร้อยครับ!")
            st.download_button(label="📥 ดาวน์โหลดไฟล์ Master V8.6", data=output_do.getvalue(), file_name="DO_Master_V8.6.xlsx")