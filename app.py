import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="DO System V8.5 Master", layout="wide")

st.title("📦 ระบบออกใบส่งสินค้า (V8.5 A4 Complete Fix)")
st.markdown("---")

uploaded_file = st.file_uploader("อัปโหลดไฟล์ Excel (Transport)", type=['xlsx'])

if uploaded_file:
    df_raw = pd.read_excel(uploaded_file, sheet_name='Transport')
    df_clean = df_raw.dropna(subset=['รหัสสาขา']).copy()
    
    if st.button("🚀 สร้างไฟล์ใบส่งสินค้า A4 (V8.5)"):
        output_do = io.BytesIO()
        with pd.ExcelWriter(output_do, engine='xlsxwriter') as writer:
            workbook = writer.book
            worksheet = workbook.add_worksheet("DO_Master")
            
            # --- Page Setup for A4 ---
            worksheet.set_paper(9) # A4
            worksheet.set_margins(0.3, 0.3, 0.3, 0.3)
            worksheet.fit_to_pages(1, 0) # บังคับความกว้างให้พอดี 1 หน้า A4
            
            # --- Styles ---
            f_title = workbook.add_format({'bold': True, 'font_size': 18, 'align': 'center', 'valign': 'vcenter'})
            f_comp = workbook.add_format({'bold': True, 'font_size': 11})
            f_std = workbook.add_format({'font_size': 10})
            f_bold = workbook.add_format({'bold': True, 'font_size': 10})
            f_right = workbook.add_format({'font_size': 10, 'align': 'right'})
            f_table_h = workbook.add_format({'border': 1, 'align': 'center', 'bold': True, 'bg_color': '#F2F2F2'})
            f_border = workbook.add_format({'border': 1, 'align': 'center'})
            f_wrap = workbook.add_format({'border': 1, 'text_wrap': True})
            f_footer = workbook.add_format({'border': 1, 'font_size': 9, 'valign': 'top'})

            # Column Widths (จัดสัดส่วนให้เต็มหน้า A4)
            worksheet.set_column('A:A', 5)   # No.
            worksheet.set_column('B:B', 12)  # Product Code
            worksheet.set_column('C:C', 35)  # Product Name
            worksheet.set_column('D:D', 15)  # Unit/UOM
            worksheet.set_column('E:E', 15)  # QTY & Right Info

            curr = 0; do_count = 0; page_breaks = []
            
            for store_code in df_clean['รหัสสาขา'].unique():
                df_store = df_clean[df_clean['รหัสสาขา'] == store_code].copy()
                if do_count > 0: page_breaks.append(curr)
                first = df_store.iloc[0]
                
                # --- Header Left (Fixing Address truncation) ---
                worksheet.write(curr, 0, 'บริษัท โมบาย โลจิสติกส์ จำกัด', f_comp)
                worksheet.merge_range(curr+1, 0, curr+1, 2, '279 หมู่ที่ 9 ตำบลบางโฉลง อำเภอบางพลี', f_std)
                worksheet.merge_range(curr+2, 0, curr+2, 2, 'จังหวัดสมุทรปราการ 10540', f_std)
                worksheet.merge_range(curr+3, 0, curr+3, 3, 'ติดต่อ/สอบถาม : ID Line Official : @505phsps (มี @ ), Tel : 099-157-3114', f_std)
                worksheet.write(curr+4, 0, 'Customer Name', f_std)
                
                # Store & Ship To
                worksheet.merge_range(curr+5, 0, curr+5, 3, f'Store Code: {store_code}', f_bold)
                worksheet.merge_range(curr+6, 0, curr+6, 3, f'Store Name: {first["Store Name"]}', f_bold)
                worksheet.merge_range(curr+7, 0, curr+7, 3, f'Ship To: {first["ที่อยู่"]}', f_std)

                # --- Header Center & Right ---
                worksheet.merge_range(curr+1, 3, curr+2, 4, 'ใบส่งสินค้าชั่วคราว', f_title)
                
                # Right Info Fix
                worksheet.write(curr, 4, f'Do. No. {first["เลขที่ DO."]}', f_right)
                worksheet.write(curr+1, 4, f'Ref. Po. {first["เลขที่ PO."]}', f_right)
                worksheet.write(curr+2, 4, 'Ref. Po. -', f_right)
                worksheet.write(curr+3, 4, 'Ref. Po. -', f_right)
                
                # Cut Off Date (Date only)
                cutoff = str(first.get('Cut Off Date', '-')).split(' ')[0]
                worksheet.write(curr+4, 4, f'Cut Off Date: {cutoff}', f_right)
                worksheet.write(curr+5, 4, f'Zone {first["โซน"]}', f_right)
                
                # Delivery Date (Date only)
                delivery = str(first["Delivery Date"]).split(' ')[0]
                worksheet.write(curr+7, 4, f'Delivery Date: {delivery}', f_right)

                # --- Table Section ---
                t_h = curr + 10
                cols = ['No.', 'Product Code', 'Product Name', 'Unit/UOM', 'QTY']
                for c_idx, val in enumerate(cols):
                    worksheet.write(t_h, c_idx, val, f_table_h)

                r_ptr = t_h + 1
                for i, (_, r) in enumerate(df_store.iterrows(), 1):
                    worksheet.write(r_ptr, 0, i, f_border)
                    worksheet.write(r_ptr, 1, r['รหัสสินค้า'], f_border)
                    worksheet.write(r_ptr, 2, r['รายการสินค้า'], f_wrap)
                    worksheet.write(r_ptr, 3, r['หน่วย'], f_border)
                    worksheet.write(r_ptr, 4, r['จำนวน'], f_border)
                    r_ptr += 1
                
                worksheet.merge_range(r_ptr, 0, r_ptr, 3, 'Total', f_table_h)
                worksheet.write(r_ptr, 4, df_store['จำนวน'].sum(), f_border)

                # --- Footer Section (Symmetrical Boxes) ---
                f_row = r_ptr + 1
                # กราฟิก Footer 3 กล่องเท่ากัน (A-B, C-D, E)
                worksheet.merge_range(f_row, 0, f_row, 1, 'ผู้รับสินค้า', f_table_h)
                worksheet.merge_range(f_row, 2, f_row, 3, 'ผู้ส่งสินค้า / ทะเบียนรถ', f_table_h)
                worksheet.write(f_row, 4, 'คลังสินค้า', f_table_h)
                
                f_row += 1
                labels = ['ชื่อ (ตัวบรรจง):', 'วันที่:', 'เวลา:', 'หมายเหตุ:']
                for label in labels:
                    worksheet.merge_range(f_row, 0, f_row, 1, label, f_footer)
                    worksheet.merge_range(f_row, 2, f_row, 3, label, f_footer)
                    worksheet.write(f_row, 4, label, f_footer)
                    f_row += 1

                curr = f_row + 2 # ระยะห่างก่อนใบถัดไป
                do_count += 1

            if page_breaks: worksheet.set_h_pagebreaks(page_breaks)

        st.success("✅ แก้ไขข้อมูลและจัด Layout A4 เรียบร้อยแล้ว!")
        st.download_button(label="📥 ดาวน์โหลดไฟล์ DO Master V8.5", data=output_do.getvalue(), file_name="DO_A4_Fixed_V8.5.xlsx")