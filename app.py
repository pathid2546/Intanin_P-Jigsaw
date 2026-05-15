import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="DO System V11.0 Clean", layout="wide")

st.title("📦 ระบบจัดการขนส่ง (V11.0 ตัดจุดและไฮไลท์สี)")
st.markdown("---")

uploaded_file = st.file_uploader("อัปโหลดไฟล์ Excel (Transport)", type=['xlsx'])

if uploaded_file:
    df_raw = pd.read_excel(uploaded_file, sheet_name='Transport')
    
    df_clean = df_raw.dropna(subset=['รหัสสาขา']).copy()
    if st.button("🚀 สร้างใบส่งสินค้า V11.0 (Clean Version)"):
        output_do = io.BytesIO()
        with pd.ExcelWriter(output_do, engine='xlsxwriter') as writer:
            workbook = writer.book
            worksheet = workbook.add_worksheet("DO_Master")
            
            # Setup A4 & Column Widths
            worksheet.set_paper(9) 
            worksheet.set_margins(0.3, 0.3, 0.3, 0.3)
            worksheet.set_column('A:A', 8)    
            worksheet.set_column('B:B', 35)   
            worksheet.set_column('C:C', 35)   
            worksheet.set_column('D:D', 18)   
            worksheet.set_column('E:E', 22)   
            
            # Styles
            f_comp = workbook.add_format({'bold': True, 'font_size': 14})
            f_title = workbook.add_format({'bold': True, 'font_size': 18, 'align': 'center', 'valign': 'vcenter'})
            f_std = workbook.add_format({'font_size': 11})
            f_bold = workbook.add_format({'bold': True, 'font_size': 11})
            f_right = workbook.add_format({'font_size': 11, 'align': 'right'})
            
            # Delivery Date Styles (No Background Color)
            f_deliv_label = workbook.add_format({
                'bold': True, 'font_size': 14, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True
            })
            f_deliv_date = workbook.add_format({
                'bold': True, 'font_size': 14, 'align': 'center', 'valign': 'vcenter' # Removed bg_color
            })

            f_border = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter'})
            f_table_h = workbook.add_format({'border': 1, 'align': 'center', 'bold': True, 'bg_color': '#F2F2F2'})
            f_wrap = workbook.add_format({'border': 1, 'text_wrap': True, 'valign': 'vcenter'})
            f_footer_h = workbook.add_format({'border': 1, 'align': 'center', 'bold': True})
            f_footer_box = workbook.add_format({'border': 1, 'font_size': 10, 'valign': 'top'})

            curr = 0; page_breaks = []
            for store_code in df_clean['รหัสสาขา'].unique():
                df_store = df_clean[df_clean['รหัสสาขา'] == store_code].copy()
                if curr > 0: page_breaks.append(curr)
                first = df_store.iloc[0]

                # --- Left Header ---
                worksheet.merge_range(curr, 0, curr, 2, 'บริษัท โมบาย โลจิสติกส์ จำกัด', f_comp)
                worksheet.merge_range(curr+1, 0, curr+1, 1, '279 หมู่ที่ 9 ตำบลบางโฉลง อำเภอบางพลี', f_std)
                worksheet.merge_range(curr+2, 0, curr+2, 1, 'จังหวัดสมุทรปราการ 10540', f_std)
                worksheet.merge_range(curr+3, 0, curr+3, 2, 'ติดต่อ/สอบถาม : ID Line Official : @505phsps (มี @ ), Tel : 099-157-3114', f_std)
                worksheet.write(curr+4, 0, 'Customer Name: ............................................................................', f_std)
                worksheet.write(curr+5, 0, f'Store Code: {store_code}', f_bold)
                worksheet.write(curr+6, 0, f'Store Name: {first["Store Name"]}', f_bold)
                worksheet.write(curr+7, 0, f'Ship To: {first["ที่อยู่"]}', f_std)

                # --- Center Title (Size 18) ---
                worksheet.merge_range(curr+1, 2, curr+2, 2, 'ใบส่งสินค้าชั่วคราว', f_title)

                # --- Right Data Section (No Dots) ---
                worksheet.write(curr, 3, 'Do. No.', f_right)
                worksheet.write(curr, 4, str(first["เลขที่ DO."]), f_std)
                
                worksheet.write(curr+1, 3, 'Ref. Po.', f_right)
                worksheet.write(curr+1, 4, str(first["เลขที่ PO."]), f_std)
                
                worksheet.write(curr+2, 3, 'Ref. Po.', f_right)
                worksheet.write(curr+2, 4, '-', f_std) 
                
                worksheet.write(curr+3, 3, 'Ref. Po.', f_right)
                worksheet.write(curr+3, 4, '-', f_std) 
                
                cutoff = pd.to_datetime(first.get('Cut Off Date')).strftime('%d/%m/%Y') if pd.notnull(first.get('Cut Off Date')) else "-"
                worksheet.write(curr+4, 3, 'Cut off Date:', f_right)
                worksheet.write(curr+4, 4, cutoff, f_std)

                worksheet.write(curr+5, 3, 'Zone:', f_right)
                worksheet.write(curr+5, 4, str(first["โซน"]), f_std)
                
                # Delivery Date (Merge Vertical Row 7-8, Center, No Highlight)
                delivery = pd.to_datetime(first["Delivery Date"]).strftime('%d/%m/%Y') if pd.notnull(first["Delivery Date"]) else "-"
                worksheet.merge_range(curr+7, 3, curr+8, 3, "Delivery\nDate", f_deliv_label) # No colon, No dots
                worksheet.merge_range(curr+7, 4, curr+8, 4, delivery, f_deliv_date)

                # --- Table Section ---
                t_h = curr + 10 
                headers = ['No.', 'Product Code', 'Product Name', 'Unit/UOM', 'QTY']
                for i, txt in enumerate(headers):
                    worksheet.write(t_h, i, txt, f_table_h)

                r_ptr = t_h + 1
                for i, (_, r) in enumerate(df_store.iterrows(), 1):
                    worksheet.write(r_ptr, 0, i, f_border)
                    worksheet.write(r_ptr, 1, str(r['รหัสสินค้า']), f_border)
                    worksheet.write(r_ptr, 2, r['รายการสินค้า'], f_wrap)
                    worksheet.write(r_ptr, 3, r['หน่วย'], f_border)
                    worksheet.write(r_ptr, 4, r['จำนวน'], f_border)
                    r_ptr += 1
                
                worksheet.merge_range(r_ptr, 0, r_ptr, 3, 'Total', f_table_h)
                worksheet.write(r_ptr, 4, df_store['จำนวน'].sum(), f_border)

                # --- Footer ---
                f_row = r_ptr + 1
                worksheet.merge_range(f_row, 0, f_row, 1, 'ผู้รับสินค้า', f_footer_h)
                worksheet.merge_range(f_row, 2, f_row, 3, 'ผู้ส่งสินค้า / ทะเบียนรถ', f_footer_h)
                worksheet.write(f_row, 4, 'คลังสินค้า', f_footer_h)
                
                for label in ['ชื่อ (ตัวบรรจง):', 'วันที่:', 'เวลา:', 'หมายเหตุ:']:
                    f_row += 1
                    worksheet.merge_range(f_row, 0, f_row, 1, label, f_footer_box)
                    worksheet.merge_range(f_row, 2, f_row, 3, label, f_footer_box)
                    worksheet.write(f_row, 4, label, f_footer_box)

                curr = f_row + 2

            worksheet.set_h_pagebreaks(page_breaks)
            worksheet.fit_to_pages(1, 0)

        st.success("✅ V11.0 เรียบร้อยครับ! เอาจุดและไฮไลท์สีออกแล้ว")
        st.download_button("📥 ดาวน์โหลด DO V11.0 (Clean)", output_do.getvalue(), "DO_Clean_V11.0.xlsx")