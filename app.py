import streamlit as st
import pandas as pd
import io
from datetime import datetime

# --- การตั้งค่าหน้ากระดาษ ---
st.set_page_config(page_title="Delivery Formatter Pro", layout="wide")

# --- 1. ส่วนการอัปโหลดไฟล์ (อยู่นอก Tab เพื่อป้องกัน Duplicate ID) ---
st.title("🚚 Delivery Formatter Pro (Mobile Logistics Edition)")
uploaded_file = st.file_uploader("อัปโหลดไฟล์ Excel (Transport)", type=['xlsx'], key="main_loader")

if uploaded_file:
    # อ่านข้อมูลดิบ (ปรับเปลี่ยนตามโครงสร้างไฟล์จริงของคุณ)
    # สมมติว่า header อยู่แถวที่ 0 หากไม่ใช่ให้ปรับ header=...
    df_raw = pd.read_excel(uploaded_file)
    
    # สร้าง Tabs
    tab1, tab2 = st.tabs(["📊 แยกซัพพลายเออร์ (Tab 1)", "📑 ออกใบส่งสินค้า (Tab 2)"])

    # --- TAB 1: แยกข้อมูลรายซัพพลายเออร์ (Format ตาม image_0b3a7c.png) ---
    with tab1:
        st.subheader("📊 แยกข้อมูลรายซัพพลายเออร์")
        if st.button("🚀 สร้างไฟล์แยกซัพพลายเออร์", key="btn_tab1"):
            output_split = io.BytesIO()
            with pd.ExcelWriter(output_split, engine='xlsxwriter') as writer:
                workbook = writer.book
                
                # Styles สำหรับ Tab 1
                header_fmt = workbook.add_format({'bold': True, 'bg_color': '#D9EAD3', 'border': 1, 'align': 'center'})
                cell_fmt = workbook.add_format({'border': 1})
                num_fmt = workbook.add_format({'border': 1, 'align': 'right'})
                total_row_fmt = workbook.add_format({'bold': True, 'border': 1, 'bg_color': '#F3F3F3', 'num_format': '#,##0'})

                for supplier, s_data in df_raw.groupby('ซัพพลายเออร์'):
                    sheet_name = str(supplier)[:31].replace('/', '-')
                    worksheet = workbook.add_worksheet(sheet_name)
                    
                    # หัวตารางฝั่งซ้าย
                    headers_left = ['รหัสสาขา', 'โซน', 'รหัสสินค้า', 'รายการสินค้า', 'ซัพพลายเออร์', 'หน่วย', 'Total']
                    for col, h in enumerate(headers_left):
                        worksheet.write(0, col, h, header_fmt)
                    
                    curr_row = 1
                    # จัดกลุ่มตามสาขาเพื่อทำ Blank Rows (image_0b3a7c.png)
                    for (branch, zone), b_group in s_data.groupby(['รหัสสาขา', 'โซน'], sort=False):
                        for i, (_, row) in enumerate(b_group.iterrows()):
                            if i == 0: # เขียนรหัสสาขาเฉพาะบรรทัดแรกของกลุ่ม
                                worksheet.write(curr_row, 0, branch, cell_fmt)
                                worksheet.write(curr_row, 1, zone, cell_fmt)
                            else:
                                worksheet.write(curr_row, 0, "", cell_fmt)
                                worksheet.write(curr_row, 1, "", cell_fmt)
                            
                            worksheet.write(curr_row, 2, row['รหัสสินค้า'], cell_fmt)
                            worksheet.write(curr_row, 3, row['รายการสินค้า'], cell_fmt)
                            worksheet.write(curr_row, 4, row['ซัพพลายเออร์'], cell_fmt)
                            worksheet.write(curr_row, 5, row['หน่วย'], cell_fmt)
                            worksheet.write(curr_row, 6, row['จำนวน'], num_fmt)
                            curr_row += 1
                    
                    # ฝั่งขวา: สรุปรายการสินค้า (Summary Table)
                    col_off = 9 # คอลัมน์ J
                    headers_right = ['รหัสสินค้า', 'รายการสินค้า', 'Total']
                    for col, h in enumerate(headers_right):
                        worksheet.write(0, col_off + col, h, header_fmt)
                    
                    summary = s_data.groupby(['รหัสสินค้า', 'รายการสินค้า'])['จำนวน'].sum().reset_index()
                    for i, row in summary.iterrows():
                        worksheet.write(i + 1, col_off, row['รหัสสินค้า'], cell_fmt)
                        worksheet.write(i + 1, col_off + 1, row['รายการสินค้า'], cell_fmt)
                        worksheet.write(i + 1, col_off + 2, row['จำนวน'], num_fmt)
                    
                    # บรรทัด Grand Total
                    t_row = len(summary) + 1
                    worksheet.merge_range(t_row, col_off, t_row, col_off + 1, f"{supplier} Total", total_row_fmt)
                    worksheet.write(t_row, col_off + 2, s_data['จำนวน'].sum(), total_row_fmt)
                    
                    worksheet.set_column('A:B', 12); worksheet.set_column('C:C', 15)
                    worksheet.set_column('D:E', 35); worksheet.set_column('J:K', 15); worksheet.set_column('K:K', 40)

            st.download_button("📥 โหลดไฟล์แยกซัพพลายเออร์", output_split.getvalue(), "Supplier_Split.xlsx")

    # --- TAB 2: ออกใบส่งสินค้า (Format ตาม image_0c2ad4.png) ---
    with tab2:
        st.subheader("📑 ออกใบส่งสินค้า (ฉบับพิมพ์แยกหน้า A4)")
        df_clean = df_raw.dropna(subset=['รหัสสาขา']).copy()
        
        if st.button("🚀 สร้างไฟล์ใบส่งสินค้า (DO Master)", key="btn_tab2"):
            output_do = io.BytesIO()
            with pd.ExcelWriter(output_do, engine='xlsxwriter') as writer:
                workbook = writer.book
                worksheet = workbook.add_worksheet("DO_Continuous")
                
                # Styles สำหรับ Tab 2
                f_title = workbook.add_format({'bold': True, 'font_size': 18, 'align': 'center'})
                f_header_bold = workbook.add_format({'bold': True, 'font_size': 11})
                f_std = workbook.add_format({'font_size': 10})
                f_right = workbook.add_format({'font_size': 10, 'align': 'right'})
                f_table_head = workbook.add_format({'border': 1, 'align': 'center', 'bold': True, 'bg_color': '#F2F2F2'})
                f_border_center = workbook.add_format({'border': 1, 'align': 'center', 'font_size': 10})
                f_wrap = workbook.add_format({'border': 1, 'text_wrap': True, 'font_size': 10})
                f_ship_to = workbook.add_format({'font_size': 10, 'text_wrap': True, 'valign': 'top'})
                f_date_val = workbook.add_format({'bold': True, 'font_size': 11, 'num_format': 'dd/mm/yyyy', 'align': 'right'})
                f_footer_box = workbook.add_format({'border': 1, 'font_size': 9, 'valign': 'top', 'text_wrap': True})

                # ตั้งค่าหน้ากระดาษ A4
                worksheet.set_paper(9); worksheet.set_margins(0.3, 0.3, 0.3, 0.3)
                worksheet.set_column('A:A', 5); worksheet.set_column('B:B', 14); worksheet.set_column('C:C', 45)
                worksheet.set_column('D:D', 12); worksheet.set_column('E:E', 15)
                
                curr = 0; page_breaks = []; do_count = 0 

                for store_code in df_clean['รหัสสาขา'].unique():
                    df_store = df_clean[df_clean['รหัสสาขา'] == store_code].copy()
                    if do_count > 0: page_breaks.append(curr) # บังคับตัดหน้าเมื่อขึ้น DO ใหม่
                    
                    first_row = df_store.iloc[0]
                    # Header Section
                    worksheet.write(curr, 0, 'บริษัท โมบาย โลจิสติกส์ จำกัด', f_header_bold)
                    worksheet.write(curr+1, 0, '279 หมู่ที่ 9 ตำบลบางโฉลง...', f_std)
                    worksheet.write(curr+1, 2, 'ใบส่งสินค้าชั่วคราว', f_title) # C2 Title
                    worksheet.write(curr, 4, f'Do. No. {first_row["เลขที่ DO."]}', f_right)
                    worksheet.write(curr+7, 4, first_row['Delivery Date'], f_date_val)

                    # Ship To Section (Merge B:C)
                    worksheet.write(curr+5, 0, f'Store Code: {store_code}', f_header_bold)
                    worksheet.merge_range(curr+7, 1, curr+8, 2, first_row['ที่อยู่'], f_ship_to)

                    # Table
                    header_row = curr + 10
                    for col, text in enumerate(['No.', 'Product Code', 'Product Name', 'Unit/UOM', 'QTY']):
                        worksheet.write(header_row, col, text, f_table_head)

                    r_ptr = header_row + 1
                    for i, (_, r) in enumerate(df_store.iterrows(), 1):
                        worksheet.write(r_ptr, 0, i, f_border_center)
                        worksheet.write(r_ptr, 2, r['รายการสินค้า'], f_wrap)
                        worksheet.write(r_ptr, 4, r['จำนวน'], f_border_center)
                        r_ptr += 1
                    
                    # Footer Section
                    f_ptr = r_ptr + 1
                    worksheet.set_row(f_ptr + 1, 60)
                    worksheet.merge_range(f_ptr, 0, f_ptr, 1, 'ผู้รับสินค้า', f_table_head)
                    worksheet.merge_range(f_ptr+1, 0, f_ptr+1, 1, 'ชื่อ:\nวันที่:', f_footer_box)
                    worksheet.write(f_ptr+1, 4, 'ชื่อ:\nวันที่:', f_footer_box)

                    curr = f_ptr + 4; do_count += 1

                if page_breaks: worksheet.set_h_pagebreaks(page_breaks)

            st.download_button("📥 โหลดไฟล์ DO Master", output_do.getvalue(), f"DO_Master_{datetime.now().strftime('%H%M')}.xlsx")