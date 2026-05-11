import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Supplier All-in-One V7.1", layout="wide")

st.title("📊 ระบบรวมข้อมูล (Summary + Invoice) ในชีทเดียว")

uploaded_file = st.file_uploader("อัปโหลดไฟล์ Excel (Transport)", type=['xlsx'])

if uploaded_file:
    df_raw = pd.read_excel(uploaded_file, sheet_name='Transport')
    df_raw = df_raw.dropna(subset=['ซัพพลายเออร์'])
    unique_suppliers = sorted(df_raw['ซัพพลายเออร์'].unique())

    # แสดงรายการซัพพลายเออร์ที่พบ
    st.info(f"พบซัพพลายเออร์ทั้งหมด {len(unique_suppliers)} ราย กำลังเตรียมสร้างรายงานรวม...")

    if st.button("สร้างไฟล์ Excel รวมทุกอย่าง"):
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            workbook = writer.book
            
            # --- กำหนด Styles ---
            title_fmt = workbook.add_format({'bold': True, 'font_size': 16, 'align': 'center'})
            total_style = workbook.add_format({'bold': True, 'bg_color': '#D9EAD3', 'border': 1, 'align': 'right'})
            table_head_fmt = workbook.add_format({'bold': True, 'border': 1, 'align': 'center', 'bg_color': '#CFE2F3'})
            cell_border = workbook.add_format({'border': 1})
            header_label_fmt = workbook.add_format({'bold': True})
            footer_box_fmt = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'top'})

            # สร้างชีทเดียวชื่อ "Combined_Report"
            worksheet = workbook.add_worksheet("Combined_Report")
            worksheet.set_zoom(80)
            curr_row = 0

            for supplier in unique_suppliers:
                df_sup = df_raw[df_raw['ซัพพลายเออร์'] == supplier].copy()
                
                # --- PART 1: ตารางแบบเดิม (Summary Table) ---
                worksheet.write(curr_row, 0, f"📌 ส่วนที่ 1: สรุปข้อมูล - {supplier}", header_label_fmt)
                curr_row += 1
                
                # เตรียมข้อมูลฝั่งซ้าย (ตาม image_8928d1.png)
                left_cols = ['รหัสสาขา', 'โซน', 'รหัสสินค้า', 'รายการสินค้า', 'ซัพพลายเออร์', 'หน่วย', 'จำนวน']
                left_final = df_sup[left_cols].copy()
                left_final.rename(columns={'จำนวน': 'Total'}, inplace=True)
                
                # เขียนหัวตาราง Summary
                for c, col_name in enumerate(left_final.columns):
                    worksheet.write(curr_row, c, col_name, table_head_fmt)
                
                start_table_row = curr_row + 1
                for i, row in left_final.iterrows():
                    curr_row += 1
                    for c, val in enumerate(row):
                        worksheet.write(curr_row, c, val, cell_border)
                
                # Grand Total ฝั่งซ้าย
                curr_row += 1
                worksheet.write(curr_row, 0, "Grand Total", total_style)
                worksheet.write(curr_row, 6, df_sup['จำนวน'].sum(), total_style)
                
                curr_row += 3 # เว้นช่องว่างระหว่างซัพพลายเออร์หรือส่วน

                # --- PART 2: รูปแบบใบส่งของ (Invoice Format) ---
                worksheet.merge_range(curr_row, 0, curr_row, 4, 'บริษัท โมบาย โลจิสติกส์ จำกัด', title_fmt)
                curr_row += 1
                worksheet.merge_range(curr_row, 0, curr_row, 4, f'ใบส่งสินค้าชั่วคราว ({supplier})', title_fmt)
                curr_row += 2
                
                # Header Invoice (ตาม image_89132c.png)
                worksheet.write(curr_row, 0, 'Customer Name:', header_label_fmt)
                worksheet.write(curr_row, 4, 'Delivery Date:', header_label_fmt)
                curr_row += 2

                # ตารางสินค้า Invoice
                inv_headers = ['No.', 'Product Code', 'Product Name', 'Unit/UOM', 'QTY']
                for c, h in enumerate(inv_headers):
                    worksheet.write(curr_row, c, h, table_head_fmt)
                
                df_inv = df_sup.groupby(['รหัสสินค้า', 'รายการสินค้า', 'หน่วย'], as_index=False)['จำนวน'].sum()
                for i, row in df_inv.iterrows():
                    curr_row += 1
                    worksheet.write(curr_row, 0, i + 1, cell_border)
                    worksheet.write(curr_row, 1, row['รหัสสินค้า'], cell_border)
                    worksheet.write(curr_row, 2, row['รายการสินค้า'], cell_border)
                    worksheet.write(curr_row, 3, row['หน่วย'], cell_border)
                    worksheet.write(curr_row, 4, row['จำนวน'], cell_border)
                
                # Total Invoice
                curr_row += 1
                worksheet.merge_range(curr_row, 0, curr_row, 3, 'Total', total_style)
                worksheet.write(curr_row, 4, df_inv['จำนวน'].sum(), total_style)

                # Footer ลายเซ็น (ตาม image_89132c.png)
                curr_row += 2
                f_heads = ['ผู้รับสินค้า', 'ผู้ส่งสินค้า', 'คลังสินค้า']
                for i, head in enumerate(f_heads):
                    worksheet.write(curr_row, i*1.5, head, header_label_fmt) # ปรับตำแหน่ง
                    worksheet.write(curr_row+1, i*1.5, 'ชื่อ..................\nวันที่................', footer_box_fmt)
                
                curr_row += 6 # เว้นระยะห่างก่อนขึ้นซัพพลายเออร์คนถัดไป
                worksheet.set_row(curr_row, 2, workbook.add_format({'bg_color': '#333333'})) # ขีดเส้นคั่นหนาๆ
                curr_row += 2

            # ปรับความกว้างคอลัมน์ให้เหมาะสมกับ "รายการสินค้า"
            worksheet.set_column('A:A', 12) # รหัสสาขา
            worksheet.set_column('C:C', 12) # รหัสสินค้า
            worksheet.set_column('D:D', 45) # รายการสินค้า (กว้างพิเศษตามรูป)
            worksheet.set_column('E:E', 25) # ซัพพลายเออร์

        st.success("✅ รวมข้อมูลทุกซัพพลายเออร์ไว้ในชีทเดียวเรียบร้อย!")
        st.download_button(label="📥 ดาวน์โหลดไฟล์ All-in-One", data=output.getvalue(), file_name="Supplier_Combined_Report.xlsx")