import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Supplier All-in-One V7.2", layout="wide")

st.title("📊 ระบบรวมข้อมูล (Summary + Invoice) - เพิ่มชื่อสาขา")

uploaded_file = st.file_uploader("อัปโหลดไฟล์ Excel (Transport)", type=['xlsx'])

if uploaded_file:
    df_raw = pd.read_excel(uploaded_file, sheet_name='Transport')
    df_raw = df_raw.dropna(subset=['ซัพพลายเออร์'])
    unique_suppliers = sorted(df_raw['ซัพพลายเออร์'].unique())

    st.info(f"พบซัพพลายเออร์ {len(unique_suppliers)} ราย กำลังจัดทำรายงานชีทเดียว...")

    if st.button("สร้างไฟล์ Excel รวมทุกส่วน"):
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

            # สร้างชีทเดียว
            worksheet = workbook.add_worksheet("All_In_One_Report")
            worksheet.set_zoom(80)
            curr_row = 0

            for supplier in unique_suppliers:
                df_sup = df_raw[df_raw['ซัพพลายเออร์'] == supplier].copy()
                
                # --- PART 1: ตารางสรุป (Summary) - เพิ่มชื่อสาขากลับมา ---
                worksheet.write(curr_row, 0, f"📍 ส่วนที่ 1: รายละเอียดการส่งสินค้า - {supplier}", header_label_fmt)
                curr_row += 1
                
                # เตรียมคอลัมน์ฝั่งซ้าย (เพิ่ม Store Name)
                left_display = df_sup.rename(columns={'Store Name': 'ชื่อสาขา', 'จำนวน': 'Total'})
                left_cols = ['รหัสสาขา', 'ชื่อสาขา', 'โซน', 'รหัสสินค้า', 'รายการสินค้า', 'ซัพพลายเออร์', 'หน่วย', 'Total']
                left_final = left_display[left_cols].copy()
                
                # เก็บไว้คำนวณ Auto-Fit
                width_ref = left_final.copy()
                
                # ทำ Mask ซ่อนค่าซ้ำ
                mask = left_final['รหัสสาขา'].duplicated()
                left_final.loc[mask, ['รหัสสาขา', 'ชื่อสาขา', 'โซน']] = ""
                
                # เขียนหัวตาราง Summary
                for c, col_name in enumerate(left_final.columns):
                    worksheet.write(curr_row, c, col_name, table_head_fmt)
                
                # เขียนข้อมูลตาราง
                for i, row in left_final.iterrows():
                    curr_row += 1
                    for c, val in enumerate(row):
                        worksheet.write(curr_row, c, val, cell_border)
                
                # Grand Total ฝั่งซ้าย (อยู่ที่คอลัมน์ Total/index 7)
                curr_row += 1
                worksheet.write(curr_row, 0, "Grand Total", total_style)
                worksheet.write(curr_row, 7, df_sup['จำนวน'].sum(), total_style)
                
                curr_row += 3 # เว้นช่องว่าง

                # --- PART 2: รูปแบบใบส่งของ (Invoice) ---
                worksheet.merge_range(curr_row, 0, curr_row, 4, 'บริษัท โมบาย โลจิสติกส์ จำกัด', title_fmt)
                curr_row += 1
                worksheet.merge_range(curr_row, 0, curr_row, 4, f'ใบส่งสินค้าชั่วคราว - {supplier}', title_fmt)
                curr_row += 2
                
                worksheet.write(curr_row, 0, 'Customer Name: .......................................', header_label_fmt)
                worksheet.write(curr_row, 4, 'Delivery Date: ....................', header_label_fmt)
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

                # Footer ช่องเซ็นชื่อ
                curr_row += 2
                f_heads = ['ผู้รับสินค้า', 'ผู้ส่งสินค้า', 'คลังสินค้า']
                for i, head in enumerate(f_heads):
                    col_pos = i * 2
                    worksheet.write(curr_row, col_pos, head, header_label_fmt)
                    worksheet.write(curr_row+1, col_pos, 'ชื่อ..................\nวันที่................', footer_box_fmt)
                
                curr_row += 6 # เว้นระยะก่อนรายต่อไป
                worksheet.set_row(curr_row, 3, workbook.add_format({'bg_color': '#CCCCCC'})) # เส้นแบ่ง
                curr_row += 2

            # ปรับความกว้างคอลัมน์ (Auto-Fit แบบประมาณค่า)
            worksheet.set_column('A:A', 10) # รหัสสาขา
            worksheet.set_column('B:B', 30) # ชื่อสาขา (เพิ่มกลับมา)
            worksheet.set_column('D:D', 12) # รหัสสินค้า
            worksheet.set_column('E:E', 45) # รายการสินค้า
            worksheet.set_column('F:F', 20) # ซัพพลายเออร์

        st.success("✅ รวมข้อมูลพร้อมชื่อสาขาและใบส่งของเรียบร้อย!")
        st.download_button(label="📥 ดาวน์โหลดไฟล์ All-in-One V7.2", data=output.getvalue(), file_name="Supplier_Summary_Invoice.xlsx")