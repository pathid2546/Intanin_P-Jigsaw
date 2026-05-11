import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Supplier Summary V7.3", layout="wide")

st.title("📊 ระบบสรุปข้อมูลซัพพลายเออร์ (All-in-One Sheet)")

uploaded_file = st.file_uploader("อัปโหลดไฟล์ Excel (Transport)", type=['xlsx'])

if uploaded_file:
    df_raw = pd.read_excel(uploaded_file, sheet_name='Transport')
    df_raw = df_raw.dropna(subset=['ซัพพลายเออร์'])
    unique_suppliers = sorted(df_raw['ซัพพลายเออร์'].unique())

    st.info(f"พบซัพพลายเออร์ {len(unique_suppliers)} ราย กำลังรวมข้อมูลลงในชีทเดียว...")

    if st.button("สร้างไฟล์ Excel"):
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            workbook = writer.book
            
            # --- Styles ---
            total_style = workbook.add_format({'bold': True, 'bg_color': '#D9EAD3', 'border': 1, 'align': 'right'})
            table_head_fmt = workbook.add_format({'bold': True, 'border': 1, 'align': 'center', 'bg_color': '#CFE2F3'})
            cell_border = workbook.add_format({'border': 1})
            sup_header_fmt = workbook.add_format({'bold': True, 'font_size': 14, 'font_color': '#1155CC'})

            # สร้างชีทเดียวชื่อ "Main_Summary"
            worksheet = workbook.add_worksheet("Main_Summary")
            worksheet.set_zoom(80)
            curr_row = 0

            for supplier in unique_suppliers:
                df_sup = df_raw[df_raw['ซัพพลายเออร์'] == supplier].copy()
                
                # เขียนชื่อซัพพลายเออร์คั่นแต่ละตาราง
                worksheet.write(curr_row, 0, f"📦 ซัพพลายเออร์: {supplier}", sup_header_fmt)
                curr_row += 1
                
                # เตรียมคอลัมน์ (ดึง ชื่อสาขา กลับมา)
                left_display = df_sup.rename(columns={'Store Name': 'ชื่อสาขา', 'จำนวน': 'Total'})
                left_cols = ['รหัสสาขา', 'ชื่อสาขา', 'โซน', 'รหัสสินค้า', 'รายการสินค้า', 'ซัพพลายเออร์', 'หน่วย', 'Total']
                left_final = left_display[left_cols].copy()
                
                # ทำ Mask ซ่อนค่าซ้ำ
                mask = left_final['รหัสสาขา'].duplicated()
                left_final.loc[mask, ['รหัสสาขา', 'ชื่อสาขา', 'โซน']] = ""
                
                # เขียนหัวตาราง
                for c, col_name in enumerate(left_final.columns):
                    worksheet.write(curr_row, c, col_name, table_head_fmt)
                
                # เขียนข้อมูลบรรทัดต่อบรรทัด
                for i, row in left_final.iterrows():
                    curr_row += 1
                    for c, val in enumerate(row):
                        worksheet.write(curr_row, c, val, cell_border)
                
                # เพิ่ม Grand Total
                curr_row += 1
                worksheet.write(curr_row, 0, "Grand Total", total_style)
                # รวมยอดที่คอลัมน์ H (index 7)
                worksheet.write(curr_row, 7, df_sup['จำนวน'].sum(), total_style)
                
                # เว้นระยะห่าง 3 บรรทัดก่อนเริ่มเจ้าถัดไป
                curr_row += 3

            # ปรับความกว้างคอลัมน์ให้กะทัดรัด (อ้างอิงตามที่คุณชอบ)
            worksheet.set_column('A:A', 10) # รหัสสาขา
            worksheet.set_column('B:B', 35) # ชื่อสาขา (ยาวหน่อย)
            worksheet.set_column('C:C', 8)  # โซน
            worksheet.set_column('D:D', 12) # รหัสสินค้า
            worksheet.set_column('E:E', 45) # รายการสินค้า
            worksheet.set_column('F:F', 20) # ซัพพลายเออร์
            worksheet.set_column('G:G', 10) # หน่วย
            worksheet.set_column('H:H', 8)  # Total

        st.success("✅ รวมข้อมูลทุกเจ้าลงในชีทเดียว (พร้อมชื่อสาขา) เรียบร้อย!")
        st.download_button(label="📥 ดาวน์โหลดไฟล์ Excel V7.3", data=output.getvalue(), file_name="Supplier_All_In_One.xlsx")