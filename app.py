import streamlit as st
import pandas as pd
import io

st.title("📦 ระบบแยกข้อมูลซัพพลายเออร์อัตโนมัติ")

uploaded_file = st.file_uploader("อัปโหลดไฟล์ Excel (ที่มีชีต Transport)", type=['xlsx'])

if uploaded_file:
    # 1. อ่านข้อมูลจากชีต Transport
    df_raw = pd.read_excel(uploaded_file, sheet_name='Transport')
    
    # ตรวจสอบชื่อคอลัมน์ให้ตรงกับ Raw Data
    # (สมมติว่า 'จำนวน' ในรูปคือค่าที่จะนำมาทำเป็น 'Total')
    
    suppliers = df_raw['ซัพพลายเออร์'].unique()
    
    # เตรียม Buffer สำหรับเขียนไฟล์ Excel
    output = io.BytesIO()
    
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for supplier in suppliers:
            # กรองข้อมูลตามซัพพลายเออร์
            df_sup = df_raw[df_raw['ซัพพลายเออร์'] == supplier].copy()
            
            # --- ฝั่งซ้าย: ข้อมูลแยกสาขา ---
            left_side = df_sup[['รหัสสาขา', 'โซน', 'รหัสสินค้า', 'รายการสินค้า', 'ซัพพลายเออร์', 'หน่วย', 'จำนวน']].copy()
            left_side.rename(columns={'จำนวน': 'Total'}, inplace=True)
            
            # --- ฝั่งขวา: ข้อมูลสรุปซัพพลายเออร์ ---
            right_side = df_sup.groupby(['ซัพพลายเออร์', 'รหัสสินค้า', 'รายการสินค้า'], as_index=False)['จำนวน'].sum()
            right_side.rename(columns={'จำนวน': 'Total'}, inplace=True)
            
            # 2. เขียนข้อมูลลงชีต (ชื่อชีตตามซัพพลายเออร์)
            # วางฝั่งซ้าย (เริ่มคอลัมน์ A)
            left_side.to_excel(writer, sheet_name=supplier, index=False, startcol=0)
            
            # วางฝั่งขวา (เริ่มคอลัมน์ I เพื่อให้เว้นระยะห่างเหมือนในรูป)
            right_side.to_excel(writer, sheet_name=supplier, index=False, startcol=9)
            
            # จัดรูปแบบเพิ่มเติม (Optional: ปรับความกว้างคอลัมน์)
            worksheet = writer.sheets[supplier]
            worksheet.set_column('A:G', 15)
            worksheet.set_column('I:L', 15)

    processed_data = output.getvalue()
    
    st.success(f"ประมวลผลเสร็จแล้ว! แยกได้ทั้งหมด {len(suppliers)} ชีต")
    
    st.download_button(
        label="📥 ดาวน์โหลดไฟล์ Excel ที่แยกชีตแล้ว",
        data=processed_data,
        file_name="Split_Supplier_Report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    st.info("กรุณาอัปโหลดไฟล์ Excel เพื่อเริ่มกระบวนการ")