import streamlit as st
import pandas as pd
import io

st.title("📦 ระบบแยกข้อมูลซัพพลายเออร์ (กำหนดชื่อชีตเอง)")

uploaded_file = st.file_uploader("อัปโหลดไฟล์ Excel (Transport)", type=['xlsx'])

if uploaded_file:
    # 1. อ่านข้อมูลเพื่อเช็คชื่อซัพพลายเออร์ก่อน
    df_raw = pd.read_excel(uploaded_file, sheet_name='Transport')
    df_raw = df_raw.dropna(subset=['ซัพพลายเออร์'])
    unique_suppliers = sorted(df_raw['ซัพพลายเออร์'].unique())

    st.subheader("📝 กำหนดชื่อตัวย่อสำหรับแต่ละชีต")
    st.info("กรุณากรอกชื่อตัวย่อ (ห้ามเกิน 31 ตัวอักษร และห้ามมีตัวอักษรพิเศษ)")

    # 2. สร้าง Dictionary เพื่อเก็บค่าตัวย่อจากหน้าเว็บ
    sheet_mapping = {}
    
    # สร้าง Form เพื่อให้กรอกทีเดียว
    with st.form("sheet_name_form"):
        cols = st.columns(2) # แบ่งเป็น 2 คอลัมน์ให้ประหยัดพื้นที่
        for i, supplier in enumerate(unique_suppliers):
            # สลับฝั่งซ้ายขวา
            with cols[i % 2]:
                # ตั้งชื่อ Default เป็นชื่อที่ตัดให้สั้นๆ หรือจะปล่อยว่างไว้ก็ได้
                default_val = str(supplier)[:10].strip()
                sheet_mapping[supplier] = st.text_input(f"ชื่อชีตสำหรับ: {supplier}", value=default_val, key=supplier)
        
        submit_button = st.form_submit_button("ยืนยันและสร้างไฟล์ Excel")

    # 3. เมื่อกดปุ่มยืนยัน จึงเริ่มสร้างไฟล์
    if submit_button:
        output = io.BytesIO()
        try:
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                for supplier, sheet_name in sheet_mapping.items():
                    # ตรวจสอบความถูกต้องของชื่อชีตที่กรอกมา
                    clean_name = sheet_name.strip()[:31]
                    if not clean_name: clean_name = "Sheet" # กันเหนียวถ้าไม่ได้กรอก
                    
                    df_sup = df_raw[df_raw['ซัพพลายเออร์'] == supplier].copy()
                    
                    # --- จัดการข้อมูลฝั่งซ้าย (แยกสาขา) ---
                    left_side = df_sup[['รหัสสาขา', 'โซน', 'รหัสสินค้า', 'รายการสินค้า', 'ซัพพลายเออร์', 'หน่วย', 'จำนวน']].copy()
                    left_side.rename(columns={'จำนวน': 'Total'}, inplace=True)
                    
                    # --- จัดการข้อมูลฝั่งขวา (สรุปยอด) ---
                    right_side = df_sup.groupby(['ซัพพลายเออร์', 'รหัสสินค้า', 'รายการสินค้า'], as_index=False)['จำนวน'].sum()
                    right_side.rename(columns={'จำนวน': 'Total'}, inplace=True)
                    
                    # เขียนลง Excel
                    left_side.to_excel(writer, sheet_name=clean_name, index=False, startcol=0)
                    right_side.to_excel(writer, sheet_name=clean_name, index=False, startcol=9)
                    
                    # ปรับความกว้างคอลัมน์อัตโนมัติ
                    worksheet = writer.sheets[clean_name]
                    worksheet.set_column('A:G', 15)
                    worksheet.set_column('I:L', 20)

            processed_data = output.getvalue()
            
            st.success("✅ สร้างไฟล์เสร็จสมบูรณ์!")
            st.download_button(
                label="📥 ดาวน์โหลดไฟล์ Excel",
                data=processed_data,
                file_name="Supplier_Split_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"เกิดข้อผิดพลาด: {str(e)}")

else:
    st.info("กรุณาอัปโหลดไฟล์ข้อมูลดิบจากชีต Transport")