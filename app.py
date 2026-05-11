import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Supplier Splitter", layout="wide")

st.title("📦 ระบบแยกข้อมูลซัพพลายเออร์ (พร้อมระบบจำชื่อตัวย่อ)")

# --- ระบบจัดการความจำ (Session State) ---
if 'name_memory' not in st.session_state:
    # สร้าง Dictionary ว่างไว้เก็บความจำในครั้งแรกที่เปิดแอป
    st.session_state['name_memory'] = {
        "บริษัท อีซี่ อินเตอร์เนชั่นแนล": "EZY",
        "บริษัท ซีนโนวา จำกัด": "SYN"
    }

uploaded_file = st.file_uploader("อัปโหลดไฟล์ Excel (Transport)", type=['xlsx'])

if uploaded_file:
    df_raw = pd.read_excel(uploaded_file, sheet_name='Transport')
    df_raw = df_raw.dropna(subset=['ซัพพลายเออร์'])
    unique_suppliers = sorted(df_raw['ซัพพลายเออร์'].unique())

    st.subheader("📝 ตรวจสอบและแก้ไขชื่อตัวย่อ")
    
    # สร้าง Form เพื่อให้กรอกข้อมูล
    with st.form("sheet_name_form"):
        cols = st.columns(3) # แบ่งเป็น 3 คอลัมน์ให้ดูง่ายขึ้น
        current_mapping = {}
        
        for i, supplier in enumerate(unique_suppliers):
            with cols[i % 3]:
                # ดึงค่าจากความจำ ถ้าไม่มีให้ใช้ชื่อเต็มตัดเหลือ 10 ตัวเป็นค่าเริ่มต้น
                remembered_name = st.session_state['name_memory'].get(supplier, str(supplier)[:10].strip())
                
                # แสดงช่องกรอกข้อมูล
                user_input = st.text_input(
                    f"ซัพพลายเออร์: {supplier}", 
                    value=remembered_name, 
                    key=f"input_{supplier}"
                )
                current_mapping[supplier] = user_input
        
        submit_button = st.form_submit_button("บันทึกชื่อตัวย่อและสร้างไฟล์ Excel")

    if submit_button:
        # อัปเดตความจำใหม่ตามที่ผู้ใช้กรอกล่าสุด
        for sup, s_name in current_mapping.items():
            st.session_state['name_memory'][sup] = s_name
        
        output = io.BytesIO()
        try:
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                for supplier, sheet_name in current_mapping.items():
                    # ทำความสะอาดชื่อ Sheet (ห้ามเกิน 31 ตัว และห้ามมีตัวอักษรพิเศษ)
                    clean_name = sheet_name.strip()[:31]
                    for char in r'[]:*?/\ ':
                        clean_name = clean_name.replace(char, ' ')
                    
                    df_sup = df_raw[df_raw['ซัพพลายเออร์'] == supplier].copy()
                    
                    # ข้อมูลฝั่งซ้าย (แยกสาขา)
                    left_side = df_sup[['รหัสสาขา', 'โซน', 'รหัสสินค้า', 'รายการสินค้า', 'ซัพพลายเออร์', 'หน่วย', 'จำนวน']].copy()
                    left_side.rename(columns={'จำนวน': 'Total'}, inplace=True)
                    
                    # ข้อมูลฝั่งขวา (สรุปยอด)
                    right_side = df_sup.groupby(['ซัพพลายเออร์', 'รหัสสินค้า', 'รายการสินค้า'], as_index=False)['จำนวน'].sum()
                    right_side.rename(columns={'จำนวน': 'Total'}, inplace=True)
                    
                    # เขียนลง Excel
                    left_side.to_excel(writer, sheet_name=clean_name, index=False, startcol=0)
                    right_side.to_excel(writer, sheet_name=clean_name, index=False, startcol=9)
                    
                    # ตกแต่ง Format เบื้องต้น
                    worksheet = writer.sheets[clean_name]
                    worksheet.set_column('A:G', 15)
                    worksheet.set_column('I:L', 20)

            st.success("✅ บันทึกชื่อตัวย่อลงในระบบจำแล้ว และสร้างไฟล์เสร็จสิ้น!")
            st.download_button(
                label="📥 ดาวน์โหลดไฟล์ Excel",
                data=output.getvalue(),
                file_name="Supplier_Split_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"Error: {e}")

else:
    st.info("กรุณาอัปโหลดไฟล์ข้อมูลดิบเพื่อเริ่มทำงาน")

# แสดงสถานะความจำปัจจุบัน (Optional - เอาไว้เช็ค)
if st.checkbox("ดูรายชื่อตัวย่อที่ระบบจำไว้"):
    st.write(st.session_state['name_memory'])