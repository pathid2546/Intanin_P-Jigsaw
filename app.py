import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Supplier Splitter Pro", layout="wide")

st.title("📦 ระบบแยกข้อมูลซัพพลายเออร์ (Layout V.2.1)")

# --- ระบบความจำชื่อตัวย่อ ---
if 'name_memory' not in st.session_state:
    st.session_state['name_memory'] = {}

uploaded_file = st.file_uploader("อัปโหลดไฟล์ Excel (ที่มีชีต Transport)", type=['xlsx'])

if uploaded_file:
    # 1. อ่านข้อมูลจาก Transport
    df_raw = pd.read_excel(uploaded_file, sheet_name='Transport')
    df_raw = df_raw.dropna(subset=['ซัพพลายเออร์'])
    unique_suppliers = sorted(df_raw['ซัพพลายเออร์'].unique())

    st.subheader("📝 กำหนดชื่อตัวย่อชีต")
    with st.form("sheet_name_form"):
        cols = st.columns(3)
        current_mapping = {}
        for i, supplier in enumerate(unique_suppliers):
            with cols[i % 3]:
                # ดึงค่าจากความจำ หรือใช้ชื่อเต็มตัด 10 ตัวแรก
                remembered_name = st.session_state['name_memory'].get(supplier, str(supplier)[:10].strip())
                current_mapping[supplier] = st.text_input(f"{supplier}:", value=remembered_name, key=f"input_{supplier}")
        submit_button = st.form_submit_button("สร้างไฟล์ Excel")

    if submit_button:
        # บันทึกความจำ
        for sup, s_name in current_mapping.items():
            st.session_state['name_memory'][sup] = s_name
            
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            workbook  = writer.book
            # เตรียม formats สำหรับการตกแต่ง
            total_format = workbook.add_format({'bold': True, 'bg_color': '#CFE2F3', 'border': 1, 'align': 'right'})
            header_format = workbook.add_format({'bold': True, 'bg_color': '#D9EAD3', 'border': 1})

            for supplier, sheet_name in current_mapping.items():
                # ตัดชื่อชีตให้ไม่เกิน 31 ตัวอักษรและลบตัวอักษรพิเศษ
                clean_name = sheet_name.strip()[:31]
                for char in r'[]:*?/\ ':
                    clean_name = clean_name.replace(char, ' ')
                
                df_sup = df_raw[df_raw['ซัพพลายเออร์'] == supplier].copy()
                
                # --- [ด้านซ้าย] แยกสาขา ---
                left_data = df_sup[['รหัสสาขา', 'โซน', 'รหัสสินค้า', 'รายการสินค้า', 'ซัพพลายเออร์', 'หน่วย', 'จำนวน']].copy()
                left_data.rename(columns={'จำนวน': 'Total'}, inplace=True)
                
                # Logic: แสดงรหัสสาขาและโซนแค่บรรทัดแรกของกลุ่ม
                left_data['รหัสสาขา'] = left_data['รหัสสาขา'].mask(left_data['รหัสสาขา'].duplicated(), "")
                left_data.loc[left_data['รหัสสาขา'] == "", 'โซน'] = ""
                
                # --- [ด้านขวา] สรุปยอด (เปลี่ยน Store Name เป็น ชื่อสาขา) ---
                # Group ข้อมูลตาม รหัสสินค้า, รายการสินค้า และ Store Name
                right_data = df_sup.groupby(['รหัสสินค้า', 'รายการสินค้า', 'Store Name'], as_index=False)['จำนวน'].sum()
                
                # เปลี่ยนชื่อคอลัมน์จาก Store Name -> ชื่อสาขา และ จำนวน -> Total
                right_data.rename(columns={'Store Name': 'ชื่อสาขา', 'จำนวน': 'Total'}, inplace=True)
                
                # เพิ่มคอลัมน์ซัพพลายเออร์ว่างไว้ข้างหน้า
                right_data.insert(0, 'ซัพพลายเออร์', "") 

                # --- เขียนข้อมูลลง Sheet ---
                left_data.to_excel(writer, sheet_name=clean_name, index=False, startcol=0)
                # ฝั่งขวาเริ่มที่คอลัมน์ I (index 8)
                right_data.to_excel(writer, sheet_name=clean_name, index=False, startcol=8)

                worksheet = writer.sheets[clean_name]
                
                # --- เพิ่มแถวสรุปผลรวม (Grand Total) ---
                # ฝั่งซ้าย (Grand Total ที่คอลัมน์ G / index 6)
                row_l = len(left_data) + 1
                worksheet.write(row_l, 0, "Grand Total", total_format)
                worksheet.write(row_l, 6, left_data['Total'].sum(), total_format)
                
                # ฝั่งขวา (Total ซัพพลายเออร์ ที่คอลัมน์ M / index 12)
                row_r = len(right_data) + 1
                worksheet.write(row_r, 8, f"{supplier} Total", total_format)
                worksheet.write(row_r, 12, right_data['Total'].sum(), total_format)

                # ปรับขนาดคอลัมน์
                worksheet.set_column('A:B', 12)
                worksheet.set_column('C:D', 30)
                worksheet.set_column('I:L', 30) # คอลัมน์ที่รวม "ชื่อสาขา"
                worksheet.set_column('M:M', 10) # คอลัมน์ Total ฝั่งขวา

        st.success("✅ เปลี่ยนชื่อคอลัมน์เป็น 'ชื่อสาขา' และจัด Layout เรียบร้อย!")
        st.download_button(
            label="📥 ดาวน์โหลดไฟล์ Excel",
            data=output.getvalue(),
            file_name="Supplier_Report_Final.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

else:
    st.info("กรุณาอัปโหลดไฟล์ข้อมูลดิบ (Transport) เพื่อเริ่มทำงาน")