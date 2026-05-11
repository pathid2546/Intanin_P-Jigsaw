import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Supplier Splitter Pro", layout="wide")

st.title("📦 ระบบแยกข้อมูลซัพพลายเออร์ (Layout แยกสาขา & สรุปยอด)")

# --- ระบบความจำชื่อตัวย่อ ---
if 'name_memory' not in st.session_state:
    st.session_state['name_memory'] = {}

uploaded_file = st.file_uploader("อัปโหลดไฟล์ Excel (Transport)", type=['xlsx'])

if uploaded_file:
    # อ่านข้อมูลจาก Transport
    df_raw = pd.read_excel(uploaded_file, sheet_name='Transport')
    df_raw = df_raw.dropna(subset=['ซัพพลายเออร์'])
    unique_suppliers = sorted(df_raw['ซัพพลายเออร์'].unique())

    st.subheader("📝 กำหนดชื่อตัวย่อชีต")
    with st.form("sheet_name_form"):
        cols = st.columns(3)
        current_mapping = {}
        for i, supplier in enumerate(unique_suppliers):
            with cols[i % 3]:
                remembered_name = st.session_state['name_memory'].get(supplier, str(supplier)[:10].strip())
                current_mapping[supplier] = st.text_input(f"{supplier}:", value=remembered_name, key=f"input_{supplier}")
        submit_button = st.form_submit_button("สร้างไฟล์ Excel")

    if submit_button:
        # บันทึกความจำ
        for sup, s_name in current_mapping.items():
            st.session_state['name_memory'][sup] = s_name
            
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            for supplier, sheet_name in current_mapping.items():
                clean_name = sheet_name.strip()[:31]
                df_sup = df_raw[df_raw['ซัพพลายเออร์'] == supplier].copy()
                
                # --- [ด้านซ้าย] เตรียมข้อมูลแยกสาขา ---
                left_data = df_sup[['รหัสสาขา', 'โซน', 'รหัสสินค้า', 'รายการสินค้า', 'ซัพพลายเออร์', 'หน่วย', 'จำนวน']].copy()
                left_data.rename(columns={'จำนวน': 'Total'}, inplace=True)
                
                # Logic: ทำให้ รหัสสาขา และ โซน แสดงแค่บรรทัดแรกของกลุ่ม
                left_data['รหัสสาขา'] = left_data['รหัสสาขา'].mask(left_data['รหัสสาขา'].duplicated(), "")
                # (หมายเหตุ: ถ้าต้องการให้โซนว่างด้วยเมื่อรหัสสาขาซ้ำ)
                left_data.loc[left_data['รหัสสาขา'] == "", 'โซน'] = ""
                
                # --- [ด้านขวา] เตรียมข้อมูลสรุปยอด ---
                right_data = df_sup.groupby(['รหัสสินค้า', 'รายการสินค้า'], as_index=False)['จำนวน'].sum()
                right_data.rename(columns={'จำนวน': 'Total'}, inplace=True)
                right_data.insert(0, 'ซัพพลายเออร์', "") # เพิ่มคอลัมน์ว่างด้านหน้าตามรูป

                # --- เขียนข้อมูลลง Sheet ---
                # เขียนฝั่งซ้าย
                left_data.to_excel(writer, sheet_name=clean_name, index=False, startrow=0, startcol=0)
                # เขียนฝั่งขวา (เริ่มคอลัมน์ I คือ col index 8)
                right_data.to_excel(writer, sheet_name=clean_name, index=False, startrow=0, startcol=8)

                # --- ตกแต่งและเพิ่ม Grand Total ด้วย XlsxWriter ---
                workbook  = writer.book
                worksheet = writer.sheets[clean_name]
                
                header_format = workbook.add_format({'bold': True, 'bg_color': '#D9EAD3', 'border': 1})
                total_format  = workbook.add_format({'bold': True, 'bg_color': '#CFE2F3', 'border': 1})
                
                # เพิ่ม Grand Total ด้านซ้าย
                last_row_left = len(left_data) + 1
                worksheet.write(last_row_left, 0, "Grand Total", total_format)
                worksheet.write(last_row_left, 6, left_data['Total'].sum(), total_format)
                
                # เพิ่ม Total ด้านขวา
                last_row_right = len(right_data) + 1
                worksheet.write(last_row_right, 8, f"{supplier} Total", total_format)
                worksheet.write(last_row_right, 10, right_data['Total'].sum(), total_format)

                # ปรับขนาดคอลัมน์
                worksheet.set_column('A:B', 10)
                worksheet.set_column('C:D', 25)
                worksheet.set_column('I:J', 25)

        st.success("✅ ประมวลผลสำเร็จ!")
        st.download_button(
            label="📥 ดาวน์โหลดไฟล์ Excel",
            data=output.getvalue(),
            file_name="Supplier_Formatted_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )