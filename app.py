import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Supplier Splitter V4", layout="wide")

st.title("📦 ระบบแยกข้อมูลซัพพลายเออร์ (ซ้าย-ขวา พร้อมชื่อสาขา)")

# --- ระบบความจำชื่อตัวย่อ ---
if 'name_memory' not in st.session_state:
    st.session_state['name_memory'] = {}

uploaded_file = st.file_uploader("อัปโหลดไฟล์ Excel (Transport)", type=['xlsx'])

if uploaded_file:
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
        for sup, s_name in current_mapping.items():
            st.session_state['name_memory'][sup] = s_name
            
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            workbook  = writer.book
            total_format = workbook.add_format({'bold': True, 'bg_color': '#CFE2F3', 'border': 1, 'align': 'right'})

            for supplier, sheet_name in current_mapping.items():
                # จัดการชื่อชีต
                clean_name = sheet_name.strip()[:31]
                for char in r'[]:*?/\ ':
                    clean_name = clean_name.replace(char, ' ')
                
                df_sup = df_raw[df_raw['ซัพพลายเออร์'] == supplier].copy()
                
                # --- [ด้านซ้าย] แยกตามสาขา (เพิ่มคอลัมน์ชื่อสาขา) ---
                # เรียงลำดับ: รหัสสาขา, ชื่อสาขา (จาก Store Name), โซน, รหัสสินค้า, รายการสินค้า, ซัพพลายเออร์, หน่วย, จำนวน
                left_data = df_sup.rename(columns={'Store Name': 'ชื่อสาขา', 'จำนวน': 'Total'})
                left_cols = ['รหัสสาขา', 'ชื่อสาขา', 'โซน', 'รหัสสินค้า', 'รายการสินค้า', 'ซัพพลายเออร์', 'หน่วย', 'Total']
                left_display = left_data[left_cols].copy()
                
                # Logic: แสดงข้อมูลสาขาแค่บรรทัดแรกของกลุ่ม
                mask = left_display['รหัสสาขา'].duplicated()
                left_display.loc[mask, ['รหัสสาขา', 'ชื่อสาขา', 'โซน']] = ""
                
                # --- [ด้านขวา] สรุปยอด GroupBy เหมือนเดิม ---
                right_summary = df_sup.groupby(['รหัสสินค้า', 'รายการสินค้า'], as_index=False)['จำนวน'].sum()
                right_summary.rename(columns={'จำนวน': 'Total'}, inplace=True)
                right_summary.insert(0, 'ซัพพลายเออร์', "") # คอลัมน์ว่างด้านหน้า

                # --- เขียนข้อมูลลง Excel (ซ้าย-ขวา) ---
                # เขียนฝั่งซ้าย (เริ่ม A1)
                left_display.to_excel(writer, sheet_name=clean_name, index=False, startcol=0)
                # เขียนฝั่งขวา (เริ่มคอลัมน์ J เพื่อเว้นระยะจากฝั่งซ้ายเล็กน้อย)
                right_summary.to_excel(writer, sheet_name=clean_name, index=False, startcol=9)

                worksheet = writer.sheets[clean_name]
                
                # --- เพิ่มแถว Grand Total ---
                # ฝั่งซ้าย (Total อยู่คอลัมน์ที่ 8 คือ index 7)
                row_l = len(left_display) + 1
                worksheet.write(row_l, 0, "Grand Total", total_format)
                worksheet.write(row_l, 7, left_display['Total'].sum(), total_format)
                
                # ฝั่งขวา (Total อยู่คอลัมน์ L คือ index 11)
                row_r = len(right_summary) + 1
                worksheet.write(row_r, 9, f"{supplier} Total", total_format)
                worksheet.write(row_r, 11, right_summary['Total'].sum(), total_format)

                # ปรับขนาดคอลัมน์
                worksheet.set_column('A:B', 15) # รหัสสาขา, ชื่อสาขา
                worksheet.set_column('D:E', 30) # รหัสสินค้า, รายการสินค้า
                worksheet.set_column('J:K', 30) # ฝั่งขวา

        st.success("✅ ระบบกลับมาเป็นแบบ ซ้าย-ขวา พร้อมเพิ่มชื่อสาขาฝั่งซ้ายเรียบร้อยครับ!")
        st.download_button(
            label="📥 ดาวน์โหลดไฟล์ Excel V4",
            data=output.getvalue(),
            file_name="Supplier_Split_V4.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )