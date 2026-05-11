import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Supplier Splitter V5.3", layout="wide")

st.title("📦 ระบบแยกข้อมูลซัพพลายเออร์ (Compact Auto-Fit)")

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
                clean_name = sheet_name.strip()[:31]
                for char in r'[]:*?/\ ':
                    clean_name = clean_name.replace(char, ' ')
                
                df_sup = df_raw[df_raw['ซัพพลายเออร์'] == supplier].copy()
                
                # --- ฝั่งซ้าย ---
                # เพิ่ม 'Store Name' เปลี่ยนเป็น 'ชื่อสาขา'
                left_data = df_sup.rename(columns={'Store Name': 'ชื่อสาขา', 'จำนวน': 'Total'})
                left_cols = ['รหัสสาขา', 'ชื่อสาขา', 'โซน', 'รหัสสินค้า', 'รายการสินค้า', 'ซัพพลายเออร์', 'หน่วย', 'Total']
                left_display = left_data[left_cols].copy()
                
                # Mask ข้อมูลที่ซ้ำ
                mask = left_display['รหัสสาขา'].duplicated()
                left_display.loc[mask, ['รหัสสาขา', 'ชื่อสาขา', 'โซน']] = ""
                
                # --- ฝั่งขวา ---
                right_summary = df_sup.groupby(['รหัสสินค้า', 'รายการสินค้า'], as_index=False)['จำนวน'].sum()
                right_summary.rename(columns={'จำนวน': 'Total'}, inplace=True)
                right_summary.insert(0, 'ซัพพลายเออร์', "") 

                # เขียนข้อมูล
                left_display.to_excel(writer, sheet_name=clean_name, index=False, startcol=0)
                right_summary.to_excel(writer, sheet_name=clean_name, index=False, startcol=9)

                worksheet = writer.sheets[clean_name]
                worksheet.set_zoom(80) 
                
                # Grand Total ซ้าย (Col H / index 7)
                row_l = len(left_display) + 1
                worksheet.write(row_l, 0, "Grand Total", total_format)
                worksheet.write(row_l, 7, left_display['Total'].sum(), total_format)
                
                # Total ขวา (Col L / index 11) - ปรับตำแหน่งให้ตรงตารางสรุป
                row_r = len(right_summary) + 1
                worksheet.write(row_r, 9, f"{supplier} Total", total_format)
                worksheet.write(row_r, 11, right_summary['Total'].sum(), total_format)

                # --- ระบบ Auto-fit แบบ Compact (อิงตามตัวยาวสุด) ---
                # ฝั่งซ้าย (0-7)
                for i, col in enumerate(left_display.columns):
                    # หาค่าความยาวสูงสุดจากข้อมูลในคอลัมน์ (แปลงเป็น string ก่อน)
                    max_val_len = left_display[col].astype(str).str.len().max()
                    header_len = len(str(col))
                    # ใช้ค่าที่ยาวที่สุด และบวกเผื่อแค่ 2 (สำหรับระยะห่างขอบเล็กน้อย)
                    final_width = max(max_val_len, header_len) + 2
                    worksheet.set_column(i, i, final_width)
                
                # ฝั่งขวา (9-11)
                for i, col in enumerate(right_summary.columns):
                    max_val_len = right_summary[col].astype(str).str.len().max()
                    header_len = len(str(col))
                    final_width = max(max_val_len, header_len) + 2
                    worksheet.set_column(i+9, i+9, final_width)

        st.success("✅ ปรับปรุงความกว้างคอลัมน์ให้กระชับอิงตามข้อมูลจริงเรียบร้อย!")
        st.download_button(
            label="📥 ดาวน์โหลดไฟล์ Excel V5.3",
            data=output.getvalue(),
            file_name="Supplier_Split_V5_3.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )