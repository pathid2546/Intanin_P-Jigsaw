import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Supplier Splitter V6.5", layout="wide")

st.title("📦 ระบบแยกข้อมูลซัพพลายเออร์ (True Auto-Fit)")

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
            workbook = writer.book
            total_format = workbook.add_format({'bold': True, 'bg_color': '#D9EAD3', 'border': 1, 'align': 'right'})

            for supplier, sheet_name in current_mapping.items():
                clean_name = sheet_name.strip()[:31]
                for char in r'[]:*?/\ ':
                    clean_name = clean_name.replace(char, ' ')
                
                df_sup = df_raw[df_raw['ซัพพลายเออร์'] == supplier].copy()
                
                # --- [เตรียมข้อมูลฝั่งซ้าย] ---
                left_display = df_sup.rename(columns={'Store Name': 'ชื่อสาขา', 'จำนวน': 'Total'})
                left_cols = ['รหัสสาขา', 'ชื่อสาขา', 'โซน', 'รหัสสินค้า', 'รายการสินค้า', 'ซัพพลายเออร์', 'หน่วย', 'Total']
                left_final = left_display[left_cols].copy()
                
                # เก็บค่าก่อนทำ Mask เพื่อใช้วัดความกว้างที่ยาวที่สุดจริงๆ
                width_ref_left = left_final.copy()
                
                # ทำ Masking ซ่อนค่าซ้ำตามรูป image_93fbee.png
                mask = left_final['รหัสสาขา'].duplicated()
                left_final.loc[mask, ['รหัสสาขา', 'ชื่อสาขา', 'โซน']] = ""
                
                # --- [เตรียมข้อมูลฝั่งขวา] ---
                right_summary = df_sup.groupby(['รหัสสินค้า', 'รายการสินค้า'], as_index=False)['จำนวน'].sum()
                right_summary.rename(columns={'จำนวน': 'Total'}, inplace=True)
                right_summary.insert(0, 'ซัพพลายเออร์', "")

                # เขียนข้อมูลลงชีต
                left_final.to_excel(writer, sheet_name=clean_name, index=False, startcol=0)
                right_summary.to_excel(writer, sheet_name=clean_name, index=False, startcol=10)

                worksheet = writer.sheets[clean_name]
                worksheet.set_zoom(80) 
                
                # Grand Total และ Summary Total (ตำแหน่ง Col N / index 13)
                row_l = len(left_final) + 1
                worksheet.write(row_l, 0, "Grand Total", total_format)
                worksheet.write(row_l, 7, df_sup['จำนวน'].sum(), total_format)
                
                row_r = len(right_summary) + 1
                worksheet.write(row_r, 10, f"{supplier} Total", total_format)
                worksheet.write(row_r, 13, right_summary['Total'].sum(), total_format)

                # --- 🎯 แก้ปัญหา Column กว้างเพี้ยน: วัดจากข้อมูลจริง ---
                def get_excel_width(df, col_name):
                    # หาค่าที่ยาวที่สุดในคอลัมน์ (รวม Header)
                    max_data = df[col_name].astype(str).map(len).max()
                    max_header = len(str(col_name))
                    # ใช้ค่าที่ยาวที่สุด และบวกเผื่อเล็กน้อยแค่ 2-3 unit สำหรับระยะขอบ
                    return max(max_data, max_header) + 3

                # ปรับความกว้างฝั่งซ้าย (0-7) อิงจากข้อมูลที่ยาวที่สุด
                for i, col in enumerate(left_final.columns):
                    w = get_excel_width(width_ref_left, col)
                    # ถ้าเป็นภาษาไทย (รายการสินค้า/ชื่อสาขา) ให้เผื่อสระเพิ่มแค่เล็กน้อย (x1.2)
                    if col in ['ชื่อสาขา', 'รายการสินค้า']:
                        w = w * 1.2
                    worksheet.set_column(i, i, w)
                
                # ปรับความกว้างฝั่งขวา (10-13)
                for i, col in enumerate(right_summary.columns):
                    w = get_excel_width(right_summary, col)
                    if col in ['รายการสินค้า']:
                        w = w * 1.2
                    worksheet.set_column(i+10, i+10, w)

        st.success("✅ ปรับขนาดคอลัมน์ Auto-Fit อิงตามข้อมูลที่ยาวที่สุดเรียบร้อย!")
        st.download_button(label="📥 ดาวน์โหลดไฟล์ Excel V6.5", data=output.getvalue(), file_name="Supplier_Report_Final.xlsx")