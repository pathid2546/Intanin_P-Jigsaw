import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Supplier Splitter V6.4", layout="wide")

st.title("📦 ระบบแยกข้อมูลซัพพลายเออร์ (Auto-Fit แก้ไขสมบูรณ์)")

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
                
                # --- [ฝั่งซ้าย] ---
                left_display = df_sup.rename(columns={'Store Name': 'ชื่อสาขา', 'จำนวน': 'Total'})
                left_cols = ['รหัสสาขา', 'ชื่อสาขา', 'โซน', 'รหัสสินค้า', 'รายการสินค้า', 'ซัพพลายเออร์', 'หน่วย', 'Total']
                left_final = left_display[left_cols].copy()
                
                # เก็บค่าดิบไว้คำนวณความกว้าง (ก่อน Mask)
                width_ref_left = left_final.copy()
                
                mask = left_final['รหัสสาขา'].duplicated()
                left_final.loc[mask, ['รหัสสาขา', 'ชื่อสาขา', 'โซน']] = ""
                
                # --- [ฝั่งขวา] ---
                right_summary = df_sup.groupby(['รหัสสินค้า', 'รายการสินค้า'], as_index=False)['จำนวน'].sum()
                right_summary.rename(columns={'จำนวน': 'Total'}, inplace=True)
                right_summary.insert(0, 'ซัพพลายเออร์', "")

                # เขียนข้อมูล (ซ้ายเริ่ม A1, ขวาเริ่ม K1)
                left_final.to_excel(writer, sheet_name=clean_name, index=False, startcol=0)
                right_summary.to_excel(writer, sheet_name=clean_name, index=False, startcol=10)

                worksheet = writer.sheets[clean_name]
                worksheet.set_zoom(80) 
                
                # Grand Total ซ้าย
                row_l = len(left_final) + 1
                worksheet.write(row_l, 0, "Grand Total", total_format)
                worksheet.write(row_l, 7, df_sup['จำนวน'].sum(), total_format)
                
                # Total ขวา (Col N / index 13)
                row_r = len(right_summary) + 1
                worksheet.write(row_r, 10, f"{supplier} Total", total_format)
                worksheet.write(row_r, 13, right_summary['Total'].sum(), total_format)

                # --- 🎯 ระบบ Auto-Fit ใหม่: อิงตามค่าที่ยาวที่สุดจริง ---
                def apply_real_auto_width(df, start_col):
                    for i, col in enumerate(df.columns):
                        # 1. หาค่าที่ยาวที่สุดในคอลัมน์ (แปลงเป็น string และจัดการค่าว่าง)
                        # 2. คำนวณความยาว Header ด้วย
                        max_data_len = df[col].astype(str).str.len().max()
                        header_len = len(str(col))
                        
                        # ใช้ตัวที่ยาวกว่าเป็นบรรทัดฐาน
                        base_len = max(max_data_len, header_len)
                        
                        # ปรับตัวคูณให้เหมาะสม (1.2 สำหรับภาษาอังกฤษ/ตัวเลข, 1.4 สำหรับภาษาไทย)
                        # เพื่อให้คอลัมน์ "กระชับ" ไม่ใหญ่เกินไปเหมือนในรูป
                        if col in ['ชื่อสาขา', 'รายการสินค้า', 'ซัพพลายเออร์']:
                            actual_width = base_len * 1.4
                        else:
                            actual_width = base_len + 2
                            
                        # ตั้งค่าความกว้างคอลัมน์โดยไม่ใช้ค่า Min/Max ที่เป็นตัวเลขคงที่
                        worksheet.set_column(start_col + i, start_col + i, actual_width)

                apply_real_auto_width(width_ref_left, 0)
                apply_real_auto_width(right_summary, 10)

        st.success("✅ แก้ไขระบบ Auto-Fit ให้กระชับตามข้อมูลจริงเรียบร้อย!")
        st.download_button(label="📥 ดาวน์โหลดไฟล์ Excel V6.4", data=output.getvalue(), file_name="Supplier_Report_AutoFit.xlsx")