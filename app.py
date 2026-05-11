import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Supplier Splitter V6.6", layout="wide")

st.title("📦 ระบบแยกข้อมูลซัพพลายเออร์ (True Auto-Fit V6.6)")

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
                
                # --- [ฝั่งซ้าย] อิงตาม image_8928d1.png ---
                left_display = df_sup.rename(columns={'จำนวน': 'Total'})
                # คอลัมน์ที่ต้องการแสดงฝั่งซ้าย
                left_cols = ['รหัสสาขา', 'โซน', 'รหัสสินค้า', 'รายการสินค้า', 'ซัพพลายเออร์', 'หน่วย', 'Total']
                left_final = left_display[left_cols].copy()
                
                # เก็บค่าดิบไว้คำนวณความกว้างก่อนทำ Mask
                width_ref_left = left_final.copy()
                
                mask = left_final['รหัสสาขา'].duplicated()
                left_final.loc[mask, ['รหัสสาขา', 'โซน']] = ""
                
                # --- [ฝั่งขวา] อิงตาม image_8a19f0.png ---
                right_summary = df_sup.groupby(['รหัสสินค้า', 'รายการสินค้า'], as_index=False)['จำนวน'].sum()
                right_summary.rename(columns={'จำนวน': 'Total'}, inplace=True)
                right_summary.insert(0, 'ซัพพลายเออร์', "")

                # เขียนข้อมูล (ซ้ายเริ่ม A1, ขวาเริ่ม J1)
                left_final.to_excel(writer, sheet_name=clean_name, index=False, startcol=0)
                right_summary.to_excel(writer, sheet_name=clean_name, index=False, startcol=9)

                worksheet = writer.sheets[clean_name]
                worksheet.set_zoom(80) 
                
                # Total ขวา (Col M / index 12) ตามภาพ image_8928d1.png
                row_r = len(right_summary) + 1
                worksheet.write(row_r, 9, f"{supplier} Total", total_format)
                worksheet.write(row_r, 12, right_summary['Total'].sum(), total_format)

                # --- 🎯 ระบบ Auto-Fit อิงจากข้อมูลที่ยาวที่สุดจริง ---
                def apply_compact_autofit(df, start_col):
                    for i, col in enumerate(df.columns):
                        # หาความยาวที่ยาวที่สุดใน Column (รวม Header) และจัดการค่าว่าง
                        max_len = max(
                            df[col].astype(str).str.len().max(),
                            len(str(col))
                        )
                        
                        # ปรับ Factor เล็กน้อยตามประเภทข้อมูล
                        if any(ord(char) > 128 for char in str(df[col].iloc[0]) if not pd.isna(df[col].iloc[0])): # เช็คว่าเป็นภาษาไทยไหม
                            final_w = max_len * 1.15
                        else:
                            final_w = max_len + 2
                            
                        worksheet.set_column(start_col + i, start_col + i, final_w)

                apply_compact_autofit(width_ref_left, 0)
                apply_compact_autofit(right_summary, 9)

        st.success("✅ แก้ไขให้ตรงตามรูปแบบที่คุณต้องการ (V6.6) เรียบร้อยแล้ว!")
        st.download_button(label="📥 ดาวน์โหลดไฟล์ Excel", data=output.getvalue(), file_name="Supplier_Split_Final.xlsx")