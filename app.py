import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Supplier Splitter V6.2", layout="wide")

st.title("📦 ระบบแยกข้อมูลซัพพลายเออร์ (Auto-Column Width Fix)")

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
                
                # --- [ฝั่งซ้าย] โครงสร้างตาม image_898a4c.png ---
                left_display = df_sup.rename(columns={'Store Name': 'ชื่อสาขา', 'จำนวน': 'Total'})
                left_cols = ['รหัสสาขา', 'ชื่อสาขา', 'โซน', 'รหัสสินค้า', 'รายการสินค้า', 'ซัพพลายเออร์', 'หน่วย', 'Total']
                left_final = left_display[left_cols].copy()
                
                # เก็บค่าต้นฉบับไว้คำนวณความกว้างก่อนทำ Masking
                raw_left_for_width = left_final.copy()
                
                mask = left_final['รหัสสาขา'].duplicated()
                left_final.loc[mask, ['รหัสสาขา', 'ชื่อสาขา', 'โซน']] = ""
                
                # --- [ฝั่งขวา] โครงสร้างตาม image_8a19f0.png ---
                right_summary = df_sup.groupby(['รหัสสินค้า', 'รายการสินค้า'], as_index=False)['จำนวน'].sum()
                right_summary.rename(columns={'จำนวน': 'Total'}, inplace=True)
                right_summary.insert(0, 'ซัพพลายเออร์', "")

                # เขียนข้อมูล (ซ้าย A1, ขวา K1)
                left_final.to_excel(writer, sheet_name=clean_name, index=False, startcol=0)
                right_summary.to_excel(writer, sheet_name=clean_name, index=False, startcol=10)

                worksheet = writer.sheets[clean_name]
                worksheet.set_zoom(80) 
                
                # Grand Total ซ้าย (Col H)
                row_l = len(left_final) + 1
                worksheet.write(row_l, 0, "Grand Total", total_format)
                worksheet.write(row_l, 7, df_sup['จำนวน'].sum(), total_format)
                
                # Total ขวา (Col N / index 13) ให้ตรงกับยอดสินค้า
                row_r = len(right_summary) + 1
                worksheet.write(row_r, 10, f"{supplier} Total", total_format)
                worksheet.write(row_r, 13, right_summary['Total'].sum(), total_format)

                # --- 🎯 แก้ไขจุดสำคัญ: ระบบ Auto-Fit อิงตามข้อมูลยาวที่สุดจริง ---
                def set_optimal_width(df, start_col_idx):
                    for i, col in enumerate(df.columns):
                        # หาค่าที่ยาวที่สุดในคอลัมน์ (รวม Header)
                        # ใช้ Logic เผื่อสระภาษาไทยโดยการคูณ 1.2 และบวกเพิ่มเล็กน้อย
                        max_len = max(
                            df[col].astype(str).map(len).max(), 
                            len(str(col))
                        ) + 2
                        # สำหรับคอลัมน์ภาษาไทย (รายการสินค้า, ชื่อสาขา) ให้กว้างขึ้นอีกนิด
                        if col in ['ชื่อสาขา', 'รายการสินค้า', 'ซัพพลายเออร์']:
                            max_len = max_len * 1.5 
                        
                        worksheet.set_column(start_col_idx + i, start_col_idx + i, max_len)

                # สั่งรัน Auto-Fit ทั้งสองฝั่ง
                set_optimal_width(raw_left_for_width, 0)
                set_optimal_width(right_summary, 10)

        st.success("✅ แก้ไขระบบ Auto-Column Width อิงตามข้อมูลที่ยาวที่สุดเรียบร้อย!")
        st.download_button(
            label="📥 ดาวน์โหลดไฟล์ Excel V6.2",
            data=output.getvalue(),
            file_name="Supplier_Report_Final.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )