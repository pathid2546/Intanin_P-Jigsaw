import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Supplier Splitter V5.1", layout="wide")

st.title("📦 ระบบแยกข้อมูลซัพพลายเออร์ (Fixed Auto-Fit)")

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
                left_data = df_sup.rename(columns={'Store Name': 'ชื่อสาขา', 'จำนวน': 'Total'})
                left_cols = ['รหัสสาขา', 'ชื่อสาขา', 'โซน', 'รหัสสินค้า', 'รายการสินค้า', 'ซัพพลายเออร์', 'หน่วย', 'Total']
                left_display = left_data[left_cols].copy()
                
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
                
                # Grand Total ซ้าย
                row_l = len(left_display) + 1
                worksheet.write(row_l, 0, "Grand Total", total_format)
                worksheet.write(row_l, 7, left_display['Total'].sum(), total_format)
                
                # Total ขวา (เลื่อนไปช่องที่ 11)
                row_r = len(right_summary) + 1
                worksheet.write(row_r, 9, f"{supplier} Total", total_format)
                worksheet.write(row_r, 11, right_summary['Total'].sum(), total_format)

                # --- แก้ไขส่วน Auto-fit (เพิ่ม Error Handling) ---
                # ฝั่งซ้าย (Col 0-7)
                for i, col in enumerate(left_display.columns):
                    column_data = left_display[col].astype(str)
                    # หาค่าความยาวสูงสุด ถ้าไม่มีข้อมูลให้ใช้ความยาวชื่อคอลัมน์แทน
                    max_len = column_data.map(len).max() if not column_data.empty else 0
                    max_len = max(max_len, len(str(col))) + 5
                    worksheet.set_column(i, i, min(max_len, 50)) # จำกัดความกว้างสูงสุดไม่เกิน 50
                
                # ฝั่งขวา (Col 9-11)
                for i, col in enumerate(right_summary.columns):
                    column_data = right_summary[col].astype(str)
                    max_len = column_data.map(len).max() if not column_data.empty else 0
                    max_len = max(max_len, len(str(col))) + 5
                    worksheet.set_column(i+9, i+9, min(max_len, 50))

        st.success("✅ แก้ไข Error เรียบร้อย! ลองดาวน์โหลดไฟล์อีกครั้งครับ")
        st.download_button(
            label="📥 ดาวน์โหลดไฟล์ Excel",
            data=output.getvalue(),
            file_name="Supplier_Split_Fixed.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )