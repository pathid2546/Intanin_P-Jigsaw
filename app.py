import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Supplier Splitter V3", layout="wide")

st.title("📦 ระบบแยกข้อมูลซัพพลายเออร์ (โครงสร้างคอลัมน์ใหม่)")

if 'name_memory' not in st.session_state:
    st.session_state['name_memory'] = {}

uploaded_file = st.file_uploader("อัปโหลดไฟล์ Excel (ที่มีชีต Transport)", type=['xlsx'])

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
                
                # --- จัดเรียงลำดับคอลัมน์ใหม่ตามที่ระบุ ---
                # ดึง Store Name มาใช้เป็น 'ชื่อสาขา'
                df_sup = df_sup.rename(columns={'Store Name': 'ชื่อสาขา', 'จำนวน': 'Total'})
                
                # เรียงลำดับ: รหัสสาขา, ชื่อสาขา, โซน, รหัสสินค้า, รายการสินค้า และตามด้วยคอลัมน์อื่นๆ ที่เหลือ
                fixed_cols = ['รหัสสาขา', 'ชื่อสาขา', 'โซน', 'รหัสสินค้า', 'รายการสินค้า']
                remaining_cols = [c for c in df_sup.columns if c not in fixed_cols and c != 'ซัพพลายเออร์']
                final_cols = fixed_cols + remaining_cols
                
                display_data = df_sup[final_cols].copy()
                
                # Logic: แสดง 'รหัสสาขา', 'ชื่อสาขา', 'โซน' แค่บรรทัดแรกของกลุ่มสาขาเดิม
                # ตรวจสอบการซ้ำโดยใช้ รหัสสาขา เป็นหลัก
                mask = display_data['รหัสสาขา'].duplicated()
                display_data.loc[mask, ['รหัสสาขา', 'ชื่อสาขา', 'โซน']] = ""

                # --- เขียนข้อมูล (มีแค่ด้านซ้าย) ---
                display_data.to_excel(writer, sheet_name=clean_name, index=False)

                worksheet = writer.sheets[clean_name]
                
                # --- เพิ่มแถว Grand Total ที่คอลัมน์สุดท้ายของข้อมูล ---
                row_idx = len(display_data) + 1
                last_col_idx = len(final_cols) - 1 # ตำแหน่งคอลัมน์ Total (ถ้าอยู่ท้ายสุด)
                
                worksheet.write(row_idx, 0, "Grand Total", total_format)
                # หาตำแหน่งคอลัมน์ชื่อ 'Total' เพื่อวางผลรวมให้ตรงช่อง
                total_col_pos = final_cols.index('Total') if 'Total' in final_cols else last_col_idx
                worksheet.write(row_idx, total_col_pos, df_sup['Total'].sum(), total_format)

                # ปรับขนาดคอลัมน์อัตโนมัติ
                for i, col in enumerate(final_cols):
                    worksheet.set_column(i, i, 18)

        st.success("✅ จัดเรียงคอลัมน์ใหม่เรียบร้อย (รหัสสาขา -> ชื่อสาขา -> โซน...)")
        st.download_button(
            label="📥 ดาวน์โหลดไฟล์ Excel V3",
            data=output.getvalue(),
            file_name="Supplier_Report_V3.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )