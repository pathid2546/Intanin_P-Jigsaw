import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Supplier Splitter V6.7.1", layout="wide")

st.title("📦 ระบบแยกข้อมูลซัพพลายเออร์ (Base V6.7 + ชื่อสาขา)")

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
            # สร้าง Format สำหรับบรรทัด Total
            total_format = workbook.add_format({
                'bold': True, 
                'bg_color': '#D9EAD3', 
                'border': 1, 
                'align': 'right'
            })
            header_format = workbook.add_format({
                'bold': True,
                'bg_color': '#CFE2F3',
                'border': 1
            })

            for supplier, sheet_name in current_mapping.items():
                clean_name = sheet_name.strip()[:31]
                for char in r'[]:*?/\ ':
                    clean_name = clean_name.replace(char, ' ')
                
                df_sup = df_raw[df_raw['ซัพพลายเออร์'] == supplier].copy()
                
                # --- [ฝั่งซ้าย] เพิ่ม 'Store Name' กลับมา ---
                left_display = df_sup.rename(columns={'Store Name': 'ชื่อสาขา', 'จำนวน': 'Total'})
                # เรียงคอลัมน์ใหม่: รหัสสาขา, ชื่อสาขา, โซน, รหัสสินค้า, รายการสินค้า, ซัพพลายเออร์, หน่วย, Total
                left_cols = ['รหัสสาขา', 'ชื่อสาขา', 'โซน', 'รหัสสินค้า', 'รายการสินค้า', 'ซัพพลายเออร์', 'หน่วย', 'Total']
                left_final = left_display[left_cols].copy()
                
                # เก็บค่าดิบไว้คำนวณความกว้าง (ให้ชื่อสาขาไม่โดนตัด)
                width_ref_left = left_final.copy()
                
                # ทำ Mask ซ่อนค่าซ้ำ (รหัสสาขา, ชื่อสาขา, โซน)
                mask = left_final['รหัสสาขา'].duplicated()
                left_final.loc[mask, ['รหัสสาขา', 'ชื่อสาขา', 'โซน']] = ""
                
                # --- [ฝั่งขวา] ตารางสรุปสินค้า ---
                right_summary = df_sup.groupby(['รหัสสินค้า', 'รายการสินค้า'], as_index=False)['จำนวน'].sum()
                right_summary.rename(columns={'จำนวน': 'Total'}, inplace=True)
                right_summary.insert(0, 'ซัพพลายเออร์', "")

                # เขียนข้อมูล (ฝั่งซ้ายเริ่มคอลัมน์ A, ฝั่งขวาขยับไปเริ่มคอลัมน์ K เพื่อไม่ให้เบียดกัน)
                left_final.to_excel(writer, sheet_name=clean_name, index=False, startcol=0)
                right_summary.to_excel(writer, sheet_name=clean_name, index=False, startcol=10)

                worksheet = writer.sheets[clean_name]
                worksheet.set_zoom(80) 
                
                # --- 🎯 Grand Total ฝั่งซ้าย ---
                last_row_l = len(left_final) + 1
                worksheet.write(last_row_l, 0, "Grand Total", total_format)
                # เขียนยอดรวมที่คอลัมน์ Total (ตอนนี้คือ index 7 เพราะเพิ่มชื่อสาขามา)
                worksheet.write(last_row_l, 7, df_sup['จำนวน'].sum(), total_format)
                
                # --- [ฝั่งขวา] Total สรุป ---
                last_row_r = len(right_summary) + 1
                # เริ่มที่คอลัมน์ K=10 ดังนั้นยอดรวมขยับไปที่ N=13
                worksheet.write(last_row_r, 10, f"{supplier} Total", total_format)
                worksheet.write(last_row_r, 13, right_summary['Total'].sum(), total_format)

                # --- ระบบ Auto-Fit (เน้นความกะทัดรัด) ---
                def apply_final_autofit(df, start_col):
                    for i, col in enumerate(df.columns):
                        # ป้องกัน Error กรณีข้อมูลว่าง
                        max_data = df[col].astype(str).str.len().max() if not df[col].empty else 0
                        max_header = len(str(col))
                        base_w = max(max_data, max_header)
                        
                        # เช็คภาษาไทยเพื่อเพิ่มระยะขอบ
                        sample_val = str(df[col].iloc[0]) if not df[col].empty else ""
                        if any(ord(c) > 128 for c in sample_val):
                            final_w = base_w * 1.25 # ปรับตัวคูณให้ชื่อสาขาไม่แน่นเกินไป
                        else:
                            final_w = base_w + 2
                        
                        worksheet.set_column(start_col + i, start_col + i, final_w)

                apply_final_autofit(width_ref_left, 0)
                apply_final_autofit(right_summary, 10)

        st.success("✅ เพิ่มชื่อสาขาและจัดเรียงข้อมูลตามโครงสร้างหลักเรียบร้อย!")
        st.download_button(label="📥 ดาวน์โหลดไฟล์ Excel", data=output.getvalue(), file_name="Supplier_Split_Final_WithStore.xlsx")