import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Supplier Splitter V6.7", layout="wide")

st.title("📦 ระบบแยกข้อมูลซัพพลายเออร์ (Restore Grand Total V6.7)")

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
            # สร้าง Format สำหรับบรรทัด Total ให้สวยงามตามรูป
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
                
                # --- [ฝั่งซ้าย] ---
                left_display = df_sup.rename(columns={'จำนวน': 'Total'})
                left_cols = ['รหัสสาขา', 'โซน', 'รหัสสินค้า', 'รายการสินค้า', 'ซัพพลายเออร์', 'หน่วย', 'Total']
                left_final = left_display[left_cols].copy()
                
                # เก็บไว้คำนวณความกว้าง
                width_ref_left = left_final.copy()
                
                # ทำ Mask ซ่อนค่าซ้ำ
                mask = left_final['รหัสสาขา'].duplicated()
                left_final.loc[mask, ['รหัสสาขา', 'โซน']] = ""
                
                # --- [ฝั่งขวา] ---
                right_summary = df_sup.groupby(['รหัสสินค้า', 'รายการสินค้า'], as_index=False)['จำนวน'].sum()
                right_summary.rename(columns={'จำนวน': 'Total'}, inplace=True)
                right_summary.insert(0, 'ซัพพลายเออร์', "")

                # เขียนข้อมูล
                left_final.to_excel(writer, sheet_name=clean_name, index=False, startcol=0)
                right_summary.to_excel(writer, sheet_name=clean_name, index=False, startcol=9)

                worksheet = writer.sheets[clean_name]
                worksheet.set_zoom(80) 
                
                # --- 🎯 [แก้ไข] นำ Grand Total ฝั่งซ้ายกลับมา ---
                last_row_l = len(left_final) + 1
                worksheet.write(last_row_l, 0, "Grand Total", total_format)
                # เขียนยอดรวมที่คอลัมน์ Total (index 6 เพราะเริ่ม 0)
                worksheet.write(last_row_l, 6, df_sup['จำนวน'].sum(), total_format)
                
                # --- [ฝั่งขวา] Total สรุป ---
                last_row_r = len(right_summary) + 1
                worksheet.write(last_row_r, 9, f"{supplier} Total", total_format)
                # เขียนยอดรวมที่คอลัมน์ Total ของฝั่งขวา (เริ่มที่ J=9 ดังนั้น M=12)
                worksheet.write(last_row_r, 12, right_summary['Total'].sum(), total_format)

                # --- ระบบ Auto-Fit (เน้นความกะทัดรัดตามรูป) ---
                def apply_final_autofit(df, start_col):
                    for i, col in enumerate(df.columns):
                        max_data = df[col].astype(str).str.len().max()
                        max_header = len(str(col))
                        base_w = max(max_data, max_header)
                        
                        # เผื่อสระภาษาไทยเล็กน้อย
                        if any(ord(c) > 128 for c in str(df[col].iloc[0]) if not pd.isna(df[col].iloc[0])):
                            final_w = base_w * 1.2
                        else:
                            final_w = base_w + 2
                        
                        worksheet.set_column(start_col + i, start_col + i, final_w)

                apply_final_autofit(width_ref_left, 0)
                apply_final_autofit(right_summary, 9)

        st.success("✅ นำบรรทัด Grand Total กลับมาและปรับขนาดคอลัมน์เรียบร้อย!")
        st.download_button(label="📥 ดาวน์โหลดไฟล์ Excel V6.7", data=output.getvalue(), file_name="Supplier_Split_Fixed_Final.xlsx")