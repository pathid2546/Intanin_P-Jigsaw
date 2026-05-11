import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Supplier Splitter V6.8", layout="wide")

st.title("📦 ระบบแยกข้อมูลซัพพลายเออร์ (Full Option V6.8)")

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
            # Format สำหรับบรรทัดผลรวม
            total_style = workbook.add_format({'bold': True, 'bg_color': '#D9EAD3', 'border': 1, 'align': 'right'})

            for supplier, sheet_name in current_mapping.items():
                clean_name = sheet_name.strip()[:31]
                for char in r'[]:*?/\ ':
                    clean_name = clean_name.replace(char, ' ')
                
                df_sup = df_raw[df_raw['ซัพพลายเออร์'] == supplier].copy()
                
                # --- [ฝั่งซ้าย: รายละเอียด] ---
                left_df = df_sup.rename(columns={'จำนวน': 'Total'})
                # เรียงคอลัมน์ตาม image_8928d1.png
                cols_order = ['รหัสสาขา', 'โซน', 'รหัสสินค้า', 'รายการสินค้า', 'ซัพพลายเออร์', 'หน่วย', 'Total']
                left_final = left_df[cols_order].copy()
                
                # เก็บค่าดิบไว้ใช้คำนวณ Auto-Fit (ก่อนทำ Mask)
                raw_left_for_width = left_final.copy()
                
                # ซ่อนค่าซ้ำในคอลัมน์ รหัสสาขา, โซน
                mask = left_final['รหัสสาขา'].duplicated()
                left_final.loc[mask, ['รหัสสาขา', 'โซน']] = ""
                
                # --- [ฝั่งขวา: สรุปยอด] ---
                right_summary = df_sup.groupby(['รหัสสินค้า', 'รายการสินค้า'], as_index=False)['จำนวน'].sum()
                right_summary.rename(columns={'จำนวน': 'Total'}, inplace=True)
                right_summary.insert(0, 'ซัพพลายเออร์', "")

                # เขียนข้อมูล (ซ้ายเริ่ม A1, ขวาเริ่ม J1)
                left_final.to_excel(writer, sheet_name=clean_name, index=False, startcol=0)
                right_summary.to_excel(writer, sheet_name=clean_name, index=False, startcol=9)

                worksheet = writer.sheets[clean_name]
                worksheet.set_zoom(80) 
                
                # --- 🎯 เพิ่ม Grand Total ฝั่งซ้าย (แถวสุดท้ายของตารางซ้าย) ---
                row_l = len(left_final) + 1
                worksheet.write(row_l, 0, "Grand Total", total_style)
                worksheet.write(row_l, 6, df_sup['จำนวน'].sum(), total_style) # คอลัมน์ G (index 6)
                
                # --- 🎯 เพิ่ม Total ฝั่งขวา (แถวสุดท้ายของตารางขวา) ---
                row_r = len(right_summary) + 1
                worksheet.write(row_r, 9, f"{supplier} Total", total_style) # คอลัมน์ J (index 9)
                worksheet.write(row_r, 12, right_summary['Total'].sum(), total_style) # คอลัมน์ M (index 12)

                # --- ระบบ Auto-Fit แบบแม่นยำ ---
                def set_optimal_width(df, start_col_idx):
                    for i, col in enumerate(df.columns):
                        # หาค่าที่ยาวที่สุดในคอลัมน์ (รวม Header)
                        max_len = max(
                            df[col].astype(str).map(len).max(), 
                            len(str(col))
                        )
                        # ปรับ Factor สำหรับภาษาไทยให้พอดี ไม่กว้างเกินไป
                        is_thai = any(ord(c) > 128 for c in str(df[col].iloc[0]) if not pd.isna(df[col].iloc[0]))
                        final_w = max_len * (1.2 if is_thai else 1.0) + 2
                        worksheet.set_column(start_col_idx + i, start_col_idx + i, final_w)

                set_optimal_width(raw_left_for_width, 0)
                set_optimal_width(right_summary, 9)

        st.success("✅ สร้างไฟล์เรียบร้อย พร้อมบรรทัด Grand Total และ Auto-Fit ที่สมบูรณ์!")
        st.download_button(label="📥 ดาวน์โหลดไฟล์ Excel", data=output.getvalue(), file_name="Supplier_Split_Final_V6.8.xlsx")