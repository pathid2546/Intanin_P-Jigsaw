import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Supplier & DO System V7.0", layout="wide")

# --- ส่วนหัวของโปรแกรม ---
st.title("📦 ระบบจัดการข้อมูลขนส่ง (Splitter & DO Generator)")
st.markdown("---")

if 'name_memory' not in st.session_state:
    st.session_state['name_memory'] = {}

uploaded_file = st.file_uploader("อัปโหลดไฟล์ Excel (Transport)", type=['xlsx'])

if uploaded_file:
    # อ่านข้อมูลจาก Sheet 'Transport' เป็นหลัก
    df_raw = pd.read_excel(uploaded_file, sheet_name='Transport')
    
    # สร้าง Tab เพื่อแยกการทำงาน
    tab1, tab2 = st.tabs(["✂️ แยกข้อมูลซัพพลายเออร์ (V6.7.1)", "📄 สร้างใบส่งสินค้า (GENPrint DO)"])

    # --- TAB 1: ระบบเดิม V6.7.1 ---
    with tab1:
        df_split = df_raw.dropna(subset=['ซัพพลายเออร์'])
        unique_suppliers = sorted(df_split['ซัพพลายเออร์'].unique())

        st.subheader("📝 กำหนดชื่อตัวย่อชีต")
        with st.form("sheet_name_form"):
            cols = st.columns(3)
            current_mapping = {}
            for i, supplier in enumerate(unique_suppliers):
                with cols[i % 3]:
                    remembered_name = st.session_state['name_memory'].get(supplier, str(supplier)[:10].strip())
                    current_mapping[supplier] = st.text_input(f"{supplier}:", value=remembered_name, key=f"input_{supplier}")
            submit_split = st.form_submit_button("สร้างไฟล์แยกซัพพลายเออร์")

        if submit_split:
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                workbook = writer.book
                total_format = workbook.add_format({'bold': True, 'bg_color': '#D9EAD3', 'border': 1, 'align': 'right'})
                
                for supplier, sheet_name in current_mapping.items():
                    # บันทึกชื่อจำไว้ใน Session
                    st.session_state['name_memory'][supplier] = sheet_name
                    
                    clean_name = "".join([c if c not in r'[]:*?/\ ' else ' ' for c in sheet_name.strip()[:31]])
                    df_sup = df_split[df_split['ซัพพลายเออร์'] == supplier].copy()
                    
                    # ฝั่งซ้าย (มีชื่อสาขา)
                    left_display = df_sup.rename(columns={'Store Name': 'ชื่อสาขา', 'จำนวน': 'Total'})
                    left_cols = ['รหัสสาขา', 'ชื่อสาขา', 'โซน', 'รหัสสินค้า', 'รายการสินค้า', 'ซัพพลายเออร์', 'หน่วย', 'Total']
                    left_final = left_display[left_cols].copy()
                    width_ref_left = left_final.copy()
                    
                    mask = left_final['รหัสสาขา'].duplicated()
                    left_final.loc[mask, ['รหัสสาขา', 'ชื่อสาขา', 'โซน']] = ""
                    
                    # ฝั่งขวา
                    right_summary = df_sup.groupby(['รหัสสินค้า', 'รายการสินค้า'], as_index=False)['จำนวน'].sum()
                    right_summary.rename(columns={'จำนวน': 'Total'}, inplace=True)
                    right_summary.insert(0, 'ซัพพลายเออร์', "")

                    left_final.to_excel(writer, sheet_name=clean_name, index=False, startcol=0)
                    right_summary.to_excel(writer, sheet_name=clean_name, index=False, startcol=10)

                    worksheet = writer.sheets[clean_name]
                    worksheet.set_zoom(80)
                    worksheet.write(len(left_final) + 1, 0, "Grand Total", total_format)
                    worksheet.write(len(left_final) + 1, 7, df_sup['จำนวน'].sum(), total_format)
                    worksheet.write(len(right_summary) + 1, 10, f"{supplier} Total", total_format)
                    worksheet.write(len(right_summary) + 1, 13, right_summary['Total'].sum(), total_format)

            st.success("✅ สร้างไฟล์แยกซัพพลายเออร์เรียบร้อย!")
            st.download_button(label="📥 ดาวน์โหลดไฟล์ Splitter", data=output.getvalue(), file_name="Supplier_Splitter_V7.xlsx")

# --- TAB 2: รวมทุกสาขาแบบต่อเนื่อง (Multiple DOs per A4 Page) ---
    with tab2:
        st.subheader("📑 ออกใบส่งสินค้า (ฉบับสมบูรณ์ - ต่อเนื่องในหน้าเดียว)")
        
        df_clean = df_raw.dropna(subset=['รหัสสาขา']).copy()
        
        if st.button("🚀 สร้างไฟล์ใบส่งสินค้า (Continuous Layout)"):
            output_do = io.BytesIO()
            with pd.ExcelWriter(output_do, engine='xlsxwriter') as writer:
                workbook = writer.book
                worksheet = workbook.add_worksheet("DO_Continuous")
                
                # --- [Styles] ---
                comp_name_fmt = workbook.add_format({'bold': True, 'font_size': 11})
                doc_title_fmt = workbook.add_format({'bold': True, 'font_size': 16, 'align': 'center'})
                std_fmt = workbook.add_format({'font_size': 9})
                std_right = workbook.add_format({'font_size': 9, 'align': 'right'})
                
                # Ship To แบบ Merge B:C ตามที่ตกลงกันไว้
                address_fmt = workbook.add_format({'font_size': 9, 'text_wrap': True, 'valign': 'top'})
                
                # Delivery Date ตัวหนาตามตัวอย่าง
                delivery_fmt = workbook.add_format({'bold': True, 'font_size': 10, 'align': 'right'})
                date_val_fmt = workbook.add_format({'bold': True, 'font_size': 10, 'num_format': 'dd/mm/yyyy', 'align': 'right'})
                
                table_head = workbook.add_format({'border': 1, 'align': 'center', 'bold': True, 'font_size': 9})
                border_left = workbook.add_format({'border': 1, 'font_size': 9, 'text_wrap': True})
                border_center = workbook.add_format({'border': 1, 'font_size': 9, 'align': 'center'})
                footer_box = workbook.add_format({'border': 1, 'font_size': 8, 'valign': 'top', 'text_wrap': True})

                # ตั้งค่าคอลัมน์
                worksheet.set_column('A:A', 4)   # No.
                worksheet.set_column('B:B', 12)  # Code
                worksheet.set_column('C:C', 40)  # Name
                worksheet.set_column('D:D', 10)  # Unit
                worksheet.set_column('E:E', 12)  # QTY
                
                worksheet.set_paper(9) # A4
                worksheet.set_margins(0.3, 0.3, 0.3, 0.3)

                current_row = 0
                
                for store_code in df_clean['รหัสสาขา'].unique():
                    df_store = df_clean[df_clean['รหัสสาขา'] == store_code].copy()
                    first_row = df_store.iloc[0]

                    # --- ตรวจสอบพื้นที่หน้ากระดาษ (ถ้าเหลือไม่พอสำหรับ Header + Data 1-2 แถว ให้ขึ้นหน้าใหม่) ---
                    # ประมาณการว่า 1 DO ใช้พื้นที่อย่างน้อย 15-20 แถว
                    if current_row > 35: 
                        worksheet.set_h_pagebreaks([current_row])
                        current_row += 1 

                    # --- [Header Section] ---
                    worksheet.write(current_row, 0, 'บริษัท โมบาย โลจิสติกส์ จำกัด', comp_name_fmt)
                    worksheet.write(current_row + 1, 0, '279 หมู่ที่ 9 ตำบลบางโฉลง...', std_fmt)
                    # Title ที่ C2 ตามสั่ง
                    worksheet.write(current_row + 1, 2, 'ใบส่งสินค้าชั่วคราว', doc_title_fmt) 
                    
                    # ข้อมูลฝั่งขวา
                    worksheet.write(current_row, 4, f'Do. No. {first_row["เลขที่ DO."]}', std_right)
                    worksheet.write(current_row + 6, 3, 'Delivery Date', delivery_fmt)
                    worksheet.write(current_row + 6, 4, first_row['Delivery Date'], date_val_fmt)

                    # --- [Customer Section] ---
                    worksheet.write(current_row + 4, 0, f'Store Code {store_code}', std_fmt)
                    worksheet.write(current_row + 5, 0, f'Store Name {first_row["Store Name"]}', std_fmt)
                    worksheet.write(current_row + 6, 0, 'Ship To', std_fmt)
                    # Merge B:C สำหรับที่อยู่
                    worksheet.merge_range(current_row + 6, 1, current_row + 7, 2, first_row['ที่อยู่'], address_fmt)

                    # --- [Table Section] ---
                    tbl_header_row = current_row + 9
                    headers = ['No.', 'Product Code', 'Product Name', 'Unit/UOM', 'QTY']
                    for col, h in enumerate(headers):
                        worksheet.write(tbl_header_row, col, h, table_head)
                    
                    r_idx = tbl_header_row + 1
                    for i, (_, r) in enumerate(df_store.iterrows(), 1):
                        worksheet.write(r_idx, 0, i, border_center)
                        worksheet.write(r_idx, 1, r['รหัสสินค้า'], border_center)
                        worksheet.write(r_idx, 2, r['รายการสินค้า'], border_left)
                        worksheet.write(r_idx, 3, r['หน่วย'], border_center)
                        worksheet.write(r_idx, 4, r['จำนวน'], border_center)
                        r_idx += 1
                    
                    # Total
                    worksheet.merge_range(r_idx, 0, r_idx, 3, 'Total', table_head)
                    worksheet.write(r_idx, 4, df_store['จำนวน'].sum(), border_center)
                    
                    # --- [Footer Section] ---
                    f_row = r_idx + 1
                    worksheet.set_row(f_row + 1, 40) # ปรับความสูงช่องเซ็นชื่อ
                    worksheet.merge_range(f_row, 0, f_row, 1, 'ผู้รับสินค้า', table_head)
                    worksheet.merge_range(f_row+1, 0, f_row+1, 1, 'ชื่อ...\nวันที่...', footer_box)
                    worksheet.merge_range(f_row, 2, f_row, 3, 'ผู้ส่งสินค้า', table_head)
                    worksheet.merge_range(f_row+1, 2, f_row+1, 3, 'ชื่อ...\nวันที่...', footer_box)
                    worksheet.write(f_row, 4, 'คลังสินค้า', table_head)
                    worksheet.write(f_row+1, 4, 'ชื่อ...\nวันที่...', footer_box)

                    # เว้นระยะห่างก่อนเริ่ม DO ถัดไป
                    current_row = f_row + 4 

            st.success("✅ ปรับ Layout เป็นแบบต่อเนื่อง (Multiple DOs/Page) เรียบร้อยแล้ว!")
            st.download_button("📥 ดาวน์โหลดไฟล์", output_do.getvalue(), "DO_Final_Continuous.xlsx")