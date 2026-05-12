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

# --- TAB 2: ระบบ GENPrint DO (ฉบับแก้ไขตำแหน่ง Title ให้ตรงต้นฉบับ) ---
    with tab2:
        st.subheader("📑 ออกใบส่งสินค้า (A4 - Fixed Title Position)")
        
        df_clean = df_raw.dropna(subset=['รหัสสาขา']).copy()
        
        if st.button("🚀 สร้างไฟล์ใบส่งสินค้า (Correct Layout)"):
            output_do = io.BytesIO()
            with pd.ExcelWriter(output_do, engine='xlsxwriter') as writer:
                workbook = writer.book
                
                # --- กำหนด Styles ---
                comp_name_fmt = workbook.add_format({'bold': True, 'font_size': 13})
                doc_title_fmt = workbook.add_format({'bold': True, 'font_size': 18, 'align': 'center', 'valign': 'vcenter'})
                std_left = workbook.add_format({'font_size': 10, 'align': 'left'})
                std_right = workbook.add_format({'font_size': 10, 'align': 'right'})
                
                # Format วันที่
                date_norm_fmt = workbook.add_format({'num_format': 'dd/mm/yyyy', 'font_size': 10, 'align': 'right'})
                date_bold_fmt = workbook.add_format({'bold': True, 'num_format': 'dd/mm/yyyy', 'font_size': 11, 'align': 'right'})
                
                table_head = workbook.add_format({'border': 1, 'align': 'center', 'bold': True})
                border_left = workbook.add_format({'border': 1, 'align': 'left', 'text_wrap': True})
                border_center = workbook.add_format({'border': 1, 'align': 'center'})
                footer_box = workbook.add_format({'border': 1, 'font_size': 9, 'valign': 'top', 'text_wrap': True})

                for store_code in df_clean['รหัสสาขา'].unique():
                    df_store = df_clean[df_clean['รหัสสาขา'] == store_code].copy()
                    if df_store.empty: continue
                    
                    first_row = df_store.iloc[0]
                    sheet_name = str(store_code).strip()[:31]
                    worksheet = workbook.add_worksheet(sheet_name)
                    
                    # ตั้งค่าหน้ากระดาษ A4
                    worksheet.set_paper(9) 
                    worksheet.set_margins(0.3, 0.3, 0.3, 0.3)
                    worksheet.fit_to_pages(1, 0)
                    
                    # กำหนดความกว้างคอลัมน์มาตรฐาน (ไม่ขยายเพิ่ม)
                    worksheet.set_column('A:A', 5)   # No.
                    worksheet.set_column('B:B', 14)  # Product Code
                    worksheet.set_column('C:C', 45)  # Product Name
                    worksheet.set_column('D:D', 12)  # Unit/UOM
                    worksheet.set_column('E:E', 15)  # QTY & Right Info

                    # --- [SECTION 1: HEADER - เน้นตำแหน่ง Title] ---
                    # ฝั่งซ้าย: บริษัท
                    worksheet.write('A1', 'บริษัท โมบาย โลจิสติกส์ จำกัด', comp_name_fmt)
                    worksheet.write('A2', '279 หมู่ที่ 9 ตำบลบางโฉลง', std_left)
                    worksheet.write('A3', 'อำเภอบางพลี จังหวัดสมุทรปราการ 10540', std_left)
                    worksheet.write('A4', 'ติดต่อ/สอบถาม : Tel : 099-157-3114', std_left)

                    # ตรงกลาง: Title อยู่ที่ C2:C4 เท่านั้น (เพื่อให้ตรงกับแกนกลางต้นฉบับ)
                    worksheet.merge_range('C2:C4', 'ใบส่งสินค้าชั่วคราว', doc_title_fmt)

                    # ฝั่งขวา: ข้อมูลเอกสาร
                    worksheet.write('E1', f'Do. No. {first_row["เลขที่ DO."]}', std_right)
                    worksheet.write('E2', f'Ref. Po. {first_row["เลขที่ PO."]}', std_right)
                    worksheet.write('E3', 'Ref. Po. -', std_right)
                    worksheet.write('E4', f'Zone {first_row["โซน"]}', std_right)
                    worksheet.write('E5', f'Cutoff Date {first_row["Cut Off Date"]}', date_norm_fmt)
                    
                    # วันที่ส่งของ (จัดให้ตรงตำแหน่งต้นฉบับที่ D8-E8)
                    worksheet.write('D8', 'Delivery Date', workbook.add_format({'bold': True, 'align': 'right'}))
                    worksheet.write('E8', first_row['Delivery Date'], date_bold_fmt)

                    # --- [SECTION 2: CUSTOMER INFO] ---
                    worksheet.write('A5', 'Customer Name', std_left)
                    worksheet.write('A6', f'Store Code: {first_row["รหัสสาขา"]}', std_left)
                    worksheet.write('A7', f'Store Name: {first_row["Store Name"]}', std_left)
                    worksheet.write('A8', 'Ship To:', std_left)
                    worksheet.merge_range('B8:C9', first_row['ที่อยู่'], footer_box)

                    # --- [SECTION 3: TABLE] ---
                    headers = ['No.', 'Product Code', 'Product Name', 'Unit/UOM', 'QTY']
                    for col_num, head in enumerate(headers):
                        worksheet.write(10, col_num, head, table_head)

                    row_idx = 11
                    for i, (_, r) in enumerate(df_store.iterrows(), 1):
                        worksheet.write(row_idx, 0, i, border_center)
                        worksheet.write(row_idx, 1, r['รหัสสินค้า'], border_center)
                        worksheet.write(row_idx, 2, r['รายการสินค้า'], border_left)
                        worksheet.write(row_idx, 3, r['หน่วย'], border_left)
                        worksheet.write(row_idx, 4, r['จำนวน'], border_center)
                        row_idx += 1

                    # --- [SECTION 4: FOOTER] ---
                    worksheet.merge_range(row_idx, 0, row_idx, 3, 'Total', workbook.add_format({'bold': True, 'border': 1, 'align': 'center'}))
                    worksheet.write(row_idx, 4, df_store['จำนวน'].sum(), border_center)
                    
                    f_row = row_idx + 1
                    worksheet.set_row(f_row+1, 70) # ความสูงช่องเซ็นชื่อ
                    
                    worksheet.merge_range(f_row, 0, f_row, 1, 'ผู้รับสินค้า', table_head)
                    worksheet.merge_range(f_row+1, 0, f_row+1, 1, 'ชื่อ (ตัวบรรจง)\nวันที่\nเวลา\nหมายเหตุ', footer_box)
                    
                    worksheet.merge_range(f_row, 2, f_row, 3, 'ผู้ส่งสินค้า / ทะเบียนรถ', table_head)
                    worksheet.merge_range(f_row+1, 2, f_row+1, 3, 'ชื่อ (ตัวบรรจง)\nวันที่\nเวลา\nหมายเหตุ', footer_box)
                    
                    worksheet.write(f_row, 4, 'คลังสินค้า', table_head)
                    worksheet.write(f_row+1, 4, 'ชื่อ:\nวันที่:', footer_box)

            st.success("✅ สร้างไฟล์ใบส่งสินค้าที่มีตำแหน่ง Title ถูกต้องตามต้นฉบับแล้ว!")
            st.download_button("📥 ดาวน์โหลดไฟล์ DO", output_do.getvalue(), "DO_Final_A4.xlsx")