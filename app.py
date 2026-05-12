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

    # --- TAB 2: ระบบใหม่ GENPrint DO (ใบส่งสินค้าชั่วคราว) ---
    with tab2:
        st.subheader("📑 ออกใบส่งสินค้าชั่วคราว (แยกตามสาขา)")
    
    required_cols = ['รหัสสาขา', 'Store Name', 'ที่อยู่', 'เลขที่ DO.', 'เลขที่ PO.', 'โซน', 'Cut Off Date', 'Delivery Date']
    missing = [c for c in required_cols if c not in df_raw.columns]
    
    if missing:
        st.error(f"❌ ไฟล์ของคุณขาดคอลัมน์: {', '.join(missing)}")
    else:
        # ลบแถวที่คอลัมน์สำคัญเป็นค่าว่างออกก่อนประมวลผล
        df_clean = df_raw.dropna(subset=['รหัสสาขา']).copy()
        
        if st.button("🚀 สร้างไฟล์ใบส่งสินค้า (GENPrint DO)"):
            if df_clean.empty:
                st.warning("⚠️ ไม่พบข้อมูลรหัสสาขาในไฟล์")
            else:
                output_do = io.BytesIO()
                with pd.ExcelWriter(output_do, engine='xlsxwriter') as writer:
                    workbook = writer.book
                    
                    # --- Styles (เหมือนเดิมเพื่อความเป๊ะของฟอร์ม) ---
                    header_bold = workbook.add_format({'bold': True, 'font_size': 14})
                    normal_font = workbook.add_format({'font_size': 11})
                    total_row = workbook.add_format({'bold': True, 'border': 1, 'align': 'center'})
                    table_header = workbook.add_format({'bold': True, 'border': 1, 'align': 'center'})
                    table_cell = workbook.add_format({'border': 1, 'align': 'left'})
                    num_cell = workbook.add_format({'border': 1, 'align': 'center'})
                    footer_style = workbook.add_format({'border': 1, 'font_size': 10})

                    # ดึงรหัสสาขาที่ไม่ซ้ำกัน
                    unique_stores = df_clean['รหัสสาขา'].unique()

                    for store_code in unique_stores:
                        # กรองข้อมูลเฉพาะสาขานั้นๆ
                        df_store = df_clean[df_clean['รหัสสาขา'] == store_code].copy()
                        
                        # ป้องกัน IndexError: ถ้ากรองแล้วมีข้อมูลค่อยทำต่อ
                        if not df_store.empty:
                            first_row = df_store.iloc[0]
                            # ใช้รหัสสาขาเป็นชื่อ Sheet (จำกัด 31 ตัวอักษรตามกฎ Excel)
                            sheet_name_do = str(store_code).strip()[:31]

                            worksheet = workbook.add_worksheet(sheet_name_do)
                            
                            # --- ส่วนหัวและข้อมูล (ดึงตาม Logic ที่คุณเจาะจง) ---
                            worksheet.write('A1', 'บริษัท โมบาย โลจิสติกส์ จำกัด', header_bold) #
                            worksheet.write('A6', f'Store Code  {first_row["รหัสสาขา"]}', normal_font) #
                            worksheet.write('A7', f'Store Name {first_row["Store Name"]}', normal_font) #
                            worksheet.write('B8', first_row['ที่อยู่'], workbook.add_format({'text_wrap': True})) #
                            
                            # ฝั่งขวา
                            worksheet.write('G1', f'Do. No. {first_row["เลขที่ DO."]}', normal_font)
                            worksheet.write('G2', f'Ref. Po. {first_row["เลขที่ PO."]}', normal_font)
                            worksheet.write('G6', f'Zone {first_row["โซน"]}', normal_font)
                            worksheet.write('G7', f'Cutoff Date {first_row["Cut Off Date"]}', normal_font)
                            worksheet.write('H8', first_row['Delivery Date'], workbook.add_format({'bold': True}))

                            # --- ตารางสินค้า (Body) ---
                            headers = ['No.', 'Product Code', 'Product Name', 'Unit/UOM', 'QTY']
                            for col_num, header in enumerate(headers):
                                worksheet.write(9, col_num, header, table_header)

                            row_idx = 10
                            for i, (_, row) in enumerate(df_store.iterrows(), 1):
                                worksheet.write(row_idx, 0, i, num_cell)
                                worksheet.write(row_idx, 1, row['รหัสสินค้า'], num_cell)
                                worksheet.write(row_idx, 2, row['รายการสินค้า'], table_cell)
                                worksheet.write(row_idx, 3, row['หน่วย'], table_cell)
                                worksheet.write(row_idx, 4, row['จำนวน'], num_cell)
                                row_idx += 1

                            # --- ส่วนท้าย (Footer) ---
                            worksheet.merge_range(row_idx, 0, row_idx, 3, 'Total', total_row)
                            worksheet.write(row_idx, 4, df_store['จำนวน'].sum(), total_row)
                            
                            # ตารางเซ็นชื่อ 3 ช่อง
                            f_row = row_idx + 1
                            f_headers = ['ผู้รับสินค้า', 'ผู้ส่งสินค้า / ทะเบียนรถ', 'คลังสินค้า']
                            for i, head in enumerate(f_headers):
                                c_start = i * 2
                                worksheet.merge_range(f_row, c_start, f_row, c_start+1, head, total_row)
                                worksheet.merge_range(f_row+1, c_start, f_row+1, c_start+1, 'ชื่อ (ตัวบรรจง)', footer_style)
                                worksheet.merge_range(f_row+2, c_start, f_row+2, c_start+1, 'วันที่', footer_style)
                                worksheet.merge_range(f_row+3, c_start, f_row+3, c_start+1, 'เวลา', footer_style)
                                worksheet.merge_range(f_row+4, c_start, f_row+4, c_start+1, 'หมายเหตุ', footer_style)

                st.success("✅ แก้ไขและสร้างไฟล์ DO เรียบร้อย!")
                st.download_button(label="📥 ดาวน์โหลดไฟล์ GENPrint DO (Fixed)", data=output_do.getvalue(), file_name="GENPrint_DO_Fixed.xlsx")