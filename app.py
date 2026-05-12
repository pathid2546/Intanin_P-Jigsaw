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

# --- TAB 2: ระบบ GENPrint DO (Fixed Overlapping) ---
    with tab2:
        st.subheader("📑 ออกใบส่งสินค้า (A4 Format - 5 Columns)")
        
        # กรองแถวว่าง (ป้องกัน IndexError)
        df_clean = df_raw.dropna(subset=['รหัสสาขา']).copy()
        
        if st.button("🚀 สร้างไฟล์ใบส่งสินค้า (Fixed Layout)"):
            output_do = io.BytesIO()
            with pd.ExcelWriter(output_do, engine='xlsxwriter') as writer:
                workbook = writer.book
                
                # --- กำหนด Styles ---
                head_comp = workbook.add_format({'bold': True, 'font_size': 13})
                title_fmt = workbook.add_format({'bold': True, 'font_size': 17, 'align': 'center', 'valign': 'vcenter', 'border': 1})
                normal_left = workbook.add_format({'font_size': 10, 'align': 'left'})
                normal_right = workbook.add_format({'font_size': 10, 'align': 'right'})
                date_bold = workbook.add_format({'bold': True, 'num_format': 'dd/mm/yyyy', 'font_size': 10, 'align': 'right'})
                date_norm = workbook.add_format({'num_format': 'dd/mm/yyyy', 'font_size': 10, 'align': 'right'})
                table_head = workbook.add_format({'border': 1, 'align': 'center', 'bold': True, 'bg_color': '#FFFFFF'})
                cell_border = workbook.add_format({'border': 1})
                cell_center = workbook.add_format({'border': 1, 'align': 'center'})
                total_fmt = workbook.add_format({'bold': True, 'border': 1, 'align': 'center', 'bg_color': '#F2F2F2'})
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
                    
                    # ความกว้าง 5 คอลัมน์ (A-E)
                    worksheet.set_column('A:A', 6)   # No.
                    worksheet.set_column('B:B', 15)  # Code
                    worksheet.set_column('C:C', 45)  # Name
                    worksheet.set_column('D:D', 15)  # Unit
                    worksheet.set_column('E:E', 12)  # QTY

                    # --- [แก้ไขจุดที่ทับซ้อน] Header Layout ---
                    # ฝั่งซ้าย (A-B): ข้อมูลบริษัท
                    worksheet.merge_range('A1:B1', 'บริษัท โมบาย โลจิสติกส์ จำกัด', head_comp)
                    worksheet.merge_range('A2:B2', '279 หมู่ที่ 9 ตำบลบางโฉลง', normal_left)
                    worksheet.merge_range('A3:B3', 'อำเภอบางพลี จังหวัดสมุทรปราการ 10540', normal_left)
                    worksheet.merge_range('A4:B4', 'Tel : 099-157-3114', normal_left)
                    
                    # ตรงกลาง (C-D): ชื่อเอกสาร (ไม่ทับกับ A-B และ E)
                    worksheet.merge_range('C2:D4', 'ใบส่งสินค้าชั่วคราว', title_fmt)
                    
                    # ฝั่งขวา (E): ข้อมูลเลขที่เอกสาร
                    worksheet.write('E1', f'Do. No. {first_row["เลขที่ DO."]}', normal_right)
                    worksheet.write('E2', f'Ref. Po. {first_row["เลขที่ PO."]}', normal_right)
                    worksheet.write('E3', 'Ref. Po. -', normal_right)
                    worksheet.write('E4', f'Zone {first_row["โซน"]}', normal_right)
                    worksheet.write('E5', f'Cutoff: {first_row["Cut Off Date"]}', date_norm)
                    worksheet.write('E6', f'Delivery: {first_row["Delivery Date"]}', date_bold)

                    # ข้อมูลสาขา (A-C)
                    worksheet.merge_range('A6:C6', f'Store Code: {first_row["รหัสสาขา"]}', normal_left)
                    worksheet.merge_range('A7:C7', f'Store Name: {first_row["Store Name"]}', normal_left)
                    worksheet.write('A8', 'Ship To:', normal_left)
                    worksheet.merge_range('B8:D9', first_row['ที่อยู่'], footer_box)

                    # --- Table Body ---
                    cols = ['No.', 'Product Code', 'Product Name', 'Unit/UOM', 'QTY']
                    for c_idx, val in enumerate(cols):
                        worksheet.write(10, c_idx, val, table_head)

                    r_idx = 11
                    for i, (_, row) in enumerate(df_store.iterrows(), 1):
                        worksheet.write(r_idx, 0, i, cell_center)
                        worksheet.write(r_idx, 1, row['รหัสสินค้า'], cell_center)
                        worksheet.write(r_idx, 2, row['รายการสินค้า'], cell_border)
                        worksheet.write(r_idx, 3, row['หน่วย'], cell_border)
                        worksheet.write(r_idx, 4, row['จำนวน'], cell_center)
                        r_idx += 1

                    # --- Footer ---
                    worksheet.merge_range(r_idx, 0, r_idx, 3, 'Total', total_fmt)
                    worksheet.write(r_idx, 4, df_store['จำนวน'].sum(), total_fmt)
                    
                    f_row = r_idx + 1
                    # ปรับกล่องเซ็นชื่อให้สมดุล 5 คอลัมน์
                    worksheet.merge_range(f_row, 0, f_row, 1, 'ผู้รับสินค้า', table_head)
                    worksheet.merge_range(f_row+1, 0, f_row+4, 1, 'ชื่อ (ตัวบรรจง):\nวันที่:\nเวลา:', footer_box)
                    
                    worksheet.merge_range(f_row, 2, f_row, 3, 'ผู้ส่งสินค้า / ทะเบียนรถ', table_head)
                    worksheet.merge_range(f_row+1, 2, f_row+4, 3, 'ชื่อ (ตัวบรรจง):\nวันที่:\nเวลา:', footer_box)
                    
                    worksheet.write(f_row, 4, 'คลังสินค้า', table_head)
                    worksheet.write(f_row+1, 4, 'ชื่อ:\nวันที่:', footer_box)
                    worksheet.set_row(f_row+1, 50)

            st.success("✅ แก้ไขปัญหา Overlapping และปรับปรุง Layout เรียบร้อย!")
            st.download_button("📥 ดาวน์โหลดไฟล์ DO", output_do.getvalue(), "DO_Final_A4.xlsx")