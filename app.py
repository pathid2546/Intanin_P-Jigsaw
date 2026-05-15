import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="DO System V8.8 Master", layout="wide")

st.title("📦 ระบบจัดการขนส่ง (V8.8 Final Master Fix)")
st.markdown("---")

if 'name_memory' not in st.session_state:
    st.session_state['name_memory'] = {}

uploaded_file = st.file_uploader("อัปโหลดไฟล์ Excel (Transport)", type=['xlsx'])

if uploaded_file:
    df_raw = pd.read_excel(uploaded_file, sheet_name='Transport')
    
    tab1, tab2 = st.tabs(["✂️ 1. แยกซัพพลายเออร์", "📄 2. ออกใบส่งสินค้า (A4 Fix)"])

    # --- TAB 1: ระบบแยกซัพพลายเออร์ (คงไว้ครบถ้วน) ---
    with tab1:
        df_split = df_raw.dropna(subset=['ซัพพลายเออร์'])
        unique_suppliers = sorted(df_split['ซัพพลายเออร์'].unique())
        with st.form("sheet_name_form"):
            cols = st.columns(3)
            current_mapping = {}
            for i, supplier in enumerate(unique_suppliers):
                with cols[i % 3]:
                    remembered_name = st.session_state['name_memory'].get(supplier, str(supplier)[:10].strip())
                    current_mapping[supplier] = st.text_input(f"{supplier}:", value=remembered_name, key=f"input_{supplier}")
            if st.form_submit_button("สร้างไฟล์แยกซัพพลายเออร์"):
                output_split = io.BytesIO()
                with pd.ExcelWriter(output_split, engine='xlsxwriter') as writer:
                    for supplier, sheet_name in current_mapping.items():
                        st.session_state['name_memory'][supplier] = sheet_name
                        df_sup = df_split[df_split['ซัพพลายเออร์'] == supplier].copy()
                        df_sup.to_excel(writer, sheet_name=sheet_name, index=False)
                st.download_button("📥 ดาวน์โหลดไฟล์แยกชีต", output_split.getvalue(), "Supplier_Split.xlsx")

    # --- TAB 2: ออกใบ DO (แก้ปัญหาตามจุดที่คุณทักท้วง) ---
    with tab2:
        df_clean = df_raw.dropna(subset=['รหัสสาขา']).copy()
        if st.button("🚀 สร้างใบส่งสินค้า V8.8"):
            output_do = io.BytesIO()
            with pd.ExcelWriter(output_do, engine='xlsxwriter') as writer:
                workbook = writer.book
                worksheet = workbook.add_worksheet("DO_Master")
                
                # Setup A4 & 5-Column Logic
                worksheet.set_paper(9) 
                worksheet.set_margins(0.3, 0.3, 0.3, 0.3)
                worksheet.set_column('A:A', 8)   # No.
                worksheet.set_column('B:B', 15)  # Code
                worksheet.set_column('C:C', 42)  # Title / Name (กว้างพอสำหรับที่อยู่บริษัท)
                worksheet.set_column('D:D', 15)  # Label
                worksheet.set_column('E:E', 18)  # Data
                
                # Styles
                f_comp = workbook.add_format({'bold': True, 'font_size': 14})
                f_title = workbook.add_format({'bold': True, 'font_size': 26, 'align': 'center', 'valign': 'vcenter'})
                f_std = workbook.add_format({'font_size': 11})
                f_bold = workbook.add_format({'bold': True, 'font_size': 11})
                f_right = workbook.add_format({'font_size': 11, 'align': 'right'})
                f_border = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter'})
                f_table_h = workbook.add_format({'border': 1, 'align': 'center', 'bold': True, 'bg_color': '#F2F2F2'})
                f_wrap = workbook.add_format({'border': 1, 'text_wrap': True, 'valign': 'vcenter'})
                f_footer_h = workbook.add_format({'border': 1, 'align': 'center', 'bold': True})
                f_footer_box = workbook.add_format({'border': 1, 'font_size': 10, 'valign': 'top'})

                curr = 0; page_breaks = []
                for store_code in df_clean['รหัสสาขา'].unique():
                    df_store = df_clean[df_clean['รหัสสาขา'] == store_code].copy()
                    if curr > 0: page_breaks.append(curr)
                    first = df_store.iloc[0]

                    # --- Row 1: ชื่อบริษัท ---
                    worksheet.merge_range(curr, 0, curr, 4, 'บริษัท โมบาย โลจิสติกส์ จำกัด', f_comp)
                    
                    # --- Row 2: ที่อยู่ (แก้ให้แสดงครบ), ใบส่งสินค้า (ใหญ่), Do. No. ---
                    worksheet.set_row(curr+1, 40) # เพิ่มความสูงแถว 2
                    worksheet.merge_range(curr+1, 0, curr+1, 1, '279 หมู่ที่ 9 ตำบลบางโฉลง อำเภอบางพลี', f_std) # Merge A-B
                    worksheet.write(curr+1, 2, 'ใบส่งสินค้าชั่วคราว', f_title)
                    worksheet.write(curr+1, 3, 'Do. No.', f_right)
                    worksheet.write(curr+1, 4, str(first["เลขที่ DO."]), f_std)

                    # --- Row 3: ที่อยู่บรรทัดต่อมา และ Ref. Po. ---
                    worksheet.merge_range(curr+2, 0, curr+2, 1, 'จังหวัดสมุทรปราการ 10540', f_std)
                    worksheet.write(curr+2, 3, 'Ref. Po.', f_right)
                    worksheet.write(curr+2, 4, str(first["เลขที่ PO."]), f_std)

                    # --- Row 4: ข้อมูลติดต่อ และ Customer Name (ลบจุดไข่ปลาออก) ---
                    worksheet.merge_range(curr+3, 0, curr+3, 2, 'ติดต่อ/สอบถาม : ID Line Official : @505phsps (มี @ ), Tel : 099-157-3114', f_std)
                    worksheet.write(curr+4, 0, 'Customer Name', f_std) # แก้ตามสั่ง: ลบจุดทิ้ง

                    # วันที่ (บังคับ Format ไม่เอาเวลา)
                    cutoff = pd.to_datetime(first.get('Cut Off Date')).strftime('%d/%m/%Y') if pd.notnull(first.get('Cut Off Date')) else "-"
                    delivery = pd.to_datetime(first["Delivery Date"]).strftime('%d/%m/%Y') if pd.notnull(first["Delivery Date"]) else "-"
                    
                    worksheet.write(curr+4, 3, 'Cutoff Date:', f_right)
                    worksheet.write(curr+4, 4, cutoff, f_std)
                    
                    # --- Row 5-7: Store & Delivery ---
                    worksheet.write(curr+5, 0, f'Store Code: {store_code}', f_bold)
                    worksheet.write(curr+5, 3, 'Zone:', f_right)
                    worksheet.write(curr+5, 4, str(first["โซน"]), f_std)
                    
                    worksheet.write(curr+6, 0, f'Store Name: {first["Store Name"]}', f_bold)
                    worksheet.write(curr+7, 0, f'Ship To: {first["ที่อยู่"]}', f_std)
                    worksheet.write(curr+7, 3, 'Delivery Date:', f_right)
                    worksheet.write(curr+7, 4, delivery, f_bold)

                    # --- Table Section ---
                    t_h = curr + 9
                    for i, txt in enumerate(['No.', 'Product Code', 'Product Name', 'Unit/UOM', 'QTY']):
                        worksheet.write(t_h, i, txt, f_table_h)

                    r_ptr = t_h + 1
                    for i, (_, r) in enumerate(df_store.iterrows(), 1):
                        worksheet.write(r_ptr, 0, i, f_border)
                        worksheet.write(r_ptr, 1, str(r['รหัสสินค้า']), f_border)
                        worksheet.write(r_ptr, 2, r['รายการสินค้า'], f_wrap)
                        worksheet.write(r_ptr, 3, r['หน่วย'], f_border)
                        worksheet.write(r_ptr, 4, r['จำนวน'], f_border)
                        r_ptr += 1
                    
                    # Total
                    worksheet.merge_range(r_ptr, 0, r_ptr, 3, 'Total', f_table_h)
                    worksheet.write(r_ptr, 4, df_store['จำนวน'].sum(), f_border)

                    # --- Footer Section (ปรับให้กว้างเท่ากัน 3 บล็อก) ---
                    f_row = r_ptr + 1
                    worksheet.merge_range(f_row, 0, f_row, 1, 'ผู้รับสินค้า', f_footer_h)
                    worksheet.merge_range(f_row, 2, f_row, 3, 'ผู้ส่งสินค้า / ทะเบียนรถ', f_footer_h)
                    worksheet.write(f_row, 4, 'คลังสินค้า', f_footer_h) # บล็อกทางขวากว้างตาม Column E
                    
                    labels = ['ชื่อ (ตัวบรรจง):', 'วันที่:', 'เวลา:', 'หมายเหตุ:']
                    for label in labels:
                        f_row += 1
                        worksheet.merge_range(f_row, 0, f_row, 1, label, f_footer_box)
                        worksheet.merge_range(f_row, 2, f_row, 3, label, f_footer_box)
                        worksheet.write(f_row, 4, label, f_footer_box)

                    curr = f_row + 2

                worksheet.set_h_pagebreaks(page_breaks)
                worksheet.fit_to_pages(1, 0)

            st.success("✅ แก้ไขที่อยู่และ Customer Name เรียบร้อยแล้วครับ!")
            st.download_button("📥 ดาวน์โหลด DO Master V8.8", output_do.getvalue(), "DO_Final_V8.8.xlsx")