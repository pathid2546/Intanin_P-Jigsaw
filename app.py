import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="DO System V12.0 Full Fix", layout="wide")

st.title("📦 ระบบจัดการขนส่ง (V12.0 - Tab 1 Fix & Clean Layout)")
st.markdown("---")

# ระบบจำชื่อชีต (Memory)
if 'name_memory' not in st.session_state:
    st.session_state['name_memory'] = {}

uploaded_file = st.file_uploader("อัปโหลดไฟล์ Excel (Transport)", type=['xlsx'])

if uploaded_file:
    # อ่านข้อมูล Transport
    df_raw = pd.read_excel(uploaded_file, sheet_name='Transport')
    
    tab1, tab2 = st.tabs(["✂️ 1. แยกซัพพลายเออร์ (ใช้งานได้แล้ว)", "📄 2. ออกใบส่งสินค้า (V12.0 Clean)"])

    # --- TAB 1: ระบบแยกซัพพลายเออร์ (แก้ไขให้กลับมาใช้งานได้) ---
    with tab1:
        df_split = df_raw.dropna(subset=['ซัพพลายเออร์'])
        unique_suppliers = sorted(df_split['ซัพพลายเออร์'].unique())
        
        st.subheader("ตั้งชื่อชีตสำหรับแต่ละซัพพลายเออร์")
        with st.form("sheet_name_form"):
            cols = st.columns(3)
            current_mapping = {}
            for i, supplier in enumerate(unique_suppliers):
                with cols[i % 3]:
                    # ดึงชื่อเดิมจาก Memory หรือใช้ชื่อย่อซัพพลายเออร์
                    default_name = st.session_state['name_memory'].get(supplier, str(supplier)[:20].strip())
                    current_mapping[supplier] = st.text_input(f"ชื่อเดิม: {supplier}", value=default_name, key=f"inp_{supplier}")
            
            submit_split = st.form_submit_button("สร้างไฟล์แยกชีตตามชื่อที่ตั้ง")
            
            if submit_split:
                output_split = io.BytesIO()
                with pd.ExcelWriter(output_split, engine='xlsxwriter') as writer:
                    for supplier, sheet_name in current_mapping.items():
                        # บันทึกค่าลง Memory
                        st.session_state['name_memory'][supplier] = sheet_name
                        # กรองข้อมูลและเขียนลงชีต
                        df_sup = df_split[df_split['ซัพพลายเออร์'] == supplier].copy()
                        # ล้างอักขระพิเศษที่ Excel ไม่ยอมรับในชื่อชีต
                        clean_sheet_name = "".join([c for c in sheet_name if c.isalnum() or c in " _-"])[:31]
                        df_sup.to_excel(writer, sheet_name=clean_sheet_name, index=False)
                
                st.success("✅ แยกชีตสำเร็จ!")
                st.download_button("📥 ดาวน์โหลดไฟล์แยกชีต", output_split.getvalue(), "Supplier_Split_V12.xlsx")

    # --- TAB 2: ออกใบ DO (V12.0 Master Clean) ---
    with tab2:
        df_clean = df_raw.dropna(subset=['รหัสสาขา']).copy()
        if st.button("🚀 สร้างใบส่งสินค้า V12.0"):
            output_do = io.BytesIO()
            with pd.ExcelWriter(output_do, engine='xlsxwriter') as writer:
                workbook = writer.book
                worksheet = workbook.add_worksheet("DO_Master")
                
                worksheet.set_paper(9) # A4
                worksheet.set_margins(0.3, 0.3, 0.3, 0.3)
                worksheet.set_column('A:A', 8)    
                worksheet.set_column('B:B', 35)   # 250 pixels
                worksheet.set_column('C:C', 35)   
                worksheet.set_column('D:D', 18)   
                worksheet.set_column('E:E', 22)   
                
                # Styles
                f_comp = workbook.add_format({'bold': True, 'font_size': 14})
                f_title = workbook.add_format({'bold': True, 'font_size': 18, 'align': 'center', 'valign': 'vcenter'})
                f_std = workbook.add_format({'font_size': 11})
                f_bold = workbook.add_format({'bold': True, 'font_size': 11})
                f_right = workbook.add_format({'font_size': 11, 'align': 'right'})
                
                # Delivery Date (Clean - No BG Color)
                f_deliv_label = workbook.add_format({'bold': True, 'font_size': 14, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True})
                f_deliv_date = workbook.add_format({'bold': True, 'font_size': 14, 'align': 'center', 'valign': 'vcenter'})

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

                    # Header
                    worksheet.merge_range(curr, 0, curr, 2, 'บริษัท โมบาย โลจิสติกส์ จำกัด', f_comp)
                    worksheet.merge_range(curr+1, 0, curr+1, 1, '279 หมู่ที่ 9 ตำบลบางโฉลง อำเภอบางพลี', f_std)
                    worksheet.merge_range(curr+2, 0, curr+2, 1, 'จังหวัดสมุทรปราการ 10540', f_std)
                    worksheet.merge_range(curr+3, 0, curr+3, 2, 'ติดต่อ/สอบถาม : ID Line Official : @505phsps (มี @ ), Tel : 099-157-3114', f_std)
                    worksheet.write(curr+4, 0, 'Customer Name', f_std)
                    worksheet.write(curr+5, 0, f'Store Code: {store_code}', f_bold)
                    worksheet.write(curr+6, 0, f'Store Name: {first["Store Name"]}', f_bold)
                    worksheet.write(curr+7, 0, f'Ship To: {first["ที่อยู่"]}', f_std)

                    # Title Size 18
                    worksheet.merge_range(curr+1, 2, curr+2, 2, 'ใบส่งสินค้าชั่วคราว', f_title)

                    # Right Section (Clean No Dots)
                    worksheet.write(curr, 3, 'Do. No.', f_right)
                    worksheet.write(curr, 4, str(first["เลขที่ DO."]), f_std)
                    worksheet.write(curr+1, 3, 'Ref. Po.', f_right)
                    worksheet.write(curr+1, 4, str(first["เลขที่ PO."]), f_std)
                    worksheet.write(curr+2, 3, 'Ref. Po.', f_right)
                    worksheet.write(curr+2, 4, '-', f_std)
                    worksheet.write(curr+3, 3, 'Ref. Po.', f_right)
                    worksheet.write(curr+3, 4, '-', f_std)
                    
                    cutoff = pd.to_datetime(first.get('Cut Off Date')).strftime('%d/%m/%Y') if pd.notnull(first.get('Cut Off Date')) else "-"
                    worksheet.write(curr+4, 3, 'Cut off Date', f_right)
                    worksheet.write(curr+4, 4, cutoff, f_std)
                    worksheet.write(curr+5, 3, 'Zone', f_right)
                    worksheet.write(curr+5, 4, str(first["โซน"]), f_std)
                    
                    # Delivery Date (Vertical Merge 7-8)
                    delivery = pd.to_datetime(first["Delivery Date"]).strftime('%d/%m/%Y') if pd.notnull(first["Delivery Date"]) else "-"
                    worksheet.merge_range(curr+7, 3, curr+8, 3, "Delivery\nDate", f_deliv_label)
                    worksheet.merge_range(curr+7, 4, curr+8, 4, delivery, f_deliv_date)

                    # Table
                    t_h = curr + 10
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
                    
                    worksheet.merge_range(r_ptr, 0, r_ptr, 3, 'Total', f_table_h)
                    worksheet.write(r_ptr, 4, df_store['จำนวน'].sum(), f_border)

                    # Footer
                    f_row = r_ptr + 1
                    worksheet.merge_range(f_row, 0, f_row, 1, 'ผู้รับสินค้า', f_footer_h)
                    worksheet.merge_range(f_row, 2, f_row, 3, 'ผู้ส่งสินค้า / ทะเบียนรถ', f_footer_h)
                    worksheet.write(f_row, 4, 'คลังสินค้า', f_footer_h)
                    for label in ['ชื่อ (ตัวบรรจง):', 'วันที่:', 'เวลา:', 'หมายเหตุ:']:
                        f_row += 1
                        worksheet.merge_range(f_row, 0, f_row, 1, label, f_footer_box)
                        worksheet.merge_range(f_row, 2, f_row, 3, label, f_footer_box)
                        worksheet.write(f_row, 4, label, f_footer_box)

                    curr = f_row + 2

                worksheet.set_h_pagebreaks(page_breaks)
                worksheet.fit_to_pages(1, 0)

            st.success("✅ V12.0: ซ่อม Tab 1 และล้าง Layout เรียบร้อยครับ!")
            st.download_button("📥 ดาวน์โหลด DO V12.0", output_do.getvalue(), "DO_V12_FullFix.xlsx")