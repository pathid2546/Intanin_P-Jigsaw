import streamlit as st
import pandas as pd
import io

# --- CONFIG & OFFICIAL CSS (Shell UI) ---
st.set_page_config(page_title="Intanin Receipt Convert", layout="wide", page_icon="📦")

st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Sarabun:wght=300;400;700&display=swap');
    html, body, [class*="css"] { font-family: 'Sarabun', sans-serif; }
    .main-header { padding: 1.5rem; border-bottom: 3px solid #2e7d32; margin-bottom: 2rem; }
    .stAlert p { margin-bottom: 0; font-weight: bold; }
    </style>
    """, unsafe_allow_html=True)

st.markdown("""
    <div class="main-header">
        <h1 style='text-align: center;'>🚛 Intanin Receipt Convert</h1>
        <p style='text-align: center; color: #666;'>Official Logistics System (V14.0 - Data Validation Dashboard)</p>
    </div>
    """, unsafe_allow_html=True)

if 'name_memory' not in st.session_state:
    st.session_state['name_memory'] = {}

uploaded_file = st.file_uploader("📂 อัปโหลดไฟล์ Excel (Transport)", type=['xlsx'])

if uploaded_file:
    with st.spinner('กำลังอ่านข้อมูลจากไฟล์...'):
        df_raw = pd.read_excel(uploaded_file, sheet_name='Transport')
    
    # ==========================================
    # 🔍 SYSTEM: DATA VALIDATION & PREVIEW
    # ==========================================
    st.markdown("### 🔍 ระบบตรวจสอบข้อมูลคอลัมน์ วันที่ (Data Validation)")
    
    # ตรวจสอบชื่อคอลัมน์ในไฟล์จริง (ป้องกันกรณีพิมพ์ผิดตัวเล็กตัวใหญ่)
    col_cutoff = [c for c in df_raw.columns if 'cut' in c.lower() and 'off' in c.lower()]
    col_delivery = [c for c in df_raw.columns if 'deliv' in c.lower() and 'date' in c.lower()]
    
    name_cutoff = col_cutoff[0] if col_cutoff else 'Cut Off Date'
    name_delivery = col_delivery[0] if col_delivery else 'Delivery Date'
    
    # ค้นหาแถวที่คอลัมน์ Cut Off Date หรือ Delivery Date เป็นค่าว่าง
    df_invalid = df_raw[df_raw[name_cutoff].isna() | df_raw[name_delivery].isna()].copy()
    
    if len(df_invalid) > 0:
        st.error(f"⚠️ พบข้อมูลไม่ครบถ้วน! มีช่องว่างในคอลัมน์วันที่ ทั้งหมด {len(df_invalid)} แถว (กรุณาตรวจสอบตารางพรีวิวด้านล่าง)")
        
        # ปรับเลย์เอาต์ดึงคอลัมน์หลักมาพรีวิวให้อ่านง่าย
        preview_cols = []
        for c in ['รหัสสาขา', 'Store Name', 'โซน', 'รหัสสินค้า', 'รายการสินค้า', name_cutoff, name_delivery]:
            if c in df_raw.columns:
                preview_cols.append(c)
                
        df_preview = df_invalid[preview_cols].copy()
        # แปลง Index แสดงเป็นเลขแถวใน Excel เพื่อให้หาเจอง่ายๆ (Excel แถวแรกเริ่มที่ 2 เพราะมีหัวตาราง)
        df_preview.index = df_preview.index + 2
        df_preview.index.name = 'แถวใน Excel (Row)'
        
        # แสดงตารางพรีวิวแถวที่มีปัญหา
        st.dataframe(df_preview, use_container_width=True)
    else:
        st.success("🎉 ไม่พบช่องว่างในคอลัมน์ Cut Off Date และ Delivery Date ข้อมูลถูกต้องครบถ้วนพร้อมใช้งาน!")
    
    st.markdown("---")
    
    # ==========================================
    # WORKSPACE TABS (V13.1 CORE)
    # ==========================================
    tab1, tab2 = st.tabs(["✂️ 1. แยกซัพพลายเออร์ (Full Option)", "📄 2. ออกใบส่งสินค้า (V11.0 Clean)"])

    # --- TAB 1: ระบบแยกซัพพลายเออร์ ---
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
            submit_split = st.form_submit_button("ประมวลผลแยกซัพพลายเออร์")

        if submit_split:
            output_split = io.BytesIO()
            with pd.ExcelWriter(output_split, engine='xlsxwriter') as writer:
                workbook = writer.book
                header_fmt = workbook.add_format({'bold': True, 'bg_color': '#D9EAD3', 'border': 1, 'align': 'center'})
                cell_fmt = workbook.add_format({'border': 1})
                num_fmt = workbook.add_format({'border': 1, 'align': 'right'})
                total_fmt = workbook.add_format({'bold': True, 'bg_color': '#F3F3F3', 'border': 1, 'align': 'right', 'num_format': '#,##0'})
                
                for supplier, sheet_name in current_mapping.items():
                    st.session_state['name_memory'][supplier] = sheet_name
                    clean_name = "".join([c if c not in r'[]:*?/\ ' else ' ' for c in sheet_name.strip()[:31]])
                    df_sup = df_split[df_split['ซัพพลายเออร์'] == supplier].copy()
                    worksheet = workbook.add_worksheet(clean_name)
                    
                    # --- ส่วนที่ 1: ตารางฝั่งซ้าย (8 คอลัมน์ โครงสร้าง V13.1 เป๊ะ) ---
                    left_headers = ['รหัสสาขา', 'สาขา', 'โซน', 'รหัสสินค้า', 'รายการสินค้า', 'ซัพพลายเออร์', 'หน่วย', 'Total']
                    for col, h in enumerate(left_headers): 
                        worksheet.write(0, col, h, header_fmt)

                    curr_row = 1
                    for (branch, store_name, zone), b_group in df_sup.groupby(['รหัสสาขา', 'Store Name', 'โซน'], sort=False):
                        for i, (_, row) in enumerate(b_group.iterrows()):
                            worksheet.write(curr_row, 0, branch if i == 0 else "", cell_fmt)
                            worksheet.write(curr_row, 1, store_name if i == 0 else "", cell_fmt)
                            worksheet.write(curr_row, 2, zone if i == 0 else "", cell_fmt)
                            worksheet.write(curr_row, 3, row['รหัสสินค้า'], cell_fmt)
                            worksheet.write(curr_row, 4, row['รายการสินค้า'], cell_fmt)
                            worksheet.write(curr_row, 5, row['ซัพพลายเออร์'], cell_fmt)
                            worksheet.write(curr_row, 6, row['หน่วย'], cell_fmt)
                            worksheet.write(curr_row, 7, row['จำนวน'], num_fmt)
                            curr_row += 1
                    
                    worksheet.write(curr_row, 0, "Grand Total", total_fmt)
                    worksheet.write(curr_row, 7, df_sup['จำนวน'].sum(), total_fmt)
                    
                    # --- ส่วนที่ 2: ตารางฝั่งขวา ---
                    col_offset = 10  
                    right_headers = ['ซัพพลายเออร์', 'รหัสสินค้า', 'รายการสินค้า', 'Total']
                    for col, h in enumerate(right_headers): 
                        worksheet.write(0, col_offset + col, h, header_fmt)
                    
                    right_summary = df_sup.groupby(['รหัสสินค้า', 'รายการสินค้า'], as_index=False)['จำนวน'].sum()
                    for i, row in right_summary.iterrows():
                        worksheet.write(i + 1, col_offset, "", cell_fmt)
                        worksheet.write(i + 1, col_offset + 1, row['รหัสสินค้า'], cell_fmt)
                        worksheet.write(i + 1, col_offset + 2, row['รายการสินค้า'], cell_fmt)
                        worksheet.write(i + 1, col_offset + 3, row['จำนวน'], num_fmt)
                    
                    sum_row = len(right_summary) + 1
                    worksheet.write(sum_row, col_offset, f"{supplier} Total", total_fmt)
                    worksheet.write(sum_row, col_offset + 3, right_summary['จำนวน'].sum(), total_fmt)
                    
                    # ความกว้างคอลัมน์ฝั่งซ้าย (A - H)
                    worksheet.set_column('A:A', 15)  
                    worksheet.set_column('B:B', 30)  
                    worksheet.set_column('C:C', 15)  
                    worksheet.set_column('D:D', 15)  
                    worksheet.set_column('E:E', 35)  
                    worksheet.set_column('F:H', 15)  
                    
                    # ความกว้างคอลัมน์ฝั่งขวา (K - N) บังคับความกว้างเท่ากันสมมาตร
                    worksheet.set_column('K:K', 15)  
                    worksheet.set_column('L:L', 15)  # รหัสสินค้า ขวา = 15 เท่ากับซ้าย
                    worksheet.set_column('M:M', 35)  # รายการสินค้า ขวา = 35 เท่ากับซ้าย
                    worksheet.set_column('N:N', 15)  

            st.success("✅ ประมวลผลแยกซัพพลายเออร์สำเร็จ!")
            st.download_button(label="📥 ดาวน์โหลดไฟล์ Splitter V14.0", data=output_split.getvalue(), file_name="Intanin_Splitter_V14.0.xlsx")

    # --- TAB 2: ออกใบ DO (โครงสร้างล็อกเดิม) ---
    with tab2:
        df_clean = df_raw.dropna(subset=['รหัสสาขา']).copy()
        if st.button("🚀 สร้างใบส่งสินค้า Official V14.0"):
            output_do = io.BytesIO()
            with pd.ExcelWriter(output_do, engine='xlsxwriter') as writer:
                workbook = writer.book
                worksheet = workbook.add_worksheet("DO_Master")
                worksheet.set_paper(9); worksheet.set_margins(0.3, 0.3, 0.3, 0.3)
                worksheet.set_column('A:A', 8); worksheet.set_column('B:B', 35); worksheet.set_column('C:C', 35); worksheet.set_column('D:D', 18); worksheet.set_column('E:E', 22)
                
                f_comp = workbook.add_format({'bold': True, 'font_size': 14})
                f_title = workbook.add_format({'bold': True, 'font_size': 18, 'align': 'center', 'valign': 'vcenter'})
                f_std = workbook.add_format({'font_size': 11}); f_bold = workbook.add_format({'bold': True, 'font_size': 11})
                f_right = workbook.add_format({'font_size': 11, 'align': 'right'})
                f_deliv_label = workbook.add_format({'bold': True, 'font_size': 11, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True})
                f_deliv_date = workbook.add_format({'bold': True, 'font_size': 11, 'align': 'center', 'valign': 'vcenter'})
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
                    worksheet.merge_range(curr, 0, curr, 2, 'บริษัท โมบาย โลจิสติกส์ จำกัด', f_comp)
                    worksheet.merge_range(curr+1, 0, curr+1, 1, '279 หมู่ที่ 9 ตำบลบางโฉลง อำเภอบางพลี', f_std)
                    worksheet.merge_range(curr+2, 0, curr+2, 1, 'จังหวัดสมุทรปราการ 10540', f_std)
                    worksheet.merge_range(curr+3, 0, curr+3, 2, 'ติดต่อ/สอบถาม : ID Line Official : @505phsps (มี @ ), Tel : 099-157-3114', f_std)
                    worksheet.write(curr+4, 0, 'Customer Name', f_std); worksheet.write(curr+5, 0, f'Store Code: {store_code}', f_bold)
                    worksheet.write(curr+6, 0, f'Store Name: {first["Store Name"]}', f_bold); worksheet.write(curr+7, 0, f'Ship To: {first["ที่อยู่"]}', f_std)
                    worksheet.merge_range(curr+1, 2, curr+2, 2, 'ใบส่งสินค้าชั่วคราว', f_title)
                    worksheet.write(curr, 3, 'Do. No.', f_right); worksheet.write(curr, 4, str(first["เลขที่ DO."]), f_std)
                    worksheet.write(curr+1, 3, 'Ref. Po.', f_right); worksheet.write(curr+1, 4, str(first["เลขที่ PO."]), f_std)
                    worksheet.write(curr+2, 3, 'Ref. Po.', f_right); worksheet.write(curr+2, 4, '-', f_std)
                    worksheet.write(curr+3, 3, 'Ref. Po.', f_right); worksheet.write(curr+3, 4, '-', f_std)
                    
                    cutoff_val = first.get(name_cutoff)
                    cutoff = pd.to_datetime(cutoff_val).strftime('%d/%m/%Y') if pd.notnull(cutoff_val) else "-"
                    worksheet.write(curr+4, 3, 'Cut off Date', f_right); worksheet.write(curr+4, 4, cutoff, f_std)
                    worksheet.write(curr+5, 3, 'Zone', f_right); worksheet.write(curr+5, 4, str(first["โซน"]), f_std)
                    
                    delivery_val = first.get(name_delivery)
                    delivery = pd.to_datetime(delivery_val).strftime('%d/%m/%Y') if pd.notnull(delivery_val) else "-"
                    worksheet.merge_range(curr+7, 3, curr+8, 3, "Delivery\nDate", f_deliv_label); worksheet.merge_range(curr+7, 4, curr+8, 4, delivery, f_deliv_date)
                    
                    t_h = curr + 10 
                    for i, txt in enumerate(['No.', 'Product Code', 'Product Name', 'Unit/UOM', 'QTY']): worksheet.write(t_h, i, txt, f_table_h)
                    r_ptr = t_h + 1
                    for i, (_, r) in enumerate(df_store.iterrows(), 1):
                        worksheet.write(r_ptr, 0, i, f_border); worksheet.write(r_ptr, 1, str(r['รหัสสินค้า']), f_border)
                        worksheet.write(r_ptr, 2, r['รายการสินค้า'], f_wrap); worksheet.write(r_ptr, 3, r['หน่วย'], f_border)
                        worksheet.write(r_ptr, 4, r['จำนวน'], f_border); r_ptr += 1
                    worksheet.merge_range(r_ptr, 0, r_ptr, 3, 'Total', f_table_h); worksheet.write(r_ptr, 4, df_store['จำนวน'].sum(), f_border)
                    f_row = r_ptr + 1
                    worksheet.merge_range(f_row, 0, f_row, 1, 'ผู้รับสินค้า', f_footer_h); worksheet.merge_range(f_row, 2, f_row, 3, 'ผู้ส่งสินค้า / ทะเบียนรถ', f_footer_h); worksheet.write(f_row, 4, 'คลังสินค้า', f_footer_h)
                    for label in ['ชื่อ (ตัวบรรจง):', 'วันที่:', 'เวลา:', 'หมายเหตุ:']:
                        f_row += 1
                        worksheet.merge_range(f_row, 0, f_row, 1, label, f_footer_box); worksheet.merge_range(f_row, 2, f_row, 3, label, f_footer_box); worksheet.write(f_row, 4, label, f_footer_box)
                    curr = f_row + 2
                worksheet.set_h_pagebreaks(page_breaks); worksheet.fit_to_pages(1, 0)

            st.success("✅ สร้างไฟล์ DO สำเร็จ!")
            st.download_button("📥 ดาวน์โหลด DO Master", output_do.getvalue(), "Intanin_DO_Final_V14.0.xlsx")
