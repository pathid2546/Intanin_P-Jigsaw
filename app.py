import streamlit as st
import pandas as pd
import io

# --- CONFIG & APPLE PREMIUM UI CSS ---
st.set_page_config(page_title="Intanin Convert Pro", layout="wide", page_icon="🍏")

# บรรจุชุดคำสั่ง CSS ดีไซน์ระดับหรูสไตล์ Apple (iOS / macOS)
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Sarabun:wght@300;400;500;700&family=SF+Pro+Display:wght@300;400;600&display=swap');
    
    /* สไตล์ฟอนต์หลักผสมผสานระหว่างความสากลและภาษาไทย */
    html, body, [class*="css"] { 
        font-family: 'SF Pro Display', 'Sarabun', sans-serif; 
        background-color: #f5f5f7;
        color: #1d1d1f;
    }
    
    /* ส่วนหัวแบบ Glassmorphism หน้าจอกระจกฝ้า */
    .apple-header { 
        background: rgba(255, 255, 255, 0.75);
        backdrop-filter: blur(20px);
        -webkit-backdrop-filter: blur(20px);
        padding: 2.5rem 1.5rem; 
        border-radius: 24px;
        box-shadow: 0 4px 30px rgba(0, 0, 0, 0.03);
        border: 1px solid rgba(255, 255, 255, 0.6);
        margin-bottom: 2.5rem;
        text-align: center;
        transition: all 0.4s ease;
    }
    .apple-header:hover {
        box-shadow: 0 10px 40px rgba(0, 0, 0, 0.06);
        transform: translateY(-2px);
    }
    .apple-title {
        font-weight: 600;
        font-size: 2.4rem;
        letter-spacing: -0.03em;
        background: linear-gradient(135deg, #1e3c1a 0%, #2e7d32 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        margin-bottom: 0.3rem;
    }
    .apple-subtitle {
        color: #86868b;
        font-size: 1rem;
        font-weight: 400;
    }
    
    /* ปรับแต่งรูปแบบการอัปโหลดไฟล์ให้ดูมินิมอล */
    .stFileUploader {
        background: #ffffff;
        padding: 1.5rem;
        border-radius: 20px;
        box-shadow: 0 2px 12px rgba(0,0,0,0.02);
        border: 1px solid #e8e8ed;
    }
    
    /* กล่อง Alert เตือนของความบกพร่องข้อมูล (iOS Style) */
    .stAlert {
        border-radius: 16px !important;
        border: none !important;
        box-shadow: 0 4px 15px rgba(0,0,0,0.02) !important;
    }
    
    /* ปรับแต่งปุ่มกดดึงดูดสายตาและการตอบสนองแบบสมูท (Micro-interactions) */
    div.stButton > button, div.stDownloadButton > button {
        background: linear-gradient(180deg, #34c759 0%, #28cd41 100%) !important;
        color: white !important;
        border-radius: 12px !important;
        border: none !important;
        padding: 0.6rem 2rem !important;
        font-weight: 500 !important;
        box-shadow: 0 4px 12px rgba(52, 199, 89, 0.2) !important;
        transition: all 0.25s cubic-bezier(0.4, 0, 0.2, 1) !important;
    }
    div.stButton > button:hover, div.stDownloadButton > button:hover {
        transform: translateY(-2px) scale(1.02) !important;
        box-shadow: 0 6px 20px rgba(52, 199, 89, 0.35) !important;
        background: linear-gradient(180deg, #30b351 0%, #26b93b 100%) !important;
    }
    div.stButton > button:active, div.stDownloadButton > button:active {
        transform: translateY(1px) scale(0.98) !important;
    }
    
    /* แบบฟอร์มกรอกข้อมูลแบบเรียบหรู */
    div[data-testid="stForm"] {
        background: #ffffff !important;
        border-radius: 20px !important;
        border: 1px solid #e8e8ed !important;
        box-shadow: 0 4px 20px rgba(0,0,0,0.02) !important;
        padding: 2rem !important;
    }
    
    /* การออกแบบแท็บ (Tabs) สไตล์เมนูตั้งค่า iOS */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
        background-color: #e3e3e6;
        padding: 6px;
        border-radius: 14px;
    }
    .stTabs [data-baseweb="tab"] {
        height: 40px;
        white-space: pre;
        background-color: transparent;
        border-radius: 10px;
        color: #1d1d1f;
        font-weight: 500;
        border: none !important;
        transition: all 0.2s ease;
    }
    .stTabs [data-baseweb="tab"]:hover {
        background-color: rgba(255, 255, 255, 0.4);
    }
    .stTabs [aria-selected="true"] {
        background-color: #ffffff !important;
        box-shadow: 0px 2px 8px rgba(0, 0, 0, 0.08) !important;
    }
    </style>
    """, unsafe_allow_html=True)

# พ่นส่วนหัวหน้าเว็บแอป (Apple Premium Glass Container)
st.markdown("""
    <div class="apple-header">
        <div class="apple-title">Intanin Receipt Convert Pro</div>
        <div class="apple-subtitle">Logistics Automation Engine • Version 15.0 Premium Edition</div>
    </div>
    """, unsafe_allow_html=True)

if 'name_memory' not in st.session_state:
    st.session_state['name_memory'] = {}

# ย้ายโซนอัปโหลดไฟล์มาตรงกลางจอแบบหล่อๆ
uploaded_file = st.file_uploader("📂 Drop your Transport Excel file here or browse", type=['xlsx'])

# ฟังก์ชันทำความสะอาดรหัสและลบเศษ .0 ออกอย่างแม่นยำ
def clean_code_to_str(val):
    if pd.isna(val):
        return ""
    val_str = str(val).strip()
    if val_str.endswith('.0'):
        val_str = val_str[:-2]
    return val_str

if uploaded_file:
    with st.spinner('Reading data securely...'):
        df_raw = pd.read_excel(uploaded_file, sheet_name='Transport')
    
    # ========================================================
    # 🔍 SYSTEM: DATA VALIDATION & SMART PREVIEW (CORE V14.3)
    # ========================================================
    df_clean_rows = df_raw.dropna(how='all').copy()
    
    col_cutoff = [c for c in df_clean_rows.columns if 'cut' in c.lower() and 'off' in c.lower()]
    col_delivery = [c for c in df_clean_rows.columns if 'deliv' in c.lower() and 'date' in c.lower()]
    
    name_cutoff = col_cutoff[0] if col_cutoff else 'Cut Off Date'
    name_delivery = col_delivery[0] if col_delivery else 'Delivery Date'
    
    df_invalid = df_clean_rows[df_clean_rows[name_cutoff].isna() | df_clean_rows[name_delivery].isna()].copy()
    
    if len(df_invalid) > 0:
        st.markdown("<h3 style='font-size:1.3rem; font-weight:600; color:#ff3b30;'>🔍 Issues Detected (Data Validation)</h3>", unsafe_allow_html=True)
        st.error(f"พบช่องว่างในคอลัมน์วันที่ ทั้งหมด {len(df_invalid)} แถว ระบบกรองแถวว่างเปล่าออกให้แล้ว")
        
        preview_cols = []
        for c in ['รหัสสาขา', 'Store Name', 'โซน', 'รหัสสินค้า', 'รายการสินค้า', name_cutoff, name_delivery]:
            if c in df_clean_rows.columns:
                preview_cols.append(c)
                
        df_preview = df_invalid[preview_cols].copy()
        
        if 'รหัสสินค้า' in df_preview.columns:
            df_preview['รหัสสินค้า'] = df_preview['รหัสสินค้า'].apply(clean_code_to_str)
        if 'รหัสสาขา' in df_preview.columns:
            df_preview['รหัสสาขา'] = df_preview['รหัสสาขา'].apply(clean_code_to_str)
            
        df_preview.index = df_preview.index + 2
        df_preview.index.name = 'Excel Row'
        
        st.dataframe(df_preview, use_container_width=True)
        st.markdown("---")
    
    # ==========================================
    # WORKSPACE TABS (iOS Navigation Style)
    # ==========================================
    tab1, tab2 = st.tabs(["✂️ Supplier Splitter", "📄 Delivery Order Master"])

    # --- TAB 1: ระบบแยกซัพพลายเออร์ ---
    with tab1:
        df_split = df_raw.dropna(subset=['ซัพพลายเออร์']).copy()
        unique_suppliers = sorted(df_split['ซัพพลายเออร์'].unique())
        
        st.markdown("<h4 style='font-weight:600; margin-bottom:1rem; color:#1d1d1f;'>📝 Sheet Name Customization</h4>", unsafe_allow_html=True)
        with st.form("sheet_name_form"):
            cols = st.columns(3)
            current_mapping = {}
            for i, supplier in enumerate(unique_suppliers):
                with cols[i % 3]:
                    remembered_name = st.session_state['name_memory'].get(supplier, str(supplier)[:10].strip())
                    current_mapping[supplier] = st.text_input(f"{supplier}:", value=remembered_name, key=f"input_{supplier}")
            submit_split = st.form_submit_button("Generate Supplier Sheets")

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
                    
                    # ตารางฝั่งซ้าย (8 คอลัมน์ ล็อคโครงสร้าง V13.1+)
                    left_headers = ['รหัสสาขา', 'สาขา', 'โซน', 'รหัสสินค้า', 'รายการสินค้า', 'ซัพพลายเออร์', 'หน่วย', 'Total']
                    for col, h in enumerate(left_headers): 
                        worksheet.write(0, col, h, header_fmt)

                    curr_row = 1
                    for (branch, store_name, zone), b_group in df_sup.groupby(['รหัสสาขา', 'Store Name', 'โซน'], sort=False):
                        for i, (_, row) in enumerate(b_group.iterrows()):
                            branch_clean = clean_code_to_str(branch)
                            p_code_clean = clean_code_to_str(row['รหัสสินค้า'])
                            
                            worksheet.write(curr_row, 0, branch_clean if i == 0 else "", cell_fmt)
                            worksheet.write(curr_row, 1, store_name if i == 0 else "", cell_fmt)
                            worksheet.write(curr_row, 2, zone if i == 0 else "", cell_fmt)
                            worksheet.write(curr_row, 3, p_code_clean, cell_fmt)
                            worksheet.write(curr_row, 4, row['รายการสินค้า'], cell_fmt)
                            worksheet.write(curr_row, 5, row['ซัพพลายเออร์'], cell_fmt)
                            worksheet.write(curr_row, 6, row['หน่วย'], cell_fmt)
                            worksheet.write(curr_row, 7, row['จำนวน'], num_fmt)
                            curr_row += 1
                    
                    worksheet.write(curr_row, 0, "Grand Total", total_fmt)
                    worksheet.write(curr_row, 7, df_sup['จำนวน'].sum(), total_fmt)
                    
                    # ตารางฝั่งขวา (Summary ล็อคระยะห่าง)
                    col_offset = 10  
                    right_headers = ['ซัพพลายเออร์', 'รหัสสินค้า', 'รายการสินค้า', 'Total']
                    for col, h in enumerate(right_headers): 
                        worksheet.write(0, col_offset + col, h, header_fmt)
                    
                    right_summary = df_sup.groupby(['รหัสสินค้า', 'รายการสินค้า'], as_index=False)['จำนวน'].sum()
                    for i, row in right_summary.iterrows():
                        p_code_right_clean = clean_code_to_str(row['รหัสสินค้า'])
                        worksheet.write(i + 1, col_offset, "", cell_fmt)
                        worksheet.write(i + 1, col_offset + 1, p_code_right_clean, cell_fmt)
                        worksheet.write(i + 1, col_offset + 2, row['รายการสินค้า'], cell_fmt)
                        worksheet.write(i + 1, col_offset + 3, row['จำนวน'], num_fmt)
                    
                    sum_row = len(right_summary) + 1
                    worksheet.write(sum_row, col_offset, f"{supplier} Total", total_fmt)
                    worksheet.write(sum_row, col_offset + 3, right_summary['จำนวน'].sum(), total_fmt)
                    
                    worksheet.set_column('A:A', 15)  
                    worksheet.set_column('B:B', 30)  
                    worksheet.set_column('C:C', 15)  
                    worksheet.set_column('D:D', 15)  
                    worksheet.set_column('E:E', 35)  
                    worksheet.set_column('F:H', 15)  
                    worksheet.set_column('K:K', 15)  
                    worksheet.set_column('L:L', 15)  
                    worksheet.set_column('M:M', 35)  
                    worksheet.set_column('N:N', 15)  

            st.success("🎉 Splitter conversion ready!")
            st.download_button(label="📥 Download Splitter Excel", data=output_split.getvalue(), file_name="Intanin_Splitter_Pro.xlsx")

    # --- TAB 2: ออกใบ DO (แก้ปัญหาทศนิยม .0 ค้างจากรูปตัวอย่างเรียบร้อย) ---
    with tab2:
        df_clean = df_raw.dropna(subset=['รหัสสาขา']).copy()
        st.markdown("<h4 style='font-weight:600; margin-bottom:1rem; color:#1d1d1f;'>📄 Batch Document Processing</h4>", unsafe_allow_html=True)
        if st.button("🚀 Compile Official Delivery Orders"):
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
                    
                    store_code_clean = clean_code_to_str(store_code)
                    
                    worksheet.merge_range(curr, 0, curr, 2, 'บริษัท โมบาย โลจิสติกส์ จำกัด', f_comp)
                    worksheet.merge_range(curr+1, 0, curr+1, 1, '279 หมู่ที่ 9 ตำบลบางโฉลง อำเภอบางพลี', f_std)
                    worksheet.merge_range(curr+2, 0, curr+2, 1, 'จังหวัดสมุทรปราการ 10540', f_std)
                    worksheet.merge_range(curr+3, 0, curr+3, 2, 'ติดต่อ/สอบถาม : ID Line Official : @505phsps (มี @ ), Tel : 099-157-3114', f_std)
                    worksheet.write(curr+4, 0, 'Customer Name', f_std); worksheet.write(curr+5, 0, f'Store Code: {store_code_clean}', f_bold)
                    worksheet.write(curr+6, 0, f'Store Name: {first["Store Name"]}', f_bold); worksheet.write(curr+7, 0, f'Ship To: {first["ที่อยู่"]}', f_std)
                    worksheet.merge_range(curr+1, 2, curr+2, 2, 'ใบส่งสินค้าชั่วคราว', f_title)
                    worksheet.write(curr, 3, 'Do. No.', f_right); worksheet.write(curr, 4, clean_code_to_str(first["เลขที่ DO."]), f_std)
                    worksheet.write(curr+1, 3, 'Ref. Po.', f_right); worksheet.write(curr+1, 4, clean_code_to_str(first["เลขที่ PO."]), f_std)
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
                        product_code_clean = clean_code_to_str(r['รหัสสินค้า'])
                        
                        worksheet.write(r_ptr, 0, i, f_border)
                        worksheet.write(r_ptr, 1, product_code_clean, f_border) 
                        worksheet.write(r_ptr, 2, r['รายการสินค้า'], f_wrap)
                        worksheet.write(r_ptr, 3, r['หน่วย'], f_border)
                        worksheet.write(r_ptr, 4, r['จำนวน'], f_border)
                        r_ptr += 1
                        
                    worksheet.merge_range(r_ptr, 0, r_ptr, 3, 'Total', f_table_h); worksheet.write(r_ptr, 4, df_store['จำนวน'].sum(), f_border)
                    f_row = r_ptr + 1
                    worksheet.merge_range(f_row, 0, f_row, 1, 'ผู้รับสินค้า', f_footer_h); worksheet.merge_range(f_row, 2, f_row, 3, 'ผู้ส่งสินค้า / ทะเบียนรถ', f_footer_h); worksheet.write(f_row, 4, 'คลังสินค้า', f_footer_h)
                    for label in ['ชื่อ (ตัวบรรจง):', 'วันที่:', 'เวลา:', 'หมายเหตุ:']:
                        f_row += 1
                        worksheet.merge_range(f_row, 0, f_row, 1, label, f_footer_box); worksheet.merge_range(f_row, 2, f_row, 3, label, f_footer_box); worksheet.write(f_row, 4, label, f_footer_box)
                    curr = f_row + 2
                worksheet.set_h_pagebreaks(page_breaks); worksheet.fit_to_pages(1, 0)

            st.success("🎉 Delivery Orders compiled flawlessly!")
            st.download_button("📥 Download DO Master Package", output_do.getvalue(), "Intanin_DO_Master_Pro.xlsx")
