import streamlit as st
import pandas as pd
import io

# --- CONFIG & DYNAMIC SYSTEM UI (Support Dark/Light Mode) ---
st.set_page_config(page_title="Intanin Receipt Convert", layout="wide", page_icon="📦")

# ปรับ CSS ให้ยืดหยุ่นตามโหมดสีของระบบ (Auto Dark/Light)
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Sarabun:wght@300;400;500;700&display=swap');
    
    html, body, [class*="css"] { 
        font-family: 'Sarabun', sans-serif; 
    }
    
    /* ส่วนหัวคลีนๆ สไตล์ iOS Settings - รองรับ Dark/Light Mode */
    .bd-header { 
        background: rgba(128, 128, 128, 0.08);
        backdrop-filter: blur(15px);
        -webkit-backdrop-filter: blur(15px);
        padding: 1.8rem 1rem; 
        border-radius: 16px;
        border: 1px solid rgba(128, 128, 128, 0.15);
        margin-bottom: 2rem;
        text-align: center;
    }
    .bd-title {
        font-weight: 700;
        font-size: 2.2rem;
        letter-spacing: -0.02em;
        /* ใช้ไล่เฉดสีเขียวที่มองเห็นได้ชัดทั้งบนพื้นขาวและพื้นดำ */
        background: linear-gradient(135deg, #2e7d32 0%, #4caf50 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        margin-bottom: 0.2rem;
    }
    .bd-subtitle {
        color: var(--text-color);
        opacity: 0.6;
        font-size: 0.95rem;
        font-weight: 400;
    }
    
    /* สไตล์ปุ่มกด Micro-interaction แบบพรีเซตสีไม่ทับซ้อนสีฟอนต์ระบบ */
    div.stButton > button, div.stDownloadButton > button {
        background: linear-gradient(180deg, #34c759 0%, #28cd41 100%) !important;
        color: #ffffff !important; /* บังคับให้ปุ่มเป็นสีขาวเสมอเพื่อความชัดเจน */
        border-radius: 10px !important;
        border: none !important;
        padding: 0.5rem 1.8rem !important;
        font-weight: 500 !important;
        box-shadow: 0 4px 10px rgba(52, 199, 89, 0.15) !important;
        transition: all 0.2s ease !important;
    }
    div.stButton > button:hover, div.stDownloadButton > button:hover {
        transform: translateY(-1px) !important;
        box-shadow: 0 6px 15px rgba(52, 199, 89, 0.25) !important;
    }
    
    /* ปรับแต่งแท็บให้เข้ากับธีมปัจจุบัน */
    .stTabs [data-baseweb="tab-list"] {
        gap: 6px;
        background-color: rgba(128, 128, 128, 0.12);
        padding: 4px;
        border-radius: 12px;
    }
    .stTabs [data-baseweb="tab"] {
        border-radius: 8px;
        color: var(--text-color) !important;
        border: none !important;
    }
    </style>
    """, unsafe_allow_html=True)

# ส่วนหัวฉบับมินิมอลทางการโดยทีม Business Development
st.markdown("""
    <div class="bd-header">
        <div class="bd-title">Intanin Receipt Convert</div>
        <div class="bd-subtitle">Mobile logistics @Business Development</div>
    </div>
    """, unsafe_allow_html=True)

if 'name_memory' not in st.session_state:
    st.session_state['name_memory'] = {}

uploaded_file = st.file_uploader("📂 อัปโหลดไฟล์ Excel (Transport)", type=['xlsx'])

# ฟังก์ชันจัดการแปลงรูปแบบและลบเศษทศนิยม .0 ออกจากรหัสต่างๆ
def clean_code_to_str(val):
    if pd.isna(val):
        return ""
    val_str = str(val).strip()
    if val_str.endswith('.0'):
        val_str = val_str[:-2]
    return val_str

if uploaded_file:
    with st.spinner('กำลังอ่านข้อมูล...'):
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
        st.markdown("<h4 style='font-weight:600; color:#ff4b4b;'>🔍 ตรวจพบข้อมูลไม่ครบถ้วนในแถวงาน</h4>", unsafe_allow_html=True)
        st.error(f"พบช่องว่างในคอลัมน์วันที่ ทั้งหมด {len(df_invalid)} แถว (ระบบละเว้นแถวเปล่าท้ายไฟล์ให้แล้ว)")
        
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
    # WORKSPACE TABS
    # ==========================================
    tab1, tab2 = st.tabs(["✂️ 1. แยกซัพพลายเออร์ (Supplier Splitter)", "📄 2. ออกใบส่งสินค้า (DO Master)"])

    # --- TAB 1: ระบบแยกซัพพลายเออร์ ---
    with tab1:
        df_split = df_raw.dropna(subset=['ซัพพลายเออร์']).copy()
        unique_suppliers = sorted(df_split['ซัพพลายเออร์'].unique())
        
        st.markdown("<h5 style='font-weight:600; margin-bottom:1rem;'>📝 ตั้งชื่อตัวย่อชีตซัพพลายเออร์</h5>", unsafe_allow_html=True)
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
                    
                    # ตารางฝั่งซ้าย (8 คอลัมน์)
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
                    
                    # ตารางฝั่งขวา (Summary)
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

            st.success("✅ แยกข้อมูลซัพพลายเออร์สำเร็จ!")
            st.download_button(label="📥 ดาวน์โหลดไฟล์แยกซัพพลายเออร์", data=output_split.getvalue(), file_name="Intanin_Splitter_BD.xlsx")

    # --- TAB 2: ออกใบ DO ---
    with tab2:
        df_clean = df_raw.dropna(subset=['รหัสสาขา']).copy()
        st.markdown("<h5 style='font-weight:600; margin-bottom:1rem;'>📄 ออกใบส่งสินค้าชุดใหญ่ (DO)</h5>", unsafe_allow_html=True)
        if st.button("🚀 ประมวลผลสร้างใบส่งสินค้า (DO)"):
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

            st.success("✅ สร้างใบส่งสินค้าเรียบร้อย")
            st.download_button("📥 ดาวน์โหลดไฟล์ DO Master", output_do.getvalue(), "Intanin_DO_BD_Edition.xlsx")
