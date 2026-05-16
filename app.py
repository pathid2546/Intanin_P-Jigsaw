import streamlit as st
import pandas as pd
import io
import plotly.express as px  # เพิ่มไลบรารีกราฟ Interactive สวยงามสไตล์โมเดิร์น

# --- CONFIG & DYNAMIC SYSTEM UI (Support Dark/Light Mode) ---
st.set_page_config(page_title="Intanin BD System", layout="wide", page_icon="📊")

# ปรับ CSS รองรับ Dark/Light Mode และแต่งสไตล์ Sidebar
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
    
    /* สไตล์ปุ่มกด Micro-interaction */
    div.stButton > button, div.stDownloadButton > button {
        background: linear-gradient(180deg, #34c759 0%, #28cd41 100%) !important;
        color: #ffffff !important;
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
    
    /* สรุป KPI การ์ดในหน้า Dashboard */
    .kpi-card {
        background: rgba(128, 128, 128, 0.05);
        border: 1px solid rgba(128, 128, 128, 0.12);
        padding: 1.2rem;
        border-radius: 12px;
        text-align: center;
    }
    
    /* ปรับแต่งแท็บ */
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
        <div class="bd-subtitle">Business Development Logistics & Analytics System</div>
    </div>
    """, unsafe_allow_html=True)

if 'name_memory' not in st.session_state:
    st.session_state['name_memory'] = {}

# ฟังก์ชันจัดการแปลงรูปแบบและลบเศษทศนิยม .0 ออกจากรหัสต่างๆ
def clean_code_to_str(val):
    if pd.isna(val):
        return ""
    val_str = str(val).strip()
    if val_str.endswith('.0'):
        val_str = val_str[:-2]
    return val_str

# ==========================================
# 🧭 SIDEBAR NAVIGATION
# ==========================================
st.sidebar.markdown("### 🍏 เมนูควบคุมระบบ")
app_mode = st.sidebar.radio(
    "เลือกหน้าต่างการใช้งาน:",
    ["📄 Convert Data (จัดการเอกสาร)", "📊 Analytics Dashboard (วิเคราะห์ข้อมูล)"]
)

st.sidebar.markdown("---")
st.sidebar.caption("Powered by Business Development Team • 2026")

# 📂 อัปโหลดไฟล์ส่วนกลาง (เก็บลง session_state เพื่อให้สลับหน้าแล้วข้อมูลไม่หาย)
uploaded_file = st.file_uploader("📂 อัปโหลดไฟล์ Excel (Transport)", type=['xlsx'])

if uploaded_file:
    if 'cached_file_name' not in st.session_state or st.session_state['cached_file_name'] != uploaded_file.name:
        with st.spinner('กำลังอ่านข้อมูล...'):
            st.session_state['df_raw'] = pd.read_excel(uploaded_file, sheet_name='Transport')
            st.session_state['cached_file_name'] = uploaded_file.name

if 'df_raw' in st.session_state:
    df_raw = st.session_state['df_raw']
    
    # ดึงแถวที่มีข้อมูลออกมา ลบแถวขยะว่างๆ ออกไปก่อนตรวจ
    df_clean_rows = df_raw.dropna(how='all').copy()
    
    # ตรวจหาชื่อคอลัมน์วันที่
    col_cutoff = [c for c in df_clean_rows.columns if 'cut' in c.lower() and 'off' in c.lower()]
    col_delivery = [c for c in df_clean_rows.columns if 'deliv' in c.lower() and 'date' in c.lower()]
    name_cutoff = col_cutoff[0] if col_cutoff else 'Cut Off Date'
    name_delivery = col_delivery[0] if col_delivery else 'Delivery Date'

    # ========================================================
    # 📄 MODE 1: CONVERT DATA
    # ========================================================
    if app_mode == "📄 Convert Data (จัดการเอกสาร)":
        
        # ตรวจสอบ Validation พรีวิวแถวที่ขาดวันที่ (ละเว้นแถวเปล่า)
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
        
        # แท็บงานย่อย
        tab1, tab2 = st.tabs(["✂️ 1. แยกซัพพลายเออร์ (Supplier Splitter)", "📄 2. ออกใบส่งสินค้า (DO Master)"])

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
                        
                        left_headers = ['รหัสสาขา', 'สาขา', 'โซน', 'รหัสสินค้า', 'รายการสินค้า', 'ซัพพลายเออร์', 'หน่วย', 'Total']
                        for col, h in enumerate(left_headers): worksheet.write(0, col, h, header_fmt)

                        curr_row = 1
                        for (branch, store_name, zone), b_group in df_sup.groupby(['รหัสสาขา', 'Store Name', 'โซน'], sort=False):
                            for i, (_, row) in enumerate(b_group.iterrows()):
                                worksheet.write(curr_row, 0, clean_code_to_str(branch) if i == 0 else "", cell_fmt)
                                worksheet.write(curr_row, 1, store_name if i == 0 else "", cell_fmt)
                                worksheet.write(curr_row, 2, zone if i == 0 else "", cell_fmt)
                                worksheet.write(curr_row, 3, clean_code_to_str(row['รหัสสินค้า']), cell_fmt)
                                worksheet.write(curr_row, 4, row['รายการสินค้า'], cell_fmt)
                                worksheet.write(curr_row, 5, row['ซัพพลายเออร์'], cell_fmt)
                                worksheet.write(curr_row, 6, row['หน่วย'], cell_fmt)
                                worksheet.write(curr_row, 7, row['จำนวน'], num_fmt)
                                curr_row += 1
                        
                        worksheet.write(curr_row, 0, "Grand Total", total_fmt)
                        worksheet.write(curr_row, 7, df_sup['จำนวน'].sum(), total_fmt)
                        
                        col_offset = 10  
                        right_headers = ['ซัพพลายเออร์', 'รหัสสินค้า', 'รายการสินค้า', 'Total']
                        for col, h in enumerate(right_headers): worksheet.write(0, col_offset + col, h, header_fmt)
                        
                        right_summary = df_sup.groupby(['รหัสสินค้า', 'รายการสินค้า'], as_index=False)['จำนวน'].sum()
                        for i, row in right_summary.iterrows():
                            worksheet.write(i + 1, col_offset, "", cell_fmt)
                            worksheet.write(i + 1, col_offset + 1, clean_code_to_str(row['รหัสสินค้า']), cell_fmt)
                            worksheet.write(i + 1, col_offset + 2, row['รายการสินค้า'], cell_fmt)
                            worksheet.write(i + 1, col_offset + 3, row['จำนวน'], num_fmt)
                        
                        sum_row = len(right_summary) + 1
                        worksheet.write(sum_row, col_offset, f"{supplier} Total", total_fmt)
                        worksheet.write(sum_row, col_offset + 3, right_summary['จำนวน'].sum(), total_fmt)
                        
                        worksheet.set_column('A:A', 15); worksheet.set_column('B:B', 30); worksheet.set_column('C:C', 15)
                        worksheet.set_column('D:D', 15); worksheet.set_column('E:E', 35); worksheet.set_column('F:H', 15)
                        worksheet.set_column('K:K', 15); worksheet.set_column('L:L', 15); worksheet.set_column('M:M', 35); worksheet.set_column('N:N', 15)

                st.success("✅ แยกข้อมูลซัพพลายเออร์สำเร็จ!")
                st.download_button(label="📥 ดาวน์โหลดไฟล์แยกซัพพลายเออร์", data=output_split.getvalue(), file_name="Intanin_Splitter_BD.xlsx")

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
                        
                        worksheet.merge_range(curr, 0, curr, 2, 'บริษัท โมบาย โลจิสติกส์ จำกัด', f_comp)
                        worksheet.merge_range(curr+1, 0, curr+1, 1, '279 หมู่ที่ 9 ตำบลบางโฉลง อำเภอบางพลี', f_std)
                        worksheet.merge_range(curr+2, 0, curr+2, 1, 'จังหวัดสมุทรปราการ 10540', f_std)
                        worksheet.merge_range(curr+3, 0, curr+3, 2, 'ติดต่อ/สอบถาม : ID Line Official : @505phsps (มี @ ), Tel : 099-157-3114', f_std)
                        worksheet.write(curr+4, 0, 'Customer Name', f_std); worksheet.write(curr+5, 0, f'Store Code: {clean_code_to_str(store_code)}', f_bold)
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
                            worksheet.write(r_ptr, 0, i, f_border)
                            worksheet.write(r_ptr, 1, clean_code_to_str(r['รหัสสินค้า']), f_border) 
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

                st.success("✅ สร้างใบส่งสินค้าเรียบร้อย รหัสไม่มีจุดทศนิยมกวนใจ!")
                st.download_button("📥 ดาวน์โหลดไฟล์ DO Master", output_do.getvalue(), "Intanin_DO_BD_Edition.xlsx")

    # ========================================================
    # 📊 MODE 2: ANALYTICS DASHBOARD (NEW FEATURES)
    # ========================================================
    elif app_mode == "📊 Analytics Dashboard (วิเคราะห์ข้อมูล)":
        st.markdown("<h3 style='font-weight:600;'>📊 Business Development Data Analysis</h3>", unsafe_allow_html=True)
        
        # ทำความสะอาดข้อมูลเบื้องต้นเพื่อใช้ในการคำนวณและวาดกราฟ
        df_dash = df_clean_rows.copy()
        df_dash['รหัสสาขา'] = df_dash['รหัสสาขา'].apply(clean_code_to_str)
        df_dash['รหัสสินค้า'] = df_dash['รหัสสินค้า'].apply(clean_code_to_str)
        # ตรวจเช็คชื่อคอลัมน์ปริมาณ (จำนวน/QTY)
        col_qty = 'จำนวน' if 'จำนวน' in df_dash.columns else ('QTY' if 'QTY' in df_dash.columns else df_dash.columns[-1])
        df_dash[col_qty] = pd.to_numeric(df_dash[col_qty], errors='coerce').fillna(0)
        
        # --- 1. ส่วนบนสุด: สรุปตัวเลขสำคัญ (KPI Cards) ---
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            st.markdown(f"<div class='kpi-card'>🏷️ <b>จำนวนซัพพลายเออร์</b><br><span style='font-size:1.8rem; font-weight:700; color:#2e7d32;'>{df_dash['ซัพพลายเออร์'].nunique() if 'ซัพพลายเออร์' in df_dash.columns else 0}</span> ราย</div>", unsafe_allow_html=True)
        with c2:
            st.markdown(f"<div class='kpi-card'>🏪 <b>สาขาที่สั่งซื้อทั้งหมด</b><br><span style='font-size:1.8rem; font-weight:700; color:#2e7d32;'>{df_dash['Store Name'].nunique() if 'Store Name' in df_dash.columns else 0}</span> สาขา</div>", unsafe_allow_html=True)
        with c3:
            st.markdown(f"<div class='kpi-card'>📦 <b>รายการสินค้าทั้งหมด</b><br><span style='font-size:1.8rem; font-weight:700; color:#2e7d32;'>{df_dash['รายการสินค้า'].nunique() if 'รายการสินค้า' in df_dash.columns else 0}</span> SKUs</div>", unsafe_allow_html=True)
        with c4:
            st.markdown(f"<div class='kpi-card'>🚚 <b>ปริมาณการสั่งซื้อรวม</b><br><span style='font-size:1.8rem; font-weight:700; color:#2e7d32;'>{int(df_dash[col_qty].sum()):,}</span> ชิ้น</div>", unsafe_allow_html=True)
            
        st.markdown("<br>", unsafe_allow_html=True)
        
        # --- 2. ส่วนควบคุมฟิลเตอร์และประเภทกราฟ ---
        col_ctrl1, col_ctrl2 = st.columns([1, 2])
        with col_ctrl1:
            st.markdown("##### ⚙️ ปรับแต่งเครื่องมือกราฟ")
            chart_type = st.radio("รูปแบบการแสดงผลที่ต้องการ:", ["📊 กราฟแท่ง (Bar Chart)", "🍩 กราฟโดนัท (Donut Chart)"])
            top_n = st.slider("จำนวนสาขาอันดับแรกที่ต้องการดู (Top N):", min_value=3, max_value=30, value=10)
            
        with col_ctrl2:
            st.markdown("##### 🎯 ฟิลเตอร์กรองเลือกข้อมูลเฉพาะกลุ่ม")
            # ถ้ามีคอลัมน์โซน ให้เลือกกรองได้
            zone_options = ["ทั้งหมด"] + sorted(list(df_dash['โซน'].dropna().unique())) if 'โซน' in df_dash.columns else ["ทั้งหมด"]
            selected_zone = st.selectbox("เลือกดูข้อมูลเฉพาะ โซน:", zone_options)
            
            # กรองซัพพลายเออร์
            sup_options = ["ทั้งหมด"] + sorted(list(df_dash['ซัพพลายเออร์'].dropna().unique())) if 'ซัพพลายเออร์' in df_dash.columns else ["ทั้งหมด"]
            selected_sup = st.selectbox("เลือกดูข้อมูลเฉพาะ ซัพพลายเออร์:", sup_options)

        # นำฟิลเตอร์ที่เลือกไปกรองข้อมูลจริง
        if selected_zone != "ทั้งหมด":
            df_dash = df_dash[df_dash['โซน'] == selected_zone]
        if selected_sup != "ทั้งหมด":
            df_dash = df_dash[df_dash['ซัพพลายเออร์'] == selected_sup]

        st.markdown("---")

        # --- 3. ส่วนพล็อตกราฟจัดอันดับสาขาที่สั่งเยอะที่สุด ---
        st.markdown(f"##### 🏪 อันดับสาขาที่มียอดสั่งซื้อสูงสุด {top_n} อันดับแรก")
        
        # จับกลุ่มข้อมูลปริมาณการสั่งซื้อรายสาขา
        df_branch_summary = df_dash.groupby('Store Name', as_index=False)[col_qty].sum()
        df_branch_summary = df_branch_summary.sort_values(by=col_qty, ascending=False).head(top_n)
        
        if not df_branch_summary.empty:
            if chart_type == "📊 กราฟแท่ง (Bar Chart)":
                fig = px.bar(
                    df_branch_summary, 
                    x=col_qty, 
                    y='Store Name', 
                    orientation='h',
                    text_auto=',',
                    labels={col_qty: 'ปริมาณการสั่งซื้อ (ชิ้น)', 'Store Name': 'ชื่อสาขา'},
                    color=col_qty,
                    color_continuous_scale='Greens'
                )
                fig.update_layout(yaxis={'categoryorder':'total ascending'}, margin=dict(l=20, r=20, t=20, b=20), height=450)
                st.plotly_chart(fig, use_container_width=True)
                
            else:  # Donut Chart
                fig = px.pie(
                    df_branch_summary, 
                    values=col_qty, 
                    names='Store Name', 
                    hole=0.5,
                    labels={col_qty: 'ปริมาณ', 'Store Name': 'ชื่อสาขา'}
                )
                fig.update_traces(textinfo='percent+label', marker=dict(colors=px.colors.sequential.Greens_r))
                fig.update_layout(margin=dict(l=20, r=20, t=20, b=20), height=450)
                st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("💡 ไม่พบข้อมูลตามฟิลเตอร์ที่คุณเลือก กรุณาปรับตัวเลือกตัวกรองใหม่")

        # --- 4. ส่วนวิเคราะห์เพิ่มเติมที่ทีม BD ควรมี (Recommended Insight) ---
        st.markdown("---")
        st.markdown("##### 💡 มุมมองข้อมูลเพิ่มเติมที่แนะนำสำหรับงาน BD: สินค้าตัวไหนขายดีที่สุด (Top Products Summary)")
        
        # ค้นหาสินค้าขายดีเรียงลำดับลงมา
        df_prod_summary = df_dash.groupby(['รหัสสินค้า', 'รายการสินค้า'], as_index=False)[col_qty].sum()
        df_prod_summary = df_prod_summary.sort_values(by=col_qty, ascending=False).rename(columns={col_qty: 'ปริมาณสั่งซื้อรวม (ชิ้น)'}).reset_index(drop=True)
        df_prod_summary.index = df_prod_summary.index + 1
        
        col_table, col_pie = st.columns([3, 2])
        with col_table:
            st.caption("ตารางจัดลำดับความนิยมของสินค้าภายในข้อมูลที่อัปโหลด")
            st.dataframe(df_prod_summary, use_container_width=True)
            
        with col_pie:
            st.caption("สัดส่วนซัพพลายเออร์ที่ครองตารางงานส่ง ณ ขณะนี้")
            if 'ซัพพลายเออร์' in df_dash.columns:
                df_sup_share = df_dash.groupby('ซัพพลายเออร์')[col_qty].sum().reset_index()
                fig_sup = px.pie(df_sup_share, values=col_qty, names='ซัพพลายเออร์', hole=0.3)
                fig_sup.update_layout(margin=dict(l=10, r=10, t=10, b=10), height=280, showlegend=True)
                st.plotly_chart(fig_sup, use_container_width=True)

else:
    st.info("👋 ยินดีต้อนรับ! กรุณาทำการอัปโหลดไฟล์เอกสาร Excel ที่ช่องด้านบนก่อน เพื่อเปิดใช้งานระบบประมวลผลและบอร์ดวิเคราะห์ข้อมูล")
