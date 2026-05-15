import streamlit as st
import pandas as pd
import io
import time

# --- CONFIG & THEME ---
st.set_page_config(page_title="Intanin Receipt Convert", layout="wide", page_icon="📦")

# ฉีด CSS สำหรับ Custom UI, Skeleton Loading และ Dark/Light Mode
st.markdown("""
    <style>
    /* Global Style */
    @import url('https://fonts.googleapis.com/css2?family=Sarabun:wght@300;400;700&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Sarabun', sans-serif;
    }

    /* Skeleton Loading Animation */
    @keyframes skeleton-loading {
        0% { background-position: 100% 50%; }
        100% { background-position: 0% 50%; }
    }
    
    .skeleton {
        background: linear-gradient(90deg, #f0f0f0 25%, #e0e0e0 50%, #f0f0f0 75%);
        background-size: 200% 100%;
        animation: skeleton-loading 1.5s infinite;
        border-radius: 8px;
        margin-bottom: 10px;
    }
    
    .skeleton-text { height: 20px; width: 80%; }
    .skeleton-title { height: 40px; width: 40%; }
    .skeleton-card { height: 150px; width: 100%; }

    /* Theme-specific adjustments */
    [data-theme="light"] {
        --bg-header: #f8f9fa;
        --text-main: #2c3e50;
    }
    [data-theme="dark"] {
        --bg-header: #1e1e1e;
        --text-main: #ecf0f1;
    }

    .main-header {
        padding: 2rem;
        border-bottom: 2px solid #2e7d32;
        margin-bottom: 2rem;
        background-color: transparent;
    }
    </style>
    """, unsafe_allow_html=True)

# --- HEADER SECTION ---
st.markdown("""
    <div class="main-header">
        <h1 style='text-align: center;'>🚛 Intanin Receipt Convert</h1>
        <p style='text-align: center; color: #666;'>Logistics Management & Document Processing System V12.6</p>
    </div>
    """, unsafe_allow_html=True)

# ระบบจำชื่อ (Memory)
if 'name_memory' not in st.session_state:
    st.session_state['name_memory'] = {}

# --- FILE UPLOADER ---
with st.container():
    uploaded_file = st.file_uploader("📂 เลือกไฟล์ Excel (Transport) เพื่อเริ่มดำเนินการ", type=['xlsx'])

if uploaded_file:
    # แสดง Skeleton Loading ขณะอ่านไฟล์
    with st.empty():
        st.markdown('<div class="skeleton skeleton-title"></div>', unsafe_allow_html=True)
        st.markdown('<div class="skeleton skeleton-card"></div>', unsafe_allow_html=True)
        df_raw = pd.read_excel(uploaded_file, sheet_name='Transport')
        time.sleep(0.5) # เพื่อให้เห็น Skeleton สั้นๆ ให้ความรู้สึกระบบกำลังประมวลผล
        st.write("") # Clear skeleton

    tab1, tab2 = st.tabs(["✂️ Supplier Splitter", "📄 DO Generator"])

    # --- TAB 1: Splitter ---
    with tab1:
        st.subheader("📦 จัดการแยกซัพพลายเออร์")
        df_split = df_raw.dropna(subset=['ซัพพลายเออร์'])
        unique_suppliers = sorted(df_split['ซัพพลายเออร์'].unique())
        
        with st.form("sheet_name_form"):
            cols = st.columns(3)
            current_mapping = {}
            for i, supplier in enumerate(unique_suppliers):
                with cols[i % 3]:
                    remembered_name = st.session_state['name_memory'].get(supplier, str(supplier)[:10].strip())
                    current_mapping[supplier] = st.text_input(f"📍 {supplier}", value=remembered_name, key=f"input_{supplier}")
            submit_split = st.form_submit_button("⚙️ ประมวลผลและสร้างไฟล์")

        if submit_split:
            with st.status("กำลังเตรียมข้อมูล...", expanded=True) as status:
                st.write("แยกข้อมูลตามซัพพลายเออร์...")
                output_split = io.BytesIO()
                with pd.ExcelWriter(output_split, engine='xlsxwriter') as writer:
                    workbook = writer.book
                    header_fmt = workbook.add_format({'bold': True, 'bg_color': '#E8F5E9', 'border': 1, 'align': 'center'})
                    cell_fmt = workbook.add_format({'border': 1})
                    num_fmt = workbook.add_format({'border': 1, 'align': 'right'})
                    total_fmt = workbook.add_format({'bold': True, 'bg_color': '#F5F5F5', 'border': 1, 'align': 'right', 'num_format': '#,##0'})
                    
                    for supplier, sheet_name in current_mapping.items():
                        st.session_state['name_memory'][supplier] = sheet_name
                        clean_name = "".join([c if c not in r'[]:*?/\ ' else ' ' for c in sheet_name.strip()[:31]])
                        df_sup = df_split[df_split['ซัพพลายเออร์'] == supplier].copy()
                        worksheet = workbook.add_worksheet(clean_name)
                        
                        # ฝั่งซ้าย
                        left_headers = ['รหัสสาขา', 'โซน', 'รหัสสินค้า', 'รายการสินค้า', 'ซัพพลายเออร์', 'หน่วย', 'Total']
                        for col, h in enumerate(left_headers): worksheet.write(0, col, h, header_fmt)
                        curr_row = 1
                        for (branch, zone), b_group in df_sup.groupby(['รหัสสาขา', 'โซน'], sort=False):
                            for i, (_, row) in enumerate(b_group.iterrows()):
                                worksheet.write(curr_row, 0, branch if i == 0 else "", cell_fmt)
                                worksheet.write(curr_row, 1, zone if i == 0 else "", cell_fmt)
                                worksheet.write(curr_row, 2, row['รหัสสินค้า'], cell_fmt)
                                worksheet.write(curr_row, 3, row['รายการสินค้า'], cell_fmt)
                                worksheet.write(curr_row, 4, row['ซัพพลายเออร์'], cell_fmt)
                                worksheet.write(curr_row, 5, row['หน่วย'], cell_fmt)
                                worksheet.write(curr_row, 6, row['จำนวน'], num_fmt)
                                curr_row += 1
                        worksheet.write(curr_row, 0, "Grand Total", total_fmt); worksheet.write(curr_row, 6, df_sup['จำนวน'].sum(), total_fmt)
                        
                        # ฝั่งขวา
                        col_offset = 9
                        right_headers = ['ซัพพลายเออร์', 'รหัสสินค้า', 'รายการสินค้า', 'Total']
                        for col, h in enumerate(right_headers): worksheet.write(0, col_offset + col, h, header_fmt)
                        right_summary = df_sup.groupby(['รหัสสินค้า', 'รายการสินค้า'], as_index=False)['จำนวน'].sum()
                        for i, row in right_summary.iterrows():
                            worksheet.write(i + 1, col_offset, "", cell_fmt); worksheet.write(i + 1, col_offset + 1, row['รหัสสินค้า'], cell_fmt)
                            worksheet.write(i + 1, col_offset + 2, row['รายการสินค้า'], cell_fmt); worksheet.write(i + 1, col_offset + 3, row['จำนวน'], num_fmt)
                        sum_row = len(right_summary) + 1
                        worksheet.write(sum_row, col_offset, f"{supplier} Total", total_fmt); worksheet.write(sum_row, col_offset + 3, right_summary['จำนวน'].sum(), total_fmt)
                        
                        worksheet.set_column('A:B', 15); worksheet.set_column('C:C', 15); worksheet.set_column('D:D', 35)
                        worksheet.set_column('J:K', 15); worksheet.set_column('L:L', 35); worksheet.set_column('M:M', 15)
                status.update(label="✅ ประมวลผลสำเร็จ!", state="complete")

            st.download_button(label="📥 Download Splitter File", data=output_split.getvalue(), file_name="Intanin_Splitter.xlsx")

    # --- TAB 2: DO Generator ---
    with tab2:
        st.subheader("📑 ออกใบส่งสินค้า (Official Form)")
        df_clean = df_raw.dropna(subset=['รหัสสาขา']).copy()
        if st.button("🚀 Generate DO Master"):
            with st.status("กำลังสร้างเอกสาร DO...", expanded=False):
                output_do = io.BytesIO()
                with pd.ExcelWriter(output_do, engine='xlsxwriter') as writer:
                    workbook = writer.book
                    worksheet = workbook.add_worksheet("DO_Master")
                    worksheet.set_paper(9); worksheet.set_margins(0.3, 0.3, 0.3, 0.3)
                    worksheet.set_column('A:A', 8); worksheet.set_column('B:B', 35); worksheet.set_column('C:C', 35)
                    worksheet.set_column('D:D', 18); worksheet.set_column('E:E', 22)
                    
                    f_comp = workbook.add_format({'bold': True, 'font_size': 14})
                    f_title = workbook.add_format({'bold': True, 'font_size': 18, 'align': 'center', 'valign': 'vcenter'})
                    f_std = workbook.add_format({'font_size': 11})
                    f_bold = workbook.add_format({'bold': True, 'font_size': 11})
                    f_right = workbook.add_format({'font_size': 11, 'align': 'right'})
                    f_deliv = workbook.add_format({'bold': True, 'font_size': 11, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True})
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
                        worksheet.merge_range(curr+3, 0, curr+3, 2, 'ติดต่อ : ID Line : @505phsps, Tel : 099-157-3114', f_std)
                        worksheet.write(curr+5, 0, f'Store Code: {store_code}', f_bold)
                        worksheet.write(curr+6, 0, f'Store Name: {first["Store Name"]}', f_bold)
                        worksheet.write(curr+7, 0, f'Ship To: {first["ที่อยู่"]}', f_std)
                        worksheet.merge_range(curr+1, 2, curr+2, 2, 'ใบส่งสินค้าชั่วคราว', f_title)

                        worksheet.write(curr, 3, 'Do. No.', f_right); worksheet.write(curr, 4, str(first["เลขที่ DO."]), f_std)
                        worksheet.write(curr+1, 3, 'Ref. Po.', f_right); worksheet.write(curr+1, 4, str(first["เลขที่ PO."]), f_std)
                        
                        cutoff = pd.to_datetime(first.get('Cut Off Date')).strftime('%d/%m/%Y') if pd.notnull(first.get('Cut Off Date')) else "-"
                        worksheet.write(curr+4, 3, 'Cut off Date', f_right); worksheet.write(curr+4, 4, cutoff, f_std)
                        worksheet.write(curr+5, 3, 'Zone', f_right); worksheet.write(curr+5, 4, str(first["โซน"]), f_std)
                        
                        delivery = pd.to_datetime(first["Delivery Date"]).strftime('%d/%m/%Y') if pd.notnull(first["Delivery Date"]) else "-"
                        worksheet.merge_range(curr+7, 3, curr+8, 3, "Delivery\nDate", f_deliv)
                        worksheet.merge_range(curr+7, 4, curr+8, 4, delivery, f_deliv)

                        t_h = curr + 10 
                        for i, txt in enumerate(['No.', 'Product Code', 'Product Name', 'Unit/UOM', 'QTY']): worksheet.write(t_h, i, txt, f_table_h)
                        r_ptr = t_h + 1
                        for i, (_, r) in enumerate(df_store.iterrows(), 1):
                            worksheet.write(r_ptr, 0, i, f_border); worksheet.write(r_ptr, 1, str(r['รหัสสินค้า']), f_border)
                            worksheet.write(r_ptr, 2, r['รายการสินค้า'], f_wrap); worksheet.write(r_ptr, 3, r['หน่วย'], f_border)
                            worksheet.write(r_ptr, 4, r['จำนวน'], f_border); r_ptr += 1
                        
                        worksheet.merge_range(r_ptr, 0, r_ptr, 3, 'Total', f_table_h); worksheet.write(r_ptr, 4, df_store['จำนวน'].sum(), f_border)
                        
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
                    worksheet.set_h_pagebreaks(page_breaks); worksheet.fit_to_pages(1, 0)
            
            st.success("✅ Generate DO Master เรียบร้อย!")
            st.download_button("📥 Download DO Master", output_do.getvalue(), "Intanin_DO_Final.xlsx")

else:
    # หน้าแรกตอนยังไม่อัปโหลดไฟล์ (Show Guide)
    st.info("กรุณาอัปโหลดไฟล์ Excel เพื่อเริ่มต้นการใช้งานระบบ Intanin Receipt Convert")
    st.markdown("""
    ### 🛠 ฟีเจอร์เด่น
    - **Splitter:** แยกไฟล์ตามซัพพลายเออร์ พร้อมสรุปยอดรายสินค้า
    - **DO Generator:** ออกเอกสารใบส่งสินค้ามาตรฐาน (A4) สะอาดตา
    - **Fast Processing:** ประมวลผลรวดเร็ว รองรับข้อมูลจำนวนมาก
    """)
