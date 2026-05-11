import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Invoice Generator V7.0", layout="wide")

st.title("📄 ระบบสร้างใบส่งสินค้าชั่วคราว (New Format)")

if 'name_memory' not in st.session_state:
    st.session_state['name_memory'] = {}

uploaded_file = st.file_uploader("อัปโหลดไฟล์ Excel (Transport)", type=['xlsx'])

if uploaded_file:
    # อ่านข้อมูลจากชีต Transport
    df_raw = pd.read_excel(uploaded_file, sheet_name='Transport')
    df_raw = df_raw.dropna(subset=['ซัพพลายเออร์'])
    unique_suppliers = sorted(df_raw['ซัพพลายเออร์'].unique())

    st.subheader("📝 กำหนดชื่อตัวย่อชีต (ยังคงระบบเดิมไว้)")
    with st.form("sheet_name_form"):
        cols = st.columns(3)
        current_mapping = {}
        for i, supplier in enumerate(unique_suppliers):
            with cols[i % 3]:
                remembered_name = st.session_state['name_memory'].get(supplier, str(supplier)[:10].strip())
                current_mapping[supplier] = st.text_input(f"{supplier}:", value=remembered_name, key=f"input_{supplier}")
        submit_button = st.form_submit_button("สร้างไฟล์ใบส่งสินค้า")

    if submit_button:
        for sup, s_name in current_mapping.items():
            st.session_state['name_memory'][sup] = s_name
            
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            workbook = writer.book
            
            # --- กำหนด Styles ตามรูปแบบใหม่ ---
            title_fmt = workbook.add_format({'bold': True, 'font_size': 16, 'align': 'center'})
            header_label_fmt = workbook.add_format({'bold': True, 'font_size': 11})
            table_head_fmt = workbook.add_format({'bold': True, 'border': 1, 'align': 'center', 'bg_color': '#F4F4F4'})
            cell_border_fmt = workbook.add_format({'border': 1})
            total_fmt = workbook.add_format({'bold': True, 'border': 1, 'align': 'center', 'bg_color': '#EEEEEE'})
            footer_label_fmt = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'top'})

            for supplier, sheet_name in current_mapping.items():
                clean_name = sheet_name.strip()[:31]
                for char in r'[]:*?/\ ':
                    clean_name = clean_name.replace(char, ' ')
                
                # กรองข้อมูลซัพพลายเออร์
                df_sup = df_raw[df_raw['ซัพพลายเออร์'] == supplier].copy()
                
                # แยกตามสาขา (Store Name) เพื่อสร้างใบส่งของแยกแต่ละสาขาในชีตเดียว (หรือแยกตามความเหมาะสม)
                # ในที่นี้จะสร้างสรุปยอดรวมของซัพพลายเออร์นั้นๆ ในรูปแบบใบส่งของ
                
                worksheet = workbook.add_worksheet(clean_name)
                worksheet.set_zoom(90)

                # --- ส่วนหัวเอกสาร (Header) อิงตาม image_89132c.png ---
                worksheet.merge_range('A1:E1', 'บริษัท โมบาย โลจิสติกส์ จำกัด', title_fmt)
                worksheet.merge_range('A2:E2', 'ใบส่งสินค้าชั่วคราว', title_fmt)
                
                worksheet.write('A4', 'Customer Name:', header_label_fmt)
                worksheet.write('A5', 'Store Code:', header_label_fmt)
                worksheet.write('A6', 'Store Name:', header_label_fmt)
                worksheet.write('A7', 'Ship To:', header_label_fmt)
                
                worksheet.write('E4', f'Do. No. {supplier[:5]}', header_label_fmt)
                worksheet.write('E5', 'Zone:', header_label_fmt)
                worksheet.write('E6', 'Delivery Date:', header_label_fmt)

                # --- ส่วนตารางสินค้า (Table) ---
                headers = ['No.', 'Product Code', 'Product Name', 'Unit/UOM', 'QTY']
                for col_num, header in enumerate(headers):
                    worksheet.write(9, col_num, header, table_head_fmt)

                # ดึงข้อมูลสินค้า (Group รวมทุกสาขาของซัพพลายเออร์รายนั้น)
                df_items = df_sup.groupby(['รหัสสินค้า', 'รายการสินค้า', 'หน่วย'], as_index=False)['จำนวน'].sum()
                
                curr_row = 10
                for index, row in df_items.iterrows():
                    worksheet.write(curr_row, 0, index + 1, cell_border_fmt)
                    worksheet.write(curr_row, 1, row['รหัสสินค้า'], cell_border_fmt)
                    worksheet.write(curr_row, 2, row['รายการสินค้า'], cell_border_fmt)
                    worksheet.write(curr_row, 3, row['หน่วย'], cell_border_fmt)
                    worksheet.write(curr_row, 4, row['จำนวน'], cell_border_fmt)
                    curr_row += 1

                # --- ส่วนท้ายตาราง (Total) ---
                worksheet.merge_range(curr_row, 0, curr_row, 3, 'Total', total_fmt)
                worksheet.write(curr_row, 4, df_items['จำนวน'].sum(), total_fmt)
                
                # --- ส่วนลายเซ็น (Footer) ---
                curr_row += 2
                footer_headers = ['ผู้รับสินค้า', 'ผู้ส่งสินค้า / ทะเบียนรถ', 'คลังสินค้า']
                for i, head in enumerate(footer_headers):
                    worksheet.merge_range(curr_row, i*2, curr_row, i*2+1, head, table_head_fmt)
                    worksheet.merge_range(curr_row+1, i*2, curr_row+4, i*2+1, 'ชื่อ....................\nวันที่....................', footer_label_fmt)

                # ปรับขนาดคอลัมน์ให้พอดี
                worksheet.set_column('A:A', 5)   # No.
                worksheet.set_column('B:B', 15)  # Code
                worksheet.set_column('C:C', 45)  # Name (ยาวพิเศษตามรูป)
                worksheet.set_column('D:D', 12)  # Unit
                worksheet.set_column('E:E', 10)  # QTY

        st.success("✅ สร้างใบส่งสินค้าชั่วคราวรูปแบบใหม่เรียบร้อย!")
        st.download_button(label="📥 ดาวน์โหลดไฟล์ใบส่งสินค้า V7.0", data=output.getvalue(), file_name="Temporary_Invoice.xlsx")