import streamlit as st
import pandas as pd
import openpyxl
import datetime
import io
import requests

# Page Configuration
st.set_page_config(page_title="Sales Order Automation Hub", page_icon="🚀", layout="centered")

st.title("📊 Sales Order Automation Hub")
st.markdown("Upload multiple **Inbound Demand Files** to process. (Template is automatically loaded from GitHub).")
st.markdown("---")

# Session State for persistent download buttons
if 'processed_files' not in st.session_state:
    st.session_state.processed_files = []

# File Upload Section
uploaded_inputs = st.file_uploader("Upload Multiple Demand Excel Files", type=["xlsx", "xls"], accept_multiple_files=True, key="inputs")

if st.button("🚀 Process Batch Orders", type="primary"):
    if uploaded_inputs:
        st.session_state.processed_files = []
        with st.spinner("⚡ Fetching template and processing files..."):
            try:
                # GitHub RAW link for template
                url = "https://raw.githubusercontent.com/aasthasarvan-ui/SaleOrder1/main/Output.xlsx"
                response = requests.get(url)
                if response.status_code != 200:
                    st.error("❌ GitHub se template file load nahi ho payi. Kripya URL check karein.")
                    st.stop()
                
                template_bytes = response.content
                today_date = datetime.date.today().strftime("%Y-%m-%d")
                timestamp = datetime.datetime.now().strftime("%H%M%S")

                for uploaded_file in uploaded_inputs:
                    file_bytes = uploaded_file.getvalue()
                    df_input = pd.read_excel(io.BytesIO(file_bytes), header=None)

                    # 1. Find FG Row & Col
                    fg_row, fg_col = -1, -1
                    for r in range(df_input.shape[0]):
                        for c in range(df_input.shape[1]):
                            val = str(df_input.iloc[r, c]).strip().upper()
                            if val.startswith("FG"):
                                fg_row, fg_col = r, c
                                break
                        if fg_row != -1:
                            break

                    if fg_row == -1:
                        continue

                    # 2. Route Number Search (26 Columns)
                    route_num = ""
                    for r in range(min(fg_row, 6)):
                        for c in range(min(df_input.shape[1], 26)):
                            cell_val = str(df_input.iloc[r, c]).strip()
                            upper_val = cell_val.upper()
                            if cell_val != "" and len(cell_val) <= 6 and any(char.isdigit() for char in cell_val):
                                if upper_val not in ["SALES PERSON", "CONTACT NO:", "RT DR", "ROUTE", "MATERIAL CODE"]:
                                    route_num = cell_val
                                    break
                        if route_num != "":
                            break

                    if route_num == "":
                        route_num = "22"

                    safe_route_num = "".join(c if c.isalnum() or c in ('-', '_') else "-" for c in str(route_num))

                    # 3. Agency Detection Logic (FG ke Left se pehla valid numerical column)
                    agency_col = -1
                    
                    # FG column ke theek left se start karke piche ki taraf (0 tak) jayenge
                    for cSearch in range(fg_col - 1, -1, -1):
                        # Header text check karenge ki kahin ye Sr No ya Serial toh nahi
                        head_cell = str(df_input.iloc[fg_row - 1, cSearch] if fg_row > 0 else "")
                        head_clean = head_cell.upper().replace(" ", "")
                        
                        if "SR" in head_clean or "SERIAL" in head_clean or "S.NO" in head_clean:
                            continue # Serial number wale column ko chhod denge
                        
                        # Check karenge ki is column ke andar neeche numerical values (agency IDs) maujud hain ya nahi
                        numeric_count = 0
                        for rCheck in range(fg_row + 1, df_input.shape[0]):
                            vTest = df_input.iloc[rCheck, cSearch]
                            if pd.notna(vTest) and str(vTest).strip() != "":
                                clean_v = str(vTest).replace('.0', '').strip()
                                if clean_v.isdigit():
                                    numeric_count += 1
                        
                        # Agar is column mein numerical values milti hain, toh yehi hamara Agency Column hai!
                        if numeric_count > 0:
                            agency_col = cSearch
                            break

                    # Agar loop se bhi na mile, toh default FG ke turant pehla column (fg_col - 1) le lenge
                    if agency_col == -1 and fg_col > 0:
                        agency_col = fg_col - 1

                    # 4. Load Template Workbook via openpyxl
                    wb_out = openpyxl.load_workbook(io.BytesIO(template_bytes))
                    ws_out = wb_out["Order Data"] if "Order Data" in wb_out.sheetnames else wb_out.active

                    current_row = 6
                    sales_order_num = 1
                    agency_counts = {}

                    # 5. Total Column Detection
                    total_col = -1
                    for search_col in range(fg_col, min(df_input.shape[1], 50)):
                        h_text = str(df_input.iloc[fg_row, search_col] if fg_row >= 0 else "").strip().upper()
                        f_formula = ""
                        try:
                            f_formula = str(df_input.iloc[fg_row + 1, search_col]).strip().upper()
                        except:
                            pass
                        
                        if "TOTAL" in h_text or "SUM" in h_text or "SUM" in f_formula:
                            total_col = search_col
                            break

                    end_col_limit = total_col if total_col != -1 else df_input.shape[1]

                    # 6. Data Mapping and Injection
                    for r in range(fg_row + 1, df_input.shape[0]):
                        agency = df_input.iloc[r, agency_col] if agency_col >= 0 and agency_col < df_input.shape[1] else None
                        
                        if pd.notna(agency) and str(agency).strip() != "":
                            agency_str = str(agency).replace('.0', '').strip()
                            if agency_str.isdigit():
                                agency_val = int(agency_str)
                                
                                # Unique Reference Logic (Repeat Agency)
                                agency_counts[agency_val] = agency_counts.get(agency_val, 0) + 1
                                seq = agency_counts[agency_val]
                                ref_number = f"REF-{agency_val}-{today_date}" if seq == 1 else f"REF-{agency_val}-{today_date}-{seq}"

                                item_id = 10
                                row_has_items = False
                                
                                for c in range(fg_col, end_col_limit):
                                    sku_qty = df_input.iloc[r, c]
                                    if pd.notna(sku_qty) and str(sku_qty).strip() != "":
                                        try:
                                            qty_val = float(sku_qty)
                                            if qty_val > 0:
                                                row_has_items = True
                                                
                                                current_fg = str(df_input.iloc[fg_row, c]).strip()
                                                if current_fg == "" or current_fg.lower() == "nan" or current_fg.upper() == "NONE":
                                                    current_fg = "FG500014"
                                                
                                                ws_out.cell(row=current_row, column=2, value=sales_order_num)
                                                ws_out.cell(row=current_row, column=3, value="OR")
                                                ws_out.cell(row=current_row, column=4, value="SO20")
                                                ws_out.cell(row=current_row, column=5, value=10)
                                                ws_out.cell(row=current_row, column=6, value=20)
                                                ws_out.cell(row=current_row, column=7, value=f"DR{agency_val}")
                                                ws_out.cell(row=current_row, column=8, value=f"DR{agency_val}")
                                                ws_out.cell(row=current_row, column=9, value=ref_number)
                                                ws_out.cell(row=current_row, column=10, value=today_date)
                                                ws_out.cell(row=current_row, column=11, value=today_date)
                                                ws_out.cell(row=current_row, column=15, value=item_id)
                                                ws_out.cell(row=current_row, column=16, value=current_fg)
                                                ws_out.cell(row=current_row, column=19, value=qty_val)
                                                ws_out.cell(row=current_row, column=20, value="Bag")
                                                ws_out.cell(row=current_row, column=22, value=2100)
                                                ws_out.cell(row=current_row, column=26, value=str(route_num))
                                                ws_out.cell(row=current_row, column=27, value=agency_val)
                                                
                                                item_id += 10
                                                current_row += 1
                                        except ValueError:
                                            pass
                                if row_has_items:
                                    sales_order_num += 1

                    # Save output file to memory buffer
                    output_buffer = io.BytesIO()
                    wb_out.save(output_buffer)
                    output_buffer.seek(0)

                    out_filename = f"{safe_route_num}_{today_date}_{timestamp}.xlsx"
                    
                    st.session_state.processed_files.append({
                        "name": uploaded_file.name,
                        "data": output_buffer.getvalue(),
                        "filename": out_filename
                    })
                
                st.success("✅ Batch Processing Complete!")

            except Exception as e:
                st.error(f"❌ Error aagaya: {e}")

# Display persistent download buttons from Session State
if st.session_state.processed_files:
    st.markdown("### 📥 Download Processed Files:")
    for item in st.session_state.processed_files:
        st.download_button(
            label=f"📥 Download Output for {item['name']}",
            data=item['data'],
            file_name=item['filename'],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=item['filename']
        )
