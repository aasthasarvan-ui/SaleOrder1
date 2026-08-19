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
                    # Read input file via pandas
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

                    # 2. Route Number Search (26 Columns, Rows up to fg_row)
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

                    # Popup/Warning Alert if Route not found
                    if route_num == "":
                        route_num = "22"
                        st.warning(f"⚠️ Warning: '{uploaded_file.name}' mein Route Number nahi mila. Default Route '22' use kiya gaya hai.")

                    safe_route_num = "".join(c if c.isalnum() or c in ('-', '_') else "-" for c in str(route_num))

                    # 3. Advanced Agency Detection (FG ke Left se Search, ignoring SR No.)
                    agency_col = 0
                    for cSearch in range(fg_col - 1, -1, -1):
                        head_cell = str(df_input.iloc[fg_row - 1, cSearch] if fg_row > 0 else "")
                        curr_cell = str(df_input.iloc[fg_row, cSearch])
                        head_clean = (head_cell + " " + curr_cell).upper().replace("\n", "").replace(" ", "")
                        
                        if ("AG" in head_clean or "AGENCY" in head_clean) and not ("SR" in head_clean or "SERIAL" in head_clean or "S.NO" in head_clean or "NO" in head_clean):
                            agency_col = cSearch
                            break

                    # Fallback to FG col - 1 if header name not matched
                    if agency_col == 0 and fg_col > 0:
                        agency_col = fg_col - 1

                    # 4. Load Template Workbook via openpyxl
                    wb_out = openpyxl.load_workbook(io.BytesIO(template_bytes))
                    ws_out = wb_out["Order Data"] if "Order Data" in wb_out.sheetnames else wb_out.active

                    current_row = 6
                    sales_order_num = 1
                    agency_counts = {} # Unique Reference logic for repeat agencies

                    # 5. Total Column Detection (To ignore Sum/Total quantities)
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
                        agency = df_input.iloc[r, agency_col] if agency_col < df_input.shape[1] else None
                        if pd.notna(agency) and str(agency).strip() != "" and str(agency).replace('.0','').isdigit():
                            agency_val = int(float(agency))
                            
                            # Unique Reference Logic
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
                                            
                                            # Default FG Code Logic (FG500014 if blank between FG and Total)
                                            current_fg = str(df_input.iloc[fg_row, c]).strip()
                                            if current_fg == "" or current_fg.lower() == "nan":
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
                    
                    # Store in Session State
                    st.session_state.processed_files.append({
                        "name": uploaded_file.name,
                        "data": output_buffer.getvalue(),
                        "filename": out_filename
                    })
                
                st.success(f"✅ Batch Processing Complete! Total files processed: {len(st.session_state.processed_files)}")

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
