import streamlit as st
import pandas as pd
import openpyxl
import datetime
import io
import requests

# Page Configuration & Styling
st.set_page_config(
    page_title="Sales Order Automation Hub", 
    page_icon="🚀", 
    layout="centered"
)

st.markdown("""
    <style>
        .stButton>button {
            width: 100%;
            background-color: #10b981 !important;
            color: #ffffff !important;
            font-size: 16px;
            font-weight: 700;
            padding: 14px;
            border-radius: 8px;
            border: none;
        }
        .stButton>button:hover {
            background-color: #059669 !important;
        }
    </style>
""", unsafe_allow_html=True)

st.title("📊 Sales Order Automation Hub")
st.markdown("Upload multiple **Inbound Demand Files** to process orders in batch (Template is loaded from GitHub automatically).")
st.markdown("---")

# Session State for persistent download buttons
if 'processed_files' not in st.session_state:
    st.session_state.processed_files = []

# File Upload Section
uploaded_inputs = st.file_uploader("Upload Multiple Demand Excel Files", type=["xlsx", "xls"], accept_multiple_files=True, key="inputs")

st.markdown("<br>", unsafe_allow_html=True)

if st.button("🚀 Process Batch Orders", type="primary"):
    if uploaded_inputs:
        st.session_state.processed_files = []
        with st.spinner("⚡ Fetching template from GitHub and processing files... Please wait."):
            try:
                # GitHub se Output.xlsx ka RAW link
                github_template_url = "https://raw.githubusercontent.com/aasthasarvan-ui/SaleOrder1/main/Output.xlsx"
                
                response = requests.get(github_template_url)
                if response.status_code == 200:
                    template_bytes = response.content
                else:
                    st.error("❌ GitHub se template file download nahi ho payi. Kripya URL check karein.")
                    st.stop()
                
                total_processed = 0
                total_orders_created = 0
                
                today_date = datetime.date.today().strftime("%Y-%m-%d")
                timestamp = datetime.datetime.now().strftime("%H%M%S")

                for uploaded_file in uploaded_inputs:
                    short_filename = uploaded_file.name
                    if short_filename.lower() == "output.xlsx":
                        continue

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

                    # 2. Route Number Finding
                    route_num = "22"
                    for r in range(min(fg_row, 5)):
                        for c in range(min(df_input.shape[1], 10)):
                            cell_val = str(df_input.iloc[r, c]).strip()
                            upper_val = cell_val.upper()
                            if cell_val != "" and len(cell_val) <= 6 and any(char.isdigit() for char in cell_val) and upper_val not in ["SALES PERSON", "CONTACT NO:", "RT DR", "ROUTE", "MATERIAL CODE"]:
                                route_num = cell_val
                                break

                    safe_route_num = "".join(c if c.isalnum() or c in ('-', '_') else "-" for c in str(route_num))

                    # 3. Smart Agency Detection (1 to 5 characters length check, ignoring mobile numbers)
                    agency_col = -1
                    for cSearch in range(fg_col - 1, -1, -1):
                        valid_agency_count = 0
                        for rCheck in range(fg_row + 1, df_input.shape[0]):
                            v = df_input.iloc[rCheck, cSearch]
                            if pd.notna(v) and str(v).strip() != "":
                                clean_v = str(v).replace('.0', '').strip()
                                # Agency number must be numeric and between 1 to 5 characters long
                                if clean_v.isdigit() and 1 <= len(clean_v) <= 5:
                                    valid_agency_count += 1
                        
                        # Agar is column mein 1-5 digit wale valid agency numbers mil jate hain
                        if valid_agency_count > 0:
                            agency_col = cSearch
                            break

                    # Fallback agar column na mile toh FG se theek pehla column
                    if agency_col == -1 and fg_col > 0:
                        agency_col = fg_col - 1

                    # 4. Strict Valid FG Columns (Excluding Total & Remarks)
                    valid_cols = []
                    for c in range(fg_col, df_input.shape[1]):
                        fg_code = str(df_input.iloc[fg_row, c] if fg_row >= 0 else "").strip()
                        header_name = str(df_input.iloc[fg_row - 1, c] if fg_row > 0 else "").strip().upper()
                        
                        if "TOTAL" in header_name or "REMARK" in header_name or "SUM" in header_name or not fg_code.upper().startswith("FG"):
                            break
                        valid_cols.append((c, fg_code))

                    # 5. Load Template Workbook via openpyxl
                    wb_out = openpyxl.load_workbook(io.BytesIO(template_bytes))
                    ws_out = wb_out["Order Data"] if "Order Data" in wb_out.sheetnames else wb_out.active

                    current_row = 6
                    sales_order_num = 1

                    # 6. Data Mapping and Injection (with Unique Reference & 1-5 char Agency validation)
                    agency_counts = {}
                    file_orders_count = 0

                    for r in range(fg_row + 1, df_input.shape[0]):
                        agency = df_input.iloc[r, agency_col] if agency_col >= 0 else None
                        
                        if pd.notna(agency) and str(agency).strip() != "":
                            agency_str = str(agency).replace('.0','').strip()
                            
                            # Strict Check: Agency number must be 1 to 5 digits only (ignores mobile numbers/long IDs)
                            if agency_str.isdigit() and 1 <= len(agency_str) <= 5:
                                agency_val = int(agency_str)
                                
                                if agency_val in agency_counts:
                                    agency_counts[agency_val] += 1
                                else:
                                    agency_counts[agency_val] = 1
                                
                                current_agency_seq = agency_counts[agency_val]
                                
                                if current_agency_seq == 1:
                                    ref_number = f"REF-{agency_val}-{today_date}"
                                else:
                                    ref_number = f"REF-{agency_val}-{today_date}-{current_agency_seq}"

                                item_id = 10
                                row_has_items = False
                                
                                for c, fg_code in valid_cols:
                                    sku_qty = df_input.iloc[r, c]
                                    if pd.notna(sku_qty) and str(sku_qty).strip() != "":
                                        try:
                                            qty_val = float(sku_qty)
                                            if qty_val > 0:
                                                row_has_items = True
                                                
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
                                                ws_out.cell(row=current_row, column=16, value=fg_code)
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
                                    file_orders_count += 1
                                    total_orders_created += 1

                    # Save output file to memory buffer
                    output_buffer = io.BytesIO()
                    wb_out.save(output_buffer)
                    output_buffer.seek(0)

                    out_filename = f"{safe_route_num}_{today_date}_{timestamp}.xlsx"
                    
                    st.session_state.processed_files.append({
                        "name": short_filename,
                        "data": output_buffer.getvalue(),
                        "filename": out_filename,
                        "orders": file_orders_count
                    })
                    total_processed += 1

                st.success("✅ Batch Processing Complete!")

            except Exception as e:
                st.error(f"❌ Error aagaya: {e}")

# Display persistent download buttons and summary from Session State
if st.session_state.processed_files:
    st.markdown("---")
    for item in st.session_state.processed_files:
        st.success(f"✅ Processed: {item['name']} -> Orders created: {item['orders']}")
        st.download_button(
            label=f"📥 Download Output for {item['name']}",
            data=item['data'],
            file_name=item['filename'],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=item['filename']
        )
    
    st.markdown("---")
    total_files_count = len(st.session_state.processed_files)
    total_gen_orders = sum(item['orders'] for item in st.session_state.processed_files)
    st.info(f"📊 **Batch Summary:** Total Files Processed: {total_files_count} | Total Orders Generated: {total_gen_orders}")
else:
    st.warning("⚠️ Kripya pehle demand files upload karein!")
