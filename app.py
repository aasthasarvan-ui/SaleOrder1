import streamlit as st
import pandas as pd
import openpyxl
import datetime
import os
import io
import glob

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
st.markdown("Upload multiple **Inbound Demand Files** and an **Output Template** to process orders in batch.")
st.markdown("---")

# File Upload Section
uploaded_inputs = st.file_uploader("Upload Multiple Demand Excel Files", type=["xlsx", "xls"], accept_multiple_files=True, key="inputs")
uploaded_template = st.file_uploader("Upload Output.xlsx Template", type=["xlsx", "xls"], key="template")

st.markdown("<br>", unsafe_allow_html=True)

if st.button("🚀 Process Batch Orders", type="primary"):
    if uploaded_inputs and uploaded_template is not None:
        with st.spinner("⚡ Processing files in batch and preserving original themes... Please wait."):
            try:
                total_processed = 0
                total_skipped = 0
                total_orders_created = 0
                
                template_bytes = uploaded_template.getvalue()
                today_date = datetime.date.today().strftime("%Y-%m-%d")
                timestamp = datetime.datetime.now().strftime("%H%M%S")

                summary_logs = []

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

                    # 3. Advanced Agency Detection
                    agency_col = 0
                    for cSearch in range(fg_col):
                        head_cell = str(df_input.iloc[fg_row - 1, cSearch] if fg_row > 0 else "")
                        head_clean = head_cell.upper().replace("\n", "").replace(" ", "")
                        if "AG" in head_clean and "SALES" not in head_clean:
                            agency_col = cSearch
                            break

                    if agency_col == 0:
                        for col in range(fg_col):
                            header_text = str(df_input.iloc[fg_row - 1, col] if fg_row > 0 else "") + " " + str(df_input.iloc[fg_row, col])
                            header_text = header_text.upper()
                            if any(x in header_text for x in ["SR", "SERIAL", "S.NO", "NO"]):
                                continue
                            
                            is_numeric = True
                            count_num = 0
                            for i in range(fg_row + 1, df_input.shape[0]):
                                v = df_input.iloc[i, col]
                                if pd.notna(v) and str(v).strip() != "":
                                    if not str(v).replace('.0','').isdigit():
                                        is_numeric = False
                                        break
                                    count_num += 1
                            if is_numeric and count_num > 0:
                                agency_col = col
                                break

                    if agency_col == 0:
                        agency_col = 2  # default fallback

                    # 4. Strict Valid FG Columns (Excluding Total & Remarks)
                    valid_cols = []
                    for c in range(fg_col, df_input.shape[1]):
                        fg_code = str(df_input.iloc[fg_row, c] if fg_row >= 0 else "").strip()
                        header_name = str(df_input.iloc[fg_row - 1, c] if fg_row > 0 else "").strip().upper()
                        
                        if "TOTAL" in header_name or "REMARK" in header_name or "SUM" in header_name or not fg_code.upper().startswith("FG"):
                            break
                        valid_cols.append((c, fg_code))

                    # 5. Load Template Workbook via openpyxl (Keeps theme/formatting 100% untouched)
                    wb_out = openpyxl.load_workbook(io.BytesIO(template_bytes))
                    ws_out = wb_out["Order Data"] if "Order Data" in wb_out.sheetnames else wb_out.active

                    current_row = 6
                    sales_order_num = 1

                    # 6. Data Mapping and Injection
                    for r in range(fg_row + 1, df_input.shape[0]):
                        agency = df_input.iloc[r, agency_col]
                        if pd.notna(agency) and str(agency).strip() != "" and str(agency).replace('.0','').isdigit():
                            agency_val = int(float(agency))
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
                                            ws_out.cell(row=current_row, column=9, value=f"REF-{agency_val}-{today_date}")
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
                                total_orders_created += 1

                    # Save output file to memory buffer
                    output_buffer = io.BytesIO()
                    wb_out.save(output_buffer)
                    output_buffer.seek(0)

                    out_filename = f"{safe_routeNum}_{today_date}_{timestamp}.xlsx"
                    
                    st.success(f"✅ Processed: {shortfilename} -> Orders created: {sales_order_num-1}")
                    
                    # Individual download buttons for each processed file
                    st.download_button(
                        label=f"📥 Download Output for {shortfilename}",
                        data=output_buffer,
                        file_name=out_filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=out_filename
                    )
                    total_processed += 1

                st.markdown("---")
                st.info(f"📊 **Batch Summary:** Total Files Processed: {total_processed} | Total Orders Generated: {total_orders_created}")

            except Exception as e:
                st.error(f"❌ Error aagaya: {e}")
    else:
        st.warning("⚠️ Kripya pehle multiple demand files aur template file upload karein!")
