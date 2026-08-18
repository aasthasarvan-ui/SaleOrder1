import streamlit as st
import pandas as pd
import openpyxl
import datetime
import io

st.set_page_config(page_title="Sales Order Automation Tool", page_icon="📊", layout="centered")

st.title("📊 Sales Order Automation Tool")
st.write("Apni **Inbound Demand Excel File** aur **Output.xlsx (Template)** yahan upload karein:")

# File uploaders
uploaded_input = st.file_uploader("1. Inbound Demand Excel File upload karein", type=["xlsx", "xls"])
uploaded_template = st.file_uploader("2. Output.xlsx Template file upload karein", type=["xlsx", "xls"])

if st.button("Process & Generate Orders", type="primary"):
    if uploaded_input is not None and uploaded_template is not None:
        with st.spinner("Processing ho rahi hai, kripya intezaar karein..."):
            try:
                # Read input file
                df_input = pd.read_excel(uploaded_input, header=None)

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
                    st.error("❌ Error: Input file mein 'FG' material code row nahi mila!")
                else:
                    # 2. Route Number Finding
                    route_num = "22"
                    for r in range(min(fg_row, 5)):
                        for c in range(min(df_input.shape[1], 10)):
                            cell_val = str(df_input.iloc[r, c]).strip()
                            upper_val = cell_val.upper()
                            if cell_val != "" and len(cell_val) <= 6 and any(char.isdigit() for char in cell_val) and upper_val not in ["SALES PERSON", "CONTACT NO:", "RT DR", "ROUTE", "MATERIAL CODE"]:
                                route_num = cell_val
                                break

                    # 3. Agency Column Detection
                    agency_col = -1
                    for c in range(fg_col):
                        header_text = str(df_input.iloc[fg_row - 1, c] if fg_row > 0 else "").upper().replace(" ", "")
                        if "AG" in header_text and "SALES" not in header_text:
                            agency_col = c
                            break
                    if agency_col == -1:
                        agency_col = 2

                    # 4. Strict Valid FG Columns (Excluding Total & Remarks)
                    valid_cols = []
                    for c in range(fg_col, df_input.shape[1]):
                        fg_code = str(df_input.iloc[fg_row, c] if fg_row >= 0 else "").strip()
                        header_name = str(df_input.iloc[fg_row - 1, c] if fg_row > 0 else "").strip().upper()
                        
                        if "TOTAL" in header_name or "REMARK" in header_name or not fg_code.upper().startswith("FG"):
                            break
                        valid_cols.append(c)

                    # 5. Load Template Workbook via openpyxl (Keeps theme/formatting 100% untouched)
                    template_bytes = uploaded_template.getvalue()
                    wb_out = openpyxl.load_workbook(io.BytesIO(template_bytes))
                    ws_out = wb_out["Order Data"] if "Order Data" in wb_out.sheetnames else wb_out.active

                    current_row = 6
                    sales_order_num = 1
                    today_date = datetime.date.today().strftime("%Y-%m-%d")

                    # 6. Data Mapping and Injection
                    for r in range(fg_row + 1, df_input.shape[0]):
                        agency = df_input.iloc[r, agency_col]
                        if pd.notna(agency) and str(agency).strip() != "" and str(agency).replace('.0','').isdigit():
                            agency_val = int(float(agency))
                            item_id = 10
                            row_has_items = False
                            
                            for c in valid_cols:
                                fg_code = str(df_input.iloc[fg_row, c]).strip()
                                sku_qty = df_input.iloc[r, c]
                                
                                if pd.notna(sku_qty) and str(sku_qty).strip() != "":
                                    try:
                                        qty_val = float(sku_qty)
                                        if qty_val > 0:
                                            row_has_items = True
                                            fg_name = fg_code
                                            
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
                                            ws_out.cell(row=current_row, column=16, value=fg_name)
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

                    # Save output to memory buffer
                    output_buffer = io.BytesIO()
                    wb_out.save(output_buffer)
                    output_buffer.seek(0)

                    output_filename = f"Generated_Sales_Order_{route_num}.xlsx"
                    
                    st.success(f"🎉 Success! Total {sales_order_num-1} orders generated successfully with 100% original theme.")
                    
                    # Download Button
                    st.download_button(
                        label="📥 Download Generated Sales Order Excel",
                        data=output_buffer,
                        file_name=output_filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            except Exception as e:
                st.error(f"❌ Error aagaya: {e}")
    else:
                    st.warning("⚠️ Kripya pehle dono files (Inbound Excel file aur Output.xlsx template) upload karein!")
