import streamlit as st
import pandas as pd
import openpyxl
import datetime
import io

# Page Configuration & Styling
st.set_page_config(
    page_title="Sales Order Automation Hub", 
    page_icon="🚀", 
    layout="centered"
)

# Custom CSS for Modern UI & Button Color Customization
st.markdown("""
    <style>
        .main-container {
            background: #ffffff;
            padding: 30px;
            border-radius: 16px;
            box-shadow: 0 4px 20px rgba(0, 0, 0, 0.05);
        }
        /* Customizing Process & Generate Button Color (Bright Emerald Green with dark text for high contrast) */
        .stButton>button {
            width: 100%;
            background-color: #10b981 !important;
            color: #ffffff !important;
            font-size: 16px;
            font-weight: 700;
            padding: 14px;
            border-radius: 8px;
            border: none;
            transition: background 0.3s ease;
        }
        .stButton>button:hover {
            background-color: #059669 !important;
        }
        h1 {
            color: #1e293b;
            font-size: 28px;
            font-weight: 700;
        }
        p {
            color: #64748b;
        }
    </style>
""", unsafe_allow_html=True)

# App Header
st.title("📊 Sales Order Automation Hub")
st.markdown("Upload your **Inbound Demand File** and **Output Template** below to generate formatted orders instantly.")
st.markdown("---")

# File Upload Section
col1, col2 = st.columns(2)

with col1:
    st.subheader("📥 Inbound File")
    uploaded_input = st.file_uploader("Upload Demand Excel", type=["xlsx", "xls"], key="input")

with col2:
    st.subheader("📋 Template File")
    uploaded_template = st.file_uploader("Upload Output.xlsx", type=["xlsx", "xls"], key="template")

st.markdown("<br>", unsafe_allow_html=True)

# Session state initialization
if "processed_data" not in st.session_state:
    st.session_state["processed_data"] = None
if "filename" not in st.session_state:
    st.session_state["filename"] = ""
if "total_orders" not in st.session_state:
    st.session_state["total_orders"] = 0

# Process Button & Action Logic
if st.button("🚀 Process & Generate Orders", type="primary"):
    if uploaded_input is not None and uploaded_template is not None:
        with st.spinner("⚡ Processing files and preserving theme... Please wait."):
            try:
                # Read input file
                df_input = pd.read_excel(uploaded_input, header=None)

                # 1. Find FG Row & Col (Fixed python startswith syntax)
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

                    # Unique file naming with Route, Date, and Exact Timestamp
                    timestamp = datetime.datetime.now().strftime("%H%M%S")
                    unique_filename = f"Generated_Sales_Order_Route_{route_num}_{today_date}_{timestamp}.xlsx"

                    st.session_state["processed_data"] = output_buffer.getvalue()
                    st.session_state["filename"] = unique_filename
                    st.session_state["total_orders"] = sales_order_num - 1
            except Exception as e:
                st.error(f"❌ Error aagaya: {e}")
    else:
        st.warning("⚠️ Kripya pehle dono files upload karein!")

# Automatically show download section if processing is done
if st.session_state["processed_data"] is not None:
    st.markdown("---")
    st.success(f"🎉 Success! Total {st.session_state['total_orders']} orders generated with original theme preserved.")
    
    st.markdown(f"**Generated File Name:** `{st.session_state['filename']}`")
    st.download_button(
        label="🔥 DOWNLOAD GENERATED FILE NOW 🔥",
        data=st.session_state["processed_data"],
        file_name=st.session_state["filename"],
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True
    )
