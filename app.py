import streamlit as st
import pandas as pd
import os

# Page Configuration
st.set_page_config(page_title="Premium Figures Dashboard", layout="wide", page_icon="📊")
st.title("📊 Complete Premium Figures Dashboard")

# 1. Sidebar Data Source & Cache Control
st.sidebar.header("📁 Data Source")
uploaded_file = st.sidebar.file_uploader("Upload your Excel File", type=['xls', 'xlsx'])

file_path = uploaded_file if uploaded_file else "for github stats.xls"

# Get modification timestamp for local file caching
file_mtime = None
if isinstance(file_path, str) and os.path.exists(file_path):
    file_mtime = os.path.getmtime(file_path)

# Sidebar Clear Cache Button
if st.sidebar.button("🔄 Refresh Data / Clear Cache"):
    st.cache_data.clear()
    st.rerun()

try:
    # 2. READ ALL SHEETS AT ONCE 
    @st.cache_data(show_spinner=False)
    def load_all_excel_sheets(file, header_row=1, mtime=None):
        xls_dict = pd.read_excel(file, sheet_name=None, header=header_row)
        cleaned_dict = {}
        
        for sheet_name, df in xls_dict.items():
            df = df.copy()
            # Drop completely empty 'Unnamed' columns
            df = df.loc[:, ~df.columns.astype(str).str.contains('^Unnamed') | df.notna().any()]
            
            # Clean header names (Datetime & whitespace)
            new_cols = []
            for c in df.columns:
                if isinstance(c, pd.Timestamp) or hasattr(c, 'strftime'):
                    new_cols.append(pd.to_datetime(c).strftime('%b-%y').upper())
                else:
                    new_cols.append(str(c).strip())
            df.columns = new_cols
            
            # Force numeric conversion for data columns
            for col in df.columns:
                if col not in [df.columns[0], 'Party Code', 'Party Name', 'AGENT', 'BROKER', 'POSP', "MISP's NAME", 'DEALER CODE', 'LOB Name']:
                    df[col] = pd.to_numeric(df[col], errors='coerce')
            
            cleaned_dict[sheet_name] = df.round(2)
            
        return cleaned_dict

    # Load all sheets dictionary
    all_sheets = load_all_excel_sheets(file_path, header_row=1, mtime=file_mtime)
    sheet_names = list(all_sheets.keys())
    
    st.sidebar.success(f"Successfully loaded {len(sheet_names)} sheets!")

    # 3. Sidebar Navigation Selector (This now controls what renders on the screen)
    st.sidebar.markdown("---")
    st.sidebar.header("📑 View Sheets")
    selected_sheet = st.sidebar.radio("Select the sheet you want to view:", sheet_names)

    # 4. Global Executive Metrics (Top Cards)
    # Default metric placeholders
    total_prem_val = "N/A"
    accretion_label = "Accretion"
    accretion_val = "N/A"
    icr_tp_val = "N/A"
    icr_ep_val = "N/A"

    # Extract Total Premium
    if '26 27' in all_sheets:
        df_kpi = all_sheets['26 27']
        dept_col = df_kpi.columns[0]
        total_row = df_kpi[df_kpi[dept_col].astype(str).str.contains('Sum for all', case=False, na=False)]
        if not total_row.empty and 'TOTAL' in total_row.columns:
            total_prem_val = f"₹ {total_row['TOTAL'].values[0]:,.2f}"

    # Extract Accretion
    if '25 26 26 27 Up to the month' in all_sheets:
        df_upto = all_sheets['25 26 26 27 Up to the month']
        accretion_cols = [c for c in df_upto.columns if 'ACCRETION' in str(c).upper()]
        if accretion_cols:
            latest_accretion_col = accretion_cols[-1] # Gets the latest/last accretion column
            accretion_label = latest_accretion_col
            dept_col = df_upto.columns[0]
            total_row = df_upto[df_upto[dept_col].astype(str).str.contains('Sum for all', case=False, na=False)]
            if not total_row.empty:
                accretion_val = f"{total_row[latest_accretion_col].values[0]:.2f}%"

    # Extract ICR Data
    if 'ICR on Total Premium and EP' in all_sheets:
        df_icr = all_sheets['ICR on Total Premium and EP']
        lob_col = df_icr.columns[0]  # First actual column with names
        grand_total_row = df_icr[df_icr[lob_col].astype(str).str.contains('Grand Total', case=False, na=False)]
        
        if not grand_total_row.empty:
            if 'ICR On Total Premium' in grand_total_row.columns:
                icr_tp_val = f"{grand_total_row['ICR On Total Premium'].values[0]:.2f}%"
            if 'ICR On Earned Premium' in grand_total_row.columns:
                icr_ep_val = f"{grand_total_row['ICR On Earned Premium'].values[0]:.2f}%"

    # Render top metrics
    k1, k2, k3, k4 = st.columns(4)
    with k1:
        st.metric("Total Premium (FY 26-27)", total_prem_val)
    with k2:
        st.metric(accretion_label.title(), accretion_val)
    with k3:
        st.metric("ICR On Total Premium", icr_tp_val)
    with k4:
        st.metric("ICR On Earned Premium", icr_ep_val)

    st.markdown("---")

    # 5. Display Only the Selected Sheet
    st.header(f"📄 {selected_sheet}")
    df = all_sheets[selected_sheet]
    
    # Auto-generate column formatting config
    col_format_config = {}
    for c in df.columns:
        c_upper = str(c).upper()
        if 'ICR' in c_upper or 'ACCRETION' in c_upper or 'SHARE' in c_upper:
            col_format_config[c] = st.column_config.NumberColumn(format="%.2f")
        elif any(k in c_upper for k in ['PREMIUM', 'TOTAL', 'APR-', 'MAY-', 'JUN-', 'JUL-', 'AUG-', 'SEP-', 'OCT-', 'NOV-', 'DEC-', 'JAN-', 'FEB-', 'MAR-']):
            col_format_config[c] = st.column_config.NumberColumn(format="₹ %,.2f")

    final_df = df.copy()

    # -------------------------------------------------------------
    # DEPARTMENT & METRICS SHEETS FILTERS
    # -------------------------------------------------------------
    if selected_sheet in ['25 26', '26 27', '25 26 26 27 For the month', '25 26 26 27 Up to the month', 'ICR on Total Premium and EP']:
        st.subheader("Filters")
        
        dept_col = df.columns[0]
        col1, col2 = st.columns([1, 2])
        
        with col1:
            departments = df[dept_col].dropna().unique()
            selected_dept = st.selectbox(
                "Filter by Department/LOB:", 
                options=["All"] + list(departments)
            )
        with col2:
            available_columns = df.columns[1:].tolist()
            selected_columns = st.multiselect(
                "Select Columns to Display:", 
                options=available_columns, 
                default=available_columns
            )
        
        if selected_dept != "All":
            final_df = final_df[final_df[dept_col] == selected_dept]
        
        columns_to_show = [dept_col] + selected_columns
        final_df = final_df[columns_to_show]

    # -------------------------------------------------------------
    # CHANNEL SHEETS FILTERS
    # -------------------------------------------------------------
    elif selected_sheet in ['Channel wise 25 26', 'Channel wise 26 27']:
        st.subheader("Filters")
        
        col1, col2, col3 = st.columns(3)
        search_agent, search_broker, search_posp = "", "", ""
        
        agent_col = 'AGENT' if 'AGENT' in df.columns else ('Party Name' if 'Party Name' in df.columns else None)
        
        with col1:
            if agent_col:
                search_agent = st.text_input(f"Search by {agent_col}:")
        with col2:
            if 'BROKER' in df.columns:
                search_broker = st.text_input("Search by BROKER:")
        with col3:
            if 'POSP' in df.columns:
                search_posp = st.text_input("Search by POSP:")
        
        if search_agent and agent_col:
            final_df = final_df[final_df[agent_col].astype(str).str.contains(search_agent, case=False, na=False)]
        if search_broker and 'BROKER' in final_df.columns:
            final_df = final_df[final_df['BROKER'].astype(str).str.contains(search_broker, case=False, na=False)]
        if search_posp and 'POSP' in final_df.columns:
            final_df = final_df[final_df['POSP'].astype(str).str.contains(search_posp, case=False, na=False)]

    # -------------------------------------------------------------
    # MISP SHEET FILTER
    # -------------------------------------------------------------
    elif selected_sheet == 'MISP':
        st.subheader("Filters")
        search_misp = ""
        
        misp_col = "MISP's NAME" if "MISP's NAME" in df.columns else ("MISP" if "MISP" in df.columns else None)
        if misp_col:
            search_misp = st.text_input(f"Search by {misp_col}:")
        
        if search_misp and misp_col:
            final_df = final_df[final_df[misp_col].astype(str).str.contains(search_misp, case=False, na=False)]

    # Display Data Table
    st.dataframe(
        final_df, 
        use_container_width=True, 
        hide_index=True,
        column_config=col_format_config
    )
    
    # Export CSV Option
    csv_data = final_df.to_csv(index=False).encode('utf-8')
    st.download_button(
        label=f"📥 Download '{selected_sheet}' Data as CSV",
        data=csv_data,
        file_name=f"{selected_sheet.replace(' ', '_')}_filtered.csv",
        mime='text/csv'
    )

except FileNotFoundError:
    st.error("Could not find the Excel file. Please use the sidebar to upload the file manually.")
except Exception as e:
    st.error(f"An error occurred while loading the dashboard: {e}")
