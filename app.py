import streamlit as st
import pandas as pd
import os

# Page Configuration
st.set_page_config(page_title="Premium Figures Dashboard", layout="wide", page_icon="📊")
st.title("📊 Complete Premium Figures Dashboard")

# 1. Sidebar Data Source Selection
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
    # 2. READ ALL SHEETS AT ONCE (Fixes single-sheet bug)
    @st.cache_data(show_spinner=False)
    def load_all_excel_sheets(file, header_row=1, mtime=None):
        # Setting sheet_name=None reads ALL sheets into a dict of DataFrames
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
                if col not in [df.columns[0], 'Party Code', 'Party Name', 'AGENT', 'BROKER', 'POSP', "MISP's NAME", 'DEALER CODE']:
                    df[col] = pd.to_numeric(df[col], errors='coerce')
            
            cleaned_dict[sheet_name] = df.round(2)
            
        return cleaned_dict

    # Load all sheets dictionary
    all_sheets = load_all_excel_sheets(file_path, header_row=1, mtime=file_mtime)
    sheet_names = list(all_sheets.keys())
    
    st.sidebar.success(f"Successfully loaded {len(sheet_names)} sheets!")

    # 3. Sidebar Navigation Selector
    st.sidebar.markdown("---")
    st.sidebar.header("📑 Select Sheet")
    selected_sheet_sidebar = st.sidebar.radio("Jump to Sheet:", sheet_names)

    # 4. Executive Metrics Header (Top Cards)
    if '26 27' in all_sheets:
        df_kpi = all_sheets['26 27']
        dept_col = df_kpi.columns[0]
        
        data_rows = df_kpi[~df_kpi[dept_col].astype(str).str.contains('Sum for all|Total', case=False, na=False)]
        total_row = df_kpi[df_kpi[dept_col].astype(str).str.contains('Sum for all', case=False, na=False)]
        
        total_prem = total_row['TOTAL'].values[0] if not total_row.empty and 'TOTAL' in total_row.columns else data_rows['TOTAL'].sum()
        
        top_lob_name = "N/A"
        if not data_rows.empty and 'TOTAL' in data_rows.columns:
            top_row = data_rows.sort_values(by='TOTAL', ascending=False).iloc[0]
            top_lob_name = str(top_row[dept_col])

        health_share_str = "N/A"
        health_row = data_rows[data_rows[dept_col].astype(str).str.contains('Health Insurance - Retail|061', case=False, na=False)]
        if not health_row.empty and 'Dept share %' in health_row.columns:
            health_share_str = f"{health_row['Dept share %'].values[0]:.2f}%"

        k1, k2, k3, k4 = st.columns(4)
        with k1:
            st.metric("Total Premium (FY 26-27)", f"₹ {total_prem:,.2f}" if total_prem else "N/A")
        with k2:
            st.metric("Total Sheets Available", f"{len(sheet_names)}")
        with k3:
            st.metric("Top LOB", top_lob_name)
        with k4:
            st.metric("Retail Health Share", health_share_str)

    st.markdown("---")

    # 5. Tabbed Navigation View
    sheet_tabs = st.tabs([f"📄 {s}" for s in sheet_names])
    
    # Map selected sidebar sheet to active tab
    active_index = sheet_names.index(selected_sheet_sidebar) if selected_sheet_sidebar in sheet_names else 0

    for i, (tab, sheet) in enumerate(zip(sheet_tabs, sheet_names)):
        with tab:
            df = all_sheets[sheet]
            
            # Configure formatting for numbers, currency, and percentages
            col_format_config = {}
            for c in df.columns:
                c_upper = str(c).upper()
                if 'ICR' in c_upper or 'ACCRETION' in c_upper or 'SHARE' in c_upper:
                    col_format_config[c] = st.column_config.NumberColumn(format="%.2f")
                elif any(k in c_upper for k in ['PREMIUM', 'TOTAL', 'APR-', 'MAY-', 'JUN-', 'JUL-', 'AUG-', 'SEP-', 'OCT-', 'NOV-', 'DEC-', 'JAN-', 'FEB-', 'MAR-']):
                    col_format_config[c] = st.column_config.NumberColumn(format="₹ %,.2f")

            # -------------------------------------------------------------
            # DEPARTMENT & METRICS SHEETS FILTERS
            # -------------------------------------------------------------
            if sheet in ['25 26', '26 27', '25 26 26 27 For the month', '25 26 26 27 Up to the month', 'ICR on Total Premium and EP']:
                st.subheader(f"Filters for {sheet}")
                
                dept_col = df.columns[0]
                col1, col2 = st.columns([1, 2])
                
                with col1:
                    departments = df[dept_col].dropna().unique()
                    selected_dept = st.selectbox(
                        "Filter by Department/LOB:", 
                        options=["All"] + list(departments), 
                        key=f"dept_{sheet}"
                    )
                with col2:
                    available_columns = df.columns[1:].tolist()
                    selected_columns = st.multiselect(
                        "Select Columns to Display:", 
                        options=available_columns, 
                        default=available_columns, 
                        key=f"cols_{sheet}_{len(available_columns)}"
                    )
                
                filtered_df = df.copy()
                if selected_dept != "All":
                    filtered_df = filtered_df[filtered_df[dept_col] == selected_dept]
                
                columns_to_show = [dept_col] + selected_columns
                final_df = filtered_df[columns_to_show]

            # -------------------------------------------------------------
            # CHANNEL SHEETS FILTERS
            # -------------------------------------------------------------
            elif sheet in ['Channel wise 25 26', 'Channel wise 26 27']:
                st.subheader(f"Filters for {sheet}")
                
                col1, col2, col3 = st.columns(3)
                search_agent, search_broker, search_posp = "", "", ""
                
                agent_col = 'AGENT' if 'AGENT' in df.columns else ('Party Name' if 'Party Name' in df.columns else None)
                
                with col1:
                    if agent_col:
                        search_agent = st.text_input(f"Search by {agent_col}:", key=f"agent_{sheet}")
                with col2:
                    if 'BROKER' in df.columns:
                        search_broker = st.text_input("Search by BROKER:", key=f"broker_{sheet}")
                with col3:
                    if 'POSP' in df.columns:
                        search_posp = st.text_input("Search by POSP:", key=f"posp_{sheet}")
                
                filtered_df = df.copy()
                if search_agent and agent_col:
                    filtered_df = filtered_df[filtered_df[agent_col].astype(str).str.contains(search_agent, case=False, na=False)]
                if search_broker and 'BROKER' in filtered_df.columns:
                    filtered_df = filtered_df[filtered_df['BROKER'].astype(str).str.contains(search_broker, case=False, na=False)]
                if search_posp and 'POSP' in filtered_df.columns:
                    filtered_df = filtered_df[filtered_df['POSP'].astype(str).str.contains(search_posp, case=False, na=False)]
                
                final_df = filtered_df

            # -------------------------------------------------------------
            # MISP SHEET FILTER
            # -------------------------------------------------------------
            elif sheet == 'MISP':
                st.subheader(f"Filters for {sheet}")
                search_misp = ""
                
                misp_col = "MISP's NAME" if "MISP's NAME" in df.columns else ("MISP" if "MISP" in df.columns else None)
                if misp_col:
                    search_misp = st.text_input(f"Search by {misp_col}:", key=f"misp_{sheet}")
                
                filtered_df = df.copy()
                if search_misp and misp_col:
                    filtered_df = filtered_df[filtered_df[misp_col].astype(str).str.contains(search_misp, case=False, na=False)]
                
                final_df = filtered_df

            # -------------------------------------------------------------
            # FALLBACK FOR OTHER SHEETS
            # -------------------------------------------------------------
            else:
                st.subheader(f"Data for {sheet}")
                final_df = df

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
                label=f"📥 Download {sheet} Data as CSV",
                data=csv_data,
                file_name=f"{sheet.replace(' ', '_')}_filtered.csv",
                mime='text/csv',
                key=f"download_{sheet}"
            )

except FileNotFoundError:
    st.error("Could not find 'for github stats.xls'. Please use the sidebar to upload the file manually.")
except Exception as e:
    st.error(f"An error occurred while loading the dashboard: {e}")
