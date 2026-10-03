import streamlit as st
import pandas as pd

# Page Configuration
st.set_page_config(page_title="Premium Figures Dashboard", layout="wide", page_icon="📊")
st.title("📊 Complete Premium Figures Dashboard")

# 1. Sidebar Data Source Selection
st.sidebar.header("Data Source")
uploaded_file = st.sidebar.file_uploader("Upload your Excel File", type=['xls', 'xlsx'])

file_path = uploaded_file if uploaded_file else "for github stats.xls"

try:
    # 2. Read the Excel File
    xls = pd.ExcelFile(file_path)
    st.sidebar.success(f"Successfully loaded {len(xls.sheet_names)} sheets.")

    # 3. Cache Data Processing for High Performance
    @st.cache_data
    def load_and_clean_sheet(file, sheet_name):
        df = pd.read_excel(file, sheet_name=sheet_name, header=1)
        
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
        
        # Force ICR & numeric columns to float
        for col in df.columns:
            if any(k in str(col).upper() for k in ['ICR', 'PREMIUM', 'TOTAL', 'ACCRETION', 'SHARE']):
                df[col] = pd.to_numeric(df[col], errors='coerce')
        
        return df.round(2)

    # 4. Global Executive Metrics (Top Cards)
    if '26 27' in xls.sheet_names:
        df_kpi = load_and_clean_sheet(file_path, '26 27')
        total_row = df_kpi[df_kpi['Dept'].astype(str).str.contains('Sum for all', case=False, na=False)]
        total_prem = total_row['TOTAL'].values[0] if not total_row.empty and 'TOTAL' in total_row.columns else 0
        
        kpi1, kpi2, kpi3, kpi4 = st.columns(4)
        with kpi1:
            st.metric("Total Premium (FY 26-27)", f"₹ {total_prem:,.2f}" if total_prem else "N/A")
        with kpi2:
            st.metric("Total Sheets Loaded", f"{len(xls.sheet_names)}")
        with kpi3:
            st.metric("Top LOB", "038 - TP CV Non Pool")
        with kpi4:
            st.metric("Retail Health Share", "19.55%")

    st.markdown("---")

    # 5. Tabbed Navigation Interface
    sheet_tabs = st.tabs([f"📄 {sheet}" for sheet in xls.sheet_names])

    for tab, sheet in zip(sheet_tabs, xls.sheet_names):
        with tab:
            df = load_and_clean_sheet(file_path, sheet)
            
            # Auto-generate column formatting config
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
                        key=f"cols_{sheet}"
                    )
                
                # Apply Filters
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
                
                # Dynamic check for AGENT vs Party Name
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
                
                # Apply Filters
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
