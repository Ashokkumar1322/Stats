import streamlit as st
import pandas as pd

st.set_page_config(page_title="Premium Figures Dashboard", layout="wide")
st.title("📊 Complete Premium Dashboard")

# 1. File Upload or Default File
st.sidebar.header("Data Source")
uploaded_file = st.sidebar.file_uploader("Upload your Excel File", type=['xls', 'xlsx'])

file_path = uploaded_file if uploaded_file else "for github stats.xls"

try:
    # 2. Read the Excel File
    xls = pd.ExcelFile(file_path)
    
    st.sidebar.info(f"Successfully loaded {len(xls.sheet_names)} sheets.")
    
    # 3. Iterate through all sheets
    for sheet in xls.sheet_names:
        
        # Read data and keep empty future months
        df = pd.read_excel(xls, sheet_name=sheet, header=1)
        df = df.loc[:, ~df.columns.astype(str).str.contains('^Unnamed') | df.notna().any()]
        
        # Clean datetime column headers (Months/Years)
        new_cols = []
        for c in df.columns:
            if isinstance(c, pd.Timestamp) or hasattr(c, 'strftime'):
                new_cols.append(pd.to_datetime(c).strftime('%b-%y').upper())
            else:
                new_cols.append(str(c).strip())
        df.columns = new_cols
        
        # Force ICR columns to numeric and round to 2 decimals
        for col in df.columns:
            if 'ICR' in str(col).upper():
                df[col] = pd.to_numeric(df[col], errors='coerce')
        
        # Broadly round all numerical columns to 2 decimal places in the dataframe
        df = df.round(2)

        st.divider()
        
        with st.expander(f"📁 View Data & Filters for: {sheet}", expanded=True):
            
            # --- CREATE COLUMN CONFIG TO FORCE STREAMLIT TO SHOW 2 DECIMALS ---
            col_format_config = {}
            for c in df.columns:
                if 'ICR' in str(c).upper():
                    col_format_config[c] = st.column_config.NumberColumn(format="%.2f")

            # -------------------------------------------------------------
            # DEPARTMENT & METRICS SHEETS FILTERS
            # -------------------------------------------------------------
            if sheet in ['25 26', '26 27', '25 26 26 27 For the month', '25 26 26 27 Up to the month', 'ICR on Total Premium and EP']:
                st.markdown(f"### Filters for {sheet}")
                
                dept_col = df.columns[0] 
                
                col1, col2 = st.columns([1, 2])
                with col1:
                    departments = df[dept_col].dropna().unique()
                    selected_dept = st.selectbox(
                        "Filter by Dept:", 
                        options=["All"] + list(departments), 
                        key=f"dept_{sheet}"
                    )
                with col2:
                    available_columns = df.columns[1:].tolist()
                    selected_columns = st.multiselect(
                        "Select columns/headers to show:", 
                        options=available_columns, 
                        default=available_columns, 
                        key=f"cols_{sheet}"
                    )
                
                # Apply Filters
                filtered_df = df.copy()
                if selected_dept != "All":
                    filtered_df = filtered_df[filtered_df[dept_col] == selected_dept]
                
                columns_to_show = [dept_col] + selected_columns
                
                # Read-Only Table (Changed from data_editor to dataframe)
                st.dataframe(
                    filtered_df[columns_to_show], 
                    use_container_width=False, 
                    hide_index=True,
                    column_config=col_format_config
                )

            # -------------------------------------------------------------
            # CHANNEL SHEETS FILTERS 
            # -------------------------------------------------------------
            elif sheet in ['Channel wise 25 26', 'Channel wise 26 27']:
                st.markdown(f"### Filters for {sheet}")
                
                col1, col2, col3 = st.columns(3)
                search_agent, search_broker, search_posp = "", "", ""
                
                with col1:
                    if 'AGENT' in df.columns:
                        search_agent = st.text_input("Search by AGENT:", key=f"agent_{sheet}")
                with col2:
                    if 'BROKER' in df.columns:
                        search_broker = st.text_input("Search by BROKER:", key=f"broker_{sheet}")
                with col3:
                    if 'POSP' in df.columns:
                        search_posp = st.text_input("Search by POSP:", key=f"posp_{sheet}")
                
                # Apply Filters
                filtered_df = df.copy()
                if search_agent and 'AGENT' in filtered_df.columns:
                    filtered_df = filtered_df[filtered_df['AGENT'].astype(str).str.contains(search_agent, case=False, na=False)]
                if search_broker and 'BROKER' in filtered_df.columns:
                    filtered_df = filtered_df[filtered_df['BROKER'].astype(str).str.contains(search_broker, case=False, na=False)]
                if search_posp and 'POSP' in filtered_df.columns:
                    filtered_df = filtered_df[filtered_df['POSP'].astype(str).str.contains(search_posp, case=False, na=False)]
                    
                # Read-Only Table (Changed from data_editor to dataframe)
                st.dataframe(
                    filtered_df, 
                    use_container_width=False, 
                    hide_index=True,
                    column_config=col_format_config
                )

except FileNotFoundError:
    st.error("Could not find 'for github stats.xls'. Please use the sidebar to upload the file manually.")
except Exception as e:
    st.error(f"An error occurred: {e}")
