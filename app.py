import streamlit as st
import pandas as pd

st.set_page_config(page_title="Premium Figures Dashboard", layout="wide")
st.title("📊 Complete Premium Dashboard")

# 1. File Upload or Default File
st.sidebar.header("Data Source")
uploaded_file = st.sidebar.file_uploader("Upload your Excel File", type=['xls', 'xlsx'])

# Use the uploaded file, or default to the local one
file_path = uploaded_file if uploaded_file else "for github stats.xls"

try:
    # 2. Read the Excel File
    xls = pd.ExcelFile(file_path)
    
    st.sidebar.info(f"Successfully loaded {len(xls.sheet_names)} sheets.")
    
    # 3. Iterate through all sheets
    for sheet in xls.sheet_names:
        
        # Read data
        df = pd.read_excel(xls, sheet_name=sheet, header=1).dropna(axis=1, how='all')
        
        # Clean datetime column headers (Months/Years)
        new_cols = []
        for c in df.columns:
            if isinstance(c, pd.Timestamp) or hasattr(c, 'strftime'):
                new_cols.append(pd.to_datetime(c).strftime('%b-%y').upper())
            else:
                # Strip extra spaces just in case (e.g., 'Dept ' becomes 'Dept')
                new_cols.append(str(c).strip())
        df.columns = new_cols

        st.divider()
        
        # CHANGED: expanded=True so you see the filters immediately!
        with st.expander(f"📁 View Data & Filters for: {sheet}", expanded=True):
            
            # -------------------------------------------------------------
            # DEPARTMENT SHEETS FILTERS
            # -------------------------------------------------------------
            if sheet in ['25 26', '26 27', '25 26 26 27 For the month', '25 26 26 27 Up to the month']:
                st.markdown(f"### Filters for {sheet}")
                
                # Using 'Dept' as requested (handling any column name variations)
                dept_col = df.columns[0] 
                
                col1, col2 = st.columns([1, 2])
                with col1:
                    departments = df[dept_col].dropna().unique()
                    # Filter for 'Dept'
                    selected_dept = st.selectbox(
                        "Filter by Dept:", 
                        options=["All"] + list(departments), 
                        key=f"dept_{sheet}"
                    )
                with col2:
                    # Filter for Month/Year (The header columns)
                    available_months = df.columns[1:].tolist()
                    selected_months = st.multiselect(
                        "Filter by Month and Year (Select headers to show):", 
                        options=available_months, 
                        default=available_months, 
                        key=f"cols_{sheet}"
                    )
                
                # Apply Filters
                filtered_df = df.copy()
                if selected_dept != "All":
                    filtered_df = filtered_df[filtered_df[dept_col] == selected_dept]
                
                # Show chosen columns
                columns_to_show = [dept_col] + selected_months
                st.dataframe(filtered_df[columns_to_show], use_container_width=True, hide_index=True)

            # -------------------------------------------------------------
            # CHANNEL SHEETS FILTERS
            # -------------------------------------------------------------
            elif sheet in ['Channel wise 25 26', 'Channel wise 26 27']:
                st.markdown(f"### Filters for {sheet}")
                
                col1, col2, col3 = st.columns(3)
                
                search_agent = ""
                search_broker = ""
                search_posp = ""
                
                # Explicit filters for AGENT, BROKER, POSP
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
                    
                st.dataframe(filtered_df, use_container_width=True, hide_index=True)

except FileNotFoundError:
    st.error("Could not find 'for github stats.xls'. Please use the sidebar to upload the file manually.")
except Exception as e:
    st.error(f"An error occurred: {e}")
