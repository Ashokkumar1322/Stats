import streamlit as st
import pandas as pd
import os

st.set_page_config(page_title="Premium Figures Dashboard", layout="wide")
st.title("📊 Complete Premium Dashboard")
st.write("Expand any section below to view and filter the data.")

# 1. File Upload or Default File
st.sidebar.header("Data Source")
uploaded_file = st.sidebar.file_uploader("Upload your Excel File", type=['xls', 'xlsx'])

# FIXED: Use the correct file name here
file_path = uploaded_file if uploaded_file else "for github stats.xls"

try:
    # 2. Read the Excel File
    xls = pd.ExcelFile(file_path)
    
    st.sidebar.info(f"Loaded {len(xls.sheet_names)} sheets successfully.")
    
    # 3. Iterate through all sheets dynamically
    for sheet in xls.sheet_names:
        
        # Read the data for the current sheet
        df = pd.read_excel(xls, sheet_name=sheet, header=1).dropna(axis=1, how='all')
        
        # Clean datetime column headers into strings (e.g., 'APR-25', 'MAY-26')
        new_cols = []
        for c in df.columns:
            if isinstance(c, pd.Timestamp) or hasattr(c, 'strftime'):
                new_cols.append(pd.to_datetime(c).strftime('%b-%y').upper())
            else:
                new_cols.append(str(c))
        df.columns = new_cols

        st.divider()
        
        # Create an expander for EACH sheet
        with st.expander(f"📁 View Data: {sheet}", expanded=False):
            st.subheader(f"{sheet} Data")
            
            # -------------------------------------------------------------
            # FILTER LOGIC FOR DEPARTMENT & MONTH SHEETS
            # -------------------------------------------------------------
            if sheet in ['25 26', '26 27', '25 26 26 27 For the month', '25 26 26 27 Up to the month']:
                # The first column is dynamically grabbed (usually 'Dept ')
                first_col = df.columns[0] 
                
                col1, col2 = st.columns([1, 2])
                with col1:
                    departments = df[first_col].dropna().unique()
                    selected_dept = st.selectbox(
                        f"Filter by {first_col}:", 
                        options=["All"] + list(departments), 
                        key=f"dept_{sheet}"
                    )
                with col2:
                    # Allow users to filter columns (Months/Years)
                    available_cols = df.columns[1:].tolist()
                    selected_cols = st.multiselect(
                        "Select Columns (Months/Years) to display:", 
                        options=available_cols, 
                        default=available_cols, 
                        key=f"cols_{sheet}"
                    )
                
                # Apply the filters
                filtered_df = df.copy()
                if selected_dept != "All":
                    filtered_df = filtered_df[filtered_df[first_col] == selected_dept]
                
                # Keep the first column, plus whatever months they selected
                columns_to_show = [first_col] + selected_cols
                st.dataframe(filtered_df[columns_to_show], use_container_width=True, hide_index=True)

            # -------------------------------------------------------------
            # FILTER LOGIC FOR CHANNEL SHEETS
            # -------------------------------------------------------------
            elif sheet in ['Channel wise 25 26', 'Channel wise 26 27']:
                st.write("**Search Filters**")
                col1, col2, col3 = st.columns(3)
                
                search_agent = ""
                search_broker = ""
                search_posp = ""
                
                # Display search inputs only if the column exists in the Excel sheet
                if 'AGENT' in df.columns:
                    with col1:
                        search_agent = st.text_input("Search by AGENT:", key=f"agent_{sheet}")
                if 'BROKER' in df.columns:
                    with col2:
                        search_broker = st.text_input("Search by BROKER:", key=f"broker_{sheet}")
                if 'POSP' in df.columns:
                    with col3:
                        search_posp = st.text_input("Search by POSP:", key=f"posp_{sheet}")
                
                filtered_df = df.copy()
                
                # Apply text search (Case-insensitive)
                if search_agent and 'AGENT' in filtered_df.columns:
                    filtered_df = filtered_df[filtered_df['AGENT'].astype(str).str.contains(search_agent, case=False, na=False)]
                if search_broker and 'BROKER' in filtered_df.columns:
                    filtered_df = filtered_df[filtered_df['BROKER'].astype(str).str.contains(search_broker, case=False, na=False)]
                if search_posp and 'POSP' in filtered_df.columns:
                    filtered_df = filtered_df[filtered_df['POSP'].astype(str).str.contains(search_posp, case=False, na=False)]
                    
                st.dataframe(filtered_df, use_container_width=True, hide_index=True)
                
            # -------------------------------------------------------------
            # DEFAULT FALLBACK (If any other new sheets are added)
            # -------------------------------------------------------------
            else:
                st.dataframe(df, use_container_width=True, hide_index=True)

except FileNotFoundError:
    st.error(f"Could not find the file `{file_path if isinstance(file_path, str) else 'uploaded file'}`. Please use the sidebar to manually upload the file!")
except Exception as e:
    st.error(f"An error occurred while processing the data: {e}")
