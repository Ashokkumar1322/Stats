import streamlit as st
import pandas as pd
import numpy as np 

st.set_page_config(page_title="Premium Figures Dashboard", layout="wide")
st.title("📊 Department-Wise Premium & Accretion Dashboard")

# 1. File Upload or Default File
uploaded_file = st.sidebar.file_uploader("Upload your Excel File", type=['xls', 'xlsx'])
file_path = uploaded_file if uploaded_file else "for github stats.xls"

try:
    # 2. Read the Data
    xls = pd.ExcelFile(file_path)
    
    # Read Department Sheets
    df_25_26 = pd.read_excel(xls, sheet_name='25 26', header=1).dropna(axis=1, how='all')
    df_26_27 = pd.read_excel(xls, sheet_name='26 27', header=1).dropna(axis=1, how='all')
    
    # Read Channel Sheets
    df_chan_25_26 = pd.read_excel(xls, sheet_name='Channel wise 25 26', header=1).dropna(axis=1, how='all')
    df_chan_26_27 = pd.read_excel(xls, sheet_name='Channel wise 26 27', header=1).dropna(axis=1, how='all')
    
    # 3. Clean up column names (Convert timestamps to Month Names & handle first column)
    def clean_columns(df, entity_name="Department"):
        new_cols = []
        for i, c in enumerate(df.columns):
            if i == 0: # Force the first column to be the entity name (Department or Channel)
                new_cols.append(entity_name)
            else:
                try:
                    # Convert any date-like header to a 3-letter uppercase month
                    new_cols.append(pd.to_datetime(c).strftime('%b').upper())
                except:
                    new_cols.append(str(c))
        df.columns = new_cols
        return df

    df_25_26 = clean_columns(df_25_26, "Department")
    df_26_27 = clean_columns(df_26_27, "Department")
    
    df_chan_25_26 = clean_columns(df_chan_25_26, "Channel")
    df_chan_26_27 = clean_columns(df_chan_26_27, "Channel")

    # 4. Show FY 25-26 Data 
    st.header("Historical Month-Wise Figures (FY 25-26)")
    st.dataframe(df_25_26.set_index('Department'), use_container_width=False)

    # 5. Show FY 26-27 Editable Data 
    st.header("Enter New Month Figures (FY 26-27)")
    st.write("Edit the table below to add figures for August through March:")
    edited_df_26_27 = st.data_editor(df_26_27.set_index('Department'), num_rows="dynamic").reset_index()

    # 6. Filters & Dropdowns (For Departments)
    st.sidebar.header("Filters")
    departments = edited_df_26_27['Department'].dropna().unique()
    selected_dept = st.sidebar.selectbox("Select Department", ["All"] + list(departments))
    
    months = ['APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC', 'JAN', 'FEB', 'MAR']
    
    selected_month = st.sidebar.selectbox("Calculate YTD Up To Month:", months, index=3) 
    selected_month_for = st.sidebar.selectbox("Show 'For the Month' Results Up To:", months, index=3) 

    # -------------------------------------------------------------------
    # 7. Calculate YTD Data Dynamically (Side-by-Side Cumulative)
    # -------------------------------------------------------------------
    idx = months.index(selected_month)
    ytd_months = months[:idx+1]
    
    merged_ytd = df_25_26[['Department']].copy()
    
    for i, m in enumerate(ytd_months):
        curr_months = months[:i+1]
        merged_ytd[f'Up to {m} 25'] = df_25_26[curr_months].apply(pd.to_numeric, errors='coerce').sum(axis=1)
        merged_ytd[f'Up to {m} 26'] = edited_df_26_27[curr_months].apply(pd.to_numeric, errors='coerce').sum(axis=1)

    merged_ytd['Accretion'] = merged_ytd[f'Up to {selected_month} 26'] - merged_ytd[f'Up to {selected_month} 25']
    merged_ytd['Accretion %'] = np.where(
        merged_ytd[f'Up to {selected_month} 25'] != 0, 
        (merged_ytd['Accretion'] / merged_ytd[f'Up to {selected_month} 25']) * 100, 
        0
    )

    if selected_dept != "All":
        merged_ytd = merged_ytd[merged_ytd['Department'] == selected_dept]
        final_table_ytd = merged_ytd.sort_values(by="Accretion", ascending=False).reset_index(drop=True)
    else:
        sum_row = merged_ytd[merged_ytd['Department'] == 'Sum for all departments']
        other_rows = merged_ytd[merged_ytd['Department'] != 'Sum for all departments']
        other_rows = other_rows.sort_values(by="Accretion", ascending=False)
        final_table_ytd = pd.concat([other_rows, sum_row]).reset_index(drop=True)
        
    final_table_ytd.index = final_table_ytd.index + 1

    # 8. Display YTD Results
    st.header(f"YTD Accretion Results (Up to {selected_month})")
    
    st.dataframe(
        final_table_ytd,
        use_container_width=False,
        column_config={
            "Accretion %": st.column_config.NumberColumn("Accretion (%)", format="%.2f%%")
        }
    )
    
    if selected_dept == "All" and len(final_table_ytd) > 1:
        st.subheader(f"YTD Insights (Up to {selected_month})")
        best = final_table_ytd.iloc[0]
        worst = final_table_ytd.iloc[-2]
        st.success(f"📈 **Best Performing:** {best['Department']} with an accretion of {best['Accretion']:,.0f} ({best['Accretion %']:.2f}%)")
        st.error(f"📉 **Needs Attention:** {worst['Department']} with an accretion of {worst['Accretion']:,.0f} ({worst['Accretion %']:.2f}%)")

    # -------------------------------------------------------------------
    # 9. "For the Month" Dynamically (Side-by-Side Individuals)
    # -------------------------------------------------------------------
    start_month = months[0]
    idx_for = months.index(selected_month_for)
    for_months = months[:idx_for+1]
    
    merged_range = df_25_26[['Department']].copy()

    for m in for_months:
        merged_range[f'For {m} 25'] = pd.to_numeric(df_25_26[m], errors='coerce').fillna(0)
        merged_range[f'For {m} 26'] = pd.to_numeric(edited_df_26_27[m], errors='coerce').fillna(0)

    merged_range['Accretion'] = merged_range[f'For {selected_month_for} 26'] - merged_range[f'For {selected_month_for} 25']
    merged_range['Accretion %'] = np.where(
        merged_range[f'For {selected_month_for} 25'] != 0, 
        (merged_range['Accretion'] / merged_range[f'For {selected_month_for} 25']) * 100, 
        0
    )

    if selected_dept != "All":
        merged_range = merged_range[merged_range['Department'] == selected_dept]
        final_table_range = merged_range.sort_values(by="Accretion", ascending=False).reset_index(drop=True)
    else:
        sum_row_range = merged_range[merged_range['Department'] == 'Sum for all departments']
        other_rows_range = merged_range[merged_range['Department'] != 'Sum for all departments']
        other_rows_range = other_rows_range.sort_values(by="Accretion", ascending=False)
        final_table_range = pd.concat([other_rows_range, sum_row_range]).reset_index(drop=True)
        
    final_table_range.index = final_table_range.index + 1

    st.header(f"Monthly Accretion Results (For {start_month} to {selected_month_for})")
    
    st.dataframe(
        final_table_range,
        use_container_width=False,
        column_config={
            "Accretion %": st.column_config.NumberColumn("Accretion (%)", format="%.2f%%")
        }
    )
    
    if selected_dept == "All" and len(final_table_range) > 1:
        st.subheader(f"Monthly Insights (For {selected_month_for} alone)")
        best_range = final_table_range.iloc[0]
        worst_range = final_table_range.iloc[-2]
        st.success(f"📈 **Best Performing ({selected_month_for}):** {best_range['Department']} with an accretion of {best_range['Accretion']:,.0f} ({best_range['Accretion %']:.2f}%)")
        st.error(f"📉 **Needs Attention ({selected_month_for}):** {worst_range['Department']} with an accretion of {worst_range['Accretion']:,.0f} ({worst_range['Accretion %']:.2f}%)")

    # -------------------------------------------------------------------
    # 10. NEW: Channel-Wise Data in an Expander (The << feature)
    # -------------------------------------------------------------------
    st.divider()
    
    # st.expander creates a collapsable section that users can open
    with st.expander("📁 View Channel-Wise Data (Click to Expand)", expanded=False):
        st.subheader("Channel-Wise Figures")
        
        # Add a filter specific to the Channel data
        channels = df_chan_25_26['Channel'].dropna().unique()
        selected_channel = st.selectbox("Filter by Channel:", ["All"] + list(channels))
        
        # Apply filter logic
        if selected_channel != "All":
            show_chan_25_26 = df_chan_25_26[df_chan_25_26['Channel'] == selected_channel]
            show_chan_26_27 = df_chan_26_27[df_chan_26_27['Channel'] == selected_channel]
        else:
            show_chan_25_26 = df_chan_25_26
            show_chan_26_27 = df_chan_26_27

        # Display DataFrames
        st.write("**Historical (FY 25-26)**")
        st.dataframe(show_chan_25_26.set_index('Channel'), use_container_width=False)
        
        st.write("**Enter New Figures (FY 26-27)**")
        st.data_editor(show_chan_26_27.set_index('Channel'), num_rows="dynamic")

except Exception as e:
    st.error(f"Error loading data. Please ensure 'for github stats.xls' is formatted correctly. Details: {e}")
# -------------------------------------------------------------------
    # 10. NEW: Channel-Wise Data in an Expander (With Specific Filters)
    # -------------------------------------------------------------------
    st.divider()
    
    with st.expander("📁 View Channel-Wise Data (Click to Expand)", expanded=False):
        st.subheader("Channel-Wise Figures")
        
        # Create a layout with 3 columns for our specific expander filters
        col1, col2, col3 = st.columns(3)
        
        with col1:
            channels = df_chan_25_26['Channel'].dropna().unique()
            selected_channel = st.selectbox("Filter by Channel:", ["All"] + list(channels), key="chan_dropdown")
            
        with col2:
            # Specific search for the 'Name' column
            search_name = st.text_input("Search by Name (Agent):", "", key="name_search")
            
        with col3:
            # Specific search for the 'Party Code' column
            search_code = st.text_input("Search by Party Code:", "", key="code_search")
            
        # Start with a copy of the full data
        show_chan_25_26 = df_chan_25_26.copy()
        show_chan_26_27 = df_chan_26_27.copy()
        
        # 1. Apply the Channel Dropdown filter
        if selected_channel != "All":
            show_chan_25_26 = show_chan_25_26[show_chan_25_26['Channel'] == selected_channel]
            show_chan_26_27 = show_chan_26_27[show_chan_26_27['Channel'] == selected_channel]
            
        # 2. Apply the Name Search filter
        if search_name:
            if 'Name' in show_chan_25_26.columns:
                show_chan_25_26 = show_chan_25_26[show_chan_25_26['Name'].astype(str).str.contains(search_name, case=False, na=False)]
            if 'Name' in show_chan_26_27.columns:
                show_chan_26_27 = show_chan_26_27[show_chan_26_27['Name'].astype(str).str.contains(search_name, case=False, na=False)]
                
        # 3. Apply the Party Code Search filter
        if search_code:
            if 'Party Code' in show_chan_25_26.columns:
                show_chan_25_26 = show_chan_25_26[show_chan_25_26['Party Code'].astype(str).str.contains(search_code, case=False, na=False)]
            if 'Party Code' in show_chan_26_27.columns:
                show_chan_26_27 = show_chan_26_27[show_chan_26_27['Party Code'].astype(str).str.contains(search_code, case=False, na=False)]

        # Display the filtered DataFrames
        st.write("**Historical (FY 25-26)**")
        st.dataframe(show_chan_25_26.set_index('Channel'), use_container_width=False)
        
        st.write("**Enter New Figures (FY 26-27)**")
        st.data_editor(show_chan_26_27.set_index('Channel'), num_rows="dynamic")
