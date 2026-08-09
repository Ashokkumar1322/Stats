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
    df_25_26 = pd.read_excel(xls, sheet_name='25 26', header=1).dropna(axis=1, how='all')
    df_26_27 = pd.read_excel(xls, sheet_name='26 27', header=1).dropna(axis=1, how='all')
    
    # 3. Clean up column names (Convert timestamps to Month Names)
    def clean_columns(df):
        new_cols = []
        for c in df.columns:
            if 'Dept' in str(c):
                new_cols.append('Department')
            else:
                try:
                    # Convert any date-like header to a 3-letter uppercase month
                    new_cols.append(pd.to_datetime(c).strftime('%b').upper())
                except:
                    new_cols.append(str(c))
        df.columns = new_cols
        return df

    df_25_26 = clean_columns(df_25_26)
    df_26_27 = clean_columns(df_26_27)

    # 4. Show FY 25-26 Data 
    st.header("Historical Month-Wise Figures (FY 25-26)")
    # use_container_width=False forces a horizontal scrollbar instead of squishing columns
    st.dataframe(df_25_26.set_index('Department'), use_container_width=False)

    # 5. Show FY 26-27 Editable Data 
    st.header("Enter New Month Figures (FY 26-27)")
    st.write("Edit the table below to add figures for August through March:")
    edited_df_26_27 = st.data_editor(df_26_27.set_index('Department'), num_rows="dynamic").reset_index()

    # 6. Filters & Dropdowns
    st.sidebar.header("Filters")
    departments = edited_df_26_27['Department'].dropna().unique()
    selected_dept = st.sidebar.selectbox("Select Department", ["All"] + list(departments))
    
    months = ['APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC', 'JAN', 'FEB', 'MAR']
    
    # Filter for the top section (YTD)
    selected_month = st.sidebar.selectbox("Calculate YTD Up To Month:", months, index=3) 
    # NEW: Independent filter for the bottom section (For the Month)
    selected_month_for = st.sidebar.selectbox("Show 'For the Month' Results Up To:", months, index=3) 

    # -------------------------------------------------------------------
    # 7. Calculate YTD Data Dynamically (Side-by-Side Cumulative)
    # -------------------------------------------------------------------
    idx = months.index(selected_month)
    ytd_months = months[:idx+1]
    
    merged_ytd = df_25_26[['Department']].copy()
    
    # Loop through each month up to the selected one to build cumulative columns
    for i, m in enumerate(ytd_months):
        curr_months = months[:i+1]
        merged_ytd[f'Up to {m} 25'] = df_25_26[curr_months].apply(pd.to_numeric, errors='coerce').sum(axis=1)
        merged_ytd[f'Up to {m} 26'] = edited_df_26_27[curr_months].apply(pd.to_numeric, errors='coerce').sum(axis=1)

    # Calculate final accretion comparing the very last month selected
    merged_ytd['Accretion'] = merged_ytd[f'Up to {selected_month} 26'] - merged_ytd[f'Up to {selected_month} 25']
    merged_ytd['Accretion %'] = np.where(
        merged_ytd[f'Up to {selected_month} 25'] != 0, 
        (merged_ytd['Accretion'] / merged_ytd[f'Up to {selected_month} 25']) * 100, 
        0
    )

    # Apply Department Filter and format the final YTD table
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
        use_container_width=False, # Allows horizontal scroll if wide
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
    # 9. NEW SECTION: "For the Month" Dynamically (Side-by-Side Individuals)
    # -------------------------------------------------------------------
    start_month = months[0]
    idx_for = months.index(selected_month_for)
    for_months = months[:idx_for+1]
    
    merged_range = df_25_26[['Department']].copy()

    # Loop to grab the individual month figures instead of summing them
    for m in for_months:
        merged_range[f'For {m} 25'] = pd.to_numeric(df_25_26[m], errors='coerce').fillna(0)
        merged_range[f'For {m} 26'] = pd.to_numeric(edited_df_26_27[m], errors='coerce').fillna(0)

    # Calculate Accretion specifically for the end month chosen in the dropdown
    merged_range['Accretion'] = merged_range[f'For {selected_month_for} 26'] - merged_range[f'For {selected_month_for} 25']
    merged_range['Accretion %'] = np.where(
        merged_range[f'For {selected_month_for} 25'] != 0, 
        (merged_range['Accretion'] / merged_range[f'For {selected_month_for} 25']) * 100, 
        0
    )

    # Apply Department Filter and format the final range table
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
        use_container_width=False, # Allows horizontal scroll if wide
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

except Exception as e:
    st.error(f"Error loading data. Please ensure 'for github stats.xls' is formatted correctly. Details: {e}")
