import streamlit as st
import pandas as pd
import numpy as np # Added for the percentage calculation

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

    # 4. Show FY 25-26 Data (Setting index freezes the Department column!)
    st.header("Historical Month-Wise Figures (FY 25-26)")
    st.dataframe(df_25_26.set_index('Department'))

    # 5. Show FY 26-27 Editable Data (Setting index freezes the Department column!)
    st.header("Enter New Month Figures (FY 26-27)")
    st.write("Edit the table below to add figures for August through March:")
    # We set the index to freeze it, then reset it after editing so calculations still work
    edited_df_26_27 = st.data_editor(df_26_27.set_index('Department'), num_rows="dynamic").reset_index()

    # 6. Filters & Dropdowns
    st.sidebar.header("Filters")
    departments = edited_df_26_27['Department'].dropna().unique()
    selected_dept = st.sidebar.selectbox("Select Department", ["All"] + list(departments))
    
    months = ['APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC', 'JAN', 'FEB', 'MAR']
    selected_month = st.sidebar.selectbox("Calculate YTD Up To Month:", months, index=3) # Defaults to JUL

    # 7. Calculate Accretion up to selected month
    idx = months.index(selected_month)
    ytd_months = months[:idx+1]
    
    # Coercing to numeric is safer when using data_editor so string typos don't break the sum
    df_25_26['YTD_25_26'] = df_25_26[ytd_months].apply(pd.to_numeric, errors='coerce').sum(axis=1)
    edited_df_26_27['YTD_26_27'] = edited_df_26_27[ytd_months].apply(pd.to_numeric, errors='coerce').sum(axis=1)

    # Merge for comparison
    merged = pd.merge(df_25_26[['Department', 'YTD_25_26']], edited_df_26_27[['Department', 'YTD_26_27']], on='Department')
    merged['Accretion'] = merged['YTD_26_27'] - merged['YTD_25_26']
    
    # NEW: Calculate Accretion Percentage and multiply by 100 here!
    merged['Accretion %'] = np.where(
        merged['YTD_25_26'] != 0, 
        (merged['Accretion'] / merged['YTD_25_26']) * 100, 
        0
    )

    # Apply Department Filter and format the final table
    if selected_dept != "All":
        merged = merged[merged['Department'] == selected_dept]
        final_table = merged.sort_values(by="Accretion", ascending=False).reset_index(drop=True)
    else:
        # Separate the "Sum" row, sort the rest, then put "Sum" at the bottom
        sum_row = merged[merged['Department'] == 'Sum for all departments']
        other_rows = merged[merged['Department'] != 'Sum for all departments']
        
        other_rows = other_rows.sort_values(by="Accretion", ascending=False)
        final_table = pd.concat([other_rows, sum_row]).reset_index(drop=True)
        
    # Make index start from 1 instead of 0 so serial numbers are clean
    final_table.index = final_table.index + 1

    # 8. Display Results
    st.header(f"YTD Accretion Results (Up to {selected_month})")
    
    # Display dataframe with formatted percentage column
    st.dataframe(
        final_table,
        column_config={
            "Accretion %": st.column_config.NumberColumn(
                "Accretion (%)",
                format="%.2f%%"
            )
        }
    )
    
    # Best/Worst performers summary (Removed the *100 here since it's already done above)
    if selected_dept == "All" and len(final_table) > 1:
        st.subheader("Insights")
        best = final_table.iloc[0]
        worst = final_table.iloc[-2] # -2 because the last row (-1) is the Sum row!
        st.success(f"📈 **Best Performing:** {best['Department']} with an accretion of {best['Accretion']:,.0f} ({best['Accretion %']:.2f}%)")
        st.error(f"📉 **Needs Attention:** {worst['Department']} with an accretion of {worst['Accretion']:,.0f} ({worst['Accretion %']:.2f}%)")
# -------------------------------------------------------------------
    # 9. NEW SECTION: "For the Month" (Range: Apr to Selected Month)
    # -------------------------------------------------------------------
    start_month = months[0] # Always 'APR'
    
    # Calculate the sum of the range (Apr to Selected Month)
    df_25_26['Range_25_26'] = df_25_26[ytd_months].apply(pd.to_numeric, errors='coerce').sum(axis=1)
    edited_df_26_27['Range_26_27'] = edited_df_26_27[ytd_months].apply(pd.to_numeric, errors='coerce').sum(axis=1)

    # Merge for comparison
    merged_range = pd.merge(df_25_26[['Department', 'Range_25_26']], edited_df_26_27[['Department', 'Range_26_27']], on='Department')
    merged_range['Accretion'] = merged_range['Range_26_27'] - merged_range['Range_25_26']
    
    # Calculate Accretion Percentage
    merged_range['Accretion %'] = np.where(
        merged_range['Range_25_26'] != 0, 
        (merged_range['Accretion'] / merged_range['Range_25_26']) * 100, 
        0
    )

    # Apply Department Filter and format the final table
    if selected_dept != "All":
        merged_range = merged_range[merged_range['Department'] == selected_dept]
        final_table_range = merged_range.sort_values(by="Accretion", ascending=False).reset_index(drop=True)
    else:
        # Separate the "Sum" row, sort the rest, then put "Sum" at the bottom
        sum_row_range = merged_range[merged_range['Department'] == 'Sum for all departments']
        other_rows_range = merged_range[merged_range['Department'] != 'Sum for all departments']
        
        other_rows_range = other_rows_range.sort_values(by="Accretion", ascending=False)
        final_table_range = pd.concat([other_rows_range, sum_row_range]).reset_index(drop=True)
        
    # Make index start from 1 instead of 0
    final_table_range.index = final_table_range.index + 1

    # Display Results (Title dynamically updates to "For APR to JUL", "For APR to AUG", etc.)
    st.header(f"YTD Accretion Results (For {start_month} to {selected_month})")
    
    st.dataframe(
        final_table_range,
        column_config={
            "Accretion %": st.column_config.NumberColumn(
                "Accretion (%)",
                format="%.2f%%"
            )
        }
    )
    
    # Best/Worst performers summary for the range
    if selected_dept == "All" and len(final_table_range) > 1:
        st.subheader(f"Insights (For {start_month} to {selected_month})")
        best_range = final_table_range.iloc[0]
        worst_range = final_table_range.iloc[-2] # -2 because the last row is the Sum row
        st.success(f"📈 **Best Performing ({start_month} to {selected_month}):** {best_range['Department']} with an accretion of {best_range['Accretion']:,.0f} ({best_range['Accretion %']:.2f}%)")
        st.error(f"📉 **Needs Attention ({start_month} to {selected_month}):** {worst_range['Department']} with an accretion of {worst_range['Accretion']:,.0f} ({worst_range['Accretion %']:.2f}%)")
    # -------------------------------------------------------------------except Exception as e:
    st.error(f"Error loading data. Please ensure 'for github stats.xls' is formatted correctly. Details: {e}")
