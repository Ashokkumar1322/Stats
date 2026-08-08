import streamlit as st
import pandas as pd

st.set_page_config(page_title="Premium Figures Dashboard", layout="wide")
st.title("📊 Department-Wise Premium & Accretion Dashboard")

# 1. File Upload or Default File
uploaded_file = st.sidebar.file_upload("Upload your Excel File", type=['xls', 'xlsx'])
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
            elif isinstance(c, pd.Timestamp):
                new_cols.append(c.strftime('%b').upper())
            else:
                new_cols.append(str(c))
        df.columns = new_cols
        return df

    df_25_26 = clean_columns(df_25_26)
    df_26_27 = clean_columns(df_26_27)

    st.header("Enter New Month Figures (FY 26-27)")
    st.write("Edit the table below to add figures for August through March:")
    # st.data_editor allows you to type directly into the web app!
    edited_df_26_27 = st.data_editor(df_26_27, num_rows="dynamic")

    # 4. Filters & Dropdowns
    st.sidebar.header("Filters")
    departments = edited_df_26_27['Department'].dropna().unique()
    selected_dept = st.sidebar.selectbox("Select Department", ["All"] + list(departments))
    
    months = ['APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC', 'JAN', 'FEB', 'MAR']
    selected_month = st.sidebar.selectbox("Calculate YTD Up To Month:", months, index=3) # Defaults to JUL

    # 5. Calculate Accretion up to selected month
    idx = months.index(selected_month)
    ytd_months = months[:idx+1]
    
    df_25_26['YTD_25_26'] = df_25_26[ytd_months].sum(axis=1)
    edited_df_26_27['YTD_26_27'] = edited_df_26_27[ytd_months].sum(axis=1)

    # Merge for comparison
    merged = pd.merge(df_25_26[['Department', 'YTD_25_26']], edited_df_26_27[['Department', 'YTD_26_27']], on='Department')
    merged['Accretion'] = merged['YTD_26_27'] - merged['YTD_25_26']

    # Apply Department Filter
    if selected_dept != "All":
        merged = merged[merged['Department'] == selected_dept]

    # 6. Display Results
    st.header(f"YTD Accretion Results (Up to {selected_month})")
    st.dataframe(merged.sort_values(by="Accretion", ascending=False))
    
    # Best/Worst performers summary
    if selected_dept == "All":
        st.subheader("Insights")
        best = merged.iloc[0]
        worst = merged.iloc[-1]
        st.success(f"📈 **Best Performing:** {best['Department']} with an accretion of {best['Accretion']:,.0f}")
        st.error(f"📉 **Needs Attention:** {worst['Department']} with an accretion of {worst['Accretion']:,.0f}")

except Exception as e:
    st.error(f"Error loading data. Please ensure 'for github stats.xls' is formatted correctly. Details: {e}")
