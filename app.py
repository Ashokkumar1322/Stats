import streamlit as st
import pandas as pd

# --- SUPABASE & IMAGE IMPORTS ---
from supabase import create_client
import io
import datetime

# --- SUPABASE CONNECTION ---
@st.cache_resource
def init_connection():
    # These secrets must be added to your Streamlit Community Cloud dashboard!
    url = st.secrets["SUPABASE_URL"]
    key = st.secrets["SUPABASE_KEY"]
    return create_client(url, key)

try:
    supabase = init_connection()
except Exception:
    supabase = None

st.set_page_config(page_title="Premium Figures Dashboard", layout="wide")
st.title("📊 Complete Premium Dashboard")

# --- MAP SHEETS TO SUPABASE TABLES ---
TABLE_MAPPING = {
    '25 26': 'premium-comparison',
    '26 27': '', 
    '25 26 26 27 For the month': '',
    '25 26 26 27 Up to the month': '',
    'Channel wise 25 26': '',
    'Channel wise 26 27': '',
    'ICR on Total Premium and EP': ''
}

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
        
        # Get the correct Supabase table name for this specific sheet
        target_supabase_table = TABLE_MAPPING.get(sheet)
        
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
                
                st.markdown("**📝 Instructions: Double-click any cell below to edit it.**")
                
                # Editable Table
                edited_df = st.data_editor(
                    filtered_df[columns_to_show], 
                    use_container_width=False, 
                    hide_index=True,
                    column_config=col_format_config,  # <--- APPLIED THE NEW FORMATTING HERE
                    key=f"editor_{sheet}"
                )

                # Save Data to Supabase
                if st.button(f"💾 Save updates for {sheet} to Database", key=f"save_btn_{sheet}"):
                    if supabase and target_supabase_table:
                        try:
                            records = edited_df.fillna("").to_dict(orient="records")
                            response = supabase.table(target_supabase_table).upsert(records).execute()
                            st.success(f"✅ Data successfully saved to the '{target_supabase_table}' table in Supabase!")
                        except Exception as e:
                            st.error(f"❌ Error saving to Supabase: {e}")
                            st.info("💡 Hint: Make sure your Supabase table columns exactly match the Streamlit table headers, and that you have a Primary Key set.")
                    elif not target_supabase_table:
                        st.error(f"⚠️ No Supabase table mapped for the sheet '{sheet}'. Please update TABLE_MAPPING in the code.")
                    else:
                        st.error("⚠️ Supabase is not connected. Check your Streamlit secrets.")

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
                    
                st.markdown("**📝 Instructions: Double-click any cell below to edit it.**")
                
                # Editable Table
                edited_channel_df = st.data_editor(
                    filtered_df, 
                    use_container_width=False, 
                    hide_index=True,
                    column_config=col_format_config,  # <--- APPLIED THE NEW FORMATTING HERE
                    key=f"editor_ch_{sheet}"
                )

                # Save Data to Supabase
                if st.button(f"💾 Save updates for {sheet} to Database", key=f"save_btn_ch_{sheet}"):
                    if supabase and target_supabase_table:
                        try:
                            records = edited_channel_df.fillna("").to_dict(orient="records")
                            response = supabase.table(target_supabase_table).upsert(records).execute()
                            st.success(f"✅ Data successfully saved to the '{target_supabase_table}' table in Supabase!")
                        except Exception as e:
                            st.error(f"❌ Error saving to Supabase: {e}")
                            st.info("💡 Hint: Make sure your Supabase table columns exactly match the Streamlit table headers, and that you have a Primary Key set.")
                    elif not target_supabase_table:
                        st.error(f"⚠️ No Supabase table mapped for the sheet '{sheet}'. Please update TABLE_MAPPING in the code.")
                    else:
                        st.error("⚠️ Supabase is not connected. Check your Streamlit secrets.")

except FileNotFoundError:
    st.error("Could not find 'for github stats.xls'. Please use the sidebar to upload the file manually.")
except Exception as e:
    st.error(f"An error occurred: {e}")
