import streamlit as st
import pandas as pd

# --- SUPABASE & IMAGE IMPORTS ---
from supabase import create_client
import io
import datetime
import matplotlib.pyplot as plt

# --- SUPABASE CONNECTION ---
@st.cache_resource
def init_connection():
    url = st.secrets["SUPABASE_URL"]
    key = st.secrets["SUPABASE_KEY"]
    return create_client(url, key)

try:
    supabase = init_connection()
except Exception:
    supabase = None

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
        
        # FIX #1 & #2: Read data without aggressively dropping empty future months
        df = pd.read_excel(xls, sheet_name=sheet, header=1)
        
        # Drop columns that are completely empty AND unnamed (keeps empty future months)
        df = df.loc[:, ~df.columns.astype(str).str.contains('^Unnamed') | df.notna().any()]
        
        # Clean datetime column headers (Months/Years)
        new_cols = []
        for c in df.columns:
            if isinstance(c, pd.Timestamp) or hasattr(c, 'strftime'):
                new_cols.append(pd.to_datetime(c).strftime('%b-%y').upper())
            else:
                new_cols.append(str(c).strip())
        df.columns = new_cols

        st.divider()
        
        with st.expander(f"📁 View Data & Filters for: {sheet}", expanded=True):
            
            # -------------------------------------------------------------
            # DEPARTMENT SHEETS FILTERS
            # -------------------------------------------------------------
            if sheet in ['25 26', '26 27', '25 26 26 27 For the month', '25 26 26 27 Up to the month']:
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
                
                columns_to_show = [dept_col] + selected_months
                
                # FIX #3: Use data_editor for editing and horizontal scrolling
                st.markdown("**Edit your data directly in the table below:**")
                edited_df = st.data_editor(
                    filtered_df[columns_to_show], 
                    use_container_width=False, # Setting False forces horizontal scroll for many columns
                    hide_index=True,
                    key=f"editor_{sheet}"
                )

                # Save / Export Edited Data
                if st.button(f"💾 Save updates for {sheet}", key=f"save_btn_{sheet}"):
                    # Provide a quick CSV download of the edited data
                    csv = edited_df.to_csv(index=False)
                    st.download_button(
                        label="Download Updated CSV", 
                        data=csv, 
                        file_name=f"{sheet}_updated.csv", 
                        mime="text/csv", 
                        key=f"dl_{sheet}"
                    )
                    
                    # Supabase Database Save Logic (Requires a table setup)
                    if supabase:
                        st.info("💡 **To save to Supabase Database:** Ensure you have a table created. Then uncomment the Supabase insertion code in the script.")
                        # --- Uncomment and update 'your_table_name' to sync to DB ---
                        # records = edited_df.fillna("").to_dict(orient="records")
                        # response = supabase.table("your_table_name").upsert(records).execute()
                        # st.success("Saved to Supabase DB!")

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
                    
                # Interactive Editor for Channels
                edited_channel_df = st.data_editor(
                    filtered_df, 
                    use_container_width=False, 
                    hide_index=True,
                    key=f"editor_ch_{sheet}"
                )

except FileNotFoundError:
    st.error("Could not find 'for github stats.xls'. Please use the sidebar to upload the file manually.")
except Exception as e:
    st.error(f"An error occurred: {e}")


# --- GENERATE FIGURE & UPLOAD TO SUPABASE ---
st.divider()
st.header("📈 Monthly Stats Figure")

fig, ax = plt.subplots(figsize=(6, 3))
ax.plot(["Jan", "Feb", "Mar"], [100, 250, 200], marker="o", color="green")
ax.set_title("Sample Monthly Performance")
st.pyplot(fig)

if supabase:
    if st.button("Save this month's stats to Supabase"):
        buf = io.BytesIO()
        fig.savefig(buf, format="png")
        buf.seek(0)
        
        current_month = datetime.datetime.now().strftime("%Y_%m")
        file_name = f"stats_{current_month}.png"
        
        try:
            response = supabase.storage.from_("github-stats").upload(
                file=buf.getvalue(),
                path=file_name,
                file_options={"content-type": "image/png"}
            )
            st.success(f"Successfully saved '{file_name}' to Supabase!")
        except Exception as e:
            st.error(f"Error saving file to Supabase: {e}")
else:
    st.warning("⚠️ Supabase is not connected. Please add your Secrets in the Streamlit Dashboard to enable uploading.")
