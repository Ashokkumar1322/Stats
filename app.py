import os
from datetime import date, datetime
from io import BytesIO

import pandas as pd
import streamlit as st

# -----------------------------------------------------------------------------
# PAGE CONFIGURATION
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="Premium Figures Dashboard",
    layout="wide",
    page_icon="📊",
)
st.title("📊 Complete Premium Figures Dashboard")

# -----------------------------------------------------------------------------
# DATA SOURCE
# -----------------------------------------------------------------------------
st.sidebar.header("📁 Data Source")
uploaded_file = st.sidebar.file_uploader(
    "Upload your Excel file", type=["xls", "xlsx"]
)

DEFAULT_FILE = "for github stats.xls"
file_source = uploaded_file if uploaded_file is not None else DEFAULT_FILE

if isinstance(file_source, str) and not os.path.exists(file_source):
    st.error(
        f"Could not find '{DEFAULT_FILE}'. Upload the Excel file using the sidebar."
    )
    st.stop()

if st.sidebar.button("🔄 Refresh Data / Clear Cache"):
    st.cache_data.clear()
    st.rerun()


def normalise_sheet_name(name):
    """Remove accidental leading/trailing spaces from worksheet names."""
    return str(name).strip()


@st.cache_data(show_spinner="Loading Excel workbook…")
def load_all_excel_sheets(file_bytes, file_name, file_mtime=None):
    """Read every worksheet and clean its headers and data."""
    extension = os.path.splitext(file_name.lower())[1]
    engine = "xlrd" if extension == ".xls" else "openpyxl"

    workbook = pd.ExcelFile(BytesIO(file_bytes), engine=engine)
    cleaned_sheets = {}

    for original_name in workbook.sheet_names:
        sheet_name = normalise_sheet_name(original_name)
        df = pd.read_excel(
            workbook,
            sheet_name=original_name,
            header=1,
        ).copy()

        # Remove completely empty columns, including Excel's Unnamed columns.
        df = df.dropna(axis=1, how="all")
        df = df.loc[
            :,
            ~(
                df.columns.astype(str).str.startswith("Unnamed")
                & ~df.notna().any()
            ),
        ]

        # Clean column names. Convert Excel date headers to APR-26 style.
        new_columns = []
        for column in df.columns:
            # Excel date headers may arrive as pandas Timestamp, Python
            # datetime/date, or a date-like string. Display them as APR-26.
            if isinstance(column, (pd.Timestamp, datetime, date)):
                new_columns.append(pd.Timestamp(column).strftime("%b-%y").upper())
            else:
                column_text = str(column).strip()
                parsed_date = pd.to_datetime(
                    column_text, errors="coerce"
                ) if column_text else pd.NaT

                # Only convert strings that look like full Excel dates;
                # leave ordinary headings such as TOTAL or Dept unchanged.
                if (
                    pd.notna(parsed_date)
                    and any(sep in column_text for sep in ("-", "/", ":"))
                    and len(column_text) >= 7
                ):
                    new_columns.append(parsed_date.strftime("%b-%y").upper())
                else:
                    new_columns.append(column_text)
        df.columns = new_columns

        # Keep the first column and known identification/text fields as text.
        text_columns = {
            "Party Code", "Party Name", "AGENT", "BROKER", "POSP",
            "MISP's NAME", "DEALER CODE", "LOB Name", "Dept",
            "Dept ", "Dept  FOR THE MONTH", "Dept   UPTO THE MONTH",
            "Dept Code", "Sl no", "SL NO",
        }
        first_column = df.columns[0] if len(df.columns) else None

        for column in df.columns:
            if column == first_column or column in text_columns:
                continue
            # Numeric conversion is safe for figures and leaves text as NaN.
            df[column] = pd.to_numeric(df[column], errors="coerce")

        cleaned_sheets[sheet_name] = df.round(2)

    return cleaned_sheets


# Streamlit's uploader is read as bytes. Reading local files as bytes also
# ensures that a changed workbook gets a different cache key.
try:
    if uploaded_file is not None:
        file_bytes = uploaded_file.getvalue()
        file_name = uploaded_file.name
        file_mtime = None
    else:
        file_name = DEFAULT_FILE
        file_mtime = os.path.getmtime(file_name)
        with open(file_name, "rb") as excel_file:
            file_bytes = excel_file.read()

    all_sheets = load_all_excel_sheets(file_bytes, file_name, file_mtime)
    sheet_names = list(all_sheets.keys())

    if not sheet_names:
        st.error("No worksheets were found in the Excel file.")
        st.stop()

    st.sidebar.success(f"Successfully loaded {len(sheet_names)} sheets!")

    # -------------------------------------------------------------------------
    # SHEET NAVIGATION
    # -------------------------------------------------------------------------
    st.sidebar.markdown("---")
    st.sidebar.header("📑 View Sheets")
    selected_sheet = st.sidebar.radio(
        "Select the sheet you want to view:", sheet_names
    )

    # -------------------------------------------------------------------------
    # KPI HELPERS - use sheet/column matching so minor renaming is manageable.
    # -------------------------------------------------------------------------
    def find_sheet(name_candidates):
        for candidate in name_candidates:
            candidate_clean = candidate.strip().casefold()
            for actual_name in all_sheets:
                if actual_name.strip().casefold() == candidate_clean:
                    return all_sheets[actual_name]
        return None

    def find_total_row(df, terms):
        if df is None or df.empty:
            return None
        first_col = df.columns[0]
        values = df[first_col].astype(str)
        for term in terms:
            matched = df[values.str.contains(term, case=False, na=False)]
            if not matched.empty:
                return matched.iloc[0]
        return None

    total_premium_value = "N/A"
    accretion_label = "Accretion"
    accretion_value = "N/A"
    icr_total_value = "N/A"
    icr_earned_value = "N/A"

    # Total premium: current workbook has this in sheet '26 27'.
    df_current = find_sheet(["26 27"])
    if df_current is not None and "TOTAL" in df_current.columns:
        total_row = find_total_row(df_current, ["Sum for all"])
        if total_row is not None and pd.notna(total_row["TOTAL"]):
            total_premium_value = f"₹ {int(round(total_row['TOTAL'])):,}"

    # Latest cumulative accretion: find the last column containing ACCRETION.
    df_upto = find_sheet(["25 26 26 27 Up to the month"])
    if df_upto is not None:
        accretion_columns = [
            col for col in df_upto.columns
            if "ACCRETION" in str(col).upper()
        ]
        if accretion_columns:
            latest_column = accretion_columns[-1]
            accretion_label = str(latest_column)
            total_row = find_total_row(df_upto, ["Sum for all"])
            if total_row is not None and pd.notna(total_row[latest_column]):
                accretion_value = f"{total_row[latest_column]:.2f}%"

    # Updated ICR worksheet name in the uploaded workbook:
    # 'Claim ICR on Prem EP' (previously 'ICR on Total Premium and EP').
    df_icr = find_sheet([
        "Claim ICR on Prem EP",
        "ICR on Total Premium and EP",
    ])
    if df_icr is not None:
        grand_total_row = find_total_row(df_icr, ["TOTAL", "Grand Total"])
        if grand_total_row is not None:
            total_icr_col = next(
                (c for c in df_icr.columns if str(c).strip().casefold() == "icr"),
                None,
            )
            earned_icr_col = next(
                (
                    c for c in df_icr.columns
                    if "ICR ON EARNED PREMIUM" in str(c).upper()
                ),
                None,
            )
            if total_icr_col and pd.notna(grand_total_row[total_icr_col]):
                icr_total_value = f"{grand_total_row[total_icr_col]:.2f}%"
            if earned_icr_col and pd.notna(grand_total_row[earned_icr_col]):
                icr_earned_value = f"{grand_total_row[earned_icr_col]:.2f}%"

    # -------------------------------------------------------------------------
    # TOP KPI CARDS
    # -------------------------------------------------------------------------
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Total Premium (FY 26-27)", total_premium_value)
    k2.metric(accretion_label.title(), accretion_value)
    k3.metric("ICR on Total Premium", icr_total_value)
    k4.metric("ICR on Earned Premium", icr_earned_value)

    st.markdown("---")

    # -------------------------------------------------------------------------
    # SELECTED WORKSHEET AND FILTERS
    # -------------------------------------------------------------------------
    st.header(f"📄 {selected_sheet}")
    df = all_sheets[selected_sheet]
    final_df = df.copy()

    # Format figures for display.
    col_format_config = {}
    for column in df.columns:
        upper = str(column).upper()
        if any(term in upper for term in ["ICR", "ACCRETION", "SHARE"]):
            col_format_config[column] = st.column_config.NumberColumn(
                format="%.2f"
            )
        elif any(
            term in upper
            for term in [
                "PREMIUM", "TOTAL", "APR-", "MAY-", "JUN-", "JUL-",
                "AUG-", "SEP-", "OCT-", "NOV-", "DEC-", "JAN-",
                "FEB-", "MAR-",
            ]
        ):
            col_format_config[column] = st.column_config.NumberColumn(
                format="₹ %,.2f"
            )

    department_sheets = {
        "25 26",
        "26 27",
        "25 26 26 27 For the month",
        "25 26 26 27 Up to the month",
        "Claim ICR on Prem EP",
        "ICR on Total Premium and EP",
    }
    channel_sheets = {
        "Channel wise 25 26",
        "Channel wise 26 27",
    }

    if selected_sheet in department_sheets:
        st.subheader("Filters")
        department_column = df.columns[0]
        col1, col2 = st.columns([1, 2])

        with col1:
            departments = df[department_column].dropna().unique().tolist()
            selected_department = st.selectbox(
                "Filter by Department / LOB:",
                options=["All"] + departments,
            )

        with col2:
            available_columns = df.columns[1:].tolist()
            selected_columns = st.multiselect(
                "Select Columns to Display:",
                options=available_columns,
                default=available_columns,
            )

        if selected_department != "All":
            final_df = final_df[
                final_df[department_column] == selected_department
            ]
        final_df = final_df[[department_column] + selected_columns]

    elif selected_sheet in channel_sheets:
        st.subheader("Filters")
        col1, col2, col3 = st.columns(3)
        agent_search = broker_search = posp_search = ""

        agent_column = next(
            (c for c in ["AGENT", "Party Name"] if c in df.columns), None
        )
        with col1:
            if agent_column:
                agent_search = st.text_input(f"Search by {agent_column}:")
        with col2:
            if "BROKER" in df.columns:
                broker_search = st.text_input("Search by BROKER:")
        with col3:
            if "POSP" in df.columns:
                posp_search = st.text_input("Search by POSP:")

        if agent_search and agent_column:
            final_df = final_df[
                final_df[agent_column].astype(str).str.contains(
                    agent_search, case=False, na=False, regex=False
                )
            ]
        if broker_search and "BROKER" in final_df.columns:
            final_df = final_df[
                final_df["BROKER"].astype(str).str.contains(
                    broker_search, case=False, na=False, regex=False
                )
            ]
        if posp_search and "POSP" in final_df.columns:
            final_df = final_df[
                final_df["POSP"].astype(str).str.contains(
                    posp_search, case=False, na=False, regex=False
                )
            ]

    elif selected_sheet == "MISP":
        st.subheader("Filters")
        misp_column = next(
            (c for c in ["MISP's NAME", "MISP"] if c in df.columns), None
        )
        if misp_column:
            misp_search = st.text_input(f"Search by {misp_column}:")
            if misp_search:
                final_df = final_df[
                    final_df[misp_column].astype(str).str.contains(
                        misp_search, case=False, na=False, regex=False
                    )
                ]

    # Display table and CSV download.
    st.dataframe(
        final_df,
        use_container_width=True,
        hide_index=True,
        column_config=col_format_config,
    )

    csv_data = final_df.to_csv(index=False).encode("utf-8-sig")
    st.download_button(
        label=f"📥 Download '{selected_sheet}' Data as CSV",
        data=csv_data,
        file_name=f"{selected_sheet.replace(' ', '_')}_filtered.csv",
        mime="text/csv",
    )

except Exception as error:
    st.error(f"An error occurred while loading the dashboard: {error}")
    st.info(
        "If you are using an .xls file, ensure xlrd is listed in requirements.txt. "
        "For an .xlsx file, openpyxl is required."
    )
