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
    # KPI FIGURES - read the exact cells from the requested worksheets.
    # -------------------------------------------------------------------------
    def format_indian_number(value):
        """Format a number using Indian digit grouping (3, then groups of 2)."""
        number = int(round(float(value)))
        sign = "-" if number < 0 else ""
        digits = str(abs(number))
        if len(digits) <= 3:
            return sign + digits
        last_three = digits[-3:]
        remaining = digits[:-3]
        groups = []
        while remaining:
            groups.insert(0, remaining[-2:])
            remaining = remaining[:-2]
        return sign + ",".join(groups + [last_three])

    def parse_month_header(value):
        """Return a pandas Timestamp for a date-like Excel column heading."""
        if isinstance(value, (pd.Timestamp, datetime, date)):
            return pd.Timestamp(value)
        if isinstance(value, str):
            parsed = pd.to_datetime(value.strip(), errors="coerce")
            if pd.notna(parsed):
                return pd.Timestamp(parsed)
        return pd.NaT

    total_premium_value = "N/A"
    premium_label = "Up to latest month Premium"
    accretion_label = "Accretion % Up to latest month"
    accretion_value = "N/A"
    icr_total_value = "N/A"
    icr_earned_value = "N/A"

    extension = os.path.splitext(file_name.lower())[1]
    engine = "xlrd" if extension == ".xls" else "openpyxl"

    # 1 & 2. Premium and accretion are on the total row (Excel row 29)
    # of 'revised upto mon comp 2526 2627'. Find the latest FY 2026-27
    # month heading that has its matching 'Accretion upto ...' heading.
    monthly_sheet_name = next(
        (
            name for name in sheet_names
            if name.strip().casefold() == "revised upto mon comp 2526 2627"
        ),
        None,
    )

    if monthly_sheet_name is not None:
        df_monthly_raw = pd.read_excel(
            BytesIO(file_bytes),
            sheet_name=monthly_sheet_name,
            header=None,
            engine=engine,
        )

        # Excel row 2 contains the month headings; Excel row 29 contains
        # 'Sum for all departments'. These are zero-based rows 1 and 28.
        if df_monthly_raw.shape[0] >= 29:
            header_row = df_monthly_raw.iloc[1]
            total_row = df_monthly_raw.iloc[28]
            monthly_candidates = []

            for column_index, heading in enumerate(header_row):
                month_date = parse_month_header(heading)
                if (
                    pd.isna(month_date)
                    or month_date < pd.Timestamp("2026-04-01")
                    or month_date > pd.Timestamp("2027-03-31")
                ):
                    continue
                accretion_index = column_index + 1
                if accretion_index >= len(header_row):
                    continue

                accretion_heading = str(header_row.iloc[accretion_index]).strip()
                expected_month = month_date.strftime("%b %y")
                if "accretion" not in accretion_heading.casefold():
                    continue
                if expected_month.casefold() not in accretion_heading.casefold():
                    continue

                premium_cell = total_row.iloc[column_index]
                accretion_cell = total_row.iloc[accretion_index]
                if pd.notna(premium_cell) and pd.notna(accretion_cell):
                    monthly_candidates.append(
                        (month_date, premium_cell, accretion_cell)
                    )

            if monthly_candidates:
                # The latest populated month is selected automatically.
                latest_month, premium_cell, accretion_cell = max(
                    monthly_candidates, key=lambda item: item[0]
                )
                month_label = latest_month.strftime("%b %y")
                premium_label = f"Up to {month_label} Premium"
                accretion_label = f"Accretion % Up to {month_label}"
                total_premium_value = f"₹ {format_indian_number(premium_cell)}"
                accretion_value = f"{float(accretion_cell):.2f}%"
        else:
            st.warning(
                "The monthly comparison sheet does not contain the expected "
                "total row (Excel row 29)."
            )
    else:
        st.warning(
            "Worksheet 'revised upto mon comp 2526 2627' was not found."
        )

    # 3 & 4. ICR values are read directly from F25 and G25 in
    # 'dept wise icr premium Ep'.
    icr_sheet_name = next(
        (
            name for name in sheet_names
            if name.strip().casefold() == "dept wise icr premium ep"
        ),
        None,
    )
    if icr_sheet_name is not None:
        df_icr_raw = pd.read_excel(
            BytesIO(file_bytes),
            sheet_name=icr_sheet_name,
            header=None,
            engine=engine,
        )
        if df_icr_raw.shape[0] >= 25 and df_icr_raw.shape[1] >= 7:
            total_icr_cell = df_icr_raw.iloc[24, 5]   # Excel F25
            earned_icr_cell = df_icr_raw.iloc[24, 6]  # Excel G25

            if pd.notna(total_icr_cell):
                icr_total_value = f"{float(total_icr_cell):.2f}%"
            if pd.notna(earned_icr_cell):
                icr_earned_value = f"{float(earned_icr_cell):.2f}%"
    else:
        st.warning("Worksheet 'dept wise icr premium Ep' was not found.")

    # -------------------------------------------------------------------------
    # TOP KPI CARDS
    # -------------------------------------------------------------------------
    k1, k2, k3, k4 = st.columns(4)
    k1.metric(premium_label, total_premium_value)
    k2.metric(accretion_label, accretion_value)
    k3.metric("ICR % on Total Premium", icr_total_value)
    k4.metric("ICR % on Earned Premium", icr_earned_value)

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
        "Dept Wise ICR Prem EP",
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

    # Keep the first (Department) column pinned while scrolling horizontally.
    if len(final_df.columns) > 0:
        first_column = final_df.columns[0]
        col_format_config.setdefault(
            first_column,
            st.column_config.TextColumn(pinned="left"),
        )

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
