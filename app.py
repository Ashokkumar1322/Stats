
import os
import re
import html
from datetime import date, datetime
from io import BytesIO

import pandas as pd
import streamlit as st

# ---------------------------------------------------------
# PAGE CONFIGURATION
# ---------------------------------------------------------
st.set_page_config(
    page_title="Premium Figures Dashboard",
    layout="wide",
    page_icon="📊",
)
st.title("📊 Complete Premium Figures Dashboard")

# ---------------------------------------------------------
# DATA SOURCE
# ---------------------------------------------------------
st.sidebar.header("📁 Data Source")
uploaded_file = st.sidebar.file_uploader(
    "Upload your Excel file", type=["xls", "xlsx"]
)

DEFAULT_FILE = "for github stats.xls"

if uploaded_file is None and not os.path.exists(DEFAULT_FILE):
    st.error(
        f"Could not find '{DEFAULT_FILE}'. Upload the Excel file."
    )
    st.stop()

if st.sidebar.button("🔄 Refresh Data / Clear Cache"):
    st.cache_data.clear()
    st.rerun()


def normalise_sheet_name(name):
    return str(name).strip()


@st.cache_data(show_spinner="Loading Excel workbook…")
def load_all_excel_sheets(file_bytes, file_name, file_mtime=None):
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

        df = df.dropna(axis=1, how="all")
        df = df.loc[
            :,
            ~(
                df.columns.astype(str).str.startswith("Unnamed")
                & ~df.notna().any()
            ),
        ]

        new_columns = []

        for column in df.columns:
            if isinstance(column, (pd.Timestamp, datetime, date)):
                new_columns.append(
                    pd.Timestamp(column).strftime("%b-%y").upper()
                )
            else:
                column_text = str(column).strip()
                parsed_date = (
                    pd.to_datetime(column_text, errors="coerce")
                    if column_text else pd.NaT
                )

                if (
                    pd.notna(parsed_date)
                    and any(sep in column_text for sep in ("-", "/", ":"))
                    and len(column_text) >= 7
                ):
                    new_columns.append(
                        parsed_date.strftime("%b-%y").upper()
                    )
                else:
                    new_columns.append(column_text)

        # Make column headings unique.
        unique_columns = []
        seen_columns = {}

        for column in new_columns:
            base = str(column).strip()
            count = seen_columns.get(base, 0)
            seen_columns[base] = count + 1
            unique_columns.append(
                base if count == 0 else f"{base}.{count}"
            )

        df.columns = unique_columns

        text_columns = {
            "Party Code", "Party Name", "AGENT", "BROKER", "POSP",
            "MISP's NAME", "DEALER CODE", "LOB Name", "Dept",
            "Dept ", "Dept  FOR THE MONTH",
            "Dept   UPTO THE MONTH", "Dept Code", "Sl no", "SL NO",
        }

        for column_index, column_name in enumerate(df.columns):
            if column_index == 0 or column_name in text_columns:
                continue

            df.iloc[:, column_index] = pd.to_numeric(
                df.iloc[:, column_index], errors="coerce"
            )

        cleaned_sheets[sheet_name] = df.round(2)

    return cleaned_sheets


# ---------------------------------------------------------
# LOAD EXCEL FILE
# ---------------------------------------------------------
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

    all_sheets = load_all_excel_sheets(
        file_bytes, file_name, file_mtime
    )

    sheet_names = list(all_sheets.keys())

    if not sheet_names:
        st.error("No worksheets were found.")
        st.stop()

    st.sidebar.success(
        f"Successfully loaded {len(sheet_names)} sheets!"
    )

    # -----------------------------------------------------
    # SHEET NAVIGATION
    # -----------------------------------------------------
    st.sidebar.markdown("---")
    st.sidebar.header("📑 View Sheets")

    selected_sheet = st.sidebar.radio(
        "Select the sheet you want to view:",
        sheet_names
    )

    # -----------------------------------------------------
    # HELPER FUNCTIONS
    # -----------------------------------------------------
    def format_indian_number(value):
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
        if isinstance(value, (pd.Timestamp, datetime, date)):
            return pd.Timestamp(value)

        if isinstance(value, str):
            parsed = pd.to_datetime(value.strip(), errors="coerce")
            if pd.notna(parsed):
                return pd.Timestamp(parsed)

        return pd.NaT

    # -----------------------------------------------------
    # KPI DEFAULT VALUES
    # -----------------------------------------------------
    total_premium_value = "N/A"
    premium_label = "Up to latest month Premium"
    accretion_label = "Accretion % Up to latest month"
    accretion_value = "N/A"
    icr_total_value = "N/A"
    icr_earned_value = "N/A"

    extension = os.path.splitext(file_name.lower())[1]
    engine = "xlrd" if extension == ".xls" else "openpyxl"

    # -----------------------------------------------------
    # FIND MONTHLY COMPARISON SHEET
    # -----------------------------------------------------
    monthly_sheet_name = next(
        (
            name for name in sheet_names
            if name.strip().casefold()
            == "revised upto mon comp 2526 2627"
        ),
        None,
    )

    df_monthly_raw = None

    if monthly_sheet_name is not None:

        df_monthly_raw = pd.read_excel(
            BytesIO(file_bytes),
            sheet_name=monthly_sheet_name,
            header=None,
            engine=engine,
        )

        if df_monthly_raw.shape[0] >= 29:

            header_row = df_monthly_raw.iloc[1]
            total_row = df_monthly_raw.iloc[28]

            monthly_candidates = []

            # Find latest month with matching premium
            # and accretion headings.
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

                accretion_heading = str(
                    header_row.iloc[accretion_index]
                ).strip()

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

                latest_month, premium_cell, accretion_cell = max(
                    monthly_candidates,
                    key=lambda item: item[0]
                )

                month_label = latest_month.strftime("%b %y")

                premium_label = f"Up to {month_label} Premium"
                accretion_label = f"Accretion % Up to {month_label}"

                total_premium_value = (
                    f"₹ {format_indian_number(premium_cell)}"
                )

                accretion_value = f"{float(accretion_cell):.2f}%"

    else:
        st.warning(
            "Worksheet 'revised upto mon comp 2526 2627' not found."
        )

    # -----------------------------------------------------
    # ICR FIGURES: F25 AND G25
    # -----------------------------------------------------
    icr_sheet_name = next(
        (
            name for name in sheet_names
            if name.strip().casefold()
            == "dept wise icr premium ep"
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

            total_icr_cell = df_icr_raw.iloc[24, 5]
            earned_icr_cell = df_icr_raw.iloc[24, 6]

            if pd.notna(total_icr_cell):
                icr_total_value = f"{float(total_icr_cell):.2f}%"

            if pd.notna(earned_icr_cell):
                icr_earned_value = f"{float(earned_icr_cell):.2f}%"

    # -----------------------------------------------------
    # FOUR MAIN KPI CARDS
    # -----------------------------------------------------
    k1, k2, k3, k4 = st.columns(4)

    k1.metric(premium_label, total_premium_value)
    k2.metric(accretion_label, accretion_value)
    k3.metric("ICR % on Total Premium", icr_total_value)
    k4.metric("ICR % on Earned Premium", icr_earned_value)

    # -----------------------------------------------------
    # LOB / DEPARTMENT-WISE ACCRETION
    # -----------------------------------------------------
    st.markdown("#### LOB / Department-wise Accretion")

    if monthly_sheet_name is not None and df_monthly_raw is not None:

        header_row = df_monthly_raw.iloc[1]
        data_rows = df_monthly_raw.iloc[2:28].copy()

        # Find the latest Accretion upto month column.
        accretion_month_columns = {}

        month_pattern = re.compile(
            r"accretion\s+upto\s+([a-z]{3}\s+\d{2})",
            flags=re.IGNORECASE,
        )

        for col_index, heading in enumerate(header_row):

            match = month_pattern.search(str(heading).strip())

            if match:
                parsed_month = pd.to_datetime(
                    match.group(1),
                    format="%b %y",
                    errors="coerce"
                )

                if pd.notna(parsed_month):
                    accretion_month_columns[
                        parsed_month.to_period("M")
                    ] = col_index

        if accretion_month_columns:

            latest_period = max(accretion_month_columns)
            latest_accretion_col = accretion_month_columns[latest_period]
            month_text = latest_period.strftime("%b %y")

            # Find matching prior-year and current-year
            # monthly premium columns.
            premium_month_columns = {}

            for col_index, heading in enumerate(header_row):

                parsed_date = parse_month_header(heading)

                if (
                    pd.notna(parsed_date)
                    and parsed_date.month == latest_period.month
                ):
                    if parsed_date.year in (
                        latest_period.year - 1,
                        latest_period.year
                    ):
                        premium_month_columns[parsed_date.year] = col_index

            def find_code_row(code):

                prefix = re.compile(
                    r"^\s*" + re.escape(code) + r"\s*-"
                )

                matches = data_rows[
                    data_rows.iloc[:, 0].astype(str).str.match(
                        prefix, na=False
                    )
                ]

                return matches.iloc[0] if not matches.empty else None

            def motor_group_accretion(codes):

                prior_year = latest_period.year - 1
                current_year = latest_period.year

                if (
                    prior_year not in premium_month_columns
                    or current_year not in premium_month_columns
                ):
                    return None

                prior_col = premium_month_columns[prior_year]
                current_col = premium_month_columns[current_year]

                prior_total = 0.0
                current_total = 0.0

                for code in codes:

                    row = find_code_row(code)

                    if row is not None:

                        prior_total += (
                            pd.to_numeric(
                                row.iloc[prior_col],
                                errors="coerce"
                            )
                            if pd.notna(row.iloc[prior_col])
                            else 0.0
                        )

                        current_total += (
                            pd.to_numeric(
                                row.iloc[current_col],
                                errors="coerce"
                            )
                            if pd.notna(row.iloc[current_col])
                            else 0.0
                        )

                if prior_total == 0:
                    return None

                return (
                    (current_total / prior_total) - 1
                ) * 100

            def display_percent(value, truncate=False):

                if value is None or pd.isna(value):
                    return "—"

                numeric_value = float(value)

                if truncate:
                    numeric_value = (
                        int(numeric_value * 100) / 100
                    )

                return f"{numeric_value:.2f}%"

            compact_rows = []

            excluded_motor_tp_codes = {
                "032", "033", "038", "039"
            }

            # Read normal LOB accretion from the selected
            # Accretion upto month column.
            for _, row in data_rows.iterrows():

                dept_text = str(row.iloc[0]).strip()

                code_match = re.match(
                    r"^\s*(\d{3})\s*-",
                    dept_text
                )

                if not code_match:
                    continue

                code = code_match.group(1)

                # Do not display individual TP codes.
                if code in excluded_motor_tp_codes:
                    continue

                if code == "031":
                    display_name = "Motor Own Damage (031)"
                else:
                    display_name = dept_text

                source_value = row.iloc[latest_accretion_col]

                compact_rows.append(
                    (display_name, display_percent(source_value))
                )

            # Calculate the three motor figures.
            motor_od_row = find_code_row("031")

            motor_od_value = (
                motor_od_row.iloc[latest_accretion_col]
                if motor_od_row is not None
                else None
            )

            motor_tp_value = motor_group_accretion(
                ["032", "033", "038", "039"]
            )

            overall_motor_value = motor_group_accretion(
                ["031", "032", "033", "038", "039"]
            )

            # Replace individual 031 line with three motor summaries.
            compact_rows = [
                item for item in compact_rows
                if item[0] != "Motor Own Damage (031)"
            ]

            insert_at = 0

            for index, (label, _) in enumerate(compact_rows):
                if label.startswith("011 -"):
                    insert_at = index + 1

            motor_items = [
                (
                    "Motor Own Damage (031)",
                    display_percent(motor_od_value)
                ),
                (
                    "Motor TP (032+033+038+039)",
                    display_percent(motor_tp_value, True)
                ),
                (
                    "Overall Motor (OD + TP)",
                    display_percent(overall_motor_value, True)
                ),
            ]

            compact_rows[insert_at:insert_at] = motor_items

            # Display compact table in smaller font.
            table_rows_html = "".join(
                "<tr><td style='padding:4px 10px;"
                "border-bottom:1px solid #ddd;'>"
                + html.escape(label)
                + "</td><td style='padding:4px 10px;"
                "text-align:right;border-bottom:1px solid #ddd;"
                "font-variant-numeric:tabular-nums;'>"
                + html.escape(value)
                + "</td></tr>"
                for label, value in compact_rows
            )

            st.caption(f"Accretion up to {month_text}")

            st.markdown(
                "<div style='font-size:13px;max-width:760px;'>"
                "<table style='width:100%;border-collapse:collapse;'>"
                "<thead><tr>"
                "<th style='text-align:left;padding:5px 10px;'>"
                "LOB / Department</th>"
                "<th style='text-align:right;padding:5px 10px;'>"
                "Accretion %</th>"
                "</tr></thead><tbody>"
                + table_rows_html
                + "</tbody></table></div>",
                unsafe_allow_html=True,
            )

        else:
            st.info(
                "No 'Accretion upto <month>' columns were found."
            )

    else:
        st.info(
            "The monthly comparison worksheet is unavailable."
        )

    st.markdown("---")

    # -----------------------------------------------------
    # SELECTED WORKSHEET AND FILTERS
    # -----------------------------------------------------
    st.header(f"📄 {selected_sheet}")

    df = all_sheets[selected_sheet]
    final_df = df.copy()

    col_format_config = {}

    for column in df.columns:

        upper = str(column).upper()

        if any(
            term in upper
            for term in ["ICR", "ACCRETION", "SHARE"]
        ):
            col_format_config[column] = (
                st.column_config.NumberColumn(format="%.2f")
            )

        elif any(
            term in upper
            for term in [
                "PREMIUM", "TOTAL", "APR-", "MAY-", "JUN-",
                "JUL-", "AUG-", "SEP-", "OCT-", "NOV-",
                "DEC-", "JAN-", "FEB-", "MAR-",
            ]
        ):
            col_format_config[column] = (
                st.column_config.NumberColumn(format="₹ %,.2f")
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

            departments = (
                df[department_column].dropna().unique().tolist()
            )

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

        final_df = final_df[
            [department_column] + selected_columns
        ]

    elif selected_sheet in channel_sheets:

        st.subheader("Filters")

        col1, col2, col3 = st.columns(3)

        agent_search = broker_search = posp_search = ""

        agent_column = next(
            (
                c for c in ["AGENT", "Party Name"]
                if c in df.columns
            ),
            None,
        )

        with col1:
            if agent_column:
                agent_search = st.text_input(
                    f"Search by {agent_column}:"
                )

        with col2:
            if "BROKER" in df.columns:
                broker_search = st.text_input(
                    "Search by BROKER:"
                )

        with col3:
            if "POSP" in df.columns:
                posp_search = st.text_input(
                    "Search by POSP:"
                )

        if agent_search and agent_column:

            final_df = final_df[
                final_df[agent_column].astype(str).str.contains(
                    agent_search,
                    case=False,
                    na=False,
                    regex=False
                )
            ]

        if broker_search and "BROKER" in final_df.columns:

            final_df = final_df[
                final_df["BROKER"].astype(str).str.contains(
                    broker_search,
                    case=False,
                    na=False,
                    regex=False
                )
            ]

        if posp_search and "POSP" in final_df.columns:

            final_df = final_df[
                final_df["POSP"].astype(str).str.contains(
                    posp_search,
                    case=False,
                    na=False,
                    regex=False
                )
            ]

    elif selected_sheet == "MISP":

        st.subheader("Filters")

        misp_column = next(
            (
                c for c in ["MISP's NAME", "MISP"]
                if c in df.columns
            ),
            None,
        )

        if misp_column:

            misp_search = st.text_input(
                f"Search by {misp_column}:"
            )

            if misp_search:

                final_df = final_df[
                    final_df[misp_column].astype(str).str.contains(
                        misp_search,
                        case=False,
                        na=False,
                        regex=False
                    )
                ]

    # Keep first column pinned.
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

    csv_data = final_df.to_csv(
        index=False
    ).encode("utf-8-sig")

    st.download_button(
        label=f"📥 Download '{selected_sheet}' Data as CSV",
        data=csv_data,
        file_name=f"{selected_sheet.replace(' ', '_')}_filtered.csv",
        mime="text/csv",
    )

except Exception as error:

    st.error(
        f"An error occurred while loading the dashboard: {error}"
    )

    st.caption(
        "If this persists, check the Streamlit app logs "
        "for the exact line number."
    )

    st.info(
        "For .xls files, ensure xlrd is listed in requirements.txt. "
        "For .xlsx files, openpyxl is required."
    )
