
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
                    pd.Timestamp(column).strftime("%b-%y")
                )
            else:
                column_text = str(column).strip()

                # Excel month headings may arrive as strings such as
                # Apr-25 (6 characters), Apr-2025, or as actual dates.
                parsed_date = pd.NaT
                if column_text:
                    for month_format in (
                        "%b-%y", "%b-%Y", "%b %y", "%b %Y",
                        "%B-%y", "%B-%Y", "%B %y", "%B %Y",
                    ):
                        parsed_date = pd.to_datetime(
                            column_text, format=month_format, errors="coerce"
                        )
                        if pd.notna(parsed_date):
                            break

                    if pd.isna(parsed_date):
                        parsed_date = pd.to_datetime(
                            column_text, errors="coerce"
                        )

                if pd.notna(parsed_date):
                    new_columns.append(
                        parsed_date.strftime("%b-%y")
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

            # Convert safely even when Excel has formatted an otherwise
            # blank column as a date. NaT values must become NaN before
            # numeric conversion; otherwise pandas can turn them into the
            # integer sentinel -9223372036854775808.
            source_values = df[column_name]
            source_values = source_values.astype("object")
            source_values = source_values.where(
                pd.notna(source_values), None
            )

            numeric_values = pd.to_numeric(
                source_values, errors="coerce"
            )

            # Assign by column name so pandas replaces the original
            # datetime dtype instead of trying to write integers into it.
            df[column_name] = numeric_values

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
            month_text = value.strip()

            for month_format in (
                "%b-%y", "%b-%Y", "%b %y", "%b %Y",
                "%B-%y", "%B-%Y", "%B %y", "%B %Y",
            ):
                parsed = pd.to_datetime(
                    month_text, format=month_format, errors="coerce"
                )
                if pd.notna(parsed):
                    return pd.Timestamp(parsed)

            parsed = pd.to_datetime(month_text, errors="coerce")
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
    # FIND MONTHLY COMPARISON SHEET (used below for LOB accretion)
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

    # -----------------------------------------------------
    # TOP PREMIUM / ACCRETION + CHANNEL-WISE GROWTH
    # Source: Channel wise 26 27
    # Excel row 299 = total premium; row 308 = total accretion
    # Excel rows 302:307 = channel-wise growth
    # Excel columns I:N = Oct 2026 to Mar 2027
    # -----------------------------------------------------
    channel_sheet_name = next(
        (
            name for name in sheet_names
            if name.strip().casefold() == "channel wise 26 27"
        ),
        None,
    )

    channel_growth_rows = []
    channel_month_text = "latest month"
    total_premium_value = "N/A"
    premium_label = "Up to latest month Premium"
    accretion_label = "Accretion % Up to latest month"
    accretion_value = "N/A"

    if channel_sheet_name is not None:
        channel_raw = pd.read_excel(
            BytesIO(file_bytes),
            sheet_name=channel_sheet_name,
            header=None,
            engine=engine,
        )

        # Excel row 292 (zero-based 291) contains month dates.
        # Excel row 299 (zero-based 298) contains cumulative premium.
        # Excel row 308 (zero-based 307) contains total accretion.
        if channel_raw.shape[0] >= 308 and channel_raw.shape[1] >= 14:
            month_header = channel_raw.iloc[291]
            premium_total_row = channel_raw.iloc[298]
            accretion_total_row = channel_raw.iloc[307]

            available_months = []
            for col_index in range(2, min(14, len(month_header))):
                month_date = parse_month_header(month_header.iloc[col_index])
                premium_cell = pd.to_numeric(
                    premium_total_row.iloc[col_index], errors="coerce"
                )
                if pd.notna(month_date) and pd.notna(premium_cell):
                    available_months.append(
                        (month_date, col_index, premium_cell)
                    )

            if available_months:
                latest_month, latest_col, premium_cell = max(
                    available_months, key=lambda item: item[0]
                )
                channel_month_text = latest_month.strftime("%b %y")
                premium_label = f"Up to {channel_month_text} Premium"
                accretion_label = f"Accretion % Up to {channel_month_text}"
                total_premium_value = (
                    f"₹ {format_indian_number(premium_cell)}"
                )

                accretion_cell = pd.to_numeric(
                    accretion_total_row.iloc[latest_col], errors="coerce"
                )
                if pd.notna(accretion_cell):
                    accretion_value = f"{float(accretion_cell):.2f}%"

                # Excel rows 302 to 307: AGENCY through CORP AGENT SF.
                for excel_row in range(302, 308):
                    row_index = excel_row - 1
                    channel_name = str(
                        channel_raw.iloc[row_index, 1]
                    ).strip()
                    growth_value = pd.to_numeric(
                        channel_raw.iloc[row_index, latest_col],
                        errors="coerce",
                    )
                    channel_growth_rows.append(
                        (channel_name, growth_value)
                    )
            else:
                st.warning(
                    "No populated month was found in 'Channel wise 26 27'."
                )
    else:
        st.warning("Worksheet 'Channel wise 26 27' was not found.")

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
    # CHANNEL-WISE GROWTH
    # -----------------------------------------------------
    st.markdown(f"#### Channel-wise Growth up to {channel_month_text}")
    if channel_growth_rows:
        growth_columns = st.columns(6)
        for growth_column, (channel_name, growth_value) in zip(
            growth_columns, channel_growth_rows
        ):
            display_value = (
                f"{float(growth_value):.2f}%"
                if pd.notna(growth_value)
                else "—"
            )
            growth_column.metric(channel_name, display_value)
    else:
        st.info("Channel-wise growth figures are not available.")

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

        # -------------------------------------------------
        # MONTH-WISE CHANNEL SUMMARY (SECOND TABLE)
        # -------------------------------------------------
        # The monthly summary is appended below the main
        # agent-wise table in the Excel worksheet. It has
        # its own header row, so it must be read separately
        # rather than treated as part of the first table.
        st.subheader("Month-wise Channel Summary")

        raw_channel_sheet = pd.read_excel(
            BytesIO(file_bytes),
            sheet_name=selected_sheet,
            header=None,
            engine=engine,
        )

        summary_header_row = None

        for row_index in range(len(raw_channel_sheet)):
            row_values = raw_channel_sheet.iloc[row_index].tolist()
            row_text = " ".join(
                str(value).strip()
                for value in row_values
                if pd.notna(value)
            ).casefold()

            if "channel wise" in row_text and (
                "25-26" in row_text or "26-27" in row_text
            ):
                summary_header_row = row_index
                break

        if summary_header_row is not None:
            header_values = raw_channel_sheet.iloc[
                summary_header_row
            ].tolist()

            month_columns = []
            for value in header_values[2:14]:
                if isinstance(value, (pd.Timestamp, datetime, date)):
                    month_columns.append(
                        pd.Timestamp(value).strftime("%b-%y")
                    )
                elif pd.notna(value):
                    month_text = str(value).strip()
                    parsed_month = pd.NaT

                    for month_format in (
                        "%b-%y", "%b-%Y", "%b %y", "%b %Y",
                        "%B-%y", "%B-%Y", "%B %y", "%B %Y",
                    ):
                        parsed_month = pd.to_datetime(
                            month_text,
                            format=month_format,
                            errors="coerce",
                        )
                        if pd.notna(parsed_month):
                            break

                    month_columns.append(
                        pd.Timestamp(parsed_month).strftime("%b-%y")
                        if pd.notna(parsed_month)
                        else month_text
                    )
                else:
                    month_columns.append("")

            summary_rows = []
            for row_index in range(
                summary_header_row + 1,
                min(summary_header_row + 10, len(raw_channel_sheet)),
            ):
                row_values = raw_channel_sheet.iloc[row_index].tolist()
                channel_name = (
                    row_values[1]
                    if len(row_values) > 1
                    else None
                )

                if pd.isna(channel_name) or not str(channel_name).strip():
                    break

                channel_name = str(channel_name).strip()
                row_data = {"Channel": channel_name}

                for month_index, month_name in enumerate(month_columns):
                    source_index = month_index + 2
                    source_value = (
                        row_values[source_index]
                        if source_index < len(row_values)
                        else None
                    )
                    row_data[month_name or f"Month {month_index + 1}"] = (
                        pd.to_numeric(source_value, errors="coerce")
                    )

                summary_rows.append(row_data)

                if channel_name.upper() == "TOTAL":
                    break

            if summary_rows:
                summary_df = pd.DataFrame(summary_rows)
                summary_config = {
                    column: st.column_config.NumberColumn(
                        format="₹ %,.0f"
                    )
                    for column in summary_df.columns
                    if column != "Channel"
                }
                summary_config["Channel"] = st.column_config.TextColumn(
                    pinned="left"
                )

                st.dataframe(
                    summary_df,
                    use_container_width=True,
                    hide_index=True,
                    column_config=summary_config,
                )

                summary_csv = summary_df.to_csv(
                    index=False
                ).encode("utf-8-sig")

                st.download_button(
                    label="📥 Download Month-wise Channel Summary as CSV",
                    data=summary_csv,
                    file_name=(
                        f"{selected_sheet.replace(' ', '_')}_"
                        "monthly_summary.csv"
                    ),
                    mime="text/csv",
                    key=f"{selected_sheet}_monthly_summary_download",
                )
            else:
                st.warning(
                    "The monthly channel summary heading was found, "
                    "but no channel figures were found below it."
                )
        else:
            st.warning(
                "Could not find the separate month-wise channel summary "
                "table in this worksheet."
            )

        st.markdown("---")
        st.subheader("Agent-wise Details and Filters")

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
