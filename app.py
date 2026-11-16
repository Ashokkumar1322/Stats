
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


    # ---------------------------------------------------------
    # CHANNEL-WISE KPI FIGURES
    # The worksheet is read by section headings, not fixed row
    # numbers, so adding rows above a section will not break it.
    # ---------------------------------------------------------
    channel_sheet_name = next(
        (name for name in sheet_names
         if name.strip().casefold() == "channel wise 26 27"),
        None,
    )

    channel_growth_rows = []
    channel_share_rows = []
    channel_month_text = "latest month"

    if channel_sheet_name is not None:
        channel_raw = pd.read_excel(
            BytesIO(file_bytes),
            sheet_name=channel_sheet_name,
            header=None,
            engine=engine,
        )

        def find_channel_heading(text, exclude=None):
            for row_index in range(len(channel_raw)):
                value = channel_raw.iloc[row_index, 1] if channel_raw.shape[1] > 1 else None
                if isinstance(value, str) and text.casefold() in value.casefold():
                    if exclude and exclude.casefold() in value.casefold():
                        continue
                    return row_index
            return None

        def next_nonempty_row(start_row, limit=8):
            for row_index in range(start_row, min(start_row + limit, len(channel_raw))):
                if channel_raw.iloc[row_index].notna().any():
                    return row_index
            return None

        def month_label(value):
            parsed = parse_month_header(value)
            return parsed.strftime("%b %y") if pd.notna(parsed) else None

        def get_table_after_heading(heading_text, exclude=None):
            heading_row = find_channel_heading(heading_text, exclude=exclude)
            if heading_row is None:
                return None, None, None
            header_row = next_nonempty_row(heading_row + 1, limit=8)
            if header_row is None:
                return heading_row, None, None
            return heading_row, header_row, channel_raw.iloc[header_row]

        latest_col = None

        # Cumulative premium table: locate the TOTAL row and latest
        # month dynamically (currently Oct-26; next month will move on).
        _, cumulative_header_row, cumulative_header = get_table_after_heading(
            "CHANNEL WISE UPTO EACH MONTH", exclude="GROWTH"
        )
        if cumulative_header_row is not None:
            cumulative_total_row = None
            for row_index in range(cumulative_header_row + 1, len(channel_raw)):
                label = str(channel_raw.iloc[row_index, 1]).strip().upper()
                if label == "TOTAL":
                    cumulative_total_row = channel_raw.iloc[row_index]
                    break
                if row_index > cumulative_header_row + 1 and not channel_raw.iloc[row_index].notna().any():
                    break

            if cumulative_total_row is not None:
                month_candidates = []
                for col_index in range(2, len(cumulative_header)):
                    parsed = parse_month_header(cumulative_header.iloc[col_index])
                    value = pd.to_numeric(cumulative_total_row.iloc[col_index], errors="coerce")
                    if pd.notna(parsed) and pd.notna(value):
                        month_candidates.append((parsed, col_index, value))

                if month_candidates:
                    latest_month, latest_col, premium_cell = max(
                        month_candidates, key=lambda item: item[0]
                    )
                    channel_month_text = latest_month.strftime("%b %y")
                    premium_label = f"Up to {channel_month_text} Premium"
                    total_premium_value = f"₹ {format_indian_number(premium_cell)}"
        # Update this cell in Excel each month; the dashboard label follows
        # the latest month with an actual cumulative premium figure.
        accretion_cell = (
            channel_raw.iloc[285, 6]
            if channel_raw.shape[0] > 285 and channel_raw.shape[1] > 6
            else None
        )
        if pd.notna(accretion_cell):
            accretion_value = f"{float(pd.to_numeric(accretion_cell, errors='coerce')):.2f}%"
            accretion_label = f"Accretion % Up to {channel_month_text}"

        # Channel-wise growth: use the same latest populated month column
        # as the cumulative premium table (I for Oct-26, J for Nov-26, etc.).
        # The six channel rows are AGENCY through CORP AGENT SF; exclude TOTAL.
        _, growth_header_row, growth_header = get_table_after_heading(
            "CHANNEL WISE GROWTH UPTO EACH MONTH"
        )
        if growth_header_row is not None and latest_col is not None:
            channel_order = [
                "AGENCY", "BANCA", "BROKER", "DIRECT", "POSP", "CORP AGENT SF"
            ]
            growth_row_lookup = {}
            for row_index in range(growth_header_row + 1, len(channel_raw)):
                raw_name = channel_raw.iloc[row_index, 1]
                if pd.isna(raw_name):
                    continue
                channel_name = str(raw_name).strip().upper()
                if channel_name == "TOTAL":
                    break
                growth_row_lookup[channel_name] = row_index

            for channel_name in channel_order:
                row_index = growth_row_lookup.get(channel_name)
                if row_index is None or latest_col >= channel_raw.shape[1]:
                    growth_value = None
                else:
                    growth_value = pd.to_numeric(
                        channel_raw.iloc[row_index, latest_col], errors="coerce"
                    )
                channel_growth_rows.append((channel_name, growth_value))


        # Channel-wise share of premium: values are in Excel D280:D285.
        # Find the summary section and read the six channel rows below
        # its column headings; the current month label is read from Excel.
        _, share_header_row, share_header = get_table_after_heading(
            "CHANNEL WISE TOTAL PREMIUM", exclude="REVISED"
        )
        if share_header_row is not None:
            share_month_label = str(channel_raw.iloc[share_header_row, 1]).strip()
            month_match = re.search(r"UPTO\s+([A-Z]{3}\s*\d{2})", share_month_label, re.I)
            share_period = month_match.group(1).title() if month_match else channel_month_text
            share_label = f"Channel-wise share of premium up to {share_period}"

            for row_index in range(share_header_row + 1, min(share_header_row + 8, len(channel_raw))):
                channel_name = channel_raw.iloc[row_index, 1]
                if pd.isna(channel_name):
                    break
                channel_name = str(channel_name).strip()
                if channel_name.upper() == "TOTAL":
                    break
                share_value = pd.to_numeric(channel_raw.iloc[row_index, 3], errors="coerce")
                channel_share_rows.append((channel_name, share_value))
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
    # MAIN KPI CARDS
    # -----------------------------------------------------
    k1, k2, k3, k4 = st.columns(4)
    k1.metric(premium_label, total_premium_value)
    k2.metric(accretion_label, accretion_value)
    k3.metric("ICR % on Total Premium", icr_total_value)
    k4.metric("ICR % on Earned Premium", icr_earned_value)

    # -----------------------------------------------------
    # CHANNEL-WISE SHARE OF PREMIUM
    # -----------------------------------------------------
    if channel_share_rows:
        st.markdown(f"#### {share_label}")
        share_columns = st.columns(6)
        for share_column, (channel_name, share_value) in zip(
            share_columns, channel_share_rows
        ):
            display_value = (
                f"{float(share_value):.2f}%"
                if pd.notna(share_value) else "—"
            )
            share_column.metric(channel_name, display_value)

    # -----------------------------------------------------
    # CHANNEL-WISE GROWTH
    # -----------------------------------------------------
    if channel_growth_rows:
        st.markdown(f"#### Channel-wise Growth up to {channel_month_text}")
        growth_columns = st.columns(6)
        for growth_column, (channel_name, growth_value) in zip(
            growth_columns, channel_growth_rows
        ):
            display_value = (
                f"{float(growth_value):.2f}%"
                if pd.notna(growth_value) else "—"
            )
            growth_column.metric(channel_name, display_value)

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

            # -------------------------------------------------
            # SHARE OF PREMIUM: use the same latest cumulative
            # month as the accretion figure and the total row
            # from this worksheet as the denominator.
            # -------------------------------------------------
            current_premium_col = premium_month_columns.get(
                latest_period.year
            )

            total_lob_premium = None
            if current_premium_col is not None:
                total_row_mask = (
                    df_monthly_raw.iloc[:, 0]
                    .astype(str)
                    .str.strip()
                    .str.casefold()
                    .eq("sum for all departments")
                )
                total_rows = df_monthly_raw.loc[total_row_mask]

                if not total_rows.empty:
                    total_lob_premium = pd.to_numeric(
                        total_rows.iloc[0, current_premium_col],
                        errors="coerce",
                    )

            def premium_for_code(code):
                row = find_code_row(code)
                if row is None or current_premium_col is None:
                    return 0.0

                value = pd.to_numeric(
                    row.iloc[current_premium_col],
                    errors="coerce",
                )
                return float(value) if pd.notna(value) else 0.0

            def premium_share(premium_amount):
                if (
                    total_lob_premium is None
                    or pd.isna(total_lob_premium)
                    or total_lob_premium == 0
                ):
                    return None
                return (float(premium_amount) / float(total_lob_premium)) * 100

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
                premium_amount = (
                    pd.to_numeric(
                        row.iloc[current_premium_col],
                        errors="coerce",
                    )
                    if current_premium_col is not None
                    else None
                )

                compact_rows.append(
                    (
                        display_name,
                        display_percent(source_value),
                        display_percent(
                            premium_share(premium_amount)
                            if premium_amount is not None and pd.notna(premium_amount)
                            else None
                        ),
                    )
                )

            # Calculate the three motor figures.
            motor_od_row = find_code_row("031")

            motor_od_value = (
                motor_od_row.iloc[latest_accretion_col]
                if motor_od_row is not None
                else None
            )

            motor_od_premium = premium_for_code("031")
            motor_tp_premium = sum(
                premium_for_code(code)
                for code in ["032", "033", "038", "039"]
            )
            overall_motor_premium = motor_od_premium + motor_tp_premium

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

            for index, (label, _, _) in enumerate(compact_rows):
                if label.startswith("011 -"):
                    insert_at = index + 1

            motor_items = [
                (
                    "Motor Own Damage (031)",
                    display_percent(motor_od_value),
                    display_percent(premium_share(motor_od_premium)),
                ),
                (
                    "Motor TP (032+033+038+039)",
                    display_percent(motor_tp_value, True),
                    display_percent(premium_share(motor_tp_premium)),
                ),
                (
                    "Overall Motor (OD + TP)",
                    display_percent(overall_motor_value, True),
                    display_percent(premium_share(overall_motor_premium)),
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
                + html.escape(accretion_value)
                + "</td><td style='padding:4px 10px;text-align:right;"
                "border-bottom:1px solid #ddd;"
                "font-variant-numeric:tabular-nums;'>"
                + html.escape(share_value)
                + "</td></tr>"
                for label, accretion_value, share_value in compact_rows
            )

            st.caption(
                f"Accretion and Share of Premium up to {month_text}"
            )

            st.markdown(
                "<div style='font-size:13px;max-width:950px;'>"
                "<table style='width:100%;border-collapse:collapse;'>"
                "<thead><tr>"
                "<th style='text-align:left;padding:5px 10px;'>"
                "LOB / Department</th>"
                "<th style='text-align:right;padding:5px 10px;'>"
                "Accretion %</th>"
                "<th style='text-align:right;padding:5px 10px;'>"
                "Share of Premium %</th>"
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

        st.caption(
            "Each Excel section is displayed separately using its heading. "
            "Blank rows separate sections; month headings are shown as Apr-26, May-26, etc."
        )

        raw_channel_sheet = pd.read_excel(
            BytesIO(file_bytes),
            sheet_name=selected_sheet,
            header=None,
            engine=engine,
        )

        def is_section_heading(row_index):
            if raw_channel_sheet.shape[1] <= 1:
                return False
            value = raw_channel_sheet.iloc[row_index, 1]
            return (
                isinstance(value, str)
                and value.strip().endswith(":-")
            )

        heading_rows = [
            row_index for row_index in range(len(raw_channel_sheet))
            if is_section_heading(row_index)
        ]

        def display_header(value, index):
            if isinstance(value, (pd.Timestamp, datetime, date)):
                return pd.Timestamp(value).strftime("%b-%y")
            if isinstance(value, str):
                parsed = parse_month_header(value.strip())
                if pd.notna(parsed):
                    return parsed.strftime("%b-%y")
                clean = value.strip()
                return clean if clean else f"Column {index + 1}"
            if pd.isna(value):
                return f"Column {index + 1}"
            return str(value).strip()

        def make_unique_headers(headers):
            counts = {}
            result = []
            for header in headers:
                base = str(header).strip() or "Column"
                counts[base] = counts.get(base, 0) + 1
                result.append(base if counts[base] == 1 else f"{base}.{counts[base]}")
            return result

        for section_number, heading_row in enumerate(heading_rows):
            raw_heading = str(raw_channel_sheet.iloc[heading_row, 1]).strip()
            section_title = raw_heading[:-2].strip() if raw_heading.endswith(":-") else raw_heading

            next_heading = (
                heading_rows[section_number + 1]
                if section_number + 1 < len(heading_rows)
                else len(raw_channel_sheet)
            )

            # Find the actual table header. Some sections have a short
            # note row (e.g. revised premium) before their column headings.
            header_row = None
            best_score = -1
            for candidate in range(heading_row + 1, min(heading_row + 7, next_heading)):
                values = raw_channel_sheet.iloc[candidate].tolist()
                nonempty = [v for v in values if pd.notna(v)]
                if not nonempty:
                    continue

                text_values = " ".join(str(v).casefold() for v in nonempty if isinstance(v, str))
                has_date = any(
                    isinstance(v, (pd.Timestamp, datetime, date))
                    or pd.notna(parse_month_header(v))
                    for v in nonempty
                )
                score = 0
                for token in ("party code", "party name", "premium", "share", "icr", "accretion"):
                    if token in text_values:
                        score += 2
                if has_date:
                    score += 3
                if score > best_score:
                    best_score = score
                    header_row = candidate

            if header_row is None or best_score <= 0:
                continue

            # End this table at the first fully blank row or next section heading.
            data_end = next_heading
            for candidate in range(header_row + 1, next_heading):
                if not raw_channel_sheet.iloc[candidate].notna().any():
                    data_end = candidate
                    break

            # The worksheet tables begin in column B. Keep all populated
            # columns, including month columns through March.
            first_col = 1
            last_col = raw_channel_sheet.shape[1]
            while last_col > first_col:
                if raw_channel_sheet.iloc[header_row:data_end, last_col - 1].notna().any():
                    break
                last_col -= 1

            header_values = raw_channel_sheet.iloc[header_row, first_col:last_col].tolist()
            headers = make_unique_headers([
                display_header(value, index)
                for index, value in enumerate(header_values)
            ])

            table_data = raw_channel_sheet.iloc[
                header_row + 1:data_end, first_col:last_col
            ].copy()
            table_data.columns = headers
            table_data = table_data.dropna(how="all")

            if table_data.empty:
                continue

            # Keep text identifiers as text; numeric measures remain numeric.
            for column in table_data.columns[1:]:
                if any(
                    token in str(column).upper()
                    for token in ("ICR", "ACCRETION", "GROWTH", "SHARE", "%")
                ):
                    continue
                if re.match(r"^[A-Z][a-z]{2}-\d{2}$", str(column)):
                    continue

            st.markdown(f"#### {section_title}")

            # Large detail sections get a simple search box.
            search_text = st.text_input(
                f"Search {section_title}",
                key=f"channel_search_{selected_sheet}_{section_number}",
            )
            if search_text:
                mask = table_data.astype(str).apply(
                    lambda column: column.str.contains(
                        search_text, case=False, na=False, regex=False
                    )
                ).any(axis=1)
                table_data = table_data.loc[mask]

            column_config = {}
            is_growth_section = "GROWTH" in section_title.upper()
            first_table_column = table_data.columns[0]

            for column in table_data.columns:
                upper = str(column).upper()

                # Month headings in the channel-wise growth section are dates,
                # not labels containing the word GROWTH. Format every numeric
                # month column in this section as a percentage with 2 decimals.
                if is_growth_section and column != first_table_column:
                    column_config[column] = st.column_config.NumberColumn(
                        format="%.2f%%"
                    )
                elif any(token in upper for token in ("ICR", "ACCRETION", "GROWTH", "SHARE", "%")):
                    column_config[column] = st.column_config.NumberColumn(
                        format="%.2f%%"
                    )
                elif (
                    re.match(r"^[A-Z][a-z]{2}-\d{2}$", str(column))
                    or "PREMIUM" in upper
                    or "TOTAL" in upper
                ):
                    column_config[column] = st.column_config.NumberColumn(
                        format="₹ %,.0f"
                    )
                else:
                    column_config[column] = st.column_config.TextColumn()

            st.dataframe(
                table_data,
                use_container_width=True,
                hide_index=True,
                column_config=column_config,
            )

            section_csv = table_data.to_csv(index=False).encode("utf-8-sig")
            st.download_button(
                label=f"📥 Download {section_title} as CSV",
                data=section_csv,
                file_name=(
                    f"{selected_sheet.replace(' ', '_')}_"
                    f"{section_number + 1}.csv"
                ),
                mime="text/csv",
                key=f"download_{selected_sheet}_{section_number}",
            )

            st.markdown("---")

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

    # Display the ordinary single-table worksheets.
    # Channel worksheets were already rendered section-by-section above.
    if selected_sheet not in channel_sheets:
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
