import os
import re
from io import BytesIO

import pandas as pd


# =========================================================
# PREMISES / X-Y-Z CONFIGURATION
# =========================================================

def _premises_metric_from_heading(value):
    """
    Convert a Premises sheet heading into:
        (period, category, metric)

    Examples:
        Current Talent FTE Cost -> ("current", "talent", "cost")
        Future AI FTE Value     -> ("future", "ai", "value")
    """
    if pd.isna(value):
        return None

    text = normalize_column_name(value)
    if "premises" in text:
        return None

    period = None
    if "current" in text:
        period = "current"
    elif "future" in text:
        period = "future"

    category = None
    if "talent" in text:
        category = "talent"
    elif re.search(r"\bai\b", text):
        category = "ai"

    metric = None
    if "cost" in text:
        metric = "cost"
    elif "value" in text:
        metric = "value"

    if period and category and metric:
        return period, category, metric

    return None


def read_premises_values(input_file):
    """
    Read X/Y/Z cost and value assumptions from the input workbook's
    'Premises' sheet.

    The function does NOT assume fixed cell addresses. It searches the
    Premises sheet for the four dimensions:
        Current/Future
        Talent/AI
        Cost/Value

    For each matching heading it finds the X, Y and Z rows below that
    heading and takes the first numeric value to the right of each
    X/Y/Z label.

    This allows different input Excel files to contain different
    X/Y/Z numbers and still use the same calculation code.
    """
    with pd.ExcelFile(input_file) as xls:
        premises_sheet = next(
            (
                s for s in xls.sheet_names
                if normalize_column_name(s).startswith("premises")
            ),
            None,
        )

        if premises_sheet is None:
            raise ValueError(
                "Could not find the 'Premises' sheet in the uploaded Excel file."
            )

        raw = pd.read_excel(
            xls,
            sheet_name=premises_sheet,
            header=None,
        )

    premises = {
        "current": {
            "talent": {"cost": {}, "value": {}},
            "ai": {"cost": {}, "value": {}},
        },
        "future": {
            "talent": {"cost": {}, "value": {}},
            "ai": {"cost": {}, "value": {}},
        },
    }

    # Search every cell for a Current/Future + Talent/AI + Cost/Value heading.
    for row_idx in range(len(raw)):
        for col_idx in range(raw.shape[1]):
            heading = raw.iat[row_idx, col_idx]

            parsed = _premises_metric_from_heading(heading)
            if parsed is None:
                continue

            period, category, metric = parsed
            found = {}

            # The screenshot shows X/Y/Z on rows below the heading.
            # Do not depend on exact row numbers or exact columns.
            for search_row in range(
                row_idx + 1,
                min(row_idx + 10, len(raw)),
            ):
                row_label = raw.iat[search_row, col_idx]

                if pd.isna(row_label):
                    continue

                xyz = normalize_xyz(row_label)
                if xyz not in {"X", "Y", "Z"} or xyz in found:
                    continue

                # Value is normally immediately to the right, but allow
                # blank spacer columns (as in the Future Premises block).
                for value_col in range(
                    col_idx + 1,
                    min(col_idx + 5, raw.shape[1]),
                ):
                    candidate = raw.iat[search_row, value_col]
                    if pd.isna(candidate):
                        continue

                    number = pd.to_numeric(candidate, errors="coerce")
                    if not pd.isna(number):
                        found[xyz] = float(number)
                        break

                if len(found) == 3:
                    break

            if set(found) != {"X", "Y", "Z"}:
                missing = sorted({"X", "Y", "Z"} - set(found))
                raise ValueError(
                    f"Could not extract X/Y/Z values for "
                    f"{period.title()} {category.title()} {metric}. "
                    f"Missing: {', '.join(missing)}."
                )

            premises[period][category][metric] = found

    # Validate all eight maps were found.
    missing_sections = []
    for period in ("current", "future"):
        for category in ("talent", "ai"):
            for metric in ("cost", "value"):
                values = premises[period][category][metric]
                if set(values) != {"X", "Y", "Z"}:
                    missing_sections.append(
                        f"{period.title()} {category.title()} {metric.title()}"
                    )

    if missing_sections:
        raise ValueError(
            "Could not extract all required Premises values. Missing sections: "
            + "; ".join(missing_sections)
        )

    return premises


# =========================================================
# COLUMN HELPERS
# =========================================================

def normalize_column_name(value):
    value = str(value)
    value = value.replace("\n", " ").replace("\r", " ")
    value = re.sub(r"\s+", " ", value)
    return value.strip().lower()


def _find_column(columns, contains_all, exclude=None):
    exclude = exclude or []
    for column in columns:
        normalized = normalize_column_name(column)
        if all(part in normalized for part in contains_all):
            if not any(part in normalized for part in exclude):
                return column
    return None


def find_input_columns(df):
    """
    Expected service-table structure:

    Service Improvement
    Team Member
    Number of same services
    % FTE/Service
    Talent value (X, Y or Z)
    AI companion
    Number of same services
    % FTE/Service
    Talent value AI (X, Y or Z)

    Duplicate Excel headers are handled by looking at their
    position in the table.
    """
    columns = list(df.columns)

    service_improvement = _find_column(columns, ["service", "improvement"])
    team_member = _find_column(columns, ["team", "member"])
    ai_companion = _find_column(columns, ["ai", "companion"])

    if service_improvement is None:
        # Fallback to the first textual column if required.
        service_improvement = columns[0]

    if team_member is None:
        # The supplied file uses "Team Member".
        matches = [
            c for c in columns
            if "team" in normalize_column_name(c)
            and "member" in normalize_column_name(c)
        ]
        if matches:
            team_member = matches[0]

    if ai_companion is None:
        matches = [
            c for c in columns
            if "ai" in normalize_column_name(c)
            and "companion" in normalize_column_name(c)
        ]
        if matches:
            ai_companion = matches[0]

    # Pandas normally changes duplicate Excel headers to .1, .2, etc.
    service_count_columns = [
        c for c in columns
        if "number" in normalize_column_name(c)
        and "service" in normalize_column_name(c)
    ]

    fte_columns = [
        c for c in columns
        if "fte" in normalize_column_name(c)
        and "service" in normalize_column_name(c)
    ]

    value_columns = [
        c for c in columns
        if "value" in normalize_column_name(c)
    ]

    if len(service_count_columns) < 2:
        raise ValueError(
            "The service table must contain two 'Number of same services' columns."
        )

    if len(fte_columns) < 2:
        raise ValueError(
            "The service table must contain two '% FTE/Service' columns."
        )

    # First occurrence = Talent, second = AI.
    talent_service_count = service_count_columns[0]
    ai_service_count = service_count_columns[1]

    talent_fte = fte_columns[0]
    ai_fte = fte_columns[1]

    # In the supplied Excel the first value column is Talent value
    # and the second value column is Talent value AI.
    talent_value = None
    ai_value = None

    for column in value_columns:
        n = normalize_column_name(column)
        if "ai" in n:
            ai_value = column
        elif talent_value is None:
            talent_value = column

    # If duplicate/odd Excel headers make the above ambiguous,
    # use positional fallback based on the table.
    if talent_value is None and len(value_columns) >= 1:
        talent_value = value_columns[0]

    if ai_value is None and len(value_columns) >= 2:
        ai_value = value_columns[1]

    if talent_value is None or ai_value is None:
        raise ValueError(
            "The service table must contain Talent value and Talent value AI columns."
        )

    return {
        "service_improvement": service_improvement,
        "team_member": team_member,
        "talent_services": talent_service_count,
        "talent_fte": talent_fte,
        "talent_value": talent_value,
        "ai_companion": ai_companion,
        "ai_services": ai_service_count,
        "ai_fte": ai_fte,
        "ai_value": ai_value,
    }


# =========================================================
# NUMBER / VALUE HELPERS
# =========================================================

def numeric(value, default=0.0):
    if pd.isna(value):
        return default

    if isinstance(value, str):
        value = value.strip().replace(",", "")
        if value == "":
            return default

    result = pd.to_numeric(value, errors="coerce")
    return default if pd.isna(result) else float(result)


def normalize_xyz(value):
    if pd.isna(value):
        return ""

    value = str(value).strip().upper()

    # Handle values such as "X ", "Y", "Z".
    # Excel blanks can sometimes arrive as the literal string "nan"
    # after dataframe/string conversion. Treat those as empty values.
    if value in {"", "NAN", "NONE", "NULL"}:
        return ""

    if value in {"X", "Y", "Z"}:
        return value

    return value


# =========================================================
# CALCULATION ENGINE
# =========================================================

def calculate_side(
    number_of_services,
    fte_per_service,
    xyz_value,
    cost_map,
    value_map,
    service_value_base="cost",
):
    """
    Required formulas:

    Total % FTE per service
        = Number of services * % FTE/service

    Total Service Cost (K$)
        = Total % FTE per service * FTE Cost / 100

    Total Service Value (K$)
        Talent = Total Service Cost * FTE Value / 100
        AI     = Total % FTE per service * FTE Value / 100

    Yearly Margin Contribution (K$)
        = Total Service Value - Total Service Cost

    Monthly Margin Contribution (K$)
        = Yearly Margin Contribution / 12
    """

    number_of_services = numeric(number_of_services)
    fte_per_service = numeric(fte_per_service)
    xyz = normalize_xyz(xyz_value)

    total_fte = number_of_services * fte_per_service

    if xyz not in cost_map:
        # Empty side = zero calculations.
        if number_of_services == 0 and fte_per_service == 0 and xyz == "":
            fte_cost = 0.0
            fte_value = 0.0
        else:
            raise ValueError(
                f"Invalid Talent/AI value '{xyz_value}'. Expected X, Y or Z."
            )
    else:
        fte_cost = cost_map[xyz]
        fte_value = value_map[xyz]

    total_service_cost = (total_fte * fte_cost) / 100
    if service_value_base == "cost":
        total_service_value = (total_service_cost * fte_value) / 100
    elif service_value_base == "fte":
        total_service_value = (total_fte * fte_value) / 100
    else:
        raise ValueError("Invalid service_value_base. Expected 'cost' or 'fte'.")
    yearly_margin = total_service_value - total_service_cost
    monthly_margin = yearly_margin / 12

    return {
        "xyz": xyz,
        "fte_rate_cost": fte_cost,
        "fte_rate_value": fte_value,
        "total_fte": total_fte,
        "service_cost": total_service_cost,
        "service_value": total_service_value,
        "yearly_margin": yearly_margin,
        "monthly_margin": monthly_margin,
    }


def calculate_row(row, columns, service_type, premises_values):
    """
    Calculate one input row independently.

    Only the requested calculation outputs are returned.
    No ROI percentage is calculated.
    """

    period = service_type.lower()
    premises = premises_values[period]

    talent = calculate_side(
        row[columns["talent_services"]],
        row[columns["talent_fte"]],
        row[columns["talent_value"]],
        premises["talent"]["cost"],
        premises["talent"]["value"],
        service_value_base="cost",
    )

    ai = calculate_side(
        row[columns["ai_services"]],
        row[columns["ai_fte"]],
        row[columns["ai_value"]],
        premises["ai"]["cost"],
        premises["ai"]["value"],
        service_value_base="fte",
    )

    return {
        # Talent calculations
        "Talent Total % FTE": talent["total_fte"],
        "Talent Service Cost (K$)": talent["service_cost"],
        "Talent Service Value (K$)": talent["service_value"],
        "Talent Yearly Margin Contribution (K$)": talent["yearly_margin"],
        "Talent Monthly Margin Contribution (K$)": talent["monthly_margin"],

        # AI calculations
        "AI Total % FTE": ai["total_fte"],
        "AI Service Cost (K$)": ai["service_cost"],
        "AI Service Value (K$)": ai["service_value"],
        "AI Yearly Margin Contribution (K$)": ai["yearly_margin"],
        "AI Monthly Margin Contribution (K$)": ai["monthly_margin"],
    }


# =========================================================
# EXCEL TABLE EXTRACTION
# =========================================================

def read_excel_tables(input_file):
    """
    Read the supplied workbook and extract Current and Future Services.
    Ensures the Excel workbook is properly closed after reading.
    """

    with pd.ExcelFile(input_file) as xls:

        service_sheet = next(
            (
                s for s in xls.sheet_names
                if normalize_column_name(s).startswith("services project")
            ),
            None,
        )

        if service_sheet is None:
            raise ValueError(
                "Could not find the 'Services Project' sheet "
                "in the uploaded Excel file."
            )

        raw = pd.read_excel(
            xls,
            sheet_name=service_sheet,
            header=None,
        )

    # The Excel file is now closed.

    # Find the header row containing Service Improvement.
    header_row = None

    for i in range(len(raw)):
        values = [
            normalize_column_name(v)
            for v in raw.iloc[i].tolist()
            if not pd.isna(v)
        ]

        if (
            "service improvement." in values
            or "service improvement" in values
        ):
            header_row = i
            break

    if header_row is None:
        raise ValueError(
            "Could not find the service table header row."
        )

    # The supplied workbook places the service table in columns C:K.
    # The separate Cost/Schedule heading may be outside C:K, so locate it
    # on the full worksheet and exclude everything from that row onward.
    schedule_raw_row = None
    for raw_row_idx in range(header_row + 1, len(raw)):
        row_text = " ".join(
            normalize_column_name(v)
            for v in raw.iloc[raw_row_idx].tolist()
            if not pd.isna(v)
        )
        if "cost and schedule" in row_text and "future services" in row_text:
            schedule_raw_row = raw_row_idx
            break

    if schedule_raw_row is not None:
        table = raw.iloc[header_row + 1:schedule_raw_row, 2:11].copy()
    else:
        table = raw.iloc[header_row + 1:, 2:11].copy()

    table.columns = [
        "Service Improvement",
        "Team Member",
        "Talent Number of same services",
        "Talent % FTE/Service",
        "Talent value (X, Y or Z)",
        "AI companion",
        "AI Number of same services",
        "AI % FTE/Service",
        "Talent value AI (X, Y or Z)",
    ]

    # The table slice above already starts immediately below the service
    # header, so do not slice it a second time.
    table = table.reset_index(drop=True)

    # Remove completely empty rows.
    table = table.dropna(how="all").copy()

    service_col = table.columns[0]

    current_rows = []
    future_rows = []
    section = None

    for _, row in table.iterrows():

        label = (
            ""
            if pd.isna(row[service_col])
            else str(row[service_col]).strip()
        )

        lower = label.lower()

        if lower == "current services":
            section = "Current"
            continue

        if lower == "future services":
            section = "Future"
            continue

        # The Cost/Schedule table is a separate table below the service
        # sections. Do not accidentally treat its rows as Future Services
        # rows, because those rows do not contain Talent/AI X/Y/Z inputs.
        if "cost and schedule" in lower and "future services" in lower:
            section = None
            continue

        if lower == "total" or lower == "":
            continue

        if section == "Current":
            current_rows.append(row)

        elif section == "Future":
            future_rows.append(row)

    current_df = pd.DataFrame(
        current_rows,
        columns=table.columns,
    )

    future_df = pd.DataFrame(
        future_rows,
        columns=table.columns,
    )

    return (
        current_df.reset_index(drop=True),
        future_df.reset_index(drop=True),
    )


# =========================================================
# COST / SCHEDULE TO REACH FUTURE SERVICES
# =========================================================

def read_future_service_schedule(input_file):
    """
    Read the % FTE/Service inputs of the 'Cost and Schedule to reach Future
    services' table for BOTH Talent and AI.

    Returns:
        {"talent": [row1, row2, row4], "ai": [row1, row2, row4]}

    The section is searched dynamically (no fixed cells). The '% FTE/Service'
    headers can be on the same row as the section title or below it. Each
    header column with three numeric inputs beneath it is one side: the
    left-most column is Talent, the next one is AI. If only one column
    exists, it is used for both sides (backward compatible).

    Row 3 (Total Investment) is not an input; it is derived later as
    row 1 + row 2 separately for each side.
    """
    with pd.ExcelFile(input_file) as xls:
        service_sheet = next(
            (
                s for s in xls.sheet_names
                if normalize_column_name(s).startswith("services project")
            ),
            None,
        )

        if service_sheet is None:
            raise ValueError(
                "Could not find the 'Services Project' sheet "
                "in the uploaded Excel file."
            )

        raw = pd.read_excel(xls, sheet_name=service_sheet, header=None)

    section_row = None
    for row_idx in range(len(raw)):
        for col_idx in range(raw.shape[1]):
            normalized = normalize_column_name(raw.iat[row_idx, col_idx])
            if "cost and schedule" in normalized and "future services" in normalized:
                section_row = row_idx
                break
        if section_row is not None:
            break

    if section_row is None:
        raise ValueError(
            "Could not find the 'Cost and Schedule to reach Future services' "
            "section in the uploaded Excel file."
        )

    def first_three_numbers(col_idx, start_row):
        values = []
        for row_idx in range(start_row, min(start_row + 20, len(raw))):
            number = pd.to_numeric(raw.iat[row_idx, col_idx], errors="coerce")
            if not pd.isna(number):
                values.append(float(number))
                if len(values) == 3:
                    return values
        return None

    # 1) Find the header row (starting AT the section row, because the
    #    '% FTE/Service' headers usually share the title row).
    columns_found = []
    for header_row in range(section_row, min(section_row + 12, len(raw))):
        for col_idx in range(raw.shape[1]):
            normalized = normalize_column_name(raw.iat[header_row, col_idx])
            if "fte" not in normalized or "service" not in normalized:
                continue
            values = first_three_numbers(col_idx, header_row + 1)
            if values:
                columns_found.append((col_idx, values))
        if columns_found:
            break

    # 2) Fallback: merged/blank headers - any columns holding 3 numbers
    #    below the section title.
    if not columns_found:
        for col_idx in range(raw.shape[1]):
            values = first_three_numbers(col_idx, section_row + 1)
            if values:
                columns_found.append((col_idx, values))

    if not columns_found:
        raise ValueError(
            "The 'Cost and Schedule to reach Future services' table must "
            "contain three numeric % FTE/Service input values (rows 1, 2 "
            "and 4). None could be detected."
        )

    columns_found.sort(key=lambda item: item[0])
    talent_values = columns_found[0][1]
    ai_values = columns_found[1][1] if len(columns_found) > 1 else list(talent_values)

    return {"talent": talent_values, "ai": ai_values}


def calculate_future_service_schedule(input_file, premises_values):
    """
    Build the four-row 'Cost and Schedule to reach Future services' table.

    The input workbook provides three Total % FTE per service values:
      Row 1, Row 2 and Row 4.
    Row 3 is calculated as Row 1 + Row 2.

    Talent and AI each read their own % FTE/Service input column from the
    workbook and use their own Current FTE Cost from the Premises sheet.

    Outputs:
      Talent:
        Talent Total % FTE per service
        Talent - Total Service Cost* (K$)
      AI:
        AI Total % FTE per service
        AI - Total Service Cost** (K$)
      Monthly Margin Contribution is populated only on the recurring-cost
      row (row 4), separately for Talent and AI:
        monthly = recurring Total Service Cost / 12
      Talent+AI monthly = Talent monthly + AI monthly
    """
    fte_inputs = read_future_service_schedule(input_file)
    talent_fte_inputs = fte_inputs["talent"]
    ai_fte_inputs = fte_inputs["ai"]

    current_talent_fte_cost = numeric(
        premises_values["current"]["talent"]["cost"].get("X")
    )
    current_ai_fte_cost = numeric(
        premises_values["current"]["ai"]["cost"].get("X")
    )

    if current_talent_fte_cost == 0:
        raise ValueError(
            "Current Talent FTE Cost (X) could not be read from the Premises sheet."
        )
    if current_ai_fte_cost == 0:
        raise ValueError(
            "Current AI FTE Cost (X) could not be read from the Premises sheet."
        )

    # Input rows are 1, 2 and 4. Row 3 is derived from rows 1 + 2,
    # separately for Talent and AI (each has its own input column).
    talent_total_fte = [
        talent_fte_inputs[0],
        talent_fte_inputs[1],
        talent_fte_inputs[0] + talent_fte_inputs[1],
        talent_fte_inputs[2],
    ]
    ai_total_fte = [
        ai_fte_inputs[0],
        ai_fte_inputs[1],
        ai_fte_inputs[0] + ai_fte_inputs[1],
        ai_fte_inputs[2],
    ]

    talent_costs = [
        value * current_talent_fte_cost
        for value in talent_total_fte
    ]
    ai_costs = [
        value * current_ai_fte_cost
        for value in ai_total_fte
    ]

    # Monthly margin contribution is required only for the recurring-cost row.
    talent_monthly = [None, None, None, talent_costs[3] / 12]
    ai_monthly = [None, None, None, ai_costs[3] / 12]

    labels = [
        "Team and AI services % FTE (Tasks, Upskilling)",
        "Cost of NRE Services of Project ((in cards as % FTE* Cost per FTE (K$))",
        "Total Investment (%FTE*Cost/FTE+NRE)",
        "Recurring Cost of Project at % FTE",
    ]

    talent_ai_monthly = [
        None,
        None,
        None,
        talent_monthly[3] + ai_monthly[3],
    ]

    output = pd.DataFrame(
        {
            "Service Improvement": labels,
            "Talent Total % FTE per service": talent_total_fte,
            "Talent Total Service Cost* (K$)": talent_costs,
            "AI Total % FTE per service": ai_total_fte,
            "AI Total Service Cost** (K$)": ai_costs,
            "Talent Monthly Margin Contribution (K$)": talent_monthly,
            "AI Monthly Margin Contribution (K$)": ai_monthly,
            "Talent+AI Monthly Margin Contribution (K$)": talent_ai_monthly,
        }
    )

    for column in [
        "Talent Total % FTE per service",
        "Talent Total Service Cost* (K$)",
        "AI Total % FTE per service",
        "AI Total Service Cost** (K$)",
        "Talent Monthly Margin Contribution (K$)",
        "AI Monthly Margin Contribution (K$)",
        "Talent+AI Monthly Margin Contribution (K$)",
    ]:
        output[column] = pd.to_numeric(output[column], errors="coerce").round(2)

    return output


# =========================================================
# PROCESS ONE TABLE
# =========================================================

def process_table(df, service_type, premises_values):
    if df.empty:
        return pd.DataFrame()

    columns = find_input_columns(df)

    output_rows = []

    for _, row in df.iterrows():
        calculations = calculate_row(row, columns, service_type, premises_values)

        result = {
            "Service Type": service_type,
            "Service Improvement": row[columns["service_improvement"]],
            "Team Member": row[columns["team_member"]]
            if columns["team_member"] is not None
            else "",
            "Number of Services (Talent)": numeric(
                row[columns["talent_services"]]
            ),
            "% FTE/Service (Talent)": numeric(
                row[columns["talent_fte"]]
            ),
            "Talent Value (X/Y/Z)": normalize_xyz(
                row[columns["talent_value"]]
            ),
            "Talent Total % FTE": calculations["Talent Total % FTE"],
            "Talent Service Cost (K$)": calculations["Talent Service Cost (K$)"],
            "Talent Service Value (K$)": calculations["Talent Service Value (K$)"],
            "Talent Yearly Margin Contribution (K$)": calculations["Talent Yearly Margin Contribution (K$)"],
            "Talent Monthly Margin Contribution (K$)": calculations["Talent Monthly Margin Contribution (K$)"],
            "AI Companion": row[columns["ai_companion"]]
            if columns["ai_companion"] is not None
            else "",
            "Number of Services (AI)": numeric(
                row[columns["ai_services"]]
            ),
            "% FTE/Service (AI)": numeric(
                row[columns["ai_fte"]]
            ),
            "Talent Value AI (X/Y/Z)": normalize_xyz(
                row[columns["ai_value"]]
            ),
            "AI Total % FTE": calculations["AI Total % FTE"],
            "AI Service Cost (K$)": calculations["AI Service Cost (K$)"],
            "AI Service Value (K$)": calculations["AI Service Value (K$)"],
            "AI Yearly Margin Contribution (K$)": calculations["AI Yearly Margin Contribution (K$)"],
            "AI Monthly Margin Contribution (K$)": calculations["AI Monthly Margin Contribution (K$)"],
            "Talent+AI Total % FTE Cost": calculations["Talent Service Cost (K$)"] + calculations["AI Service Cost (K$)"],
            "Talent+AI Service Value": calculations["Talent Service Value (K$)"] + calculations["AI Service Value (K$)"],
            "Talent+AI Yearly Margin Contribution (K$)": calculations["Talent Yearly Margin Contribution (K$)"] + calculations["AI Yearly Margin Contribution (K$)"],
            "Talent+AI Monthly Margin Contribution (K$)": calculations["Talent Monthly Margin Contribution (K$)"] + calculations["AI Monthly Margin Contribution (K$)"],
        }

        output_rows.append(result)

    output = pd.DataFrame(output_rows)

    # Numeric output formatting.
    numeric_columns = output.select_dtypes(include="number").columns
    output[numeric_columns] = output[numeric_columns].round(2)

    # Add a final totals row. Sum numeric columns individually; leave text columns blank.
    totals_row = {column: (output[column].sum() if column in numeric_columns else "")
                  for column in output.columns}

    # Label the final totals row in the Service Improvement column.
    if "Service Improvement" in totals_row:
        totals_row["Service Improvement"] = "Total"

    output = pd.concat([output, pd.DataFrame([totals_row])], ignore_index=True)
    output[numeric_columns] = output[numeric_columns].round(2)

    return output



# def calculate_future_current_service_margin(current_df, future_df, future_schedule_df):
    """
    Build the five-row 'Future-Current Service Margin Contribution' table.

    The first two rows are the existing Net Gain/Loss and NRE rows.

    Talent:
      Row 1 Yearly  = Future Talent Yearly Margin - Current Talent Yearly Margin
      Row 1 Monthly = Future Talent Monthly Margin - Current Talent Monthly Margin

      Row 2 Yearly  = Talent Total Investment Service Cost* / 100
      Row 2 Monthly = Talent Recurring Service Cost* / 100

      Row 3 Yearly  = Row 1 Yearly - Row 2 Yearly - (Row 2 Monthly / 12)

      Row 4 Yearly  = (Row 2 Yearly / Row 3 Yearly) * 100

      Row 5 Monthly = Row 2 Yearly /
                      (Row 1 Monthly - Row 2 Monthly)

    AI uses exactly the same formulas, using the AI columns and
    AI Total Service Cost** values from the Cost/Schedule table.

    Cells for which the user did not request a value are left blank.
    """

    def total(df, col):
        if df.empty or col not in df.columns:
            return 0.0

        values = pd.to_numeric(df[col], errors="coerce").fillna(0)

        # process_table() adds a final "Total" row. Exclude it here,
        # otherwise the service values would be counted twice.
        if "Service Improvement" in df.columns:
            labels = df["Service Improvement"].astype(str).str.strip().str.lower()
            values = values[labels != "total"]

        return float(values.sum())

    def safe_divide(numerator, denominator):
        if denominator == 0:
            return None
        return numerator / denominator

    # ---------------------------------------------------------
    # Row 1: Future - Current service margin
    # ---------------------------------------------------------
    talent_net_yearly = (
        total(future_df, "Talent Yearly Margin Contribution (K$)")
        - total(current_df, "Talent Yearly Margin Contribution (K$)")
    )
    talent_net_monthly = (
        total(future_df, "Talent Monthly Margin Contribution (K$)")
        - total(current_df, "Talent Monthly Margin Contribution (K$)")
    )

    ai_net_yearly = (
        total(future_df, "AI Yearly Margin Contribution (K$)")
        - total(current_df, "AI Yearly Margin Contribution (K$)")
    )
    ai_net_monthly = (
        total(future_df, "AI Monthly Margin Contribution (K$)")
        - total(current_df, "AI Monthly Margin Contribution (K$)")
    )

    # ---------------------------------------------------------
    # Row 2: NRE / investment and recurring cost
    # ---------------------------------------------------------
    talent_cost_col = "Talent Total Service Cost* (K$)"
    ai_cost_col = "AI Total Service Cost** (K$)"

    talent_costs = (
        pd.to_numeric(
            future_schedule_df[talent_cost_col],
            errors="coerce",
        ).fillna(0)
        if talent_cost_col in future_schedule_df.columns
        else pd.Series(dtype=float)
    )

    ai_costs = (
        pd.to_numeric(
            future_schedule_df[ai_cost_col],
            errors="coerce",
        ).fillna(0)
        if ai_cost_col in future_schedule_df.columns
        else pd.Series(dtype=float)
    )

    # Cost/Schedule rows:
    # index 2 = Total Investment
    # index 3 = Recurring Cost of Project
    talent_investment = float(talent_costs.iloc[2]) if len(talent_costs) > 2 else 0.0
    talent_recurring = float(talent_costs.iloc[3]) if len(talent_costs) > 3 else 0.0

    ai_investment = float(ai_costs.iloc[2]) if len(ai_costs) > 2 else 0.0
    ai_recurring = float(ai_costs.iloc[3]) if len(ai_costs) > 3 else 0.0

    talent_nre_yearly = talent_investment / 100
    talent_nre_monthly = talent_recurring / 100

    ai_nre_yearly = ai_investment / 100
    ai_nre_monthly = ai_recurring / 100

    # ---------------------------------------------------------
    # Row 3: Net Margin Contribution Year 1
    #
    # Exact requested formula:
    # Net Gain/Loss yearly
    #   - NRE yearly
    #   - NRE monthly / 12
    # ---------------------------------------------------------
    talent_net_year_1 = (
        talent_net_yearly
        - talent_nre_monthly
        - (talent_nre_yearly / 12)
    )

    ai_net_year_1 = (
        ai_net_yearly
        - ai_nre_monthly
        - (ai_nre_yearly / 12)
    )

    # ---------------------------------------------------------
    # Row 4: ROI
    # = NRE yearly / Net Margin Contribution Year 1 * 100
    # ---------------------------------------------------------
    talent_roi = safe_divide(talent_net_year_1, talent_nre_yearly)
    ai_roi = safe_divide(ai_nre_yearly, ai_net_year_1)

    if talent_roi is not None:
        talent_roi *= 100
    if ai_roi is not None:
        ai_roi *= 100

    # ---------------------------------------------------------
    # Row 5: Months to Breakeven
    # = NRE yearly /
    #   (Net Gain/Loss monthly - NRE monthly)
    # ---------------------------------------------------------
    talent_breakeven = safe_divide(
        talent_nre_yearly,
        talent_nre_monthly - talent_net_monthly
    )

    ai_breakeven = safe_divide(
        ai_nre_yearly,
        ai_nre_monthly - ai_net_monthly
    )

    return pd.DataFrame(
        {
            "Service Improvement": [
                "Net Gain or Loss service Margin Contribution",
                "Cost of NRE Services of Project ((in cards as % FTE)",
                "Net Margin Contribution Year 1",
                "ROI (Net Margin Contribution/Investment) - %",
                "Months to Breakeven (Investment/Net Gain)",
            ],
            "Talent Yearly Margin Contribution (K$)": [
                talent_net_yearly,
                talent_nre_yearly,
                talent_net_year_1,
                talent_roi,
                None,
            ],
            "Talent Monthly Margin Contribution (K$)": [
                talent_net_monthly,
                talent_nre_monthly,
                None,
                None,
                talent_breakeven,
            ],
            "AI Yearly Margin Contribution (K$)": [
                ai_net_yearly,
                ai_nre_yearly,
                ai_net_year_1,
                ai_roi,
                None,
            ],
            "AI Monthly Margin Contribution (K$)": [
                ai_net_monthly,
                ai_nre_monthly,
                None,
                None,
                ai_breakeven,
            ],
        }
    ).round(2)

def calculate_future_current_service_margin(current_df, future_df, future_schedule_df):
    """
    Build the five-row 'Future-Current Service Margin Contribution' table.

    New columns:
      - Talent+AI Yearly Margin Contribution (K$)
      - Talent+AI Monthly Margin Contribution (K$)

    Row 1:
      Talent+AI Yearly
        = Total Current Talent+AI Yearly Margin
          - Total Future Talent+AI Yearly Margin

      Talent+AI Monthly
        = Current Talent Monthly Margin
          + Current AI Monthly Margin

    Row 2:
      Talent+AI Yearly
        = Future Current Service Margin Contribution
          Talent Yearly + AI Yearly

      Talent+AI Monthly
        = Future Current Service Margin Contribution
          Talent Monthly + AI Monthly

    Row 3:
      Talent+AI Yearly
        = Future Current Service Margin Contribution
          Talent Yearly + AI Yearly

      Talent+AI Monthly
        = Future Current Service Margin Contribution
          Talent Monthly + AI Monthly

    Row 4 and Row 5:
      Talent+AI columns are left blank.
    """

    def total(df, col):
        """Sum a numeric column while excluding the Total row."""
        if df.empty or col not in df.columns:
            return 0.0

        values = pd.to_numeric(df[col], errors="coerce").fillna(0)

        # process_table() adds a final Total row.
        if "Service Improvement" in df.columns:
            labels = df["Service Improvement"].astype(str).str.strip().str.lower()
            values = values[labels != "total"]

        return float(values.sum())

    def safe_divide(numerator, denominator):
        if denominator == 0:
            return None
        return numerator / denominator

    # =========================================================
    # ROW 1: Current - Future
    # =========================================================

    # Talent
    talent_current_yearly = total(
        current_df,
        "Talent Yearly Margin Contribution (K$)"
    )

    talent_future_yearly = total(
        future_df,
        "Talent Yearly Margin Contribution (K$)"
    )

    talent_current_monthly = total(
        current_df,
        "Talent Monthly Margin Contribution (K$)"
    )

    talent_future_monthly = total(
        future_df,
        "Talent Monthly Margin Contribution (K$)"
    )

    # AI
    ai_current_yearly = total(
        current_df,
        "AI Yearly Margin Contribution (K$)"
    )

    ai_future_yearly = total(
        future_df,
        "AI Yearly Margin Contribution (K$)"
    )

    ai_current_monthly = total(
        current_df,
        "AI Monthly Margin Contribution (K$)"
    )

    ai_future_monthly = total(
        future_df,
        "AI Monthly Margin Contribution (K$)"
    )

    # Existing Talent / AI Row 1 values
    talent_net_yearly = talent_future_yearly - talent_current_yearly
    talent_net_monthly = talent_future_monthly - talent_current_monthly

    ai_net_yearly = ai_future_yearly - ai_current_yearly
    ai_net_monthly = ai_future_monthly - ai_current_monthly

    # NEW:
    # User requested Current Talent+AI Yearly - Future Talent+AI Yearly
    talent_ai_yearly_row1 = (
        (talent_future_yearly + ai_future_yearly) - (talent_current_yearly + ai_current_yearly)
    )

    # User requested:
    # Talent monthly + AI monthly
    talent_ai_monthly_row1 = (
        (talent_future_monthly + ai_future_monthly) - (talent_current_monthly + ai_current_monthly)
    )

    # =========================================================
    # ROW 2: NRE / Investment and recurring cost
    # =========================================================

    talent_cost_col = "Talent Total Service Cost* (K$)"
    ai_cost_col = "AI Total Service Cost** (K$)"

    talent_costs = (
        pd.to_numeric(
            future_schedule_df[talent_cost_col],
            errors="coerce",
        ).fillna(0)
        if talent_cost_col in future_schedule_df.columns
        else pd.Series(dtype=float)
    )

    ai_costs = (
        pd.to_numeric(
            future_schedule_df[ai_cost_col],
            errors="coerce",
        ).fillna(0)
        if ai_cost_col in future_schedule_df.columns
        else pd.Series(dtype=float)
    )

    # Cost/Schedule rows:
    # index 2 = Total Investment
    # index 3 = Recurring Cost of Project
    talent_investment = (
        float(talent_costs.iloc[2])
        if len(talent_costs) > 2
        else 0.0
    )

    talent_recurring = (
        float(talent_costs.iloc[3])
        if len(talent_costs) > 3
        else 0.0
    )

    ai_investment = (
        float(ai_costs.iloc[2])
        if len(ai_costs) > 2
        else 0.0
    )

    ai_recurring = (
        float(ai_costs.iloc[3])
        if len(ai_costs) > 3
        else 0.0
    )

    talent_nre_yearly = talent_investment / 100
    talent_nre_monthly = talent_recurring / 100

    ai_nre_yearly = ai_investment / 100
    ai_nre_monthly = ai_recurring / 100

    # Existing Talent + AI Row 2
    talent_ai_yearly_row2 = (
        talent_nre_yearly
        + ai_nre_yearly
    )

    talent_ai_monthly_row2 = (
        talent_nre_monthly
        + ai_nre_monthly
    )

    # =========================================================
    # ROW 3: Net Margin Contribution Year 1
    # =========================================================

    talent_net_year_1 = (
        talent_net_yearly
        - talent_nre_monthly
        - (talent_nre_yearly / 12)
    )

    ai_net_year_1 = (
        ai_net_yearly
        - ai_nre_monthly
        - (ai_nre_yearly / 12)
    )

    # NEW:
    # Row 3 = Future Current Service Margin Contribution
    # Talent Yearly + AI Yearly
    talent_ai_yearly_row3 = (
        talent_net_year_1
        + ai_net_year_1
    )

    # User did not specify a Row 3 monthly formula.
    # Leave it blank.
    talent_ai_monthly_row3 = None

    # =========================================================
    # ROW 4: ROI
    # =========================================================

    talent_roi = safe_divide(
        ai_nre_yearly,
        ai_net_year_1
    )

    # Keep the original Talent ROI behavior
    talent_roi_original = safe_divide(
        talent_net_year_1,
        talent_nre_yearly
    )

    if talent_roi_original is not None:
        talent_roi_original *= 100

    if talent_roi is not None:
        talent_roi *= 100

    # Original AI ROI
    ai_roi = safe_divide(
        ai_net_yearly,
        ai_nre_yearly
    )

    print("ai yearly 3rd",ai_net_yearly)
    print("ai yearly 3rd",ai_net_year_1)
    print("ai yearly 2nd", ai_nre_yearly)
    print("ai yearly 1st row",ai_future_yearly)


    if ai_roi is not None:
        ai_roi *= 100

    talent_ai_yearly_row4 = safe_divide(
        talent_ai_yearly_row1,
        talent_ai_yearly_row2
    )

    if talent_ai_yearly_row4 is not None:
        talent_ai_yearly_row4 *= 100


    # =========================================================
    # ROW 5: Months to Breakeven
    # =========================================================

    talent_breakeven = safe_divide(
        talent_nre_yearly,
        talent_net_monthly-talent_nre_monthly 
    )

    ai_breakeven = safe_divide(
        ai_nre_yearly,
        ai_net_monthly- ai_nre_monthly 
    )

    breakeven_denominator = (
        talent_ai_monthly_row1
        - talent_ai_monthly_row2
    )

    talent_ai_monthly_row5 = safe_divide(
        talent_ai_yearly_row2,
        breakeven_denominator
    )


    # =========================================================
    # FINAL TABLE
    # =========================================================

    output = pd.DataFrame(
        {
            "Service Improvement": [
                "Net Gain or Loss service Margin Contribution",
                "Cost of NRE Services of Project ((in cards as % FTE)",
                "Net Margin Contribution Year 1",
                "ROI (Net Margin Contribution/Investment) - %",
                "Months to Breakeven (Investment/Net Gain)",
            ],

            # -------------------------
            # Existing Talent columns
            # -------------------------
            "Talent Yearly Margin Contribution (K$)": [
                talent_net_yearly,
                talent_nre_yearly,
                talent_net_year_1,
                talent_roi_original,
                None,
            ],

            "Talent Monthly Margin Contribution (K$)": [
                talent_net_monthly,
                talent_nre_monthly,
                None,
                None,
                talent_breakeven,
            ],

            # -------------------------
            # Existing AI columns
            # -------------------------
            "AI Yearly Margin Contribution (K$)": [
                ai_net_yearly,
                ai_nre_yearly,
                ai_net_year_1,
                ai_roi,
                None,
            ],

            "AI Monthly Margin Contribution (K$)": [
                ai_net_monthly,
                ai_nre_monthly,
                None,
                None,
                ai_breakeven,
            ],

            # -------------------------
            # NEW Talent + AI columns
            # -------------------------
            "Talent+AI Yearly Margin Contribution (K$)": [
                talent_ai_yearly_row1,
                talent_ai_yearly_row2,
                talent_ai_yearly_row3,
                talent_ai_yearly_row4,
                None,
            ],

            "Talent+AI Monthly Margin Contribution (K$)": [
                talent_ai_monthly_row1,
                talent_ai_monthly_row2,
                talent_ai_monthly_row3,
                None,
                talent_ai_monthly_row5,
            ],
        }
    )

    return output.round(2)



# =========================================================
# MAIN PROCESS
# =========================================================

def process_file(input_file):
    """
    Returns:
        current_df
        future_df
        combined_df
        summary_df
    """

    premises_values = read_premises_values(input_file)

    current_input, future_input = read_excel_tables(input_file)

    # Third table: Cost and Schedule to reach Future services.
    future_schedule_df = calculate_future_service_schedule(
        input_file,
        premises_values,
    )

    current_df = process_table(
        current_input,
        "Current",
        premises_values,
    )
    future_df = process_table(
        future_input,
        "Future",
        premises_values,
    )

    if not current_df.empty and not future_df.empty:
        combined_df = pd.concat(
            [current_df, future_df],
            ignore_index=True,
        )
    elif not current_df.empty:
        combined_df = current_df.copy()
    else:
        combined_df = future_df.copy()

    summary_rows = []

    for service_type, df in [
        ("Current", current_df),
        ("Future", future_df),
    ]:
        if df.empty:
            continue

        summary_rows.append({
            "Service Type": service_type,
            "Rows": len(df),
            "Talent Total % FTE": df["Talent Total % FTE"].sum(),
            "AI Total % FTE": df["AI Total % FTE"].sum(),
            "Talent Service Cost (K$)": df["Talent Service Cost (K$)"].sum(),
            "AI Service Cost (K$)": df["AI Service Cost (K$)"].sum(),
            "Talent Service Value (K$)": df["Talent Service Value (K$)"].sum(),
            "AI Service Value (K$)": df["AI Service Value (K$)"].sum(),
            "Talent Yearly Margin Contribution (K$)": (
                df["Talent Yearly Margin Contribution (K$)"].sum()
            ),
            "AI Yearly Margin Contribution (K$)": (
                df["AI Yearly Margin Contribution (K$)"].sum()
            ),
            "Talent Monthly Margin Contribution (K$)": (
                df["Talent Monthly Margin Contribution (K$)"].sum()
            ),
            "AI Monthly Margin Contribution (K$)": (
                df["AI Monthly Margin Contribution (K$)"].sum()
            ),
        })

    summary_df = pd.DataFrame(summary_rows).round(2)

    future_current_margin_df = calculate_future_current_service_margin(
        current_df, future_df, future_schedule_df
    )

    return (
        current_df,
        future_df,
        combined_df,
        summary_df,
        future_schedule_df,
        future_current_margin_df,
    )


# =========================================================
# EXCEL DOWNLOAD
# =========================================================

def create_excel_download(
    current_df,
    future_df,
    combined_df,
    summary_df,
    future_schedule_df,
    future_current_margin_df,
):
    """
    Create an Excel workbook containing Current Services, Future Services,
    and the Cost and Schedule to reach Future services table.
    """

    output = BytesIO()

    with pd.ExcelWriter(
        output,
        engine="openpyxl",
    ) as writer:
        sheet_name = "Calculated Services"

        # Current Services table.
        current_df.to_excel(
            writer,
            sheet_name=sheet_name,
            index=False,
            startrow=0,
        )

        # Future Services table starts after Current table + 3 blank rows.
        future_start_row = len(current_df) + 4
        future_df.to_excel(
            writer,
            sheet_name=sheet_name,
            index=False,
            startrow=future_start_row,
        )

        # Third table starts after Future Services + 3 blank rows.
        schedule_start_row = future_start_row + len(future_df) + 4
        future_schedule_df.to_excel(
            writer,
            sheet_name=sheet_name,
            index=False,
            startrow=schedule_start_row,
        )

        margin_start_row = schedule_start_row + len(future_schedule_df) + 4
        future_current_margin_df.to_excel(
            writer,
            sheet_name=sheet_name,
            index=False,
            startrow=margin_start_row,
        )

        worksheet = writer.book[sheet_name]
        worksheet.freeze_panes = "A2"

        for column_cells in worksheet.columns:
            max_length = 0
            column_letter = column_cells[0].column_letter
            for cell in column_cells:
                value = "" if cell.value is None else str(cell.value)
                max_length = max(max_length, len(value))
            worksheet.column_dimensions[column_letter].width = min(
                max(max_length + 2, 12),
                45,
            )

    output.seek(0)
    return output.getvalue()