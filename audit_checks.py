from collections import Counter

from audit_models import Finding
from workbook_utils import (
    EXCEL_ERROR_VALUES,
    detect_data_regions,
    formula_references_sheet,
    normalise_formula,
    classify_cell_value,
    is_identifier_header,
    normalise_record_value,
)




def check_broken_references(workbook) -> list[Finding]:
    """
    EA01 — Broken Formula References

    Flag formula cells containing #REF!.
    """

    findings = []

    for worksheet in workbook.worksheets:
        for row in worksheet.iter_rows():
            for cell in row:

                if (
                    cell.data_type == "f"
                    and isinstance(cell.value, str)
                    and "#REF!" in cell.value.upper()
                ):
                    finding = Finding(
                        finding_id="",
                        check_id="EA01",
                        check_name="Broken Formula Reference",
                        priority="High Priority",
                        worksheet=worksheet.title,
                        location=cell.coordinate,
                        summary="Broken formula reference detected",
                        detail=(
                            f"Formula '{cell.value}' contains a #REF! reference."
                        ),
                        why_review=(
                            "The formula refers to a cell or range that no longer "
                            "exists and may therefore produce an incorrect result."
                        ),
                        confidence="High",
                        affected_items=[
                            {
                                "worksheet": worksheet.title,
                                "location": cell.coordinate,
                                "formula": cell.value,
                            }
                        ],
                    )

                    findings.append(finding)

    return findings


def check_formula_pattern_inconsistencies(workbook) -> list[Finding]:
    """
    EA02 — Formula Pattern Inconsistency

    Identify isolated formulas that interrupt a repeated relative formula
    pattern within a locally comparable vertical calculation region.

    V1 requires at least four formula cells sharing the dominant pattern.
    Sustained changes to another formula family are treated as legitimate
    structural boundaries rather than formula anomalies.
    """

    findings = []

    for worksheet in workbook.worksheets:

        for column_cells in worksheet.iter_cols():

            # Build locally comparable regions.
            # A region continues through formulas and isolated numeric
            # constants, allowing EA03-style hardcodes to sit inside it.
            regions = []
            current_region = []

            for cell in column_cells:

                is_formula = (
                    cell.data_type == "f"
                    and isinstance(cell.value, str)
                    and "#REF!" not in cell.value.upper()
                )

                is_numeric_constant = (
                    cell.data_type != "f"
                    and isinstance(cell.value, (int, float))
                    and not isinstance(cell.value, bool)
                )

                if is_formula or (
                    is_numeric_constant
                    and current_region
                ):
                    current_region.append(cell)

                else:
                    if current_region:
                        regions.append(current_region)
                        current_region = []

            if current_region:
                regions.append(current_region)

            for region in regions:

                formula_cells = [
                    cell
                    for cell in region
                    if cell.data_type == "f"
                    and isinstance(cell.value, str)
                    and "#REF!" not in cell.value.upper()
                ]

                if len(formula_cells) < 4:
                    continue

                patterns = []

                for cell in formula_cells:
                    pattern = normalise_formula(cell)

                    if pattern is not None:
                        patterns.append((cell, pattern))

                if len(patterns) < 4:
                    continue

                pattern_counts = Counter(
                    pattern for _, pattern in patterns
                )

                dominant_pattern, dominant_count = (
                    pattern_counts.most_common(1)[0]
                )

                if dominant_count < 4:
                    continue

                for index, (cell, pattern) in enumerate(patterns):

                    if pattern == dominant_pattern:
                        continue

                    previous_pattern = (
                        patterns[index - 1][1]
                        if index > 0
                        else None
                    )

                    next_pattern = (
                        patterns[index + 1][1]
                        if index < len(patterns) - 1
                        else None
                    )

                    # EA02 V1 only treats a different formula as anomalous
                    # when it is isolated inside the dominant pattern.
                    is_isolated_anomaly = (
                        previous_pattern == dominant_pattern
                        and next_pattern == dominant_pattern
                    )

                    if not is_isolated_anomaly:
                        continue

                    finding = Finding(
                        finding_id="",
                        check_id="EA02",
                        check_name="Formula Pattern Inconsistency",
                        priority="Review Recommended",
                        worksheet=worksheet.title,
                        location=cell.coordinate,
                        summary=(
                            "Formula differs from the surrounding pattern"
                        ),
                        detail=(
                            f"Formula '{cell.value}' differs from the "
                            "repeated relative formula pattern immediately "
                            "before and after this cell."
                        ),
                        why_review=(
                            "The formula interrupts an otherwise consistent "
                            "calculation pattern and may represent an "
                            "unintended change to the calculation logic."
                        ),
                        confidence="Medium",
                        affected_items=[
                            {
                                "worksheet": worksheet.title,
                                "location": cell.coordinate,
                                "formula": cell.value,
                                "observed_pattern": pattern,
                                "dominant_pattern": dominant_pattern,
                            }
                        ],
                    )

                    findings.append(finding)

    return findings


def check_hardcoded_formula_overrides(workbook) -> list[Finding]:
    """
    EA03 — Hardcoded Formula Override

    Identify numeric constants that interrupt an established,
    consistent formula pattern.

    V1 requires:
    - the cell itself to contain a numeric constant;
    - formula cells immediately above and below;
    - those neighbouring formulas to share the same relative pattern.
    """

    findings = []

    for worksheet in workbook.worksheets:

        for column_cells in worksheet.iter_cols():

            cells = list(column_cells)

            for index in range(1, len(cells) - 1):

                cell = cells[index]
                previous_cell = cells[index - 1]
                next_cell = cells[index + 1]

                is_numeric_constant = (
                    cell.data_type != "f"
                    and isinstance(cell.value, (int, float))
                    and not isinstance(cell.value, bool)
                )

                if not is_numeric_constant:
                    continue

                previous_is_formula = (
                    previous_cell.data_type == "f"
                    and isinstance(previous_cell.value, str)
                    and "#REF!" not in previous_cell.value.upper()
                )

                next_is_formula = (
                    next_cell.data_type == "f"
                    and isinstance(next_cell.value, str)
                    and "#REF!" not in next_cell.value.upper()
                )

                if not (previous_is_formula and next_is_formula):
                    continue

                previous_pattern = normalise_formula(previous_cell)
                next_pattern = normalise_formula(next_cell)

                if (
                    previous_pattern is None
                    or next_pattern is None
                    or previous_pattern != next_pattern
                ):
                    continue

                finding = Finding(
                    finding_id="",
                    check_id="EA03",
                    check_name="Hardcoded Formula Override",
                    priority="Review Recommended",
                    worksheet=worksheet.title,
                    location=cell.coordinate,
                    summary="Fixed value interrupts a formula pattern",
                    detail=(
                        f"Cell {cell.coordinate} contains the fixed numeric "
                        f"value {cell.value} between formulas following the "
                        "same relative calculation pattern."
                    ),
                    why_review=(
                        "This value may have replaced a formula. If underlying "
                        "inputs change, the fixed value will not automatically "
                        "recalculate."
                    ),
                    confidence="Medium",
                    affected_items=[
                        {
                            "worksheet": worksheet.title,
                            "location": cell.coordinate,
                            "value": cell.value,
                            "previous_formula": previous_cell.value,
                            "next_formula": next_cell.value,
                            "expected_pattern": previous_pattern,
                        }
                    ],
                )

                findings.append(finding)

    return findings


def check_formula_errors(
    formula_workbook,
    cached_workbook,
) -> list[Finding]:
    """
    EA04 — Formula Errors

    Inspect cached results for formula cells and report recognised
    Excel error values.

    Related cells with the same error type are grouped when they form
    one directly connected horizontal or vertical block.

    #REF! is excluded because broken references are handled by EA01.
    """

    findings = []

    for formula_sheet in formula_workbook.worksheets:

        if formula_sheet.title not in cached_workbook.sheetnames:
            continue

        cached_sheet = cached_workbook[formula_sheet.title]

        error_cells = []

        for row in formula_sheet.iter_rows():
            for formula_cell in row:

                if (
                    formula_cell.data_type != "f"
                    or not isinstance(formula_cell.value, str)
                ):
                    continue

                cached_cell = cached_sheet[
                    formula_cell.coordinate
                ]
                cached_value = cached_cell.value

                if cached_value not in EXCEL_ERROR_VALUES:
                    continue

                # EA01 owns broken-reference findings.
                if cached_value == "#REF!":
                    continue

                error_cells.append(
                    {
                        "cell": formula_cell,
                        "error": cached_value,
                    }
                )

        # Group error cells by Excel error type first.
        errors_by_type = {}

        for error_item in error_cells:
            error_type = error_item["error"]

            errors_by_type.setdefault(
                error_type,
                [],
            ).append(error_item["cell"])

        for error_type, cells in errors_by_type.items():

            cells_by_position = {
                (cell.row, cell.column): cell
                for cell in cells
            }

            unvisited = set(cells_by_position)
            groups = []

            while unvisited:

                start_position = unvisited.pop()
                group_positions = [start_position]
                stack = [start_position]

                while stack:
                    row_index, column_index = stack.pop()

                    neighbours = [
                        (row_index - 1, column_index),
                        (row_index + 1, column_index),
                        (row_index, column_index - 1),
                        (row_index, column_index + 1),
                    ]

                    for neighbour in neighbours:
                        if neighbour not in unvisited:
                            continue

                        unvisited.remove(neighbour)
                        stack.append(neighbour)
                        group_positions.append(neighbour)

                group_cells = [
                    cells_by_position[position]
                    for position in group_positions
                ]

                group_cells.sort(
                    key=lambda cell: (
                        cell.row,
                        cell.column,
                    )
                )

                groups.append(group_cells)

            groups.sort(
                key=lambda group: (
                    group[0].row,
                    group[0].column,
                )
            )

            for group_cells in groups:

                affected_items = [
                    {
                        "worksheet": formula_sheet.title,
                        "location": cell.coordinate,
                        "formula": cell.value,
                        "cached_result": error_type,
                    }
                    for cell in group_cells
                ]

                if len(group_cells) == 1:
                    location = group_cells[0].coordinate

                    summary = (
                        f"Formula returns {error_type}"
                    )

                    detail = (
                        f"Formula '{group_cells[0].value}' has the "
                        f"cached Excel result '{error_type}'."
                    )

                else:
                    first_cell = group_cells[0]
                    last_cell = group_cells[-1]

                    location = (
                        f"{first_cell.coordinate}:"
                        f"{last_cell.coordinate}"
                    )

                    summary = (
                        f"{len(group_cells)} related formulas return "
                        f"{error_type}"
                    )

                    detail = (
                        f"{len(group_cells)} connected formula cells "
                        f"on worksheet '{formula_sheet.title}' have "
                        f"the cached Excel result '{error_type}'."
                    )

                if error_type == "#N/A":
                    priority = "Review Recommended"

                    why_review = (
                        "These formulas currently return #N/A. This can "
                        "be intentional in some lookup or fallback "
                        "scenarios, but should be reviewed to confirm "
                        "that the missing results are expected."
                    )

                else:
                    priority = "High Priority"

                    why_review = (
                        "These formulas currently return an Excel "
                        "calculation error and may affect these cells "
                        "or calculations that depend on them."
                    )

                finding = Finding(
                    finding_id="",
                    check_id="EA04",
                    check_name="Formula Error",
                    priority=priority,
                    worksheet=formula_sheet.title,
                    location=location,
                    summary=summary,
                    detail=detail,
                    why_review=why_review,
                    confidence="High",
                    affected_items=affected_items,
                )

                findings.append(finding)

    return findings


def check_hidden_calculation_sheets(workbook) -> list[Finding]:
    """
    EA05 — Hidden Calculation Sheets

    Identify hidden worksheets containing formulas.

    If a visible worksheet contains formulas that reference the
    hidden worksheet, classify the finding as Review Recommended.

    Otherwise report the hidden calculation sheet as Information.
    """

    findings = []

    for hidden_sheet in workbook.worksheets:

        if hidden_sheet.sheet_state == "visible":
            continue

        formula_cells = [
            cell
            for row in hidden_sheet.iter_rows()
            for cell in row
            if cell.data_type == "f"
            and isinstance(cell.value, str)
        ]

        if not formula_cells:
            continue

        visible_dependencies = []

        for visible_sheet in workbook.worksheets:

            if visible_sheet.sheet_state != "visible":
                continue

            for row in visible_sheet.iter_rows():
                for cell in row:

                    if (
                        cell.data_type == "f"
                        and isinstance(cell.value, str)
                        and formula_references_sheet(
                            cell.value,
                            hidden_sheet.title,
                        )
                    ):
                        visible_dependencies.append(
                            {
                                "worksheet": visible_sheet.title,
                                "location": cell.coordinate,
                                "formula": cell.value,
                            }
                        )

        if visible_dependencies:
            priority = "Review Recommended"
            confidence = "High"

            summary = (
                "Hidden calculation sheet feeds visible workbook outputs"
            )

            detail = (
                f"Hidden worksheet '{hidden_sheet.title}' contains "
                f"{len(formula_cells)} formula cell(s) and is referenced "
                f"by {len(visible_dependencies)} formula cell(s) on "
                "visible worksheets."
            )

            why_review = (
                "Important calculations are being performed on a hidden "
                "worksheet and feed visible workbook outputs. Users may "
                "not realise these hidden calculations affect the results "
                "they rely on."
            )

        else:
            priority = "Information"
            confidence = "Medium"

            summary = "Hidden worksheet contains calculations"

            detail = (
                f"Hidden worksheet '{hidden_sheet.title}' contains "
                f"{len(formula_cells)} formula cell(s), but no direct "
                "reference from a visible worksheet was identified."
            )

            why_review = (
                "Hidden calculation sheets are not necessarily a problem, "
                "but they can make workbook logic harder to understand "
                "and maintain."
            )

        finding = Finding(
            finding_id="",
            check_id="EA05",
            check_name="Hidden Calculation Sheet",
            priority=priority,
            worksheet=hidden_sheet.title,
            location="Worksheet",
            summary=summary,
            detail=detail,
            why_review=why_review,
            confidence=confidence,
            affected_items=visible_dependencies,
        )

        findings.append(finding)

    return findings


def check_suspicious_blanks(workbook) -> list[Finding]:
    """
    EA06 — Suspicious Blanks

    Flag blanks in highly populated columns within detected
    table-like data regions.

    V1 threshold:
        >= 90% of records contain a value in the column.

    A blank is only considered when the rest of that record
    contains meaningful data.
    """

    findings = []

    for worksheet in workbook.worksheets:

        regions = detect_data_regions(worksheet)

        for region in regions:

            total_rows = (
                region.end_row
                - region.start_row
                + 1
            )

            if total_rows < 10:
                continue

            for column in range(
                region.start_column,
                region.end_column + 1,
            ):

                values = [
                    worksheet.cell(row, column).value
                    for row in range(
                        region.start_row,
                        region.end_row + 1,
                    )
                ]

                populated_count = sum(
                    value not in (None, "")
                    for value in values
                )

                population_rate = (
                    populated_count / total_rows
                )

                if population_rate < 0.90:
                    continue

                suspicious_cells = []

                for row in range(
                    region.start_row,
                    region.end_row + 1,
                ):

                    cell = worksheet.cell(row, column)

                    if cell.value not in (None, ""):
                        continue

                    other_populated = sum(
                        worksheet.cell(
                            row,
                            other_column,
                        ).value not in (None, "")
                        for other_column in range(
                            region.start_column,
                            region.end_column + 1,
                        )
                        if other_column != column
                    )

                    if other_populated < 2:
                        continue

                    suspicious_cells.append(cell.coordinate)

                if not suspicious_cells:
                    continue

                header_index = (
                    column - region.start_column
                )

                header = region.headers[header_index]

                finding = Finding(
                    finding_id="",
                    check_id="EA06",
                    check_name="Suspicious Blank",
                    priority="Review Recommended",
                    worksheet=worksheet.title,
                    location="; ".join(suspicious_cells),
                    summary=(
                        f"Unexpected blank values detected in "
                        f"'{header}'"
                    ),
                    detail=(
                        f"Column '{header}' is populated in "
                        f"{population_rate:.1%} of records, but "
                        f"{len(suspicious_cells)} blank value(s) "
                        "were found within otherwise populated rows."
                    ),
                    why_review=(
                        "These blanks occur in a field that is normally "
                        "populated and may represent missing business data."
                    ),
                    confidence="Medium",
                    affected_items=[
                        {
                            "worksheet": worksheet.title,
                            "location": location,
                            "column": header,
                        }
                        for location in suspicious_cells
                    ],
                )

                findings.append(finding)

    return findings


def check_inconsistent_data_types(workbook) -> list[Finding]:
    """
    EA07 — Inconsistent Data Type

    Identify minority data types in columns where at least 90%
    of populated values share a dominant type.

    Identifier fields are excluded from V1 type consistency checks.
    """

    findings = []

    for worksheet in workbook.worksheets:

        regions = detect_data_regions(worksheet)

        for region in regions:

            total_rows = (
                region.end_row
                - region.start_row
                + 1
            )

            if total_rows < 10:
                continue

            for column in range(
                region.start_column,
                region.end_column + 1,
            ):

                header_index = column - region.start_column
                header = region.headers[header_index]

                if is_identifier_header(header):
                    continue

                typed_values = []

                for row in range(
                    region.start_row,
                    region.end_row + 1,
                ):

                    cell = worksheet.cell(row, column)

                    if cell.data_type == "f":
                        continue

                    value_type = classify_cell_value(
                        cell.value
                    )

                    if value_type is None:
                        continue

                    typed_values.append(
                        (
                            cell,
                            value_type,
                        )
                    )

                if len(typed_values) < 10:
                    continue

                type_counts = Counter(
                    value_type
                    for _, value_type in typed_values
                )

                dominant_type, dominant_count = (
                    type_counts.most_common(1)[0]
                )

                dominance_rate = (
                    dominant_count / len(typed_values)
                )

                if dominance_rate < 0.90:
                    continue

                anomalies = [
                    cell
                    for cell, value_type in typed_values
                    if value_type != dominant_type
                ]

                if not anomalies:
                    continue

                locations = [
                    cell.coordinate
                    for cell in anomalies
                ]

                finding = Finding(
                    finding_id="",
                    check_id="EA07",
                    check_name="Inconsistent Data Type",
                    priority="Review Recommended",
                    worksheet=worksheet.title,
                    location="; ".join(locations),
                    summary=(
                        f"Inconsistent data type detected in "
                        f"'{header}'"
                    ),
                    detail=(
                        f"Column '{header}' contains predominantly "
                        f"{dominant_type} values "
                        f"({dominance_rate:.1%}), but "
                        f"{len(anomalies)} value(s) use a "
                        "different data type."
                    ),
                    why_review=(
                        "Unexpected data types can interfere with "
                        "calculations, sorting, filtering, validation "
                        "and downstream data processing."
                    ),
                    confidence="Medium",
                    affected_items=[
                        {
                            "worksheet": worksheet.title,
                            "location": cell.coordinate,
                            "column": header,
                            "value": cell.value,
                            "observed_type": classify_cell_value(
                                cell.value
                            ),
                            "expected_type": dominant_type,
                        }
                        for cell in anomalies
                    ],
                )

                findings.append(finding)

    return findings


def check_duplicate_records(
    formula_workbook,
    cached_workbook,
) -> list[Finding]:
    """
    EA08 — Duplicate Records

    Identify exact duplicate business records within detected
    table-like data regions.

    Ordinary cells are compared using their stored values.
    Formula cells are compared using their last cached calculated
    result rather than their literal formula text.

    This prevents equivalent row-relative formulas such as
    =E8*F8 and =E17*F17 from making otherwise identical records
    appear different.
    """

    findings = []

    for formula_sheet in formula_workbook.worksheets:

        if formula_sheet.title not in cached_workbook.sheetnames:
            continue

        cached_sheet = cached_workbook[formula_sheet.title]

        regions = detect_data_regions(formula_sheet)

        for region in regions:

            total_rows = (
                region.end_row
                - region.start_row
                + 1
            )

            if total_rows < 2:
                continue

            records = {}

            for row in range(
                region.start_row,
                region.end_row + 1,
            ):

                record_values = []

                for column in range(
                    region.start_column,
                    region.end_column + 1,
                ):

                    formula_cell = formula_sheet.cell(
                        row,
                        column,
                    )

                    if formula_cell.data_type == "f":
                        comparison_value = cached_sheet.cell(
                            row,
                            column,
                        ).value
                    else:
                        comparison_value = formula_cell.value

                    record_values.append(
                        normalise_record_value(
                            comparison_value
                        )
                    )

                record = tuple(record_values)

                if all(
                    value in (None, "")
                    for value in record
                ):
                    continue

                records.setdefault(
                    record,
                    []
                ).append(row)

            duplicate_groups = [
                rows
                for rows in records.values()
                if len(rows) > 1
            ]

            for rows in duplicate_groups:

                row_locations = [
                    f"Row {row}"
                    for row in rows
                ]

                finding = Finding(
                    finding_id="",
                    check_id="EA08",
                    check_name="Duplicate Record",
                    priority="Review Recommended",
                    worksheet=formula_sheet.title,
                    location="; ".join(row_locations),
                    summary="Potential duplicate records detected",
                    detail=(
                        f"{len(rows)} records contain identical "
                        "business values across all columns in the "
                        "detected data region."
                    ),
                    why_review=(
                        "Exact duplicate records may represent repeated "
                        "transactions, duplicated data entry or duplicated "
                        "source records and could distort reporting or "
                        "analysis."
                    ),
                    confidence="High",
                    affected_items=[
                        {
                            "worksheet": formula_sheet.title,
                            "location": f"Row {row}",
                            "row": row,
                        }
                        for row in rows
                    ],
                )

                findings.append(finding)

    return findings