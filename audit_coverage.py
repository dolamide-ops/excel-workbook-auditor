from audit_models import AuditCoverage
from workbook_utils import (
    detect_data_regions,
    load_audit_workbook,
    load_cached_workbook,
)


def assess_audit_coverage(file_path):
    """
    Assess whether formula-dependent audit checks had
    the information required to operate fully.
    """

    formula_workbook = load_audit_workbook(
        file_path
    )

    cached_workbook = load_cached_workbook(
        file_path
    )

    formula_cells = 0
    missing_cached_results = 0
    ea08_affected = False

    for formula_sheet in formula_workbook.worksheets:

        if formula_sheet.title not in cached_workbook.sheetnames:
            continue

        cached_sheet = cached_workbook[
            formula_sheet.title
        ]

        data_regions = detect_data_regions(
            formula_sheet
        )

        for row in formula_sheet.iter_rows():
            for cell in row:

                if cell.data_type != "f":
                    continue

                formula_cells += 1

                cached_value = cached_sheet[
                    cell.coordinate
                ].value

                if cached_value is not None:
                    continue

                missing_cached_results += 1

                # EA08 only depends on cached formula
                # results when the formula occurs inside
                # a detected tabular data region.
                for region in data_regions:

                    inside_region = (
                        region.start_row
                        <= cell.row
                        <= region.end_row
                        and
                        region.start_column
                        <= cell.column
                        <= region.end_column
                    )

                    if inside_region:
                        ea08_affected = True
                        break

    if missing_cached_results == 0:
        return AuditCoverage(
            status="Complete",
            formula_cells=formula_cells,
            missing_cached_results=0,
            affected_checks=[],
            warnings=[],
        )

    affected_checks = [
        "EA04",
    ]

    if ea08_affected:
        affected_checks.append(
            "EA08"
        )

    warning = (
        f"{missing_cached_results} formula cell(s) do not "
        "have cached calculation results available. "
        "Formula-dependent audit checks may therefore "
        "have limited coverage. Recalculate and save the "
        "workbook in Excel before rerunning the audit."
    )

    return AuditCoverage(
        status="Limited",
        formula_cells=formula_cells,
        missing_cached_results=missing_cached_results,
        affected_checks=affected_checks,
        warnings=[warning],
    )