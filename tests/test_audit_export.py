from audit_engine import run_complete_audit
from audit_export import build_audit_report


def test_audit_report_contains_expected_sheets():
    result = run_complete_audit(
        "data/Apex_FY26_Forecast.xlsx"
    )

    workbook = build_audit_report(result)

    assert workbook.sheetnames == [
        "Audit Summary",
        "Findings",
        "Finding Detail",
        "Audit Information",
    ]


def test_audit_summary_contains_expected_metrics():
    result = run_complete_audit(
        "data/Apex_FY26_Forecast.xlsx"
    )

    workbook = build_audit_report(result)
    worksheet = workbook["Audit Summary"]

    assert worksheet["A1"].value == (
        "EXCEL WORKBOOK AUDIT REPORT"
    )

    assert worksheet["A3"].value == (
        "Apex_FY26_Forecast.xlsx"
    )

    assert worksheet["A7"].value == 9
    assert worksheet["C7"].value == 2
    assert worksheet["E7"].value == 7

    assert worksheet["B10"].value == 8
    assert worksheet["E10"].value == 5


def test_findings_sheet_contains_all_findings():
    result = run_complete_audit(
        "data/Apex_FY26_Forecast.xlsx"
    )

    workbook = build_audit_report(result)
    worksheet = workbook["Findings"]

    finding_ids = [
        worksheet.cell(row, 1).value
        for row in range(
            2,
            worksheet.max_row + 1,
        )
    ]

    assert finding_ids == [
    "F001",
    "F004",
    "F002",
    "F003",
    "F005",
    "F006",
    "F007",
    "F008",
    "F009",
    ]


def test_finding_detail_preserves_explanation_and_evidence():
    result = run_complete_audit(
        "data/Apex_FY26_Forecast.xlsx"
    )

    workbook = build_audit_report(result)
    worksheet = workbook["Finding Detail"]

    assert worksheet["A2"].value == "F001"
    assert worksheet["G2"].value == (
        "Broken formula reference detected"
    )

    assert "#REF!" in worksheet["H2"].value
    assert "cell or range" in worksheet["I2"].value

    assert (
        "Worksheet: Executive Summary"
        in worksheet["K2"].value
    )

    assert "Location: B10" in worksheet["K2"].value
    assert "Formula: =#REF!" in worksheet["K2"].value


def test_clean_workbook_report_contains_no_findings():
    result = run_complete_audit(
        "data/Apex_FY26_Forecast_CLEAN.xlsx"
    )

    workbook = build_audit_report(result)

    findings_sheet = workbook["Findings"]
    summary_sheet = workbook["Audit Summary"]

    assert result.total_findings == 0
    assert findings_sheet.max_row == 1

    assert summary_sheet["A13"].value == (
    "No issues were detected by the enabled "
    "audit checks."
)


def test_audit_information_shows_complete_coverage():
    from audit_engine import run_complete_audit
    from audit_export import build_audit_report

    audit_result = run_complete_audit(
        "data/Apex_FY26_Forecast.xlsx"
    )

    workbook = build_audit_report(
        audit_result
    )

    worksheet = workbook[
        "Audit Information"
    ]

    assert worksheet["A10"].value == (
        "AUDIT COVERAGE"
    )

    assert worksheet["A11"].value == "Status"
    assert worksheet["B11"].value == "Complete"

    assert worksheet["A12"].value == (
        "Formula Cells Assessed"
    )
    assert worksheet["B12"].value == 284

    assert worksheet["A13"].value == (
        "Missing Cached Results"
    )
    assert worksheet["B13"].value == 0

    assert worksheet["A14"].value == (
        "Checks with Reduced Coverage"
    )
    assert worksheet["B14"].value == "None"


def test_audit_summary_shows_coverage_status():
    from audit_engine import run_complete_audit
    from audit_export import build_audit_report

    audit_result = run_complete_audit(
        "data/Apex_FY26_Forecast.xlsx"
    )

    workbook = build_audit_report(
        audit_result
    )

    worksheet = workbook[
        "Audit Summary"
    ]

    coverage_heading_found = False
    coverage_status_found = False

    for row in worksheet.iter_rows():
        for cell in row:

            if cell.value == "AUDIT COVERAGE":
                coverage_heading_found = True

            if (
                coverage_heading_found
                and cell.value == "Complete"
            ):
                coverage_status_found = True

    assert coverage_heading_found is True
    assert coverage_status_found is True