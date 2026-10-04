from audit_checks import check_hidden_calculation_sheets
from workbook_utils import load_audit_workbook


def test_ea05_detects_hidden_active_calculation_sheet():
    workbook = load_audit_workbook(
        "data/Apex_FY26_Forecast.xlsx"
    )

    findings = check_hidden_calculation_sheets(workbook)

    matches = [
        finding
        for finding in findings
        if finding.worksheet == "Calc_Helper"
    ]

    assert len(matches) == 1
    assert matches[0].check_id == "EA05"
    assert matches[0].priority == "Review Recommended"
    assert matches[0].confidence == "High"


def test_ea05_detects_visible_dependencies():
    workbook = load_audit_workbook(
        "data/Apex_FY26_Forecast.xlsx"
    )

    findings = check_hidden_calculation_sheets(workbook)

    calc_helper = next(
        finding
        for finding in findings
        if finding.worksheet == "Calc_Helper"
    )

    dependencies = {
        (
            item["worksheet"],
            item["location"],
        )
        for item in calc_helper.affected_items
    }

    assert dependencies == {
        ("Executive Summary", "B7"),
        ("Executive Summary", "B8"),
        ("Executive Summary", "B9"),
    }


def test_ea05_returns_only_expected_apex_finding():
    workbook = load_audit_workbook(
        "data/Apex_FY26_Forecast.xlsx"
    )

    findings = check_hidden_calculation_sheets(workbook)

    assert len(findings) == 1
    assert findings[0].worksheet == "Calc_Helper"