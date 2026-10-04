from audit_checks import check_broken_references
from workbook_utils import load_audit_workbook


def test_ea01_detects_planted_broken_reference():
    workbook = load_audit_workbook(
        "data/Apex_FY26_Forecast.xlsx"
    )

    findings = check_broken_references(workbook)

    matches = [
        finding
        for finding in findings
        if finding.worksheet == "Executive Summary"
        and finding.location == "B10"
    ]

    assert len(matches) == 1
    assert matches[0].check_id == "EA01"
    assert matches[0].priority == "High Priority"
    assert matches[0].confidence == "High"


def test_ea01_clean_workbook_has_no_broken_references():
    workbook = load_audit_workbook(
        "data/Apex_FY26_Forecast_CLEAN.xlsx"
    )

    findings = check_broken_references(workbook)

    assert len(findings) == 0