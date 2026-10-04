from audit_checks import check_hardcoded_formula_overrides
from workbook_utils import load_audit_workbook


def test_ea03_detects_planted_hardcoded_override():
    workbook = load_audit_workbook(
        "data/Apex_FY26_Forecast.xlsx"
    )

    findings = check_hardcoded_formula_overrides(workbook)

    matches = [
        finding
        for finding in findings
        if finding.worksheet == "Sales Forecast"
        and finding.location == "E14"
    ]

    assert len(matches) == 1
    assert matches[0].check_id == "EA03"
    assert matches[0].priority == "Review Recommended"
    assert matches[0].confidence == "Medium"


def test_ea03_returns_only_expected_apex_finding():
    workbook = load_audit_workbook(
        "data/Apex_FY26_Forecast.xlsx"
    )

    findings = check_hardcoded_formula_overrides(workbook)

    locations = {
        (finding.worksheet, finding.location)
        for finding in findings
    }

    assert locations == {
        ("Sales Forecast", "E14")
    }


def test_ea03_clean_workbook_has_no_hardcoded_overrides():
    workbook = load_audit_workbook(
        "data/Apex_FY26_Forecast_CLEAN.xlsx"
    )

    findings = check_hardcoded_formula_overrides(workbook)

    assert len(findings) == 0