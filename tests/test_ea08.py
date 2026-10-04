from audit_checks import check_duplicate_records
from workbook_utils import (
    load_audit_workbook,
    load_cached_workbook,
)


def test_ea08_detects_planted_duplicate_records():
    path = "data/Apex_FY26_Forecast.xlsx"

    formula_workbook = load_audit_workbook(path)
    cached_workbook = load_cached_workbook(path)

    findings = check_duplicate_records(
        formula_workbook,
        cached_workbook,
    )

    matches = [
        finding
        for finding in findings
        if finding.worksheet == "Customer Orders"
        and finding.location == "Row 8; Row 17"
    ]

    assert len(matches) == 1
    assert matches[0].check_id == "EA08"
    assert matches[0].priority == "Review Recommended"
    assert matches[0].confidence == "High"


def test_ea08_returns_only_expected_apex_duplicate():
    path = "data/Apex_FY26_Forecast.xlsx"

    formula_workbook = load_audit_workbook(path)
    cached_workbook = load_cached_workbook(path)

    findings = check_duplicate_records(
        formula_workbook,
        cached_workbook,
    )

    locations = {
        (finding.worksheet, finding.location)
        for finding in findings
    }

    assert locations == {
        ("Customer Orders", "Row 8; Row 17")
    }


def test_ea08_does_not_flag_legitimate_multiline_order():
    path = "data/Apex_FY26_Forecast.xlsx"

    formula_workbook = load_audit_workbook(path)
    cached_workbook = load_cached_workbook(path)

    findings = check_duplicate_records(
        formula_workbook,
        cached_workbook,
    )

    affected_rows = {
        item["row"]
        for finding in findings
        for item in finding.affected_items
    }

    # Rows 10 and 11 contain the legitimate ORD0004
    # multi-line order and must not be treated as duplicates.
    assert not ({10, 11} <= affected_rows)


def test_ea08_duplicate_evidence_contains_both_rows():
    path = "data/Apex_FY26_Forecast.xlsx"

    formula_workbook = load_audit_workbook(path)
    cached_workbook = load_cached_workbook(path)

    findings = check_duplicate_records(
        formula_workbook,
        cached_workbook,
    )

    duplicate = next(
        finding
        for finding in findings
        if finding.worksheet == "Customer Orders"
        and finding.location == "Row 8; Row 17"
    )

    rows = {
        item["row"]
        for item in duplicate.affected_items
    }

    assert rows == {8, 17}