from audit_checks import check_suspicious_blanks
from workbook_utils import (
    detect_data_regions,
    load_audit_workbook,
)


def test_customer_orders_region_is_detected_correctly():
    workbook = load_audit_workbook(
        "data/Apex_FY26_Forecast.xlsx"
    )

    worksheet = workbook["Customer Orders"]
    regions = detect_data_regions(worksheet)

    assert len(regions) == 1

    region = regions[0]

    assert region.header_row == 6
    assert region.start_row == 7
    assert region.end_row == 256
    assert region.start_column == 1
    assert region.end_column == 10



def test_ea06_detects_customer_name_blanks():
    workbook = load_audit_workbook(
        "data/Apex_FY26_Forecast.xlsx"
    )

    findings = check_suspicious_blanks(workbook)

    matches = [
        finding
        for finding in findings
        if finding.worksheet == "Customer Orders"
        and finding.location == "C10; C13"
    ]

    assert len(matches) == 1
    assert matches[0].check_id == "EA06"
    assert matches[0].priority == "Review Recommended"
    assert matches[0].confidence == "Medium"


def test_ea06_groups_related_blanks_into_one_finding():
    workbook = load_audit_workbook(
        "data/Apex_FY26_Forecast.xlsx"
    )

    findings = check_suspicious_blanks(workbook)

    customer_name_finding = next(
        finding
        for finding in findings
        if finding.worksheet == "Customer Orders"
        and finding.location == "C10; C13"
    )

    locations = {
        item["location"]
        for item in customer_name_finding.affected_items
    }

    assert locations == {"C10", "C13"}


def test_ea06_returns_only_expected_apex_finding():
    workbook = load_audit_workbook(
        "data/Apex_FY26_Forecast.xlsx"
    )

    findings = check_suspicious_blanks(workbook)

    assert len(findings) == 1
    assert findings[0].worksheet == "Customer Orders"
    assert findings[0].location == "C10; C13"