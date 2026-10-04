from audit_engine import run_complete_audit


def test_apex_audit_summary():
    result = run_complete_audit(
        "data/Apex_FY26_Forecast.xlsx"
    )

    assert result.workbook_name == "Apex_FY26_Forecast.xlsx"
    assert result.checks_completed == 8
    assert result.total_findings == 9

    assert result.high_priority_count == 2
    assert result.review_recommended_count == 7
    assert result.information_count == 0

    assert result.worksheets_affected == 5

    assert set(result.affected_worksheets) == {
        "Calc_Helper",
        "Customer Orders",
        "Executive Summary",
        "Product Master",
        "Sales Forecast",
    }


def test_clean_workbook_audit_summary():
    result = run_complete_audit(
        "data/Apex_FY26_Forecast_CLEAN.xlsx"
    )

    assert result.checks_completed == 8
    assert result.total_findings == 0

    assert result.high_priority_count == 0
    assert result.review_recommended_count == 0
    assert result.information_count == 0

    assert result.worksheets_affected == 0
    assert result.affected_worksheets == []