from baseline_validator import validate_against_baseline


def test_complete_audit_matches_expected_findings_register():
    result = validate_against_baseline(
        "data/Apex_FY26_Forecast.xlsx",
        "tests/Expected_Findings_Register.xlsx",
    )

    assert result["passed"] is True
    assert result["expected_count"] == 9
    assert result["actual_count"] == 9
    assert result["matched_count"] == 9
    assert result["missed"] == []
    assert result["unexpected"] == []


def test_clean_workbook_returns_no_findings():
    from audit_engine import run_audit

    findings = run_audit(
        "data/Apex_FY26_Forecast_CLEAN.xlsx"
    )

    assert findings == []


def test_clean_workbook_has_complete_cached_formula_results():
    from diagnostics import inspect_cached_formula_results

    result = inspect_cached_formula_results(
        "data/Apex_FY26_Forecast_CLEAN.xlsx"
    )

    assert result["formula_count"] == 43
    assert result["missing_cached_count"] == 0