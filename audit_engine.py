from pathlib import Path
from audit_summary import build_audit_result

from audit_checks import (
    check_broken_references,
    check_duplicate_records,
    check_formula_errors,
    check_formula_pattern_inconsistencies,
    check_hardcoded_formula_overrides,
    check_hidden_calculation_sheets,
    check_inconsistent_data_types,
    check_suspicious_blanks,
)
from workbook_utils import (
    load_audit_workbook,
    load_cached_workbook,
)

from workbook_compatibility import (
    create_compatible_audit_copy,
)

import shutil

def run_audit(file_path: str | Path):
    """
    Run the complete V1 Excel Workbook Auditor.

    EA01 — Broken Formula References
    EA02 — Formula Pattern Inconsistency
    EA03 — Hardcoded Formula Override
    EA04 — Formula Errors
    EA05 — Hidden Calculation Sheets
    EA06 — Suspicious Blanks
    EA07 — Inconsistent Data Types
    EA08 — Duplicate Records
    """

    formula_workbook = load_audit_workbook(file_path)
    cached_workbook = load_cached_workbook(file_path)

    findings = []

    # Formula and workbook-structure checks
    findings.extend(
        check_broken_references(formula_workbook)
    )

    findings.extend(
        check_formula_pattern_inconsistencies(
            formula_workbook
        )
    )

    findings.extend(
        check_hardcoded_formula_overrides(
            formula_workbook
        )
    )

    findings.extend(
        check_formula_errors(
            formula_workbook,
            cached_workbook,
        )
    )

    findings.extend(
        check_hidden_calculation_sheets(
            formula_workbook
        )
    )

    # Data-quality checks
    findings.extend(
        check_suspicious_blanks(
            formula_workbook
        )
    )

    findings.extend(
        check_inconsistent_data_types(
            formula_workbook
        )
    )

    findings.extend(
        check_duplicate_records(
            formula_workbook,
            cached_workbook,
        )
    )

    # Assign stable finding IDs after all checks have run.
    for index, finding in enumerate(
        findings,
        start=1,
    ):
        finding.finding_id = f"F{index:03d}"

    return findings


def run_complete_audit(file_path):

    compatibility_result = (
        create_compatible_audit_copy(
            file_path
        )
    )

    audit_path = (
        compatibility_result["path"]
    )

    try:

        findings = run_audit(
            audit_path
        )

        audit_result = build_audit_result(
            audit_path,
            findings,
        )

        return audit_result

    finally:

        temporary_directory = (
            compatibility_result[
                "temporary_directory"
            ]
        )

        try:
            shutil.rmtree(
                temporary_directory
            )
        except PermissionError:
            pass