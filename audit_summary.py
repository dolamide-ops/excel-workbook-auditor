from pathlib import Path

from audit_models import AuditResult
from audit_coverage import assess_audit_coverage


def build_audit_result(
    file_path,
    findings,
    checks_completed=8,
):
    """
    Build the overall V1 audit result from individual findings.
    """

    workbook_name = Path(file_path).name

    high_priority_count = sum(
        finding.priority == "High Priority"
        for finding in findings
    )

    review_recommended_count = sum(
        finding.priority == "Review Recommended"
        for finding in findings
    )

    information_count = sum(
        finding.priority == "Information"
        for finding in findings
    )

    affected_worksheets = sorted(
        {
            finding.worksheet
            for finding in findings
        }
    )

    coverage = assess_audit_coverage(
    file_path
    )

    return AuditResult(
        workbook_name=workbook_name,
        findings=findings,
        checks_completed=checks_completed,
        total_findings=len(findings),
        high_priority_count=high_priority_count,
        review_recommended_count=review_recommended_count,
        information_count=information_count,
        worksheets_affected=len(affected_worksheets),
        affected_worksheets=affected_worksheets,
        coverage=coverage,
    )