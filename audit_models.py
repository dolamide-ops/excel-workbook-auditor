from dataclasses import dataclass, field
from typing import Any


@dataclass
class Finding:
    finding_id: str
    check_id: str
    check_name: str
    priority: str
    worksheet: str
    location: str
    summary: str
    detail: str
    why_review: str
    confidence: str
    affected_items: list[dict[str, Any]] = field(default_factory=list)


@dataclass
class DataRegion:
    worksheet: str
    header_row: int
    start_row: int
    end_row: int
    start_column: int
    end_column: int
    headers: list[str]


@dataclass
class AuditCoverage:
    status: str
    formula_cells: int
    missing_cached_results: int
    affected_checks: list[str]
    warnings: list[str]


@dataclass
class AuditResult:
    workbook_name: str
    findings: list[Finding]
    checks_completed: int
    total_findings: int
    high_priority_count: int
    review_recommended_count: int
    information_count: int
    worksheets_affected: int
    affected_worksheets: list[str]
    coverage: AuditCoverage