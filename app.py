import io
import os
import tempfile
import shutil
from zipfile import BadZipFile

import pandas as pd
import streamlit as st


from audit_engine import run_complete_audit
from audit_export import build_audit_report
from workbook_compatibility import (
    create_compatible_audit_copy,
)


st.set_page_config(
    page_title="Excel Workbook Auditor",
    page_icon="📊",
    layout="wide",
)


st.title("Excel Workbook Auditor")

st.write(
    "Audit Excel workbooks for formula, structure and "
    "data-quality issues before they affect your analysis "
    "or decisions."
)

st.caption(
    "Get prioritised findings, clear explanations, "
    "audit coverage information and a downloadable "
    "Excel audit report."
)

st.caption(
    "Designed and developed by **Deji Olamide**"
)


with st.expander(
    "What does the auditor check?",
    expanded=False,
):
    st.markdown(
        """
        The current version checks for:

        - Broken formula references
        - Formula pattern inconsistencies
        - Hardcoded formula overrides
        - Formula errors
        - Hidden calculation sheets
        - Suspicious blanks
        - Inconsistent data types
        - Duplicate records
        """
    )

    st.caption(
        "Checks are designed to identify items for review. "
        "A finding does not necessarily mean the workbook "
        "is incorrect."
    )


uploaded_file = st.file_uploader(
    "Upload an Excel workbook",
    type=["xlsx"],
)


current_file_name = (
    uploaded_file.name
    if uploaded_file is not None
    else None
)

previous_file_name = st.session_state.get(
    "current_uploaded_file"
)

if current_file_name != previous_file_name:

    st.session_state[
        "current_uploaded_file"
    ] = current_file_name

    st.session_state.pop(
        "audit_result",
        None,
    )

    st.session_state.pop(
        "audited_file_name",
        None,
    )

    st.session_state.pop(
        "compatibility_adjustment",
        None,
    )


if uploaded_file is None:

    st.info(
        "Upload an .xlsx workbook to begin."
    )

else:

    st.success(
        f"Workbook loaded: {uploaded_file.name}"
    )

    st.caption(
    "The auditor analyses a temporary copy. "
    "Your uploaded workbook is not modified."
    )

    compatibility_directory = None

    run_audit = st.button(
        "Run Workbook Audit",
        type="primary",
    )

    if run_audit:

        temporary_path = None
        
        st.session_state.pop(
            "audit_result",
            None,
        )

        st.session_state.pop(
            "audited_file_name",
            None,
        )

        try:

            with tempfile.NamedTemporaryFile(
                delete=False,
                suffix=".xlsx",
            ) as temporary_file:

                temporary_file.write(
                    uploaded_file.getvalue()
                )

                temporary_path = (
                    temporary_file.name
                )

            with st.spinner(
                "Auditing workbook..."
            ):

                compatibility_result = (
                    create_compatible_audit_copy(
                        temporary_path
                    )
                )

                compatibility_directory = (
                    compatibility_result[
                        "temporary_directory"
                    ]
                )

                audit_result = (
                    run_complete_audit(
                        compatibility_result["path"]
                    )
                )

                audit_result.workbook_name = (
                    uploaded_file.name
                )

            st.session_state[
                "audit_result"
            ] = audit_result

            st.session_state[
                "audited_file_name"
            ] = uploaded_file.name

            st.session_state.pop(
                "compatibility_adjustment",
                None,
            )

            st.session_state[
                "compatibility_adjustment"
            ] = (
                compatibility_result[
                    "adjustment_applied"
                ]
            )
        
        except BadZipFile:

            st.error(
                "This file is not a valid Excel workbook."
            )

            st.write(
                "The file has an `.xlsx` extension, but its "
                "contents could not be recognised as a valid "
                "Excel workbook."
            )

            st.write(
                "Check that the file opens correctly in "
                "Microsoft Excel, then save or export a valid "
                "`.xlsx` copy and upload it again."
            )

            st.caption(
                "The auditor has not changed your original file."
            )        
        
        except Exception as error:

            error_message = str(error).lower()

            workbook_read_error = (
                "unable to read workbook"
                in error_message
                or "could not read stylesheet"
                in error_message
                or "invalid xml"
                in error_message
            )

            if workbook_read_error:

                st.error(
                    "This workbook could not be read by "
                    "the audit engine."
                )

                st.write(
                    "The workbook appears to contain Excel "
                    "formatting or file information that is not "
                    "supported by the current version of the "
                    "auditor."
                )

                with st.expander(
                    "What can I do?"
                ):

                    st.markdown(
                        """
                        1. Open the workbook in **Microsoft Excel**.
                        2. Select **File → Save As**.
                        3. Save a new copy as **Excel Workbook (.xlsx)**.
                        4. Return here and upload the newly saved copy.
                        5. Select **Run Workbook Audit** again.
                        """
                    )

                    st.caption(
                        "The auditor has not changed your original "
                        "workbook."
                    )

            else:

                st.error(
                    "Something went wrong while auditing "
                    "this workbook."
                )

                st.write(
                    "The audit could not be completed. "
                    "Please check the workbook and try again."
                )

        finally:

            if (
                temporary_path
                and os.path.exists(
                    temporary_path
                )
            ):

                try:

                    os.remove(
                        temporary_path
                    )

                except PermissionError:

                    pass

    if (
        compatibility_directory
        and os.path.exists(
            compatibility_directory
        )
    ):
        shutil.rmtree(
            compatibility_directory,
            ignore_errors=True,
        )

    if "audit_result" in st.session_state:

        audit_result = st.session_state[
            "audit_result"
        ]

        st.divider()

        st.subheader("Audit Results")

        col1, col2, col3, col4 = st.columns(
            4
        )

        col1.metric(
            "Total Findings",
            audit_result.total_findings,
        )

        col2.metric(
            "High Priority",
            audit_result.high_priority_count,
        )

        col3.metric(
            "Review Recommended",
            audit_result.review_recommended_count,
        )

        col4.metric(
            "Audit Coverage",
            audit_result.coverage.status,
        )

        st.markdown(
            "### Audit Overview"
        )

        if st.session_state.get(
            "compatibility_adjustment",
            False,
        ):
            st.info(
                "Compatibility adjustment applied — "
                "a temporary copy of this workbook was adjusted "
                "so it could be audited. Your original workbook "
                "was not changed."
            )

        if audit_result.high_priority_count > 0:

            st.error(
                f"{audit_result.high_priority_count} high-priority "
                f"{'finding requires' if audit_result.high_priority_count == 1 else 'findings require'} "
                "attention. Review these first before relying on "
                "the workbook for analysis or decisions."
            )

        elif audit_result.review_recommended_count > 0:

            st.warning(
                "No high-priority issues were detected, but "
                f"{audit_result.review_recommended_count} "
                f"{'finding should' if audit_result.review_recommended_count == 1 else 'findings should'} "
                "be reviewed to confirm whether the identified "
                "patterns are intentional."
            )

        else:

            st.success(
                "No issues were detected by the enabled audit "
                "checks."
            )


        if audit_result.coverage.status == "Limited":

            st.warning(
                "Audit coverage is limited. Some formula-dependent "
                "checks could not fully assess this workbook because "
                "calculated formula results were unavailable."
            )

            with st.expander(
                "Why is audit coverage limited?"
            ):

                st.write(
                    "Some Excel formulas store their latest calculated "
                    "results inside the workbook. The auditor uses these "
                    "results for checks that need to understand what a "
                    "formula currently evaluates to."
                )

                st.write(
                    "To improve coverage, recalculate the workbook "
                    "in Microsoft Excel, save it, and then run the "
                    "audit again."
                )

                st.markdown(
                    "**How to recalculate in Excel**"
                )

                st.markdown(
                    """
                    1. Open the workbook in **Microsoft Excel**.
                    2. Select the **Formulas** tab on the ribbon.
                    3. Select **Calculate Now**.
                    4. Save the workbook.
                    5. Return here and upload the saved workbook again.
                    """
                )

                st.caption(
                    "If the workbook contains many formulas or you "
                    "still receive a Limited coverage message, use "
                    "Ctrl + Alt + F9 in Excel to force a full "
                    "recalculation, then save the workbook."
                )

                st.markdown(
                    f"**Missing calculated results:** "
                    f"{audit_result.coverage.missing_cached_results}"
                )

                check_names = {
                    "EA01": "Broken Formula Reference",
                    "EA02": "Formula Pattern Inconsistency",
                    "EA03": "Hardcoded Formula Override",
                    "EA04": "Formula Error Detection",
                    "EA05": "Hidden Calculation Sheet",
                    "EA06": "Suspicious Blank Detection",
                    "EA07": "Inconsistent Data Type",
                    "EA08": "Duplicate Record Detection",
                }

                affected_check_names = [
                    (
                        f"{check_names.get(check_id, check_id)} "
                        f"({check_id})"
                    )
                    for check_id
                    in audit_result.coverage.affected_checks
                ]

                st.markdown(
                    "**Checks that may have reduced coverage:**"
                )

                if affected_check_names:

                    for check_name in affected_check_names:

                        st.markdown(
                            f"- {check_name}"
                        )

                else:

                    st.write(
                        "None"
                    )

        else:

            st.caption(
                "Complete coverage means the enabled audit checks "
                "had the information required to perform their "
                "assessment."
            )

        st.subheader(
            "Findings"
        )

        if audit_result.findings:

            priority_order = {
                "High Priority": 0,
                "Review Recommended": 1,
                "Information": 2,
            }

            sorted_findings = sorted(
                audit_result.findings,
                key=lambda finding: (
                    priority_order.get(
                        finding.priority,
                        99,
                    ),
                    finding.finding_id,
                ),
            )

            findings_data = []

            for finding in sorted_findings:

                findings_data.append(
                    {
                        "Finding ID": finding.finding_id,
                        "Check": finding.check_id,
                        "Issue": finding.check_name,
                        "Worksheet": finding.worksheet,
                        "Location": finding.location,
                        "Priority": finding.priority,
                        "Confidence": finding.confidence,
                    }
                )

            findings_dataframe = pd.DataFrame(
                findings_data
            )

            st.dataframe(
                findings_dataframe,
                use_container_width=True,
                hide_index=True,
            )

            st.subheader(
                "Finding Details"
            )

            finding_options = {
                (
                    f"{finding.finding_id} — "
                    f"{finding.check_name} — "
                    f"{finding.worksheet} — "
                    f"{finding.location}"
                ): finding
                for finding in sorted_findings
            }

            selected_finding_label = st.selectbox(
                "Select a finding to review",
                options=list(
                    finding_options.keys()
                ),
            )

            selected_finding = finding_options[
                selected_finding_label
            ]

            st.markdown(
                f"### {selected_finding.finding_id} · "
                f"{selected_finding.check_id} — "
                f"{selected_finding.check_name}"
            )

            detail_col1, detail_col2 = st.columns(
                [2, 1]
            )

            detail_col1.markdown(
                f"**Worksheet:** "
                f"{selected_finding.worksheet}"
            )

            detail_col1.markdown(
                f"**Location:** "
                f"{selected_finding.location}"
            )

            detail_col2.markdown(
                f"**Priority:** "
                f"{selected_finding.priority}"
            )

            detail_col2.markdown(
                f"**Confidence:** "
                f"{selected_finding.confidence}"
            )

            st.markdown(
                "**What was found**"
            )

            st.write(
                selected_finding.summary
            )

            st.markdown(
                "**Technical detail**"
            )

            st.write(
                selected_finding.detail
            )

            st.markdown(
                "**Why this matters**"
            )

            st.write(
                selected_finding.why_review
            )

            if selected_finding.affected_items:

                st.markdown(
                    "**Evidence**"
                )

                evidence_dataframe = pd.DataFrame(
                    selected_finding.affected_items
                )

                evidence_dataframe.columns = [
                    str(column)
                    .replace("_", " ")
                    .title()
                    for column
                    in evidence_dataframe.columns
                ]

                st.dataframe(
                    evidence_dataframe,
                    use_container_width=True,
                    hide_index=True,
                )

        else:

            st.info(
                "There are no findings to display."
            )

        st.divider()

        st.subheader(
            "Download Excel Audit Report"
        )

        st.write(
            "Download the full Excel audit report containing "
            "the audit summary, findings, finding details and "
            "audit information."
        )

        report_workbook = build_audit_report(
            audit_result
        )

        report_buffer = io.BytesIO()

        report_workbook.save(
            report_buffer
        )

        report_buffer.seek(0)

        original_file_name = (
            st.session_state.get(
                "audited_file_name",
                "workbook.xlsx",
            )
        )

        base_file_name = os.path.splitext(
            original_file_name
        )[0]

        report_file_name = (
            f"{base_file_name}_Audit_Report.xlsx"
        )

        st.download_button(
            label="Download Excel Audit Report",
            data=report_buffer.getvalue(),
            file_name=report_file_name,
            mime=(
                "application/vnd.openxmlformats-"
                "officedocument.spreadsheetml.sheet"
            ),
            type="primary",
        )