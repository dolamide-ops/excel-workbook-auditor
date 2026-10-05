# Excel Workbook Auditor

**Audit Excel workbooks for formula, structure and data-quality issues before they affect your analysis or decisions.**

The Excel Workbook Auditor is a lightweight Streamlit application designed to help users investigate whether an Excel workbook can be trusted before relying on it for analysis, reporting or business decisions.

Rather than automatically declaring a workbook "right" or "wrong", the auditor identifies anomalies worth reviewing and provides the context and evidence needed to investigate them.

> **A finding means "take a closer look" — not necessarily "this workbook is wrong."**

---

## Why I Built It

Excel workbooks often sit behind important business decisions.

A workbook can look perfectly normal while containing:

- broken formula references
- formulas that behave differently from surrounding formulas
- hardcoded values replacing expected calculations
- formula errors
- hidden calculation sheets
- suspicious blanks
- inconsistent data types
- duplicate records

The challenge isn't simply finding differences. It's identifying the differences that are actually worth investigating.

That became the central design principle behind this project.

---

## What the Auditor Checks

The current V1 includes eight audit checks:

| Check | Description |
|---|---|
| **EA01 — Broken Formula Reference** | Detects formulas containing invalid `#REF!` references |
| **EA02 — Formula Pattern Inconsistency** | Identifies isolated formulas that differ from the surrounding calculation pattern |
| **EA03 — Hardcoded Formula Override** | Identifies values that appear to replace an expected formula |
| **EA04 — Formula Error** | Detects Excel formula error results and groups related errors |
| **EA05 — Hidden Calculation Sheet** | Identifies hidden worksheets containing calculations that influence visible workbook outputs |
| **EA06 — Suspicious Blank** | Identifies unexpected blanks within populated data regions |
| **EA07 — Inconsistent Data Type** | Identifies values that differ from the dominant data type in a column |
| **EA08 — Duplicate Record** | Identifies duplicate records within detected tabular data |

The checks are intentionally conservative. The goal is to surface useful review signals rather than maximise the number of findings.

---

## How It Works

1. Upload an `.xlsx` workbook.
2. Run the audit.
3. Review prioritised findings.
4. Investigate the location, explanation and supporting evidence.
5. Review the audit coverage status.
6. Download a structured Excel audit report.

The uploaded workbook is analysed using a temporary copy. **The original workbook is not modified.**

---

## Explainable Findings

Each finding includes:

- the audit check that detected it
- worksheet and location
- priority
- confidence
- what was found
- technical detail
- why the finding may matter
- supporting evidence

The aim is not just detection, but **explainable detection**.

---

## Audit Coverage

Some Excel checks depend on cached formula results stored when a workbook is calculated and saved by Excel.

Because Python libraries such as `openpyxl` read formulas but do not calculate them, the auditor makes this limitation visible.

The audit therefore reports one of two coverage states:

**Complete** — the enabled checks had the information required to perform their assessment.

**Limited** — the workbook was still audited, but one or more checks may have reduced coverage because required cached formula results were unavailable.

This prevents the application from presenting an incomplete assessment as though it were complete.

---

## Excel Audit Report

Audit results can be downloaded as a structured Excel report containing four worksheets:

- **Audit Summary**
- **Findings**
- **Finding Detail**
- **Audit Information**

The report is designed to support investigation, review and conversations about workbook risk — not simply produce another spreadsheet.

---

## Validation

The auditor was developed against a controlled synthetic workbook containing deliberately seeded issues.

The expected findings were defined before the audit was run.

**Expected findings:** 9  
**Detected findings:** 9

A clean baseline version of the same workbook produces:

**Detected findings:** 0

The project currently includes:

- **54 automated tests**
- **10 release acceptance tests**
- controlled expected-findings validation
- clean-baseline validation
- compatibility testing
- limited-coverage testing
- real-world workbook hardening
- performance testing

---

## Real-World Hardening

Testing against a more complex workbook exposed an important problem.

The first audit produced **314 findings**.

Although many were technically explainable, the result contained far too much noise to be useful.

The generic audit rules were subsequently refined to:

- focus formula-pattern detection on isolated anomalies
- group related formula errors
- improve interpretation of data regions and headers

The same workbook then produced **4 reviewable findings** — approximately a **98.7% reduction in noise**.

The rules were improved generically rather than being configured to ignore that specific workbook.

This reinforced one of the main lessons from the project:

> **The goal isn't to find more. It's to find what is worth reviewing.**

---

## Performance Testing

Release testing included synthetic workbooks of increasing size.

The largest contained:

- **100,000 data rows**
- **200,003 formula cells**

The auditor successfully completed the assessment.

Performance depends on workbook size, structure and formula complexity, so this should not be interpreted as a formal maximum supported workbook size.

---

## Technology

The V1 application uses:

- **Python**
- **Streamlit**
- **openpyxl**
- **pytest**
- **Git / GitHub**

---

## AI-Assisted Development

ChatGPT was used as a development partner throughout the project, including support with:

- implementation
- debugging
- automated testing
- refactoring
- exploring technical approaches

I retained ownership of:

- problem framing
- requirements
- audit rules
- acceptance criteria
- validation
- product decisions

The project was deliberately approached from a Business Analysis perspective: define the problem and expected behaviour first, then use technology and AI to help implement and test the solution.

---

## Try the Auditor

The Excel Workbook Auditor is available as a live Streamlit application:

**Live application:** https://excel-workbook-auditor.streamlit.app

Sample synthetic workbooks are included in the `data` directory for testing.

---

## Run Locally

Clone the repository and install the production dependencies:

```bash
pip install -r requirements.txt
```

Then start the application:

```bash
streamlit run app.py
```

---

## V1 Scope and Limitations

This is a V1 diagnostic tool rather than a replacement for manual workbook review or Excel itself.

Important limitations include:

- `.xlsx` workbooks only
- formula calculations are not performed by the application
- some checks depend on cached Excel formula results
- complex workbook structures may receive reduced audit coverage
- findings indicate items worth investigating rather than automatically proving an error exists
- performance varies with workbook size, structure and formula complexity

The application reports relevant coverage limitations rather than hiding them.

---

## Project Status

**V1 complete**

Core audit engine, explainable findings, audit coverage, compatibility handling, Excel reporting, automated testing and release acceptance testing are complete.

---

## Author

**Deji Olamide**

Business Analyst exploring practical applications of AI, automation and data quality in business analysis.

---

*Built as a practical exploration of a simple question: before trusting the analysis, how do you know you can trust the workbook producing it?*