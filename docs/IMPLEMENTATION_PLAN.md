# Implementation Plan - Fully Automated PhD-Level Survey Analyzer (Excel-First + UDF Macros)

Date: 2026-02-05
Owner: Codex + User
Status: Draft v1

## Objectives
- Deliver a fully automated, zero-human workflow that produces PhD-level, academically defensible survey analysis.
- Use Excel formulas wherever possible, and Excel UDF macros for statistics not natively supported in Excel.
- Maintain a strict fail-closed verification loop: if any computation fails or is unverifiable, the workflow must halt or revise.
- Produce SPSS-grade outputs, including assumptions, effect sizes, and APA 7 reporting.

## Non-Negotiables
- No human-in-the-loop at any stage after file upload.
- No academic quality reduction, no shortcuts.
- All outputs must be reproducible and auditable.
- Excel workbook must contain formula-driven results; UDF macros allowed where Excel lacks native functions.

## Scope
### In Scope
- Quantitative EDA, inferential statistics, assumptions, effect sizes, APA-style results.
- Qualitative analysis via automated coding and frequency/co-occurrence outputs.
- Full audit trails and methodology documentation.
- Excel macro-enabled output (.xlsm) with UDFs.

### Out of Scope
- Manual review or human editing.
- Interactive UI beyond upload/status/download (can be added later).

## High-Level Architecture Changes
1. Deterministic computation engine for all statistics.
2. Excel formula generator for every output sheet.
3. Macro-enabled template workbook that includes UDFs for non-native tests.
4. Plan review gate before any execution.
5. Deterministic verification gate after each task and at final audit.

## Data Flow Overview
1. Upload Excel file.
2. Load and profile data.
3. Strategist produces structured, schema-validated task plan.
4. Plan review gate validates completeness and methodological correctness.
5. Implementer executes tasks using deterministic formula engine into Excel .xlsm.
6. QC verifies formulas and recomputes expected values to compare.
7. If mismatch or missing formulas, task is rejected and revised.
8. Auditor validates entire workbook and issues certification.

## Implementation Phases

### Phase 1: Schema and Planning Hardening
1. Add a structured task schema using Pydantic.
2. Update strategist to output strict JSON that matches the schema.
3. Add plan review node with deterministic checks.

Deliverables:
- `backend/schemas/plan.py`
- Updated `backend/agents/strategist.py`
- New `backend/graph/plan_review.py`
- Updated `backend/graph/workflow.py`

Acceptance Criteria:
- Plan output validates against schema.
- Plan review blocks if any required task category is missing.

### Phase 2: Excel Macro Template and UDFs
1. Create a macro-enabled template workbook `templates/analysis_template.xlsm`.
2. Add VBA module with UDFs for non-native tests.
3. Ensure openpyxl writes using `keep_vba=True` to preserve macros.

Required UDFs (initial set):
- `SHAPIRO_WILK(range)`
- `LEVENE_TEST(range1, range2, ...)`
- `CRONBACH_ALPHA(range)`
- `FISHER_Z(r)`
- `P_VALUE_T(t, df)`
- `P_VALUE_F(f, df1, df2)`

Deliverables:
- `templates/analysis_template.xlsm`
- `backend/tools/excel_udf_writer.py` (loader utilities)

Acceptance Criteria:
- Workbook saves as .xlsm with UDFs intact.
- UDFs callable inside Excel formulas.

### Phase 3: Deterministic Formula Engine
1. Implement formula generators by task type.
2. Ensure all numeric outputs are formulas, not literals.
3. Allow labels as text only in header/label columns.

Deliverables:
- `backend/tools/formula_engine.py`
- Updates to `backend/agents/implementer.py`

Acceptance Criteria:
- For each task, formulas reference `00_RAW_DATA_LOCKED`.
- No literal numeric values in output cells.

### Phase 4: Deterministic Verification
1. Build a verification layer that recomputes expected values using Python stats.
2. Compare Excel formula outputs to expected values within strict tolerance.
3. Fail closed if any mismatch.

Deliverables:
- `backend/tools/verification.py`
- Updates to `backend/agents/qc_reviewer.py`

Acceptance Criteria:
- Any mismatch forces REJECT and revision loop.
- Verification results logged in QC trail.

### Phase 5: Quantitative Pipeline Coverage
1. Data cleaning, recoding, missingness, and normalization.
2. Descriptives, distributions, outlier flags.
3. Reliability, correlations, group comparisons, effect sizes.
4. APA tables and interpretation text.

Deliverables:
- Expanded formula templates in `backend/tools/formula_engine.py`
- Reporting generators in `backend/tools/reporting.py`

Acceptance Criteria:
- SPSS-equivalent outputs for each test.
- APA 7 formatting correctness.

### Phase 6: Qualitative Pipeline Coverage
1. Automated coding and codebook creation.
2. Frequency and co-occurrence tables.
3. Inter-rater reliability using multiple model coders.

Deliverables:
- `backend/tools/qual_tools.py`
- Updates to `backend/agents/strategist.py` to schedule qual tasks

Acceptance Criteria:
- Codebook generated with traceable definitions.
- Reliability metrics produced without human input.

### Phase 7: Auditor Hardening
1. Use deterministic checks for formula coverage and result integrity.
2. Use LLM auditor only for narrative assessment.

Deliverables:
- Updates to `backend/agents/auditor.py`

Acceptance Criteria:
- Certification based on verified metrics, not model guess.

### Phase 8: Frontend and API Consistency
1. Show .xlsm outputs for download.
2. Report QC failures and revision counts clearly.

Deliverables:
- Updates to `frontend/src/App.jsx`
- Updates to `backend/main.py`

Acceptance Criteria:
- UI indicates macro-enabled output.
- Status includes verification results.

## Verification Strategy
- Formula coverage check on all output ranges.
- UDF availability check per workbook.
- Numerical verification against Python ground truth.
- Tolerance thresholds are test-specific and documented.

## Risk Register
- Macro execution blocked by client security policies.
- Excel UDFs not available in non-desktop environments.
- Task schema drift if strategist prompt changes.

Mitigations:
- Provide signed macro template when possible.
- Include a pure-Python validation report as backup evidence.
- Version-lock strategist output format.

## Success Metrics
- 100 percent formula-driven outputs (or UDF formulas where required).
- All tasks verified without mismatch.
- Final certification >= 97 percent.
- Full audit trail generated and reproducible.

## Execution Order
1. Implement schema and plan review gate.
2. Create macro-enabled template and UDFs.
3. Build formula engine.
4. Build deterministic verification.
5. Expand quant tasks.
6. Add qual pipeline.
7. Harden audit.
8. Update frontend and API.

## Notes
- This plan assumes all inputs are Excel files.
- Output will be macro-enabled Excel (.xlsm) to preserve UDFs.
- The system will halt on any validation failure and iterate until resolved.
