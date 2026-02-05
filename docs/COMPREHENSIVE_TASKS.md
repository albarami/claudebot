# Comprehensive Implementation Tasks - Excel-First + UDF Macros

Date: 2026-02-05
Scope: Fully automated, PhD-level survey analysis with Excel formulas + macro UDFs
Output: Macro-enabled Excel (.xlsm) with full audit trail

## Principles (Non-Negotiable)
- Zero human involvement after file upload.
- Excel-first for all supported stats; UDF macros for non-native tests.
- Fail-closed verification: any mismatch triggers revision or halt.
- Academic rigor: SPSS-equivalent outputs, APA 7 formatting, defensible methodology.

## Phase 0: Project Hygiene and Baseline
### Task 0.1: Create/confirm folders
- Create: `templates/`, `backend/schemas/`, `backend/tests/`, `backend/tools/udf/`
- Acceptance: folders exist and are referenced in code.

### Task 0.2: Update documentation index
- Add links in `docs/ARCHITECTURE.md` and `docs/IMPLEMENTATION_PLAN.md` to this file.
- Acceptance: docs cross-reference tasks and plan.

### Task 0.3: Dependency pinning and environment checks
- Add lock strategy for Python deps and minimum Python version.
- Acceptance: repeatable installs with pinned versions.

### Task 0.4: File safety baseline
- Validate file size, extension, and Excel structure before processing.
- Reject files with malformed sheets or unsupported formats.
- Acceptance: invalid inputs fail fast with clear error.

## Phase 1: Structured Planning + Plan Review Gate
### Task 1.1: Define plan/task schema (Pydantic)
- File: `backend/schemas/plan.py`
- Define Task, Plan, acceptance criteria, required phases.
- Acceptance: schema validates strategist output.

### Task 1.2: Update strategist to emit JSON
- File: `backend/agents/strategist.py`
- Replace regex parsing with strict JSON parsing.
- Acceptance: strategist output parses into `Plan` with 40-60 tasks.

### Task 1.3: Add plan review node
- Files: `backend/graph/plan_review.py`, `backend/graph/workflow.py`
- Deterministic checks for missing phases, missing formulas, invalid outputs.
- Acceptance: workflow blocks execution on invalid plan.

### Task 1.4: Update state to store structured plan
- File: `backend/graph/state.py`
- Store `plan_json`, `plan_valid`, `plan_errors`.
- Acceptance: state includes and preserves plan metadata.

### Task 1.5: Plan completeness map
- Require explicit coverage of each required sheet and statistical family.
- Acceptance: missing sheet or test family fails plan review.

## Phase 2: Excel Macro Template + UDFs
### Task 2.1: Create macro template
- File: `templates/analysis_template.xlsm`
- Contains VBA project with UDFs.
- Acceptance: UDFs callable in Excel formulas.

### Task 2.2: Add VBA module source to repo
- File: `backend/tools/udf/analysis_udf.bas`
- Include UDF implementations:
- `SHAPIRO_WILK(range)`
- `LEVENE_TEST(range1, range2, ...)`
- `CRONBACH_ALPHA(range)`
- `FISHER_Z(r)`
- `P_VALUE_T(t, df)`
- `P_VALUE_F(f, df1, df2)`
- Acceptance: VBA module matches template UDFs.

### Task 2.3: Template loader utility
- File: `backend/tools/excel_template.py`
- Load `.xlsm` with `keep_vba=True` and write sheets without losing macros.
- Acceptance: output workbook preserves macros and UDFs.

### Task 2.4: Macro trust and execution checks
- Detect if macros are disabled at runtime and fail closed with guidance.
- Acceptance: system refuses to run if UDFs cannot execute.

### Task 2.5: UDF test harness
- Create a minimal Excel test sheet to validate UDF results.
- Acceptance: UDF outputs match Python ground truth within tolerance.

## Phase 3: Deterministic Formula Engine
### Task 3.1: Create formula engine
- File: `backend/tools/formula_engine.py`
- Deterministic formula builders by task type.
- Acceptance: all data cells use formulas referencing `00_RAW_DATA_LOCKED`.

### Task 3.2: Implement Excel-safe naming and ranges
- File: `backend/tools/formula_engine.py`
- Enforce range safety, column mapping, and row bounds.
- Acceptance: no invalid Excel ranges or sheet names.

### Task 3.3: Replace heuristic implementer
- File: `backend/agents/implementer.py`
- Use formula engine with deterministic task types.
- Acceptance: execution is not keyword-based; it follows plan schema.

### Task 3.4: Formula-only enforcement rules
- Allow text labels only in header/label columns.
- Numeric cells must be formulas or UDF formulas.
- Acceptance: numeric literals rejected in output ranges.

## Phase 4: Deterministic Verification Gate
### Task 4.1: Create verification module
- File: `backend/tools/verification.py`
- Compute ground-truth stats (Python) and compare to Excel output.
- Acceptance: verification returns PASS/FAIL with discrepancies.

### Task 4.2: Integrate verification into QC
- File: `backend/agents/qc_reviewer.py`
- Use deterministic verification before LLM review.
- Acceptance: QC rejects on any numeric mismatch.

### Task 4.3: Excel formula coverage validation
- File: `backend/tools/verification.py`
- Validate that numeric output cells are formulas.
- Acceptance: non-formula numeric cells cause rejection.

### Task 4.4: Tolerance policy and rounding alignment
- Define per-test tolerances and rounding to match Excel behavior.
- Acceptance: verification uses consistent rounding rules.

## Phase 5: Quantitative Pipeline (SPSS-Grade)
### Task 5.1: Data cleaning and preparation
- Tasks include: missingness, valid response filtering, recoding, normalization.
- Output sheets: `02_DATA_QUALITY`, `03_MISSING_ANALYSIS`, `04_VALID_RESPONSES`, `05_CLEAN_NUMERIC`.
- Acceptance: cleaning rules logged and reproducible.

### Task 5.2: Descriptives and distributions
- Sheets: `06_DESCRIPTIVES`, `07_NORMALITY`.
- Include skew/kurtosis, CI, and UDF Shapiro-Wilk where required.
- Acceptance: outputs match Python ground truth.

### Task 5.3: Reliability and correlations
- Sheets: `08_RELIABILITY`, `09_CORRELATIONS`.
- Cronbach’s alpha via UDF where needed.
- Acceptance: alpha and r values verified.

### Task 5.4: Group comparisons and effect sizes
- Sheets: `10_GROUP_COMPARISONS`, `11_EFFECT_SIZES`.
- T-tests, ANOVA, chi-square, effect sizes.
- Acceptance: effect sizes and p-values verified.

### Task 5.5: APA tables and interpretations
- Sheets: `12_APA_RESULTS`, `13_METHODOLOGY`.
- APA 7 formatting and narrative interpretation.
- Acceptance: APA formatting validated.

### Task 5.6: Multiple comparisons correction
- Add Bonferroni or FDR corrections where applicable.
- Acceptance: corrected p-values reported and verified.

### Task 5.7: Assumption diagnostics sheet
- Report homogeneity, normality, and variance checks.
- Acceptance: diagnostics are formula or UDF driven.

## Phase 6: Qualitative Pipeline (Zero-Human)
### Task 6.1: Automated codebook creation
- File: `backend/tools/qual_tools.py`
- Generate themes, codes, and definitions.
- Acceptance: codebook sheet produced with traceable logic.

### Task 6.2: Coding + reliability
- Use multiple model coders and compute kappa.
- Acceptance: inter-rater reliability reported with method documented.

### Task 6.3: Qualitative preprocessing
- Detect text columns, normalize language, remove noise.
- Acceptance: preprocessing pipeline documented in output.

## Phase 7: Auditor Hardening
### Task 7.1: Deterministic audit checks
- File: `backend/agents/auditor.py`
- Use verification metrics, not just LLM narrative.
- Acceptance: audit score derived from verified checks.

### Task 7.2: Certification gate
- Update routing to stop or loop when audit < threshold.
- Acceptance: `route_after_audit` no longer auto-approves.

### Task 7.3: Tamper-evident audit hashes
- Generate hashes of sheets and formula logs.
- Acceptance: audit report includes integrity hashes.

## Phase 8: API + Frontend Alignment
### Task 8.1: Backend download for `.xlsm`
- File: `backend/main.py`
- Ensure downloads return macro-enabled workbook.
- Acceptance: downloaded file opens with UDFs intact.

### Task 8.2: Frontend status clarity
- File: `frontend/src/App.jsx`
- Show macro-enabled output and verification pass/fail.
- Acceptance: UI shows verification errors explicitly.

### Task 8.3: Fail-closed API error responses
- Ensure API returns clear error codes and halts on validation failure.
- Acceptance: no partial outputs served on failure.

## Phase 9: Testing
### Task 9.1: Unit tests
- Files: `backend/tests/test_formula_engine.py`, `backend/tests/test_verification.py`
- Acceptance: tests pass locally.

### Task 9.2: Integration test
- Use `data/` workbook as fixture.
- Validate output `.xlsm` with at least 5 sheets and passing verification.
- Acceptance: full workflow completes with score >= 97.

### Task 9.3: Frontend smoke test
- Commands: `npm run build`, `npm run preview` (optional)
- Acceptance: UI loads and can upload.

### Task 9.4: UDF regression test
- Execute UDF test harness workbook and validate outputs.
- Acceptance: UDF results match ground truth within tolerance.

## Phase 10: Operational Hardening
### Task 10.1: Concurrency and session isolation
- Ensure per-session output isolation and no cross-session leakage.
- Acceptance: concurrent sessions produce isolated outputs.

### Task 10.2: Resource limits and timeouts
- Add timeouts per task and maximum revision safeguards.
- Acceptance: runaway tasks fail safely and log reasons.

### Task 10.3: Sensitive data handling
- Strip PII from logs and avoid exposing raw data in logs.
- Acceptance: logs contain no respondent-level values.

## Git Commit Plan
Commit after each phase to keep history clean:
1. `git commit -m "Add plan schema and plan review gate"`
2. `git commit -m "Add Excel macro template and UDFs"`
3. `git commit -m "Add deterministic formula engine"`
4. `git commit -m "Add verification gate and QC integration"`
5. `git commit -m "Complete quantitative pipeline"`
6. `git commit -m "Add qualitative pipeline"`
7. `git commit -m "Harden auditor and certification routing"`
8. `git commit -m "Align API and frontend for .xlsm outputs"`
9. `git commit -m "Add unit and integration tests"`
10. `git commit -m "Add operational hardening"`

## Testing Commands
- Backend unit tests: `python -m pytest -q`
- Backend lint (optional): `python -m ruff check backend`
- Frontend build: `npm run build`
- Full run (manual): start backend, upload Excel, verify `.xlsm` output

## Definition of Done
- All tasks implemented with deterministic validation.
- All outputs formula-driven or UDF-driven.
- End-to-end run completes with PASS verification and certification.
- Audit and QC logs generated for every run.
