# Implementation Status Review (Verified)

Date: 2026-02-05
Reviewer: Codex
Scope: Validate the implementation claims vs current codebase

## Verification Performed
- Ran backend unit tests: `python -m pytest -q`
  - Result: **33 passed**
  - Warning: `PytestDeprecationWarning` about `asyncio_default_fixture_loop_scope`
- Verified backend import: `python -c "import main"` (run from `backend/`)
  - Result: **import OK**

## Verified as True
- **Structured plan schema integrated** and strategist outputs JSON, with deterministic plan review gate.
- **Implementer uses deterministic formula engine** with macro-enabled `.xlsm` output.
- **No fixed template required**: macro workbook is created dynamically with UDF injection.
- **Cleaned + normalized sheets generated automatically** (`00_CLEANED_DATA`, `00_NORMALIZED_DATA`).
- **UDFs used in outputs** (Shapiro-Wilk, Cronbach alpha, Cohen's d, Levene, etc.).
- **Deterministic QC verification runs** with cleaned data alignment.
- **Audit routing fails closed** unless score >= 97 (publication-ready threshold).
- **State schema updated** for plan JSON, verification status, formula coverage, and audit revisions.

## Remaining Gaps (Not Blocking Core EDA)
1. **Qualitative pipeline not wired**: `qual_tools.py` exists but no workflow node integrates it.
2. **APA reporting not integrated**: `tools/reporting.py` exists but is not called in the workflow.
3. **Legacy formula engine (tools/formula_engine.py)** still exists and is unused (technical debt).

## Hardening Notes
- Macro trust enforcement now fails closed if UDFs cannot execute.
- Deterministic QC uses cleaned data to align with Excel formulas.

## Approval Status
**Approved for core quantitative EDA pipeline.**
Remaining items are advanced deliverables (qualitative + APA reporting) and can be scheduled next.

## Recommended Next Fixes (Optional Enhancements)
1. Integrate `qual_tools.py` into workflow for qualitative coding and reliability.
2. Wire `tools/reporting.py` to generate APA tables and narrative outputs.
3. Remove or consolidate `tools/formula_engine.py` to reduce duplication.
