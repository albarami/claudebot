# PhD-Level Survey Analyzer - System Architecture

## Overview

A **LangGraph-based multi-agent system** that produces academically rigorous, publication-ready statistical analysis from any Excel survey file. The system prioritizes **accuracy over speed**, implementing unlimited revision loops until output meets doctoral dissertation standards.

---

## Core Principles

### 1. Accuracy First
- **No shortcuts** - Every statistical test verified
- **Unlimited iterations** - QC can reject indefinitely until perfect
- **Zero tolerance** for computational errors
- **Cost is irrelevant** - Quality is the only metric

### 2. Academic Rigor
- All analyses defensible in dissertation defense
- APA 7th edition compliance mandatory
- Effect sizes for every significant finding
- Assumptions tested and documented for every statistical test

### 3. Formula-Based Excel
- **ZERO hardcoded values** in output
- Every cell contains auditable Excel formula
- Complete reproducibility from formulas alone
- Any academic reviewer can verify calculations

### 4. Comprehensive Documentation
- Full audit trail of every decision
- Methodology section ready for thesis
- Limitations honestly assessed
- Every transformation documented

---

## LangGraph Workflow Architecture

```
                    ┌──────────────────────────────────────┐
                    │           START                       │
                    │    Load Survey Excel File             │
                    └──────────────────┬───────────────────┘
                                       │
                                       ▼
                    ┌──────────────────────────────────────┐
                    │      AGENT 1: STRATEGIST             │
                    │      (Claude Opus 4.5)               │
                    │                                      │
                    │  • Analyze survey structure          │
                    │  • Detect variables, scales, types   │
                    │  • Create 40-60 task Master Plan     │
                    │  • Specify exact Excel formulas      │
                    └──────────────────┬───────────────────┘
                                       │
                                       ▼
                    ┌──────────────────────────────────────┐
                    │      PLAN REVIEW GATE                │
                    │      (Claude Opus 4.5)               │
                    │                                      │
                    │  • Verify plan completeness          │
                    │  • Check all phases covered          │
                    │  • Validate formula specifications   │
                    └──────────────────┬───────────────────┘
                                       │
                         ┌─────────────┴─────────────┐
                         │                           │
                    ❌ REJECT                    ✅ APPROVE
                         │                           │
                         ▼                           ▼
                    Back to                    ┌─────────────────┐
                    Strategist                 │  TASK ITERATOR  │
                                               │  (Loop Start)   │
                                               └────────┬────────┘
                                                        │
                    ┌───────────────────────────────────┘
                    │
                    ▼
┌──────────────────────────────────────────────────────────────────┐
│                    TASK EXECUTION LOOP                            │
│                                                                   │
│   ┌─────────────────────────────────────┐                        │
│   │      AGENT 2: IMPLEMENTER           │                        │
│   │      (Claude Sonnet 4)              │                        │
│   │                                     │                        │
│   │  • Execute SINGLE task from plan    │                        │
│   │  • Use ONLY Excel formulas          │                        │
│   │  • Document every formula used      │                        │
│   │  • Generate sheet with formulas     │                        │
│   └──────────────────┬──────────────────┘                        │
│                      │                                            │
│                      ▼                                            │
│   ┌─────────────────────────────────────┐                        │
│   │      AGENT 3: QC REVIEWER           │                        │
│   │      (Claude Opus 4.5)              │                        │
│   │                                     │                        │
│   │  VERIFICATION CHECKLIST:            │                        │
│   │  ☐ Task matches plan specification  │                        │
│   │  ☐ ALL cells contain formulas       │                        │
│   │  ☐ No hardcoded values              │                        │
│   │  ☐ Statistical method appropriate   │                        │
│   │  ☐ Assumptions verified             │                        │
│   │  ☐ Results plausible                │                        │
│   │  ☐ APA formatting correct           │                        │
│   │  ☐ Documentation complete           │                        │
│   └──────────────────┬──────────────────┘                        │
│                      │                                            │
│          ┌───────────┴───────────┐                               │
│          │                       │                                │
│     ❌ REJECT                ✅ APPROVE                           │
│          │                       │                                │
│          ▼                       ▼                                │
│     ┌─────────┐           ┌─────────────┐                        │
│     │ REVISION│           │ More tasks? │                        │
│     │  LOOP   │           └──────┬──────┘                        │
│     │         │                  │                                │
│     │ • Log   │           ┌──────┴──────┐                        │
│     │   error │           │             │                        │
│     │ • Send  │          YES           NO                        │
│     │   back  │           │             │                        │
│     │   to    │           ▼             ▼                        │
│     │   Impl. │      Next Task     Exit Loop                     │
│     └────┬────┘                                                   │
│          │                                                        │
│          └──────► Back to Implementer (with feedback)            │
│                                                                   │
└──────────────────────────────────────────────────────────────────┘
                                       │
                                       ▼
                    ┌──────────────────────────────────────┐
                    │      AGENT 4: AUDITOR                │
                    │      (Claude Opus 4.5)               │
                    │                                      │
                    │  COMPREHENSIVE AUDIT:                │
                    │                                      │
                    │  A. Methodological Soundness (30%)   │
                    │     • Appropriate tests              │
                    │     • Assumptions verified           │
                    │     • Effect sizes reported          │
                    │                                      │
                    │  B. Computational Accuracy (25%)     │
                    │     • All formulas verified          │
                    │     • Zero hardcoded values          │
                    │     • Results reproducible           │
                    │                                      │
                    │  C. Academic Standards (25%)         │
                    │     • APA 7th edition                │
                    │     • Complete reporting             │
                    │     • Publication-ready              │
                    │                                      │
                    │  D. Documentation Quality (15%)      │
                    │     • Audit trail complete           │
                    │     • Methodology documented         │
                    │     • Limitations assessed           │
                    │                                      │
                    │  E. Reproducibility (5%)             │
                    │     • Another researcher can         │
                    │       replicate from formulas        │
                    └──────────────────┬───────────────────┘
                                       │
                         ┌─────────────┴─────────────┐
                         │                           │
                    Score < 97%               Score ≥ 97%
                         │                           │
                         ▼                           ▼
                    ┌─────────┐              ┌──────────────┐
                    │ REVISION│              │ CERTIFICATION│
                    │ REQUIRED│              │              │
                    │         │              │ 🏆 PUBLISH   │
                    │ Return  │              │    READY     │
                    │ to task │              │              │
                    │ loop    │              │ Generate all │
                    │ with    │              │ deliverables │
                    │ issues  │              └──────────────┘
                    └─────────┘
```

---

## Agent Specifications

### Agent 1: Survey Strategist
**Model:** Claude Sonnet 4.5 (claude-sonnet-4-20250514)
**Temperature:** 0.3 (analytical, not creative)
**Role:** Research Methodologist & Analysis Architect

**Responsibilities:**
1. Load and analyze survey structure
2. Identify variable types (demographic, Likert, continuous, categorical)
3. Detect scale patterns from naming conventions
4. Map variable relationships
5. Create comprehensive 40-60 task Master Plan
6. Specify exact Excel formulas for every computation

**Output:** MASTER_PLAN.md with detailed task specifications

**Quality Criteria:**
- Every task has unique ID
- Every task specifies exact formulas
- All 8 analysis phases covered
- Dependencies mapped
- Acceptance criteria defined per task

---

### Agent 2: Survey Implementer
**Model:** Claude Opus 4.5 (claude-opus-4-20250514)
**Temperature:** 0.1 (precision mode)
**Role:** Statistical Programmer & Excel Specialist (Highest Capability)

**Responsibilities:**
1. Execute ONE task at a time (never batch)
2. Use ONLY Excel formulas (never hardcode)
3. Create Excel sheet with formulas in cells
4. Document every formula
5. Submit for QC review
6. WAIT for approval before next task
7. If rejected, revise based on feedback

**Critical Rules:**
- `=AVERAGE(B2:B200)` ✅
- `34.5` ❌ (hardcoded)
- Must show formula, not just result
- All formulas reference raw data sheet

**Tools Available:**
- `write_excel_formula` - Write formula to cell
- `create_sheet` - Create new worksheet
- `read_data_range` - Read data for analysis
- `calculate_statistics` - Compute stats for formula construction

---

### Agent 3: QC Reviewer - DUAL REVIEW SYSTEM
**Models:** 
- **Review 1:** Claude Sonnet 4.5 (claude-sonnet-4-20250514)
- **Review 2:** OpenAI 5.2 (gpt-5.2)
**Temperature:** 0.2 (strict scrutiny)
**Role:** Statistical Reviewer with VETO POWER

**DUAL REVIEW LOGIC:**
- Both Sonnet 4.5 AND OpenAI 5.2 review every task
- BOTH must approve for task to pass
- If EITHER rejects → task is rejected → back to Implementer (Opus 4.5)
- Maximum error catching through two independent AI reviews

**Verification Checklist:**
```
METHODOLOGICAL:
☐ Method matches Master Plan specification
☐ Statistical test appropriate for data type
☐ Sample size adequate (n ≥ 30 for parametric)
☐ Assumptions tested (normality, homogeneity)
☐ Effect sizes calculated and interpreted

COMPUTATIONAL:
☐ ALL cells contain formulas (zero hardcoding)
☐ Formulas reference correct ranges
☐ No Excel errors (#DIV/0!, #N/A, #REF!)
☐ Results plausible (r between -1 and 1, etc.)
☐ Spot-check: manually verify 3 random values

DOCUMENTATION:
☐ Formula documentation complete
☐ Variable labels clear
☐ Assumptions stated
☐ Notes explain decisions

FORMATTING:
☐ APA 7th edition (M, SD, p, r, d)
☐ Proper decimal places
☐ Publication-ready tables
```

**Decisions:**
- ✅ **APPROVE** - All criteria met → proceed to next task
- ❌ **REJECT** - Issues found → return with specific feedback
- ⚠️ **CONDITIONAL** - Minor issues → approve with notes
- 🛑 **HALT** - Critical error → stop entire workflow

**VETO POWER:** QC can reject unlimited times until perfect

---

### Agent 4: Academic Auditor
**Model:** OpenAI 5.2 (gpt-5.2)
**Temperature:** 0.1 (maximum objectivity)
**Role:** Senior Methodologist & Publication Certifier

**Scoring Matrix:**
| Dimension | Weight | Criteria |
|-----------|--------|----------|
| Methodological Soundness | 30% | Appropriate tests, assumptions, effect sizes |
| Computational Accuracy | 25% | All formulas, zero errors, reproducible |
| Academic Standards | 25% | APA compliance, complete reporting |
| Documentation Quality | 15% | Audit trail, methodology, limitations |
| Reproducibility | 5% | Another researcher can replicate |

**Certification Levels:**
- 🏆 **PUBLICATION-READY** (≥97%): Suitable for peer-reviewed journals
- ✅ **THESIS-READY** (95-96.9%): Suitable for dissertation with minor notes
- ⚠️ **NEEDS REVISION** (90-94.9%): Specific improvements required
- ❌ **MAJOR ISSUES** (<90%): Return to task loop

---

## State Management

LangGraph maintains persistent state across the workflow:

```python
class SurveyAnalysisState(TypedDict):
    # Input
    file_path: str
    research_questions: List[str]
    
    # Data
    raw_data: pd.DataFrame
    clean_data: pd.DataFrame
    
    # Planning
    master_plan: str
    tasks: List[Task]
    current_task_index: int
    
    # Execution
    workbook: Workbook
    sheets_created: List[str]
    formulas_used: Dict[str, str]
    
    # Review
    qc_decisions: List[QCDecision]
    revision_count: Dict[str, int]  # Per task
    
    # Audit
    quality_scores: Dict[str, float]
    certification_level: str
    
    # Logs
    execution_log: List[LogEntry]
    error_log: List[ErrorEntry]
    
    # Output
    output_path: str
    deliverables: List[str]
```

---

## Excel Output Structure

### Sheet Naming Convention
```
00_RAW_DATA_LOCKED      - Original data (protected)
01_CODEBOOK             - Variable definitions with formulas
02_DATA_QUALITY         - Quality metrics with formulas
03_MISSING_ANALYSIS     - Missing data patterns
04_VALID_RESPONSES      - Filtered dataset
05_CLEAN_NUMERIC        - Numeric variables only
06_DESCRIPTIVES         - M, SD, skew, kurtosis (all formulas)
07_NORMALITY            - Shapiro-Wilk results
08_RELIABILITY          - Cronbach's alpha per scale
09_CORRELATIONS         - Correlation matrix (CORREL formulas)
10_GROUP_COMPARISONS    - T-tests, ANOVA (T.TEST formulas)
11_EFFECT_SIZES         - Cohen's d, eta-squared
12_REGRESSION           - If applicable
13_QC_REPORT            - Quality control summary
14_APA_RESULTS          - Publication-ready results section
15_METHODOLOGY          - Methods section for thesis
16_AUDIT_CERTIFICATE    - Final certification
17_EXECUTION_LOG        - Complete audit trail
```

### Formula Documentation Standard
Every sheet includes a "Formula" column showing the exact formula used:

| Variable | N | M | SD | Formula |
|----------|---|---|-----|---------|
| Age | =COUNT(range) | =AVERAGE(range) | =STDEV.S(range) | AVERAGE/STDEV.S('00_RAW'!B:B) |

---

## Deliverables

1. **Excel Workbook** (17+ sheets, all formula-based)
2. **MASTER_PLAN.md** - 40-60 detailed tasks
3. **QC_AUDIT_TRAIL.md** - Every review decision
4. **AUDIT_CERTIFICATE.md** - Final quality scores
5. **METHODOLOGY_DOCUMENTATION.md** - Ready for thesis Chapter 3
6. **LIMITATIONS_ASSESSMENT.md** - Honest evaluation
7. **EXECUTION_LOG.md** - Complete task log

---

## Quality Guarantees

### Computational Accuracy: 100%
- Every formula verified by QC
- Spot-checks against manual calculation
- No Excel errors allowed

### Methodological Rigor: ≥95%
- Appropriate tests for data types
- All assumptions tested
- Effect sizes for significant findings

### Academic Standards: ≥98%
- APA 7th edition throughout
- Complete statistical reporting
- Publication-ready tables

### Reproducibility: 100%
- Any researcher can verify from formulas
- Raw data preserved and locked
- Every transformation documented

---

## Technology Stack

- **Orchestration:** LangGraph
- **LLM Provider:** Anthropic Claude API
- **Models:** Claude Opus 4.5, Claude Sonnet 4
- **Excel Generation:** openpyxl (formula mode)
- **Statistics:** scipy, pandas, numpy
- **Backend:** FastAPI
- **Frontend:** React + TailwindCSS

---

## Success Criteria

The system succeeds when:
1. ✅ Output passes rigorous academic audit (≥97%)
2. ✅ All computations are formula-based (zero hardcoding)
3. ✅ Results are reproducible from formulas alone
4. ✅ APA 7th edition compliance throughout
5. ✅ Complete documentation for thesis submission
6. ✅ Any survey Excel file produces publication-ready output
