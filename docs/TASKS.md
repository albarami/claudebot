# Implementation Tasks - PhD Survey Analyzer

## Phase 1: Project Setup

### Task 1.1: Initialize Project Structure
**Priority:** Critical
**Dependencies:** None

**Actions:**
1. Create directory structure:
   ```
   backend/
   ├── __init__.py
   ├── main.py              # FastAPI application
   ├── config.py            # Configuration management
   ├── graph/
   │   ├── __init__.py
   │   ├── state.py         # LangGraph state definition
   │   ├── nodes.py         # Agent node implementations
   │   ├── edges.py         # Conditional routing logic
   │   └── workflow.py      # Main workflow graph
   ├── agents/
   │   ├── __init__.py
   │   ├── strategist.py    # Agent 1: Master Plan creator
   │   ├── implementer.py   # Agent 2: Task executor
   │   ├── qc_reviewer.py   # Agent 3: Quality controller
   │   └── auditor.py       # Agent 4: Final certifier
   ├── tools/
   │   ├── __init__.py
   │   ├── excel_tools.py   # Excel formula writing tools
   │   ├── stats_tools.py   # Statistics calculation tools
   │   └── file_tools.py    # File management tools
   └── utils/
       ├── __init__.py
       ├── prompts.py       # System prompts for agents
       └── formatters.py    # Output formatting utilities
   ```

2. Create requirements.txt with exact versions
3. Create .env.example template

**Acceptance Criteria:**
- [ ] All directories created
- [ ] All __init__.py files present
- [ ] requirements.txt complete

---

### Task 1.2: Install Dependencies
**Priority:** Critical
**Dependencies:** Task 1.1

**Actions:**
1. Install core dependencies:
   ```
   langgraph>=0.2.0
   langchain>=0.3.0
   langchain-anthropic>=0.2.0
   anthropic>=0.40.0
   ```

2. Install data processing:
   ```
   pandas>=2.0.0
   numpy>=1.24.0
   scipy>=1.11.0
   openpyxl>=3.1.0
   ```

3. Install API framework:
   ```
   fastapi>=0.115.0
   uvicorn>=0.32.0
   python-multipart>=0.0.9
   python-dotenv>=1.0.0
   pydantic>=2.0.0
   aiofiles>=24.1.0
   ```

**Acceptance Criteria:**
- [ ] All packages install without errors
- [ ] Import test passes

---

## Phase 2: LangGraph State & Configuration

### Task 2.1: Define State Schema
**Priority:** Critical
**Dependencies:** Task 1.2

**File:** `backend/graph/state.py`

**State Definition:**
```python
class SurveyAnalysisState(TypedDict):
    # === INPUT ===
    file_path: str                    # Path to uploaded Excel
    session_id: str                   # Unique session identifier
    
    # === RAW DATA ===
    raw_dataframe: Optional[str]      # JSON serialized DataFrame
    n_rows: int                       # Number of rows
    n_cols: int                       # Number of columns
    column_info: Dict[str, Any]       # Column metadata
    
    # === PLANNING ===
    data_summary: str                 # Summary for strategist
    master_plan: str                  # Full plan text
    tasks: List[Dict]                 # Parsed task list
    current_task_idx: int             # Current task index
    total_tasks: int                  # Total task count
    
    # === EXECUTION ===
    workbook_path: str                # Path to output Excel
    sheets_created: List[str]         # List of sheet names
    current_task_output: str          # Current task result
    formulas_documented: List[Dict]   # Formula audit trail
    
    # === REVIEW ===
    qc_feedback: str                  # Latest QC feedback
    qc_decision: str                  # APPROVE/REJECT/CONDITIONAL
    revision_count: int               # Revisions for current task
    max_revisions: int                # Max before escalation (default: 5)
    qc_history: List[Dict]            # All QC decisions
    
    # === AUDIT ===
    audit_result: str                 # Full audit text
    quality_scores: Dict[str, float]  # Per-dimension scores
    overall_score: float              # Weighted overall
    certification: str                # PUBLICATION-READY, etc.
    
    # === OUTPUT ===
    deliverables: List[str]           # Generated file paths
    execution_log: List[Dict]         # Complete log
    errors: List[str]                 # Error messages
    status: str                       # RUNNING/COMPLETED/FAILED
```

**Acceptance Criteria:**
- [ ] State schema compiles
- [ ] All fields have correct types
- [ ] Serialization works for DataFrame

---

### Task 2.2: Create Agent Prompts
**Priority:** Critical
**Dependencies:** Task 2.1

**File:** `backend/utils/prompts.py`

**Prompts Required:**

1. **STRATEGIST_SYSTEM_PROMPT** (~2000 tokens)
   - Role: PhD Research Methodologist
   - Task: Create 40-60 task Master Plan
   - Output format: Structured markdown
   - Formula specification requirements
   - Quality gate definitions

2. **IMPLEMENTER_SYSTEM_PROMPT** (~1500 tokens)
   - Role: Statistical Programmer
   - Rules: ONE task, formulas only, document everything
   - Excel formula reference guide
   - APA formatting rules
   - Stop-and-wait protocol

3. **QC_REVIEWER_SYSTEM_PROMPT** (~1500 tokens)
   - Role: Quality Controller with VETO power
   - Full checklist (methodological, computational, documentation, formatting)
   - Decision criteria (APPROVE/REJECT/CONDITIONAL/HALT)
   - Feedback format specification

4. **AUDITOR_SYSTEM_PROMPT** (~1500 tokens)
   - Role: Senior Academic Reviewer
   - Scoring matrix (5 dimensions with weights)
   - Certification levels
   - Deliverable requirements

**Acceptance Criteria:**
- [ ] All 4 prompts written
- [ ] Prompts are comprehensive and specific
- [ ] Token counts reasonable

---

## Phase 3: Agent Implementation

### Task 3.1: Implement Strategist Agent
**Priority:** Critical
**Dependencies:** Task 2.2

**File:** `backend/agents/strategist.py`

**Function:** `create_master_plan(state: SurveyAnalysisState) -> dict`

**Logic:**
1. Read data summary from state
2. Construct prompt with survey details
3. Call Claude Opus 4.5
4. Parse response into task list
5. Return updated state with master_plan and tasks

**Output Format:**
```markdown
# MASTER PLAN

## Phase 1: Data Validation
### Task 1.1: Lock Raw Data
OBJECTIVE: Protect original data
METHOD: Copy to sheet, apply protection
FORMULA: N/A (data preservation)
VALIDATION: Sheet is read-only
OUTPUT: Sheet "00_RAW_DATA_LOCKED"
ACCEPTANCE: ☐ Data intact ☐ Protection enabled

### Task 1.2: Create Codebook
OBJECTIVE: Document all variables
METHOD: 
- N Valid: =COUNT(range)
- N Missing: =COUNTBLANK(range)
- % Missing: =COUNTBLANK(range)/ROWS(range)*100
...
```

**Acceptance Criteria:**
- [ ] Generates 40-60 tasks
- [ ] Every task has required fields
- [ ] Formulas specified for computational tasks

---

### Task 3.2: Implement Implementer Agent
**Priority:** Critical
**Dependencies:** Task 3.1

**File:** `backend/agents/implementer.py`

**Function:** `execute_task(state: SurveyAnalysisState) -> dict`

**Logic:**
1. Get current task from state.tasks[state.current_task_idx]
2. Construct prompt with task details
3. Call Claude Sonnet 4
4. Use tools to write Excel formulas
5. Document all formulas used
6. Return state with current_task_output

**Tools Available:**
- `write_formula(sheet, cell, formula)` - Write formula to cell
- `create_sheet(name)` - Create new worksheet
- `get_column_letter(col_name)` - Get Excel column letter
- `format_cell(sheet, cell, format_spec)` - Apply formatting

**Critical Rules:**
- NEVER write literal values (only formulas)
- ALWAYS document formulas in adjacent column
- ALWAYS reference raw data sheet for calculations

**Acceptance Criteria:**
- [ ] Executes one task at a time
- [ ] All outputs are formulas
- [ ] Complete documentation

---

### Task 3.3: Implement QC Reviewer Agent
**Priority:** Critical
**Dependencies:** Task 3.2

**File:** `backend/agents/qc_reviewer.py`

**Function:** `review_task(state: SurveyAnalysisState) -> dict`

**Logic:**
1. Get current task output from state
2. Get task specification from master plan
3. Construct review prompt
4. Call Claude Opus 4.5
5. Parse decision (APPROVE/REJECT/CONDITIONAL/HALT)
6. Return state with qc_decision and qc_feedback

**Checklist Verification:**
```python
checklist = {
    "methodological": [
        "Method matches plan specification",
        "Statistical test appropriate",
        "Sample size adequate",
        "Assumptions verified",
        "Effect sizes included"
    ],
    "computational": [
        "All cells contain formulas",
        "No hardcoded values",
        "Formulas reference correct ranges",
        "No Excel errors",
        "Results plausible"
    ],
    "documentation": [
        "Formula documentation complete",
        "Variable labels clear",
        "Notes explain decisions"
    ],
    "formatting": [
        "APA 7th edition compliance",
        "Proper decimal places",
        "Publication-ready appearance"
    ]
}
```

**Acceptance Criteria:**
- [ ] Full checklist verification
- [ ] Clear APPROVE/REJECT decision
- [ ] Specific feedback when rejecting

---

### Task 3.4: Implement Auditor Agent
**Priority:** Critical
**Dependencies:** Task 3.3

**File:** `backend/agents/auditor.py`

**Function:** `final_audit(state: SurveyAnalysisState) -> dict`

**Logic:**
1. Review entire workbook
2. Score each dimension (0-100)
3. Calculate weighted overall score
4. Determine certification level
5. Generate deliverables
6. Return state with certification

**Scoring:**
```python
weights = {
    "methodological_soundness": 0.30,
    "computational_accuracy": 0.25,
    "academic_standards": 0.25,
    "documentation_quality": 0.15,
    "reproducibility": 0.05
}

overall = sum(scores[k] * weights[k] for k in weights)

if overall >= 97:
    certification = "PUBLICATION-READY"
elif overall >= 95:
    certification = "THESIS-READY"
elif overall >= 90:
    certification = "NEEDS-REVISION"
else:
    certification = "MAJOR-ISSUES"
```

**Acceptance Criteria:**
- [ ] All 5 dimensions scored
- [ ] Correct weighted calculation
- [ ] Appropriate certification assigned

---

## Phase 4: LangGraph Workflow

### Task 4.1: Define Graph Nodes
**Priority:** Critical
**Dependencies:** Tasks 3.1-3.4

**File:** `backend/graph/nodes.py`

**Nodes:**
1. `load_data_node` - Load Excel, create summary
2. `strategist_node` - Create master plan
3. `plan_review_node` - Validate plan completeness
4. `implementer_node` - Execute current task
5. `qc_review_node` - Review task output
6. `advance_task_node` - Move to next task
7. `auditor_node` - Final certification
8. `generate_deliverables_node` - Create output files

**Acceptance Criteria:**
- [ ] All nodes implemented
- [ ] Each node returns valid state update

---

### Task 4.2: Define Conditional Edges
**Priority:** Critical
**Dependencies:** Task 4.1

**File:** `backend/graph/edges.py`

**Routing Functions:**

1. `route_after_plan_review(state) -> str`
   - If plan complete → "implementer"
   - If plan incomplete → "strategist"

2. `route_after_qc(state) -> str`
   - If APPROVE and more tasks → "advance_task"
   - If APPROVE and no more tasks → "auditor"
   - If REJECT → "implementer" (with feedback)
   - If HALT → "error_handler"

3. `route_after_audit(state) -> str`
   - If score >= 97 → "generate_deliverables"
   - If score < 97 → "revision_handler"

**Acceptance Criteria:**
- [ ] All routing logic correct
- [ ] No infinite loops possible
- [ ] Proper error handling

---

### Task 4.3: Assemble Workflow Graph
**Priority:** Critical
**Dependencies:** Tasks 4.1, 4.2

**File:** `backend/graph/workflow.py`

**Graph Structure:**
```python
from langgraph.graph import StateGraph, END

workflow = StateGraph(SurveyAnalysisState)

# Add nodes
workflow.add_node("load_data", load_data_node)
workflow.add_node("strategist", strategist_node)
workflow.add_node("plan_review", plan_review_node)
workflow.add_node("implementer", implementer_node)
workflow.add_node("qc_review", qc_review_node)
workflow.add_node("advance_task", advance_task_node)
workflow.add_node("auditor", auditor_node)
workflow.add_node("deliverables", generate_deliverables_node)

# Set entry point
workflow.set_entry_point("load_data")

# Add edges
workflow.add_edge("load_data", "strategist")
workflow.add_edge("strategist", "plan_review")
workflow.add_conditional_edges("plan_review", route_after_plan_review)
workflow.add_edge("implementer", "qc_review")
workflow.add_conditional_edges("qc_review", route_after_qc)
workflow.add_edge("advance_task", "implementer")
workflow.add_conditional_edges("auditor", route_after_audit)
workflow.add_edge("deliverables", END)

# Compile
app = workflow.compile()
```

**Acceptance Criteria:**
- [ ] Graph compiles without errors
- [ ] All paths lead to END or loop correctly
- [ ] Checkpointing enabled

---

## Phase 5: Excel Tools

### Task 5.1: Implement Excel Formula Writer
**Priority:** Critical
**Dependencies:** Task 1.2

**File:** `backend/tools/excel_tools.py`

**Functions:**

```python
def create_workbook() -> Workbook:
    """Create new workbook with formula calculation enabled."""

def create_sheet(wb: Workbook, name: str) -> Worksheet:
    """Create named worksheet with proper formatting."""

def write_formula(ws: Worksheet, cell: str, formula: str) -> None:
    """
    Write FORMULA to cell (not value).
    Example: write_formula(ws, "B2", "=AVERAGE(A2:A100)")
    """

def write_header_row(ws: Worksheet, headers: List[str]) -> None:
    """Write header row with formatting."""

def protect_sheet(ws: Worksheet, password: str = "locked") -> None:
    """Protect sheet to prevent accidental edits."""

def add_formula_documentation(ws: Worksheet, row: int, formula: str) -> None:
    """Add formula text in documentation column."""
```

**Critical Rule:** 
`write_formula` must ONLY accept formulas starting with "=".
Reject any attempt to write literal values.

**Acceptance Criteria:**
- [ ] All functions work correctly
- [ ] Formulas preserved (not evaluated)
- [ ] Protection works

---

### Task 5.2: Implement Statistics Calculator
**Priority:** Critical
**Dependencies:** Task 1.2

**File:** `backend/tools/stats_tools.py`

**Functions:**
```python
def analyze_dataframe(df: pd.DataFrame) -> Dict:
    """Create comprehensive data summary."""

def detect_variable_types(df: pd.DataFrame) -> Dict[str, str]:
    """Classify each column (numeric, categorical, ordinal)."""

def detect_scales(columns: List[str]) -> Dict[str, List[str]]:
    """Detect scale patterns from naming (Faith1, Faith2 -> Faith)."""

def calculate_descriptives(df: pd.DataFrame, col: str) -> Dict:
    """Calculate M, SD, skew, kurtosis, etc."""

def test_normality(df: pd.DataFrame, col: str) -> Dict:
    """Shapiro-Wilk test for normality."""

def calculate_reliability(df: pd.DataFrame, items: List[str]) -> float:
    """Cronbach's alpha for scale items."""

def calculate_correlation(df: pd.DataFrame, col1: str, col2: str) -> Dict:
    """Pearson r with p-value."""

def run_ttest(df: pd.DataFrame, dv: str, group: str) -> Dict:
    """Independent samples t-test with effect size."""

def run_anova(df: pd.DataFrame, dv: str, group: str) -> Dict:
    """One-way ANOVA with effect size."""
```

**Note:** These calculate VALUES that inform formula construction.
The actual Excel output uses FORMULAS, not these computed values.

**Acceptance Criteria:**
- [ ] All functions return correct statistics
- [ ] Effect sizes included
- [ ] P-values accurate

---

## Phase 6: API & Frontend

### Task 6.1: Implement FastAPI Backend
**Priority:** High
**Dependencies:** Task 4.3

**File:** `backend/main.py`

**Endpoints:**
```python
POST /api/upload          # Upload Excel file
POST /api/analyze         # Start analysis
GET  /api/status/{id}     # Get current status
GET  /api/logs/{id}       # Get execution logs
GET  /api/download/{id}   # Download results
```

**Acceptance Criteria:**
- [ ] All endpoints functional
- [ ] Proper error handling
- [ ] CORS configured

---

### Task 6.2: Implement React Frontend
**Priority:** High
**Dependencies:** Task 6.1

**Features:**
- File upload with drag-drop
- Real-time progress display
- Agent activity log
- Download results

**Acceptance Criteria:**
- [ ] Clean UI
- [ ] Real-time updates
- [ ] All functions work

---

## Phase 7: Testing & Validation

### Task 7.1: Unit Tests
**Priority:** High
**Dependencies:** All implementation tasks

**Test Coverage:**
- [ ] State serialization
- [ ] Agent responses
- [ ] Excel formula writing
- [ ] Statistics calculations
- [ ] Routing logic

---

### Task 7.2: Integration Test with Real Survey
**Priority:** Critical
**Dependencies:** Task 7.1

**Test Steps:**
1. Upload Wellbeing Questionnaire
2. Verify master plan has 40+ tasks
3. Verify each task executes
4. Verify QC catches errors
5. Verify final score ≥ 97%
6. Verify all deliverables generated

**Acceptance Criteria:**
- [ ] Full workflow completes
- [ ] Output is publication-ready
- [ ] All formulas verified

---

## Implementation Order

1. **Day 1:** Tasks 1.1, 1.2, 2.1, 2.2
2. **Day 2:** Tasks 3.1, 3.2, 3.3, 3.4
3. **Day 3:** Tasks 4.1, 4.2, 4.3
4. **Day 4:** Tasks 5.1, 5.2
5. **Day 5:** Tasks 6.1, 6.2
6. **Day 6:** Tasks 7.1, 7.2

---

## Success Metrics

| Metric | Target |
|--------|--------|
| Task completion | 100% |
| Formula-based cells | 100% |
| QC pass rate (first try) | ≥80% |
| Final audit score | ≥97% |
| Deliverables generated | All 7 |
| Processing time | <30 min |
