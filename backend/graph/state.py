"""
LangGraph State Schema for PhD Survey Analyzer.
Maintains complete state across the multi-agent workflow.
"""

from typing import TypedDict, List, Dict, Any, Optional, Annotated
from operator import add


class Task(TypedDict):
    """Individual task from Master Plan."""
    id: str
    phase: str
    task_type: str
    name: str
    objective: str
    output_sheet: str
    columns: Dict[str, Any]
    group_by: Optional[str]
    scale_items: Optional[List[str]]
    status: str  # pending, in_progress, completed, failed


class QCDecision(TypedDict):
    """QC review decision record."""
    task_id: str
    decision: str  # APPROVE, REJECT, CONDITIONAL, HALT
    feedback: str
    checklist_results: Dict[str, bool]
    timestamp: str
    revision_number: int


class LogEntry(TypedDict):
    """Execution log entry."""
    timestamp: str
    agent: str
    action: str
    details: str
    task_id: Optional[str]


class SurveyAnalysisState(TypedDict):
    """
    Complete state for the survey analysis workflow.
    LangGraph maintains this across all nodes.
    """
    
    # === SESSION ===
    session_id: str
    status: str  # initializing, planning, executing, reviewing, auditing, completed, failed
    
    # === INPUT ===
    file_path: str
    file_name: str
    research_questions: List[str]
    
    # === RAW DATA INFO ===
    n_rows: int
    n_cols: int
    columns: List[str]
    column_types: Dict[str, str]
    numeric_columns: List[str]
    categorical_columns: List[str]
    detected_scales: Dict[str, List[str]]
    data_summary: str
    
    # === PLANNING ===
    master_plan: str
    plan_json: Dict[str, Any]
    master_plan_approved: bool
    plan_revision_count: int
    plan_errors: List[str]
    tasks: List[Task]
    total_tasks: int
    
    # === EXECUTION ===
    current_task_idx: int
    current_task: Optional[Task]
    current_task_output: str
    task_revision_count: int
    
    # === WORKBOOK ===
    workbook_path: str
    sheets_created: List[str]
    formulas_documented: Annotated[List[Dict[str, str]], add]
    
    # === QC REVIEW ===
    qc_decision: str
    qc_feedback: str
    qc_history: Annotated[List[QCDecision], add]
    
    # === AUDIT ===
    audit_complete: bool
    audit_result: str
    quality_scores: Dict[str, float]
    overall_score: float
    certification: str
    audit_revision_count: int

    # === VERIFICATION ===
    verification_status: str
    formula_coverage: float
    output_type: str
    
    # === OUTPUT ===
    deliverables: List[str]
    output_excel_path: str
    
    # === LOGGING ===
    execution_log: Annotated[List[LogEntry], add]
    errors: Annotated[List[str], add]
    
    # === MESSAGES (for agent communication) ===
    messages: Annotated[List[Dict[str, Any]], add]


def create_initial_state(session_id: str, file_path: str) -> SurveyAnalysisState:
    """Create initial state for a new analysis session."""
    return SurveyAnalysisState(
        session_id=session_id,
        status="initializing",
        file_path=file_path,
        file_name="",
        research_questions=[],
        n_rows=0,
        n_cols=0,
        columns=[],
        column_types={},
        numeric_columns=[],
        categorical_columns=[],
        detected_scales={},
        data_summary="",
        master_plan="",
        plan_json={},
        master_plan_approved=False,
        plan_revision_count=0,
        plan_errors=[],
        tasks=[],
        total_tasks=0,
        current_task_idx=0,
        current_task=None,
        current_task_output="",
        task_revision_count=0,
        workbook_path="",
        sheets_created=[],
        formulas_documented=[],
        qc_decision="",
        qc_feedback="",
        qc_history=[],
        audit_complete=False,
        audit_result="",
        quality_scores={},
        overall_score=0.0,
        certification="",
        audit_revision_count=0,
        verification_status="pending",
        formula_coverage=0.0,
        output_type="unknown",
        deliverables=[],
        output_excel_path="",
        execution_log=[],
        errors=[],
        messages=[]
    )
