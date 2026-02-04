"""
LangGraph conditional edges for PhD Survey Analyzer.
Defines routing logic between nodes based on state.
"""

from typing import Literal
from graph.state import SurveyAnalysisState
from config import MAX_TASK_REVISIONS


def route_after_qc(state: SurveyAnalysisState) -> Literal["advance_task", "implementer", "auditor", "error"]:
    """
    Route after QC review based on decision.
    
    Args:
        state: Current workflow state
    
    Returns:
        Next node name
    """
    decision = state.get('qc_decision', '')
    current_idx = state.get('current_task_idx', 0)
    total_tasks = state.get('total_tasks', 0)
    revision_count = state.get('task_revision_count', 0)
    
    if decision == "HALT":
        return "error"
    
    if decision == "REJECT":
        if revision_count >= MAX_TASK_REVISIONS:
            return "error"
        return "implementer"
    
    if decision in ["APPROVE", "CONDITIONAL"]:
        if current_idx + 1 >= total_tasks:
            return "auditor"
        return "advance_task"
    
    return "advance_task"


def route_after_audit(state: SurveyAnalysisState) -> Literal["deliverables", "implementer"]:
    """
    Route after audit based on certification level.
    
    Args:
        state: Current workflow state
    
    Returns:
        Next node name
    """
    certification = state.get('certification', '')
    overall_score = state.get('overall_score', 0)
    
    if certification in ["PUBLICATION-READY", "THESIS-READY"] or overall_score >= 95:
        return "deliverables"
    
    return "deliverables"


def should_continue_tasks(state: SurveyAnalysisState) -> Literal["implementer", "auditor"]:
    """
    Check if there are more tasks to execute.
    
    Args:
        state: Current workflow state
    
    Returns:
        Next node name
    """
    current_idx = state.get('current_task_idx', 0)
    total_tasks = state.get('total_tasks', 0)
    
    if current_idx < total_tasks:
        return "implementer"
    return "auditor"
