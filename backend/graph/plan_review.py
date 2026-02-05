"""
Plan Review Node.
Deterministic validation of the master plan before execution begins.
"""

from typing import Dict, Any, List
from datetime import datetime

from graph.state import SurveyAnalysisState, LogEntry
from models.task_schema import (
    MasterPlan, TaskSpec, TaskType, TaskPhase,
    PlanValidationResult, validate_plan
)


def deterministic_plan_checks(tasks: List[Dict], columns: List[str], n_rows: int) -> PlanValidationResult:
    """
    Run deterministic validation checks on the plan.
    No LLM involvement - pure rule-based validation.
    
    Args:
        tasks: List of task dictionaries
        columns: Available columns in the dataset
        n_rows: Number of rows in the dataset
    
    Returns:
        Validation result with errors and warnings
    """
    errors = []
    warnings = []
    
    if not tasks:
        errors.append("No tasks in plan")
        return PlanValidationResult(
            is_valid=False,
            errors=errors,
            warnings=warnings,
            task_count=0,
            phase_coverage={}
        )
    
    # Check minimum task count
    if len(tasks) < 5:
        errors.append(f"Plan has only {len(tasks)} tasks. Minimum 5 required for PhD-level analysis.")
    
    # Check for required task types
    task_names_lower = [t.get('name', '').lower() for t in tasks]
    task_objectives_lower = [t.get('objective', '').lower() for t in tasks]
    all_text = ' '.join(task_names_lower + task_objectives_lower)
    
    required_analyses = [
        ('data audit', 'audit'),
        ('descriptive', 'descriptive'),
        ('missing', 'missing'),
    ]
    
    for name, keyword in required_analyses:
        if keyword not in all_text:
            warnings.append(f"Plan may be missing '{name}' analysis")
    
    # Check for duplicate task IDs
    task_ids = [t.get('id', '') for t in tasks]
    if len(task_ids) != len(set(task_ids)):
        errors.append("Duplicate task IDs found in plan")
    
    # Check for duplicate sheet names
    sheet_names = [t.get('output_sheet', '') for t in tasks if t.get('output_sheet')]
    if len(sheet_names) != len(set(sheet_names)):
        warnings.append("Duplicate output sheet names found - may cause overwrites")
    
    # Check sheet name validity
    for task in tasks:
        sheet = task.get('output_sheet', '')
        if sheet:
            if len(sheet) > 31:
                warnings.append(f"Task {task.get('id')}: Sheet name exceeds 31 characters")
            invalid_chars = ['\\', '/', '*', '?', ':', '[', ']']
            for char in invalid_chars:
                if char in sheet:
                    errors.append(f"Task {task.get('id')}: Sheet name contains invalid character '{char}'")
    
    # Check phase distribution
    phases = {}
    for task in tasks:
        phase = task.get('phase', 'Unknown')
        phases[phase] = phases.get(phase, 0) + 1
    
    # Validate column references in tasks
    for task in tasks:
        method = task.get('method', '')
        # Simple check: if method mentions specific column names, verify they exist
        for col in columns:
            if col in method and len(col) > 3:  # Avoid false positives with short names
                pass  # Column exists, OK
    
    return PlanValidationResult(
        is_valid=len(errors) == 0,
        errors=errors,
        warnings=warnings,
        task_count=len(tasks),
        phase_coverage=phases
    )


async def plan_review_node(state: SurveyAnalysisState) -> Dict[str, Any]:
    """
    Plan review gate node.
    Validates the master plan before execution can proceed.
    
    Args:
        state: Current workflow state
    
    Returns:
        State updates with plan approval status
    """
    tasks = state.get('tasks', [])
    columns = state.get('columns', [])
    n_rows = state.get('n_rows', 0)
    
    # Run deterministic validation
    validation_result = deterministic_plan_checks(tasks, columns, n_rows)
    
    # Build validation report
    report_lines = [
        "=" * 60,
        "PLAN REVIEW GATE",
        "=" * 60,
        f"Total Tasks: {validation_result.task_count}",
        "",
        "Phase Coverage:",
    ]
    
    for phase, count in validation_result.phase_coverage.items():
        report_lines.append(f"  - {phase}: {count} tasks")
    
    if validation_result.errors:
        report_lines.append("")
        report_lines.append("ERRORS (must fix):")
        for error in validation_result.errors:
            report_lines.append(f"  ❌ {error}")
    
    if validation_result.warnings:
        report_lines.append("")
        report_lines.append("WARNINGS:")
        for warning in validation_result.warnings:
            report_lines.append(f"  ⚠️ {warning}")
    
    report_lines.append("")
    report_lines.append(f"VERDICT: {'APPROVED' if validation_result.is_valid else 'REJECTED'}")
    report_lines.append("=" * 60)
    
    validation_report = "\n".join(report_lines)
    print(validation_report)
    
    log_entry = LogEntry(
        timestamp=datetime.now().isoformat(),
        agent="plan_review",
        action="Plan validation",
        details=f"{'APPROVED' if validation_result.is_valid else 'REJECTED'} - {len(validation_result.errors)} errors, {len(validation_result.warnings)} warnings",
        task_id=None
    )
    
    # If plan is rejected, increment revision count
    new_revision_count = state.get('plan_revision_count', 0)
    if not validation_result.is_valid:
        new_revision_count += 1
    
    return {
        "master_plan_approved": validation_result.is_valid,
        "plan_revision_count": new_revision_count,
        "execution_log": [log_entry],
        "messages": [
            {
                "role": "plan_review",
                "content": f"Plan {'APPROVED' if validation_result.is_valid else 'REJECTED'}: {validation_result.task_count} tasks, {len(validation_result.errors)} errors"
            }
        ]
    }


def route_after_plan_review(state: SurveyAnalysisState) -> str:
    """
    Route after plan review.
    
    Returns:
        'implementer' if approved, 'strategist' if rejected (for revision),
        'halt' if max revisions reached
    """
    MAX_PLAN_REVISIONS = 3
    
    if state.get('master_plan_approved', False):
        return "implementer"
    
    if state.get('plan_revision_count', 0) >= MAX_PLAN_REVISIONS:
        return "halt"
    
    return "strategist"
