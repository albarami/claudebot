"""
Plan Review Node.
Deterministic validation of the master plan before execution begins.
"""

from typing import Dict, Any
from datetime import datetime

from graph.state import SurveyAnalysisState, LogEntry
from models.task_schema import MasterPlan, PlanValidationResult, validate_plan


def _coerce_plan_from_state(state: SurveyAnalysisState) -> MasterPlan:
    """Build MasterPlan from state (plan_json preferred)."""
    plan_json = state.get("plan_json") or {}
    if plan_json:
        return MasterPlan.model_validate(plan_json)

    return MasterPlan.model_validate({
        "session_id": state.get("session_id", "unknown"),
        "total_variables": state.get("n_cols", 0),
        "total_observations": state.get("n_rows", 0),
        "detected_scales": list(state.get("detected_scales", {}).keys()),
        "research_questions": state.get("research_questions", []),
        "tasks": state.get("tasks", [])
    })


def build_validation_report(result: PlanValidationResult) -> str:
    """Create a human-readable plan review report."""
    report_lines = [
        "=" * 60,
        "PLAN REVIEW GATE",
        "=" * 60,
        f"Total Tasks: {result.task_count}",
        "",
        "Phase Coverage:",
    ]

    for phase, count in result.phase_coverage.items():
        report_lines.append(f"  - {phase}: {count} tasks")

    if result.errors:
        report_lines.append("")
        report_lines.append("ERRORS (must fix):")
        for error in result.errors:
            report_lines.append(f"  ? {error}")

    if result.warnings:
        report_lines.append("")
        report_lines.append("WARNINGS:")
        for warning in result.warnings:
            report_lines.append(f"  ?? {warning}")

    report_lines.append("")
    report_lines.append(f"VERDICT: {'APPROVED' if result.is_valid else 'REJECTED'}")
    report_lines.append("=" * 60)

    return "\n".join(report_lines)


async def plan_review_node(state: SurveyAnalysisState) -> Dict[str, Any]:
    """
    Plan review gate node.
    Validates the master plan before execution can proceed.
    """
    columns = state.get("columns", [])

    try:
        plan = _coerce_plan_from_state(state)
        validation_result = validate_plan(plan, columns)
    except Exception as exc:
        validation_result = PlanValidationResult(
            is_valid=False,
            errors=[f"Plan validation error: {exc}"],
            warnings=[],
            task_count=0,
            phase_coverage={}
        )

    validation_report = build_validation_report(validation_result)
    print(validation_report)

    log_entry = LogEntry(
        timestamp=datetime.now().isoformat(),
        agent="plan_review",
        action="Plan validation",
        details=f"{'APPROVED' if validation_result.is_valid else 'REJECTED'} - {len(validation_result.errors)} errors, {len(validation_result.warnings)} warnings",
        task_id=None
    )

    new_revision_count = state.get("plan_revision_count", 0)
    if not validation_result.is_valid:
        new_revision_count += 1

    return {
        "master_plan_approved": validation_result.is_valid,
        "plan_revision_count": new_revision_count,
        "plan_errors": validation_result.errors,
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

    if state.get("master_plan_approved", False):
        return "implementer"

    if state.get("plan_revision_count", 0) >= MAX_PLAN_REVISIONS:
        return "halt"

    return "strategist"

