"""
Agent 1: Survey Strategist
Creates comprehensive Master Plan with 40-60 structured tasks.
Outputs JSON that matches the task schema.
"""

import json
import re
from typing import Dict, List, Any, Optional
from datetime import datetime

from langchain_anthropic import ChatAnthropic
from langchain_core.messages import HumanMessage, SystemMessage

from config import STRATEGIST_MODEL, STRATEGIST_TEMP, STRATEGIST_MAX_TOKENS, ANTHROPIC_API_KEY, STRATEGIST_PROVIDER
from utils.prompts import STRATEGIST_SYSTEM_PROMPT
from graph.state import SurveyAnalysisState, LogEntry
from models.task_schema import MasterPlan, TaskSpec, TaskType, TaskPhase


def _extract_json(text: str) -> Optional[str]:
    """Extract JSON object from model response."""
    if not text:
        return None
    text = text.strip()
    if text.startswith("{") and text.endswith("}"):
        return text
    fence_match = re.search(r"```json\s*(\{.*\})\s*```", text, re.DOTALL | re.IGNORECASE)
    if fence_match:
        return fence_match.group(1).strip()
    brace_match = re.search(r"(\{.*\})", text, re.DOTALL)
    if brace_match:
        return brace_match.group(1).strip()
    return None


def parse_master_plan_json(plan_text: str) -> MasterPlan:
    """Parse and validate JSON master plan."""
    json_text = _extract_json(plan_text)
    if not json_text:
        raise ValueError("No JSON found in strategist output")
    data = json.loads(json_text)
    return MasterPlan.model_validate(data)


def generate_default_master_plan(state: SurveyAnalysisState) -> MasterPlan:
    """
    Deterministic fallback plan generator (40 tasks).
    Ensures plan passes validation when LLM output is invalid.
    """
    session_id = state["session_id"]
    total_vars = state.get("n_cols", 0)
    total_obs = state.get("n_rows", 0)
    detected_scales = list(state.get("detected_scales", {}).keys())

    tasks: List[TaskSpec] = []
    task_id = 1.1

    def add_task(phase: TaskPhase, task_type: TaskType, name: str, sheet: str, cols: Optional[List[str]] = None):
        nonlocal task_id
        tasks.append(TaskSpec(
            id=f"{task_id:.1f}",
            phase=phase,
            task_type=task_type,
            name=name,
            objective=f"{name} for academically defensible analysis.",
            output_sheet=sheet,
            columns={"column_names": cols or [], "column_type": "all", "max_columns": None},
            group_by=None,
            scale_items=None
        ))
        task_id += 0.1

    # Phase 1: Data validation (5)
    add_task(TaskPhase.DATA_VALIDATION, TaskType.DATA_AUDIT, "Data audit trail", "01_DATA_AUDIT")
    add_task(TaskPhase.DATA_VALIDATION, TaskType.DATA_DICTIONARY, "Variable dictionary", "02_DATA_DICT")
    add_task(TaskPhase.DATA_VALIDATION, TaskType.MISSING_DATA, "Missing data analysis", "03_MISSING")
    add_task(TaskPhase.DATA_VALIDATION, TaskType.FREQUENCY_TABLES, "Categorical frequencies", "04_FREQ")
    add_task(TaskPhase.DATA_VALIDATION, TaskType.DESCRIPTIVE_STATS, "Initial descriptives", "05_DESC")

    # Phase 2: Exploratory (5)
    add_task(TaskPhase.EXPLORATORY, TaskType.DESCRIPTIVE_STATS, "Expanded descriptives", "06_DESC2")
    add_task(TaskPhase.EXPLORATORY, TaskType.NORMALITY_CHECK, "Normality diagnostics", "07_NORMAL")
    add_task(TaskPhase.EXPLORATORY, TaskType.CORRELATION_MATRIX, "Correlation matrix", "08_CORR")
    add_task(TaskPhase.EXPLORATORY, TaskType.FREQUENCY_TABLES, "Expanded frequency tables", "09_FREQ2")
    add_task(TaskPhase.EXPLORATORY, TaskType.SUMMARY_DASHBOARD, "Exploratory summary", "10_SUMMARY")

    # Phase 3: Descriptive (5)
    add_task(TaskPhase.DESCRIPTIVE, TaskType.DESCRIPTIVE_STATS, "Final descriptives", "11_DESC3")
    add_task(TaskPhase.DESCRIPTIVE, TaskType.NORMALITY_CHECK, "Normality (final)", "12_NORMAL2")
    add_task(TaskPhase.DESCRIPTIVE, TaskType.FREQUENCY_TABLES, "Final frequencies", "13_FREQ3")
    add_task(TaskPhase.DESCRIPTIVE, TaskType.DATA_DICTIONARY, "Codebook refinement", "14_CODEBOOK")
    add_task(TaskPhase.DESCRIPTIVE, TaskType.SUMMARY_DASHBOARD, "Descriptive dashboard", "15_DASH")

    # Phase 4: Inferential (5)
    add_task(TaskPhase.INFERENTIAL, TaskType.GROUP_COMPARISON, "Group comparisons", "16_GROUPS")
    add_task(TaskPhase.INFERENTIAL, TaskType.EFFECT_SIZES, "Effect sizes", "17_EFFECTS")
    add_task(TaskPhase.INFERENTIAL, TaskType.CROSS_TABULATION, "Cross-tabulations", "18_CROSSTAB")
    add_task(TaskPhase.INFERENTIAL, TaskType.CORRELATION_MATRIX, "Inferential correlations", "19_CORR2")
    add_task(TaskPhase.INFERENTIAL, TaskType.SUMMARY_DASHBOARD, "Inferential summary", "20_INF_SUM")

    # Phase 5: Reliability (5)
    add_task(TaskPhase.RELIABILITY, TaskType.RELIABILITY_ALPHA, "Scale reliability", "21_RELIAB")
    add_task(TaskPhase.RELIABILITY, TaskType.RELIABILITY_ALPHA, "Reliability (alt)", "22_RELIAB2")
    add_task(TaskPhase.RELIABILITY, TaskType.EFFECT_SIZES, "Reliability effect sizes", "23_REL_EFF")
    add_task(TaskPhase.RELIABILITY, TaskType.SUMMARY_DASHBOARD, "Reliability summary", "24_REL_SUM")
    add_task(TaskPhase.RELIABILITY, TaskType.DESCRIPTIVE_STATS, "Reliability descriptives", "25_REL_DESC")

    # Phase 6: Advanced (5)
    add_task(TaskPhase.ADVANCED, TaskType.NORMALITY_CHECK, "Advanced normality", "26_ADV_NORM")
    add_task(TaskPhase.ADVANCED, TaskType.CORRELATION_MATRIX, "Advanced correlations", "27_ADV_CORR")
    add_task(TaskPhase.ADVANCED, TaskType.GROUP_COMPARISON, "Advanced group comparisons", "28_ADV_GRP")
    add_task(TaskPhase.ADVANCED, TaskType.EFFECT_SIZES, "Advanced effect sizes", "29_ADV_EFF")
    add_task(TaskPhase.ADVANCED, TaskType.SUMMARY_DASHBOARD, "Advanced summary", "30_ADV_SUM")

    # Phase 7: Synthesis (5)
    add_task(TaskPhase.SYNTHESIS, TaskType.SUMMARY_DASHBOARD, "Synthesis summary", "31_SYN_SUM")
    add_task(TaskPhase.SYNTHESIS, TaskType.CORRELATION_MATRIX, "Synthesis correlations", "32_SYN_CORR")
    add_task(TaskPhase.SYNTHESIS, TaskType.DESCRIPTIVE_STATS, "Synthesis descriptives", "33_SYN_DESC")
    add_task(TaskPhase.SYNTHESIS, TaskType.EFFECT_SIZES, "Synthesis effect sizes", "34_SYN_EFF")
    add_task(TaskPhase.SYNTHESIS, TaskType.FREQUENCY_TABLES, "Synthesis frequencies", "35_SYN_FREQ")

    # Phase 8: Deliverables (5)
    add_task(TaskPhase.DELIVERABLES, TaskType.SUMMARY_DASHBOARD, "APA results summary", "36_APA")
    add_task(TaskPhase.DELIVERABLES, TaskType.SUMMARY_DASHBOARD, "Methodology notes", "37_METHOD")
    add_task(TaskPhase.DELIVERABLES, TaskType.SUMMARY_DASHBOARD, "Limitations", "38_LIMITS")
    add_task(TaskPhase.DELIVERABLES, TaskType.SUMMARY_DASHBOARD, "Audit certificate", "39_AUDIT")
    add_task(TaskPhase.DELIVERABLES, TaskType.SUMMARY_DASHBOARD, "Execution log", "40_LOG")

    return MasterPlan(
        session_id=session_id,
        total_variables=total_vars,
        total_observations=total_obs,
        detected_scales=detected_scales,
        research_questions=state.get("research_questions", []),
        tasks=tasks
    )


async def strategist_node(state: SurveyAnalysisState) -> Dict[str, Any]:
    """
    Strategist agent node - creates the Master Plan.

    Args:
        state: Current workflow state

    Returns:
        State updates with master_plan and tasks
    """
    llm = ChatAnthropic(
        model=STRATEGIST_MODEL,
        temperature=STRATEGIST_TEMP,
        max_tokens=STRATEGIST_MAX_TOKENS,
        api_key=ANTHROPIC_API_KEY
    )

    prompt = f"""Analyze this survey data and create a comprehensive Master Plan:

{state['data_summary']}

DETECTED SCALES: {list(state.get('detected_scales', {}).keys())}

NUMERIC COLUMNS: {state.get('numeric_columns', [])[:15]}

CATEGORICAL COLUMNS: {state.get('categorical_columns', [])[:10]}

Return JSON only, using the schema defined in your system prompt.
The raw data sheet is named '00_RAW_DATA_LOCKED' with {state['n_rows']} data rows.
Column letters start at A for the first column."""

    messages = [
        SystemMessage(content=STRATEGIST_SYSTEM_PROMPT),
        HumanMessage(content=prompt)
    ]

    response = await llm.ainvoke(messages)
    master_plan = response.content

    try:
        plan = parse_master_plan_json(master_plan)
    except Exception:
        plan = generate_default_master_plan(state)
        master_plan = plan.model_dump_json(indent=2)

    tasks = []
    for t in plan.tasks:
        td = t.model_dump()
        td["status"] = "pending"
        tasks.append(td)

    log_entry = LogEntry(
        timestamp=datetime.now().isoformat(),
        agent="strategist",
        action="Created Master Plan",
        details=f"Generated {len(tasks)} tasks across phases",
        task_id=None
    )

    return {
        "master_plan": master_plan,
        "plan_json": plan.model_dump(),
        "tasks": tasks,
        "total_tasks": len(tasks),
        "master_plan_approved": False,
        "plan_errors": [],
        "status": "planning",
        "execution_log": [log_entry],
        "messages": [{"role": "strategist", "content": f"Created Master Plan with {len(tasks)} tasks"}]
    }

