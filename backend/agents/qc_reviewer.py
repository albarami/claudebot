"""
Agent 3: QC Reviewer - DUAL REVIEW SYSTEM
Verifies every task meets PhD-level standards with VETO POWER.

Architecture:
1. DETERMINISTIC QC: Fast, programmatic checks (formula coverage, references)
2. LLM QC: Dual review by Claude Sonnet 4.5 AND OpenAI 5.2

Reviews the ACTUAL Excel file, not just text reports.
"""

from pathlib import Path
from typing import Dict, Any, List
from datetime import datetime
import re

import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

from langchain_anthropic import ChatAnthropic
from langchain_openai import ChatOpenAI
from langchain_core.messages import HumanMessage, SystemMessage

from config import (
    QC_REVIEWER_MODEL_1, QC_REVIEWER_MODEL_2,
    QC_REVIEWER_TEMP, QC_REVIEWER_MAX_TOKENS,
    ANTHROPIC_API_KEY, OPENAI_API_KEY
)
from utils.prompts import QC_REVIEWER_SYSTEM_PROMPT
from graph.state import SurveyAnalysisState, QCDecision, LogEntry
from engines.qc_engine import run_deterministic_qc
from models.task_schema import TaskSpec, TaskType


def verify_excel_file(workbook_path: str, sheet_name: str) -> Dict[str, Any]:
    """Read and verify the Excel file."""
    result = {
        "file_exists": False,
        "sheet_exists": False,
        "total_cells": 0,
        "formula_cells": 0,
        "value_cells": 0,
        "empty_cells": 0,
        "formula_percentage": 0,
        "sample_formulas": [],
        "potential_errors": [],
        "cell_contents": []
    }

    path = Path(workbook_path)
    if not path.exists():
        result["potential_errors"].append(f"Excel file not found: {workbook_path}")
        return result

    result["file_exists"] = True

    try:
        wb = load_workbook(workbook_path, data_only=False)

        if sheet_name not in wb.sheetnames:
            result["potential_errors"].append(f"Sheet '{sheet_name}' not found. Available: {wb.sheetnames}")
            return result

        result["sheet_exists"] = True
        ws = wb[sheet_name]

        for row in ws.iter_rows(min_row=1, max_row=min(50, ws.max_row or 1),
                                 min_col=1, max_col=min(15, ws.max_column or 1)):
            for cell in row:
                result["total_cells"] += 1

                if cell.value is None:
                    result["empty_cells"] += 1
                elif isinstance(cell.value, str) and cell.value.startswith("="):
                    result["formula_cells"] += 1
                    if len(result["sample_formulas"]) < 10:
                        result["sample_formulas"].append({
                            "cell": f"{get_column_letter(cell.column)}{cell.row}",
                            "formula": cell.value
                        })
                else:
                    result["value_cells"] += 1
                    result["cell_contents"].append({
                        "cell": f"{get_column_letter(cell.column)}{cell.row}",
                        "value": str(cell.value)[:50]
                    })

        non_empty = result["total_cells"] - result["empty_cells"]
        if non_empty > 0:
            result["formula_percentage"] = (result["formula_cells"] / non_empty) * 100

        if result["formula_percentage"] < 50 and result["formula_cells"] > 0:
            result["potential_errors"].append(
                f"Low formula percentage: {result['formula_percentage']:.1f}% - expected mostly formulas"
            )

        wb.close()

    except Exception as e:
        result["potential_errors"].append(f"Error reading Excel: {str(e)}")

    return result


def clean_dataframe_for_verification(df: pd.DataFrame) -> pd.DataFrame:
    """Mirror formula-engine cleaning so verification uses the same data."""
    cleaned = df.copy()
    for col in df.columns:
        series = df[col]
        if pd.api.types.is_numeric_dtype(series):
            cleaned[col] = pd.to_numeric(series, errors="coerce")
            continue

        numeric_candidate = pd.to_numeric(series, errors="coerce")
        non_null = series.notna().sum()
        numeric_ratio = numeric_candidate.notna().sum() / max(non_null, 1)

        if non_null >= 5 and numeric_ratio >= 0.8:
            cleaned[col] = numeric_candidate
        else:
            clean_series = series.astype(str).str.strip()
            clean_series = clean_series.replace({
                "": pd.NA,
                "nan": pd.NA,
                "NaN": pd.NA,
                "None": pd.NA
            })
            cleaned[col] = clean_series

    return cleaned


def build_review_prompt(
    current_task: Dict,
    task_output: str,
    revision_count: int,
    prev_feedback: str,
    excel_verification: Dict[str, Any]
) -> str:
    """Build the review prompt with actual Excel verification data."""

    excel_report = f"""
EXCEL FILE VERIFICATION (ACTUAL FILE CONTENTS):
- File exists: {excel_verification['file_exists']}
- Sheet exists: {excel_verification['sheet_exists']}
- Total cells examined: {excel_verification['total_cells']}
- Cells with formulas: {excel_verification['formula_cells']}
- Cells with values: {excel_verification['value_cells']}
- Formula percentage: {excel_verification['formula_percentage']:.1f}%

SAMPLE FORMULAS FOUND IN EXCEL:
"""
    for f in excel_verification.get('sample_formulas', [])[:5]:
        excel_report += f"  {f['cell']}: {f['formula']}\n"

    if excel_verification.get('potential_errors'):
        excel_report += "\nPOTENTIAL ISSUES:\n"
        for err in excel_verification['potential_errors']:
            excel_report += f"  WARNING: {err}\n"

    return f"""Review this task execution for PhD-level quality:

TASK SPECIFICATION (from Master Plan):
- Task ID: {current_task['id']}
- Phase: {current_task['phase']}
- Type: {current_task.get('task_type', 'unknown')}
- Name: {current_task['name']}
- Objective: {current_task['objective']}
- Expected Output: {current_task['output_sheet']}

{excel_report}

IMPLEMENTER'S REPORT:
{task_output}

REVISION HISTORY:
This is revision attempt #{revision_count + 1}
{f"Previous QC feedback: {prev_feedback}" if revision_count > 0 else "First submission"}

YOUR VERIFICATION CHECKLIST:

A. EXCEL FILE VERIFICATION (from actual file inspection above)
- Excel file exists and is accessible
- Required sheet was created
- Cells contain formulas (check formula percentage above)
- Formulas use correct syntax (=AVERAGE, =STDEV.S, =CORREL, etc.)
- Formulas reference '00_CLEANED_DATA' (preferred) or '00_RAW_DATA_LOCKED' correctly

B. METHODOLOGICAL SOUNDNESS
- Statistical method matches the task objective
- Appropriate for the data type
- Sample size considerations addressed

C. FORMULA ACCURACY (verify from sample formulas above)
- Formulas would produce correct results
- Data ranges are appropriate
- No obvious errors in formula logic

DECISION CRITERIA:
- APPROVE: File exists, sheet created, formulas present and correct
- REJECT: Missing file/sheet, no formulas, or incorrect formulas
- CONDITIONAL: Minor issues that don't affect accuracy
- HALT: Only if task is fundamentally impossible

Make your decision: APPROVE, REJECT, CONDITIONAL, or HALT"""


def parse_decision(review_text: str) -> str:
    """Parse decision from review text."""
    if "REJECT" in review_text.upper():
        return "REJECT"
    elif "HALT" in review_text.upper():
        return "HALT"
    elif "CONDITIONAL" in review_text.upper():
        return "CONDITIONAL"
    return "APPROVE"


def build_verification_config(task: TaskSpec, state: SurveyAnalysisState) -> Dict[str, Any]:
    """Build verification config for deterministic checks."""
    config: Dict[str, Any] = {}

    if task.task_type == TaskType.DESCRIPTIVE_STATS:
        cols = task.columns.column_names or state.get("numeric_columns", [])
        if task.columns.max_columns:
            cols = cols[:task.columns.max_columns]
        cell_maps = {}
        start_row = 4
        for idx, col in enumerate(cols):
            row = start_row + idx
            cell_maps[col] = {
                "count": f"B{row}",
                "mean": f"C{row}",
                "std": f"D{row}",
                "median": f"F{row}",
                "min": f"G{row}",
                "max": f"H{row}",
                "skewness": f"J{row}",
                "kurtosis": f"K{row}"
            }
        config = {
            "columns": cols,
            "cell_maps": cell_maps,
            "data_region": (4, 2, 3 + len(cols), 11)
        }

    if task.task_type == TaskType.CORRELATION_MATRIX:
        cols = task.columns.column_names or state.get("numeric_columns", [])
        cols = cols[:15]
        config = {
            "columns": cols,
            "start_row": 4,
            "start_col": 2,
            "data_region": (4, 2, 3 + len(cols), 1 + len(cols))
        }

    return config


async def qc_reviewer_node(state: SurveyAnalysisState) -> Dict[str, Any]:
    """Dual QC Reviewer with deterministic pre-checks."""
    current_task = state.get('current_task')
    task_output = state.get('current_task_output', '')

    if not current_task:
        return {
            "qc_decision": "ERROR",
            "qc_feedback": "No task to review",
            "messages": [{"role": "qc_reviewer", "content": "Error: No task to review"}]
        }

    revision_count = state.get('task_revision_count', 0)
    prev_feedback = state.get('qc_feedback', '')

    task_spec = TaskSpec.model_validate(current_task)

    workbook_path = state.get('workbook_path', '')

    sheets_created = state.get('sheets_created', [])
    if sheets_created:
        sheet_name = sheets_created[-1]
    else:
        sheet_name = task_spec.output_sheet

    excel_verification = verify_excel_file(workbook_path, sheet_name)

    raw_df = pd.read_excel(state['file_path'])
    cleaned_df = clean_dataframe_for_verification(raw_df)
    verification_config = build_verification_config(task_spec, state)

    deterministic_result = run_deterministic_qc(
        workbook_path=Path(workbook_path),
        sheet_name=sheet_name,
        raw_data=cleaned_df,
        task_id=task_spec.id,
        task_type=task_spec.task_type.value,
        verification_config=verification_config
    )

    if not deterministic_result.get("passed", False):
        deterministic_errors = deterministic_result.get("errors", [])
        combined_feedback = f"""
DETERMINISTIC QC FAILED (no LLM review needed)

Errors:
{chr(10).join('- ' + e for e in deterministic_errors)}

Metrics:
- Formula coverage: {deterministic_result.get('metrics', {}).get('formula_percentage', 0)}%
- Formula cells: {deterministic_result.get('metrics', {}).get('formula_cells', 0)}

Fix these issues before resubmitting.
"""
        return {
            "qc_decision": "REJECT",
            "qc_feedback": combined_feedback,
            "verification_status": "fail",
            "formula_coverage": deterministic_result.get("metrics", {}).get("formula_percentage", 0),
            "task_revision_count": revision_count + 1,
            "execution_log": [LogEntry(
                timestamp=datetime.now().isoformat(),
                agent="qc_reviewer",
                action=f"Deterministic QC failed for task {task_spec.id}",
                details=f"Errors: {deterministic_errors}",
                task_id=task_spec.id
            )],
            "messages": [
                {"role": "qc_reviewer", "content": f"Deterministic QC: REJECT - {deterministic_errors[0] if deterministic_errors else 'Failed checks'}"}
            ]
        }

    prompt = build_review_prompt(current_task, task_output, revision_count, prev_feedback, excel_verification)
    messages = [
        SystemMessage(content=QC_REVIEWER_SYSTEM_PROMPT),
        HumanMessage(content=prompt)
    ]

    llm_sonnet = ChatAnthropic(
        model=QC_REVIEWER_MODEL_1,
        temperature=QC_REVIEWER_TEMP,
        max_tokens=QC_REVIEWER_MAX_TOKENS,
        api_key=ANTHROPIC_API_KEY
    )

    response_sonnet = await llm_sonnet.ainvoke(messages)
    review_sonnet = response_sonnet.content
    decision_sonnet = parse_decision(review_sonnet)

    llm_openai = ChatOpenAI(
        model=QC_REVIEWER_MODEL_2,
        temperature=QC_REVIEWER_TEMP,
        max_tokens=QC_REVIEWER_MAX_TOKENS,
        api_key=OPENAI_API_KEY
    )

    response_openai = await llm_openai.ainvoke(messages)
    review_openai = response_openai.content
    decision_openai = parse_decision(review_openai)

    if decision_sonnet == "HALT" or decision_openai == "HALT":
        final_decision = "HALT"
    elif decision_sonnet == "REJECT" or decision_openai == "REJECT":
        final_decision = "REJECT"
    elif decision_sonnet == "APPROVE" and decision_openai == "APPROVE":
        final_decision = "APPROVE"
    elif decision_sonnet == "CONDITIONAL" or decision_openai == "CONDITIONAL":
        final_decision = "CONDITIONAL"
    else:
        final_decision = "APPROVE"

    combined_feedback = f"""
=================================================================
DUAL QC REVIEW RESULTS
=================================================================

REVIEW 1: Claude Sonnet 4.5
Decision: {decision_sonnet}
{review_sonnet}

---------------------------------------------------------------

REVIEW 2: OpenAI 5.2
Decision: {decision_openai}
{review_openai}

---------------------------------------------------------------

FINAL DECISION: {final_decision}
(Both reviewers must agree for approval)
=================================================================
"""

    qc_record = QCDecision(
        task_id=task_spec.id,
        decision=final_decision,
        feedback=combined_feedback,
        checklist_results={
            "sonnet_decision": decision_sonnet,
            "openai_decision": decision_openai
        },
        timestamp=datetime.now().isoformat(),
        revision_number=revision_count + 1
    )

    log_entry = LogEntry(
        timestamp=datetime.now().isoformat(),
        agent="qc_reviewer",
        action=f"Dual review of task {task_spec.id}",
        details=f"Sonnet: {decision_sonnet}, OpenAI: {decision_openai} -> Final: {final_decision}",
        task_id=task_spec.id
    )

    if final_decision == "REJECT":
        new_revision_count = revision_count + 1
    else:
        new_revision_count = 0

    return {
        "qc_decision": final_decision,
        "qc_feedback": combined_feedback,
        "verification_status": "pass" if deterministic_result.get("passed") else "fail",
        "formula_coverage": deterministic_result.get("metrics", {}).get("formula_percentage", 0),
        "qc_history": [qc_record],
        "task_revision_count": new_revision_count,
        "execution_log": [log_entry],
        "messages": [
            {"role": "qc_reviewer_sonnet", "content": f"Sonnet 4.5: {decision_sonnet}"},
            {"role": "qc_reviewer_openai", "content": f"OpenAI 5.2: {decision_openai}"},
            {"role": "qc_reviewer", "content": f"FINAL: {final_decision}"}
        ]
    }

