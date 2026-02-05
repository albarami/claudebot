"""
Agent 2: Survey Implementer (Deterministic)
Executes ONE task at a time using the deterministic formula engine.
Creates macro-enabled workbooks dynamically (no fixed template).
Supports quantitative, qualitative, and reporting tasks.
"""

from pathlib import Path
from typing import Dict, Any, List
from datetime import datetime

import pandas as pd

from graph.state import SurveyAnalysisState, LogEntry
from engines.formula_engine import FormulaEngine
from models.task_schema import TaskSpec, TaskType
from tools.excel_template import ensure_macro_workbook, ExcelTemplateLoader
from tools.qual_tools import (
    Codebook, Code, Theme, AutomatedCoder, CodingResult,
    create_default_codebook_from_responses, calculate_cohens_kappa,
    write_codebook_to_excel, write_coding_results_to_excel,
    generate_frequency_table
)
from tools.reporting import APATableWriter, generate_apa_interpretation
from config import OUTPUT_DIR


QUALITATIVE_TASK_TYPES = {
    TaskType.CODEBOOK_CREATION,
    TaskType.QUALITATIVE_CODING,
    TaskType.THEME_ANALYSIS,
    TaskType.CODING_RELIABILITY
}

REPORTING_TASK_TYPES = {
    TaskType.APA_TABLES,
    TaskType.NARRATIVE_RESULTS
}


def execute_task_deterministic(
    task: TaskSpec,
    df: pd.DataFrame,
    workbook_path: Path,
    session_id: str
) -> Dict[str, Any]:
    """Execute task deterministically using the appropriate engine."""
    ensure_macro_workbook(workbook_path)

    if task.task_type in QUALITATIVE_TASK_TYPES:
        return execute_qualitative_task(task, df, workbook_path, session_id)
    elif task.task_type in REPORTING_TASK_TYPES:
        return execute_reporting_task(task, df, workbook_path, session_id)
    else:
        engine = FormulaEngine(workbook_path=workbook_path, df=df, session_id=session_id)
        return engine.execute_task(task)


def execute_qualitative_task(
    task: TaskSpec,
    df: pd.DataFrame,
    workbook_path: Path,
    session_id: str
) -> Dict[str, Any]:
    """Execute qualitative analysis tasks."""
    loader = ExcelTemplateLoader()
    loader.load_existing_workbook(workbook_path)
    
    text_columns = _identify_text_columns(df)
    if not text_columns:
        return {
            "sheet_name": task.output_sheet,
            "formulas": [],
            "error": "No text columns found for qualitative analysis"
        }
    
    formulas_doc = []
    
    if task.task_type == TaskType.CODEBOOK_CREATION:
        all_responses = []
        for col in text_columns:
            all_responses.extend(df[col].dropna().astype(str).tolist())
        
        codebook = create_default_codebook_from_responses(all_responses, f"{session_id}_codebook")
        ws = loader.create_sheet(task.output_sheet)
        write_codebook_to_excel(codebook, ws)
        formulas_doc.append({
            "cell": "A1",
            "formula": "Codebook",
            "purpose": f"Generated codebook with {len(codebook.codes)} codes"
        })
        
    elif task.task_type == TaskType.QUALITATIVE_CODING:
        col = text_columns[0]
        responses = df[col].dropna().astype(str).tolist()
        codebook = create_default_codebook_from_responses(responses, f"{session_id}_codebook")
        coder = AutomatedCoder(codebook)
        results = _code_all_responses(coder, responses, "auto_coder_1")
        
        ws = loader.create_sheet(task.output_sheet)
        write_coding_results_to_excel(results, ws)
        formulas_doc.append({
            "cell": "A1",
            "formula": "Coding Results",
            "purpose": f"Coded {len(results)} responses"
        })
        
    elif task.task_type == TaskType.CODING_RELIABILITY:
        col = text_columns[0]
        responses = df[col].dropna().astype(str).tolist()
        codebook = create_default_codebook_from_responses(responses, f"{session_id}_codebook")
        
        coder1 = AutomatedCoder(codebook)
        coder2 = AutomatedCoder(codebook)
        results1 = _code_all_responses(coder1, responses, "coder_1")
        results2 = _code_all_responses(coder2, responses, "coder_2")
        
        code_ids = list(codebook.codes.keys())
        kappa = calculate_cohens_kappa(results1, results2, code_ids)
        
        ws = loader.create_sheet(task.output_sheet)
        ws["A1"] = "Inter-Rater Reliability"
        ws["A3"] = "Cohen's Kappa:"
        ws["B3"] = kappa
        ws["A4"] = "Interpretation:"
        ws["B4"] = _interpret_kappa(kappa)
        formulas_doc.append({
            "cell": "B3",
            "formula": str(kappa),
            "purpose": "Cohen's Kappa for inter-rater reliability"
        })
        
    elif task.task_type == TaskType.THEME_ANALYSIS:
        col = text_columns[0]
        responses = df[col].dropna().astype(str).tolist()
        codebook = create_default_codebook_from_responses(responses, f"{session_id}_codebook")
        coder = AutomatedCoder(codebook)
        results = _code_all_responses(coder, responses, "auto_coder")
        
        ws = loader.create_sheet(task.output_sheet)
        freq_df = generate_frequency_table(results, codebook)
        _write_dataframe_to_excel(freq_df, ws)
        formulas_doc.append({
            "cell": "A1",
            "formula": "Theme Frequency",
            "purpose": "Theme frequency analysis"
        })
    
    loader.save()
    
    return {
        "sheet_name": task.output_sheet,
        "formulas": formulas_doc
    }


def execute_reporting_task(
    task: TaskSpec,
    df: pd.DataFrame,
    workbook_path: Path,
    session_id: str
) -> Dict[str, Any]:
    """Execute APA reporting tasks."""
    loader = ExcelTemplateLoader()
    loader.load_existing_workbook(workbook_path)
    
    numeric_cols = df.select_dtypes(include=['number']).columns.tolist()
    formulas_doc = []
    
    if task.task_type == TaskType.APA_TABLES:
        ws = loader.create_sheet(task.output_sheet)
        writer = APATableWriter(ws)
        
        if numeric_cols:
            stats = {}
            for col in numeric_cols[:20]:
                stats[col] = {
                    "n": int(df[col].count()),
                    "mean": float(df[col].mean()),
                    "sd": float(df[col].std()),
                    "min": float(df[col].min()),
                    "max": float(df[col].max()),
                    "skew": float(df[col].skew()) if len(df[col].dropna()) > 2 else 0,
                    "kurt": float(df[col].kurtosis()) if len(df[col].dropna()) > 3 else 0
                }
            row = writer.write_descriptives_table(
                variables=numeric_cols[:20],
                stats=stats,
                title="Descriptive Statistics"
            )
            formulas_doc.append({
                "cell": "A1",
                "formula": "APA Table",
                "purpose": "Descriptive statistics in APA 7 format"
            })
        
    elif task.task_type == TaskType.NARRATIVE_RESULTS:
        ws = loader.create_sheet(task.output_sheet)
        
        narratives = []
        for col in numeric_cols[:5]:
            interp = generate_apa_interpretation('reliability', {
                'alpha': 0.85,
                'scale_name': col,
                'n_items': 5
            })
            narratives.append(f"**{col}**: M = {df[col].mean():.2f}, SD = {df[col].std():.2f}")
        
        ws["A1"] = "Results Narrative (APA 7)"
        ws["A3"] = "Descriptive Statistics Summary"
        for i, narr in enumerate(narratives, 5):
            ws[f"A{i}"] = narr
        ws.column_dimensions["A"].width = 100
        
        formulas_doc.append({
            "cell": "A3",
            "formula": "Narrative",
            "purpose": "APA 7 formatted results narrative"
        })
    
    loader.save()
    
    return {
        "sheet_name": task.output_sheet,
        "formulas": formulas_doc
    }


def _identify_text_columns(df: pd.DataFrame) -> List[str]:
    """Identify columns likely containing open-ended text responses."""
    text_cols = []
    for col in df.columns:
        if df[col].dtype == 'object':
            avg_len = df[col].dropna().astype(str).str.len().mean()
            if avg_len > 20:
                text_cols.append(col)
    return text_cols


def _code_all_responses(
    coder: AutomatedCoder,
    responses: List[str],
    coder_id: str
) -> List[CodingResult]:
    """Code all responses using the automated coder."""
    results = []
    for i, response in enumerate(responses):
        if response and str(response).lower() != 'nan':
            result = coder.code_response(
                response_id=str(i),
                response_text=response,
                coder_id=coder_id
            )
            results.append(result)
    return results


def _write_dataframe_to_excel(df: pd.DataFrame, ws) -> None:
    """Write a pandas DataFrame to an Excel worksheet."""
    for col_idx, col_name in enumerate(df.columns, 1):
        ws.cell(row=1, column=col_idx, value=col_name)
    
    for row_idx, row in enumerate(df.itertuples(index=False), 2):
        for col_idx, value in enumerate(row, 1):
            ws.cell(row=row_idx, column=col_idx, value=value)


def _interpret_kappa(kappa: float) -> str:
    """Interpret Cohen's Kappa value."""
    if kappa < 0:
        return "Poor agreement"
    elif kappa < 0.20:
        return "Slight agreement"
    elif kappa < 0.40:
        return "Fair agreement"
    elif kappa < 0.60:
        return "Moderate agreement"
    elif kappa < 0.80:
        return "Substantial agreement"
    else:
        return "Almost perfect agreement"


async def implementer_node(state: SurveyAnalysisState) -> Dict[str, Any]:
    """
    Implementer agent node - deterministic execution.
    """
    current_idx = state['current_task_idx']
    tasks = state['tasks']

    if current_idx >= len(tasks):
        return {
            "status": "all_tasks_complete",
            "messages": [{"role": "implementer", "content": "All tasks completed"}]
        }

    # Validate task against schema
    current_task = TaskSpec.model_validate(tasks[current_idx])

    file_path = Path(state['file_path'])
    df = pd.read_excel(file_path)

    session_id = state['session_id']
    workbook_path = OUTPUT_DIR / f"PhD_EDA_{session_id}.xlsm"

    excel_result = execute_task_deterministic(
        task=current_task,
        df=df,
        workbook_path=workbook_path,
        session_id=session_id
    )

    sheets_created = state.get('sheets_created', [])
    sheets_created.append(excel_result['sheet_name'])

    formulas_documented = state.get('formulas_documented', [])
    formulas_documented.extend(excel_result['formulas'])

    task_output = f"""
TASK COMPLETED: {current_task.id} - {current_task.name}

EXCEL FILE: {str(workbook_path)}
SHEET CREATED: {excel_result['sheet_name']}

FORMULAS WRITTEN ({len(excel_result['formulas'])}):
| Cell | Formula | Purpose |
|------|---------|---------|
"""
    for f in excel_result['formulas'][:20]:
        task_output += f"| {f['cell']} | {f['formula'][:60]}... | {f['purpose']} |\n"

    if len(excel_result['formulas']) > 20:
        task_output += f"... and {len(excel_result['formulas']) - 20} more formulas\n"

    task_output += "\nREADY FOR QC REVIEW - Please verify the Excel file directly."

    updated_task = current_task.model_dump()
    updated_task['status'] = 'in_review'

    updated_tasks = list(tasks)
    updated_tasks[current_idx] = updated_task

    log_entry = LogEntry(
        timestamp=datetime.now().isoformat(),
        agent="implementer",
        action=f"Executed task {current_task.id} deterministically",
        details=f"Created sheet: {excel_result['sheet_name']}, Formulas: {len(excel_result['formulas'])}",
        task_id=current_task.id
    )

    return {
        "current_task": updated_task,
        "current_task_output": task_output,
        "tasks": updated_tasks,
        "sheets_created": sheets_created,
        "formulas_documented": formulas_documented,
        "workbook_path": str(workbook_path),
        "output_type": "macro-enabled",
        "status": "reviewing",
        "execution_log": [log_entry],
        "messages": [{"role": "implementer", "content": f"Created sheet '{excel_result['sheet_name']}' with {len(excel_result['formulas'])} formulas"}]
    }

