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
from tools.reporting import APATableWriter
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

    safe_task_id = task.id.replace(".", "_")
    data_sheet = _ensure_hidden_sheet(loader, f"00_QUAL_{safe_task_id}")
    
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
        data_ws = loader.create_sheet(data_sheet)
        write_codebook_to_excel(codebook, data_ws)
        output_ws = loader.create_sheet(task.output_sheet)
        _link_sheet_from_data(data_ws, output_ws)
        formulas_doc.append({
            "cell": "A1",
            "formula": f"='{data_sheet}'!A1",
            "purpose": f"Generated codebook with {len(codebook.codes)} codes"
        })
        
    elif task.task_type == TaskType.QUALITATIVE_CODING:
        col = text_columns[0]
        responses = df[col].dropna().astype(str).tolist()
        codebook = create_default_codebook_from_responses(responses, f"{session_id}_codebook")
        coder = AutomatedCoder(codebook)
        results = _code_all_responses(coder, responses, "auto_coder_1")
        
        data_ws = loader.create_sheet(data_sheet)
        write_coding_results_to_excel(results, data_ws)
        output_ws = loader.create_sheet(task.output_sheet)
        _link_sheet_from_data(data_ws, output_ws)
        formulas_doc.append({
            "cell": "A1",
            "formula": f"='{data_sheet}'!A1",
            "purpose": f"Coded {len(results)} responses"
        })
        
    elif task.task_type == TaskType.CODING_RELIABILITY:
        col = text_columns[0]
        responses = df[col].dropna().astype(str).tolist()
        codebook = create_default_codebook_from_responses(responses, f"{session_id}_codebook")
        
        coder1 = AutomatedCoder(codebook)
        coder2 = AutomatedCoder(_clone_codebook_without_examples(codebook))
        results1 = _code_all_responses(coder1, responses, "coder_1")
        results2 = _code_all_responses(coder2, responses, "coder_2")
        
        code_ids = list(codebook.codes.keys())
        kappa = calculate_cohens_kappa(results1, results2, code_ids)
        
        data_ws = loader.create_sheet(data_sheet)
        data_ws["A1"] = "Inter-Rater Reliability"
        data_ws["A3"] = "Cohen's Kappa:"
        data_ws["B3"] = kappa
        data_ws["A4"] = "Interpretation:"
        data_ws["B4"] = _interpret_kappa(kappa)

        output_ws = loader.create_sheet(task.output_sheet)
        _link_sheet_from_data(data_ws, output_ws)
        formulas_doc.append({
            "cell": "B3",
            "formula": f"='{data_sheet}'!B3",
            "purpose": "Cohen's Kappa for inter-rater reliability"
        })
        
    elif task.task_type == TaskType.THEME_ANALYSIS:
        col = text_columns[0]
        responses = df[col].dropna().astype(str).tolist()
        codebook = create_default_codebook_from_responses(responses, f"{session_id}_codebook")
        coder = AutomatedCoder(codebook)
        results = _code_all_responses(coder, responses, "auto_coder")
        
        data_ws = loader.create_sheet(data_sheet)
        freq_df = generate_frequency_table(results, codebook)
        _write_dataframe_to_excel(freq_df, data_ws)
        output_ws = loader.create_sheet(task.output_sheet)
        _link_sheet_from_data(data_ws, output_ws)
        formulas_doc.append({
            "cell": "A1",
            "formula": f"='{data_sheet}'!A1",
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

    formulas_doc = []
    descriptives_sheet, descriptives_ws = _find_descriptives_sheet(loader)
    if not descriptives_sheet or descriptives_ws is None:
        ws = loader.create_sheet(task.output_sheet)
        ws["A1"] = "Error: Descriptive statistics sheet not found"
        loader.save()
        return {
            "sheet_name": task.output_sheet,
            "formulas": [],
            "error": "Descriptive statistics sheet not found for reporting"
        }

    row_map = _extract_descriptive_row_map(descriptives_ws)
    if not row_map:
        ws = loader.create_sheet(task.output_sheet)
        ws["A1"] = "Error: Descriptive statistics sheet has no variables"
        loader.save()
        return {
            "sheet_name": task.output_sheet,
            "formulas": [],
            "error": "Descriptive statistics sheet has no variables"
        }

    if task.task_type == TaskType.APA_TABLES:
        ws = loader.create_sheet(task.output_sheet)
        writer = APATableWriter(ws)
        writer.write_descriptives_table_from_sheet(
            source_sheet=descriptives_sheet,
            row_map=row_map,
            title="Descriptive Statistics"
        )
        formulas_doc.append({
            "cell": "B4",
            "formula": f"=IFERROR('{descriptives_sheet}'!C4,\"\")",
            "purpose": "APA table formulas reference descriptive stats"
        })

    elif task.task_type == TaskType.NARRATIVE_RESULTS:
        ws = loader.create_sheet(task.output_sheet)
        ws["A1"] = "Results Narrative (APA 7)"
        ws["A3"] = "Descriptive Statistics Summary"

        start_row = 5
        for idx, (var, source_row) in enumerate(list(row_map.items())[:5]):
            target_row = start_row + idx
            safe_var = _escape_excel_text(var)
            formula = (
                f"=\"{safe_var}: M = \" & TEXT('{descriptives_sheet}'!C{source_row},\"0.00\") & "
                f"\", SD = \" & TEXT('{descriptives_sheet}'!D{source_row},\"0.00\")"
            )
            ws[f"A{target_row}"] = formula

        ws.column_dimensions["A"].width = 100
        first_var = list(row_map.keys())[0]
        first_row = list(row_map.values())[0]
        safe_first = _escape_excel_text(first_var)
        formulas_doc.append({
            "cell": f"A{start_row}",
            "formula": f"=\"{safe_first}: M = \" & TEXT('{descriptives_sheet}'!C{first_row},\"0.00\") & \", SD = \" & TEXT('{descriptives_sheet}'!D{first_row},\"0.00\")",
            "purpose": "APA narrative formulas reference descriptive stats"
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


def _ensure_hidden_sheet(loader: ExcelTemplateLoader, sheet_name: str) -> str:
    """Ensure a hidden data sheet exists for non-formula outputs."""
    wb = loader.workbook
    if wb is not None and sheet_name in wb.sheetnames:
        del wb[sheet_name]
    ws = loader.create_sheet(sheet_name, 0)
    ws.sheet_state = "hidden"
    return sheet_name


def _link_sheet_from_data(source_ws, target_ws) -> None:
    """Link a visible sheet to a hidden data sheet using formulas."""
    max_row = source_ws.max_row or 1
    max_col = source_ws.max_column or 1
    for r in range(1, max_row + 1):
        for c in range(1, max_col + 1):
            cell_ref = f"'{source_ws.title}'!{source_ws.cell(row=r, column=c).coordinate}"
            target_ws.cell(row=r, column=c, value=f"={cell_ref}")


def _clone_codebook_without_examples(codebook: Codebook) -> Codebook:
    """Clone codebook but remove examples to vary coder patterns."""
    clone = Codebook(name=f"{codebook.name}_no_examples")
    for code_id, code in codebook.codes.items():
        clone.add_code(Code(
            id=code.id,
            name=code.name,
            definition=code.definition,
            examples=[],
            parent_code=code.parent_code,
            frequency=code.frequency
        ))
    for theme_id, theme in codebook.themes.items():
        clone.add_theme(Theme(
            id=theme.id,
            name=theme.name,
            description=theme.description,
            codes=list(theme.codes)
        ))
    return clone


def _find_descriptives_sheet(loader: ExcelTemplateLoader):
    """Find the descriptive statistics sheet in the workbook."""
    wb = loader.workbook
    if wb is None:
        return None, None

    for name in wb.sheetnames:
        ws = wb[name]
        if str(ws["A1"].value).strip().upper().startswith("DESCRIPTIVE"):
            return name, ws

    for name in wb.sheetnames:
        if "DESC" in name.upper():
            return name, wb[name]

    return None, None


def _extract_descriptive_row_map(ws) -> Dict[str, int]:
    """Extract variable -> row map from a descriptive stats sheet."""
    row_map: Dict[str, int] = {}
    for row in range(4, (ws.max_row or 4) + 1):
        var = ws.cell(row=row, column=1).value
        if var is None or str(var).strip() == "":
            continue
        row_map[str(var)] = row
    return row_map


def _escape_excel_text(value: str) -> str:
    """Escape text for Excel formula strings."""
    return str(value).replace('"', '""')


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
