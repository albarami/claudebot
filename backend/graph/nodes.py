"""
LangGraph workflow nodes for PhD Survey Analyzer.
Each node represents a step in the multi-agent workflow.
"""

from pathlib import Path
from typing import Dict, Any
from datetime import datetime

import pandas as pd
from openpyxl.utils import get_column_letter

from graph.state import SurveyAnalysisState, LogEntry
from tools.stats_tools import SurveyDataAnalyzer
from tools.excel_tools import ExcelFormulaWorkbook, get_column_mapping
from config import OUTPUT_DIR


async def load_data_node(state: SurveyAnalysisState) -> Dict[str, Any]:
    """
    Load and analyze the survey Excel file.
    
    Args:
        state: Current workflow state with file_path
    
    Returns:
        State updates with data analysis results
    """
    file_path = Path(state['file_path'])
    
    df = pd.read_excel(file_path)
    
    analyzer = SurveyDataAnalyzer(df)
    
    data_summary = analyzer.create_data_summary()
    column_types = analyzer.get_column_types()
    numeric_cols = analyzer.get_numeric_columns()
    categorical_cols = analyzer.get_categorical_columns()
    scales = analyzer.detect_scales()
    
    log_entry = LogEntry(
        timestamp=datetime.now().isoformat(),
        agent="system",
        action="Loaded survey data",
        details=f"Loaded {len(df)} rows × {len(df.columns)} columns",
        task_id=None
    )
    
    return {
        "file_name": file_path.name,
        "n_rows": len(df),
        "n_cols": len(df.columns),
        "columns": list(df.columns),
        "column_types": column_types,
        "numeric_columns": numeric_cols,
        "categorical_columns": categorical_cols,
        "detected_scales": scales,
        "data_summary": data_summary,
        "status": "data_loaded",
        "execution_log": [log_entry],
        "messages": [{"role": "system", "content": f"Loaded survey: {len(df)} rows × {len(df.columns)} columns"}]
    }


async def advance_task_node(state: SurveyAnalysisState) -> Dict[str, Any]:
    """
    Move to the next task after QC approval.
    
    Args:
        state: Current workflow state
    
    Returns:
        State updates with incremented task index
    """
    current_idx = state['current_task_idx']
    tasks = state['tasks']
    
    if current_idx < len(tasks):
        updated_tasks = list(tasks)
        updated_tasks[current_idx]['status'] = 'completed'
        
        log_entry = LogEntry(
            timestamp=datetime.now().isoformat(),
            agent="system",
            action=f"Task {tasks[current_idx]['id']} completed",
            details="Moving to next task",
            task_id=tasks[current_idx]['id']
        )
        
        return {
            "current_task_idx": current_idx + 1,
            "tasks": updated_tasks,
            "task_revision_count": 0,
            "qc_feedback": "",
            "status": "executing",
            "execution_log": [log_entry],
            "messages": [{"role": "system", "content": f"Task {current_idx + 1}/{len(tasks)} completed"}]
        }
    
    return {"status": "all_tasks_complete"}


async def generate_deliverables_node(state: SurveyAnalysisState) -> Dict[str, Any]:
    """
    Generate all output deliverables.
    
    Args:
        state: Current workflow state with completed analysis
    
    Returns:
        State updates with deliverable paths
    """
    session_id = state['session_id']
    file_name = state.get('file_name', 'survey')
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    
    deliverables = []
    
    master_plan_path = OUTPUT_DIR / f"MASTER_PLAN_{session_id}.md"
    with open(master_plan_path, "w", encoding="utf-8") as f:
        f.write(f"# MASTER PLAN - PhD-Level Survey EDA\n\n")
        f.write(f"Survey: {file_name}\n")
        f.write(f"Session: {session_id}\n")
        f.write(f"Generated: {timestamp}\n")
        f.write(f"Agent: Survey Strategist (Claude Opus 4.5)\n\n")
        f.write("---\n\n")
        f.write(state.get('master_plan', 'No plan generated'))
    deliverables.append(str(master_plan_path))
    
    qc_trail_path = OUTPUT_DIR / f"QC_AUDIT_TRAIL_{session_id}.md"
    with open(qc_trail_path, "w", encoding="utf-8") as f:
        f.write("# QC AUDIT TRAIL\n\n")
        f.write(f"Session: {session_id}\n")
        f.write(f"Generated: {timestamp}\n\n")
        f.write("| Task | Decision | Revision # | Timestamp |\n")
        f.write("|------|----------|------------|----------|\n")
        for qc in state.get('qc_history', []):
            f.write(f"| {qc.get('task_id', '')} | {qc.get('decision', '')} | {qc.get('revision_number', '')} | {qc.get('timestamp', '')} |\n")
    deliverables.append(str(qc_trail_path))
    
    audit_path = OUTPUT_DIR / f"AUDIT_CERTIFICATE_{session_id}.md"
    with open(audit_path, "w", encoding="utf-8") as f:
        f.write("═" * 60 + "\n")
        f.write("# ACADEMIC AUDIT CERTIFICATE\n")
        f.write("═" * 60 + "\n\n")
        f.write(f"**Survey:** {file_name}\n")
        f.write(f"**Session:** {session_id}\n")
        f.write(f"**Date:** {timestamp}\n")
        f.write(f"**Auditor:** Claude Opus 4.5 (Survey Auditor Agent)\n\n")
        f.write("## QUALITY SCORES\n\n")
        quality_scores = state.get('quality_scores') or {}
        if quality_scores:
            for k, v in quality_scores.items():
                score = float(v) if v is not None else 0.0
                status = "✓" if score >= 95 else "⚠" if score >= 90 else "❌"
                f.write(f"- **{k.replace('_', ' ').title()}:** {score:.1f}% {status}\n")
        else:
            f.write("- Quality scores not yet available\n")
        overall = float(state.get('overall_score') or 0)
        f.write(f"\n**OVERALL SCORE:** {overall:.1f}%\n")
        f.write(f"\n## CERTIFICATION: {state.get('certification', 'PENDING')}\n\n")
        f.write("---\n\n")
        f.write(state.get('audit_result', ''))
    deliverables.append(str(audit_path))
    
    method_path = OUTPUT_DIR / f"METHODOLOGY_DOCUMENTATION_{session_id}.md"
    with open(method_path, "w", encoding="utf-8") as f:
        f.write("# METHODOLOGY DOCUMENTATION\n\n")
        f.write("Ready for thesis Chapter 3 (Methods)\n\n")
        f.write(f"Generated: {timestamp}\n\n")
        f.write("---\n\n")
        f.write("## Participants\n\n")
        f.write(f"The sample consisted of N = {state['n_rows']} participants.\n\n")
        f.write("## Measures\n\n")
        for scale, items in state.get('detected_scales', {}).items():
            f.write(f"### {scale}\n")
            f.write(f"This scale consisted of {len(items)} items ({', '.join(items[:3])}{'...' if len(items) > 3 else ''}).\n\n")
        f.write("## Data Analysis\n\n")
        f.write("All analyses were conducted using Excel with formula-based computations.\n")
    deliverables.append(str(method_path))
    
    limits_path = OUTPUT_DIR / f"LIMITATIONS_ASSESSMENT_{session_id}.md"
    with open(limits_path, "w", encoding="utf-8") as f:
        f.write("# STUDY LIMITATIONS - HONEST ASSESSMENT\n\n")
        f.write("## Sampling\n")
        f.write("- Convenience sample (not probability-based)\n")
        f.write("- Possible self-selection bias\n")
        f.write(f"- Sample size: N = {state['n_rows']}\n\n")
        f.write("## Measurement\n")
        f.write("- Self-report measures (social desirability bias)\n")
        f.write("- Cross-sectional design (no causal inference)\n\n")
        f.write("## Statistical\n")
        f.write(f"- {len(state.get('numeric_columns', []))} variables analyzed\n")
        f.write("- Multiple comparisons increase Type I error risk\n\n")
        f.write("## Recommendations for Future Research\n")
        f.write("- Longitudinal design for causal relationships\n")
        f.write("- Probability sampling for generalizability\n")
    deliverables.append(str(limits_path))
    
    log_entry = LogEntry(
        timestamp=datetime.now().isoformat(),
        agent="system",
        action="Generated deliverables",
        details=f"Created {len(deliverables)} deliverable files",
        task_id=None
    )
    
    return {
        "deliverables": deliverables,
        "status": "completed",
        "execution_log": [log_entry],
        "messages": [{"role": "system", "content": f"Generated {len(deliverables)} deliverables"}]
    }
