"""
Agent 1: Survey Strategist
Creates comprehensive Master Plan with 40-60 detailed tasks.
Uses Claude Opus 4.5 for highest reasoning capability.
"""

import re
from typing import Dict, List, Any
from datetime import datetime

from langchain_anthropic import ChatAnthropic
from langchain_core.messages import HumanMessage, SystemMessage

from config import STRATEGIST_MODEL, STRATEGIST_TEMP, STRATEGIST_MAX_TOKENS, ANTHROPIC_API_KEY, STRATEGIST_PROVIDER
from utils.prompts import STRATEGIST_SYSTEM_PROMPT
from graph.state import SurveyAnalysisState, Task, LogEntry


def parse_master_plan(plan_text: str) -> List[Task]:
    """
    Parse Master Plan text into structured Task objects.
    
    Args:
        plan_text: Raw master plan markdown text
    
    Returns:
        List of Task dictionaries
    """
    tasks = []
    
    task_pattern = re.compile(
        r'TASK ID:\s*([\d.]+)\s*\n'
        r'PHASE:\s*([^\n]+)\s*\n'
        r'NAME:\s*([^\n]+)',
        re.IGNORECASE
    )
    
    matches = list(task_pattern.finditer(plan_text))
    
    for i, match in enumerate(matches):
        task_id = match.group(1).strip()
        phase = match.group(2).strip()
        name = match.group(3).strip()
        
        start = match.end()
        end = matches[i + 1].start() if i + 1 < len(matches) else len(plan_text)
        task_content = plan_text[start:end]
        
        objective_match = re.search(r'OBJECTIVE:\s*\n(.*?)(?=METHOD:|$)', task_content, re.DOTALL | re.IGNORECASE)
        method_match = re.search(r'METHOD:\s*\n(.*?)(?=VALIDATION|OUTPUT|$)', task_content, re.DOTALL | re.IGNORECASE)
        output_match = re.search(r'OUTPUT:\s*\n(.*?)(?=ACCEPTANCE|$)', task_content, re.DOTALL | re.IGNORECASE)
        
        formula_pattern = re.compile(r'=\w+\([^)]+\)')
        formulas = formula_pattern.findall(task_content)
        
        task = Task(
            id=task_id,
            phase=phase,
            name=name,
            objective=objective_match.group(1).strip() if objective_match else "",
            method=method_match.group(1).strip() if method_match else "",
            formulas=formulas,
            validation="",
            output_sheet=output_match.group(1).strip() if output_match else "",
            acceptance_criteria=[],
            status="pending"
        )
        tasks.append(task)
    
    if not tasks:
        sections = plan_text.split('---')
        for i, section in enumerate(sections):
            if 'TASK' in section.upper() or 'Phase' in section:
                task = Task(
                    id=f"{i+1}.0",
                    phase="General",
                    name=f"Task {i+1}",
                    objective=section[:200],
                    method=section,
                    formulas=[],
                    validation="",
                    output_sheet="",
                    acceptance_criteria=[],
                    status="pending"
                )
                tasks.append(task)
    
    return tasks


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

Create a detailed Master Plan with 40-60 tasks following the exact format specified.
Every computational task must include EXACT Excel formulas.
The raw data sheet is named '00_RAW_DATA_LOCKED' with {state['n_rows']} data rows.
Column letters start at A for the first column."""

    messages = [
        SystemMessage(content=STRATEGIST_SYSTEM_PROMPT),
        HumanMessage(content=prompt)
    ]
    
    response = await llm.ainvoke(messages)
    master_plan = response.content
    
    tasks = parse_master_plan(master_plan)
    
    if len(tasks) < 10:
        tasks = generate_default_tasks(state)
    
    log_entry = LogEntry(
        timestamp=datetime.now().isoformat(),
        agent="strategist",
        action="Created Master Plan",
        details=f"Generated {len(tasks)} tasks across 8 phases",
        task_id=None
    )
    
    return {
        "master_plan": master_plan,
        "tasks": tasks,
        "total_tasks": len(tasks),
        "status": "planning",
        "execution_log": [log_entry],
        "messages": [{"role": "strategist", "content": f"Created Master Plan with {len(tasks)} tasks"}]
    }


def generate_default_tasks(state: SurveyAnalysisState) -> List[Task]:
    """Generate default task list if parsing fails."""
    n_rows = state['n_rows']
    numeric_cols = state.get('numeric_columns', [])[:10]
    
    tasks = [
        Task(id="1.1", phase="Data Validation", name="Lock Raw Data",
             objective="Protect original data from modification",
             method="Copy data to sheet and apply protection",
             formulas=[], validation="Sheet is read-only", output_sheet="00_RAW_DATA_LOCKED",
             acceptance_criteria=["Data intact", "Protection enabled"], status="pending"),
        
        Task(id="1.2", phase="Data Validation", name="Create Codebook",
             objective="Document all variables with formulas",
             method=f"For each column: =COUNT(range), =COUNTBLANK(range), =MIN(range), =MAX(range), =AVERAGE(range)",
             formulas=["=COUNT()", "=COUNTBLANK()", "=MIN()", "=MAX()", "=AVERAGE()"],
             validation="All variables documented", output_sheet="01_CODEBOOK",
             acceptance_criteria=["All columns included", "Formulas correct"], status="pending"),
        
        Task(id="2.1", phase="Data Quality", name="Assess Data Quality",
             objective="Calculate quality metrics with formulas",
             method="Calculate missing %, complete cases using COUNTBLANK formulas",
             formulas=["=COUNTBLANK()", "=COUNT()"],
             validation="Metrics accurate", output_sheet="02_DATA_QUALITY",
             acceptance_criteria=["All metrics formula-based"], status="pending"),
        
        Task(id="2.2", phase="Data Quality", name="Analyze Missing Data",
             objective="Document missing patterns",
             method="For each variable: =COUNTBLANK(range), =COUNTBLANK(range)/ROWS(range)*100",
             formulas=["=COUNTBLANK()"],
             validation="Patterns identified", output_sheet="03_MISSING_ANALYSIS",
             acceptance_criteria=["Per-variable analysis complete"], status="pending"),
        
        Task(id="3.1", phase="Data Cleaning", name="Create Valid Responses",
             objective="Filter rows with >30% missing",
             method="=IF(COUNTBLANK(row)/COLUMNS>0.3, 'EXCLUDE', 'INCLUDE')",
             formulas=["=IF(COUNTBLANK()>0.3)"],
             validation="Exclusions documented", output_sheet="04_VALID_RESPONSES",
             acceptance_criteria=["Exclusion criteria clear"], status="pending"),
        
        Task(id="4.1", phase="Feature Engineering", name="Extract Numeric Data",
             objective="Create numeric-only dataset",
             method="Copy numeric columns to separate sheet",
             formulas=[],
             validation="Only numeric columns", output_sheet="05_CLEAN_NUMERIC",
             acceptance_criteria=["All numeric variables included"], status="pending"),
        
        Task(id="5.1", phase="EDA", name="Descriptive Statistics",
             objective="Calculate M, SD, skew, kurtosis for all numeric variables",
             method="=AVERAGE(range), =STDEV.S(range), =SKEW(range), =KURT(range)",
             formulas=["=AVERAGE()", "=STDEV.S()", "=SKEW()", "=KURT()"],
             validation="All variables analyzed", output_sheet="06_DESCRIPTIVES",
             acceptance_criteria=["All stats formula-based", "APA format"], status="pending"),
        
        Task(id="5.2", phase="EDA", name="Normality Tests",
             objective="Test distribution normality",
             method="Document Shapiro-Wilk results, =SKEW(range), =KURT(range)",
             formulas=["=SKEW()", "=KURT()"],
             validation="All variables tested", output_sheet="07_NORMALITY",
             acceptance_criteria=["Results documented"], status="pending"),
        
        Task(id="6.1", phase="Statistical Testing", name="Scale Reliability",
             objective="Calculate Cronbach's alpha for each scale",
             method="α = (k/(k-1)) * (1 - Σvar_items/var_total)",
             formulas=["=VAR.S()"],
             validation="All scales analyzed", output_sheet="08_RELIABILITY",
             acceptance_criteria=["Alpha values documented"], status="pending"),
        
        Task(id="6.2", phase="Statistical Testing", name="Correlation Matrix",
             objective="Compute correlations between numeric variables",
             method="=CORREL(range1, range2) for each pair",
             formulas=["=CORREL()"],
             validation="Matrix complete", output_sheet="09_CORRELATIONS",
             acceptance_criteria=["All CORREL formulas"], status="pending"),
        
        Task(id="6.3", phase="Statistical Testing", name="Group Comparisons",
             objective="T-tests for categorical groupings",
             method="=T.TEST(group1, group2, 2, 2)",
             formulas=["=T.TEST()"],
             validation="Significant differences flagged", output_sheet="10_GROUP_COMPARISONS",
             acceptance_criteria=["Effect sizes included"], status="pending"),
        
        Task(id="6.4", phase="Statistical Testing", name="Effect Sizes",
             objective="Calculate Cohen's d and eta-squared",
             method="d = (M1-M2)/pooled_SD",
             formulas=[],
             validation="All significant results have effect sizes", output_sheet="11_EFFECT_SIZES",
             acceptance_criteria=["Interpretation included"], status="pending"),
        
        Task(id="7.1", phase="QC", name="Quality Control Review",
             objective="Verify all computations",
             method="Check all sheets for formula integrity",
             formulas=[],
             validation="No hardcoded values", output_sheet="12_QC_REPORT",
             acceptance_criteria=["All checks passed"], status="pending"),
        
        Task(id="8.1", phase="Reporting", name="APA Results",
             objective="Write APA 7th edition results section",
             method="Format all findings in APA style",
             formulas=[],
             validation="APA compliant", output_sheet="13_APA_RESULTS",
             acceptance_criteria=["Publication-ready"], status="pending"),
        
        Task(id="8.2", phase="Reporting", name="Methodology Documentation",
             objective="Document methods for thesis",
             method="Write comprehensive methods section",
             formulas=[],
             validation="Complete and accurate", output_sheet="14_METHODOLOGY",
             acceptance_criteria=["Ready for Chapter 3"], status="pending"),
        
        Task(id="8.3", phase="Reporting", name="Final Audit",
             objective="Comprehensive quality audit",
             method="Score all dimensions, issue certification",
             formulas=[],
             validation="Score >= 97%", output_sheet="15_AUDIT_CERTIFICATE",
             acceptance_criteria=["Publication-ready certification"], status="pending"),
    ]
    
    return tasks
