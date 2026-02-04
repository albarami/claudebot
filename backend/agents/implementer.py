"""
Agent 2: Survey Implementer
Executes ONE task at a time using ONLY Excel formulas.
Uses Claude Opus 4.5 for highest capability execution.
"""

from typing import Dict, Any
from datetime import datetime

from langchain_anthropic import ChatAnthropic
from langchain_core.messages import HumanMessage, SystemMessage

from config import IMPLEMENTER_MODEL, IMPLEMENTER_TEMP, IMPLEMENTER_MAX_TOKENS, ANTHROPIC_API_KEY, IMPLEMENTER_PROVIDER
from utils.prompts import IMPLEMENTER_SYSTEM_PROMPT
from graph.state import SurveyAnalysisState, LogEntry


async def implementer_node(state: SurveyAnalysisState) -> Dict[str, Any]:
    """
    Implementer agent node - executes current task.
    
    Args:
        state: Current workflow state
    
    Returns:
        State updates with task output
    """
    current_idx = state['current_task_idx']
    tasks = state['tasks']
    
    if current_idx >= len(tasks):
        return {
            "status": "all_tasks_complete",
            "messages": [{"role": "implementer", "content": "All tasks completed"}]
        }
    
    current_task = tasks[current_idx]
    
    llm = ChatAnthropic(
        model=IMPLEMENTER_MODEL,
        temperature=IMPLEMENTER_TEMP,
        max_tokens=IMPLEMENTER_MAX_TOKENS,
        api_key=ANTHROPIC_API_KEY
    )
    
    revision_context = ""
    if state.get('qc_feedback') and state.get('qc_decision') == "REJECT":
        revision_context = f"""
REVISION REQUIRED - Previous attempt was REJECTED.
QC FEEDBACK:
{state['qc_feedback']}

You MUST address all issues identified above.
"""
    
    prompt = f"""Execute this task from the Master Plan:

TASK ID: {current_task['id']}
PHASE: {current_task['phase']}
NAME: {current_task['name']}

OBJECTIVE:
{current_task['objective']}

METHOD:
{current_task['method']}

EXPECTED OUTPUT SHEET: {current_task['output_sheet']}

DATA CONTEXT:
- Raw data sheet: '00_RAW_DATA_LOCKED'
- Number of data rows: {state['n_rows']}
- Columns: {state['columns'][:20]}{'...' if len(state['columns']) > 20 else ''}
- Numeric columns: {state.get('numeric_columns', [])[:15]}

{revision_context}

Execute this task using ONLY Excel formulas. Remember:
1. Every cell must contain a formula (=...)
2. Never hardcode values
3. Reference '00_RAW_DATA_LOCKED' for all data
4. Document every formula used

Provide your output in this format:
TASK COMPLETED: [Task ID] - [Task Name]

ACTIONS TAKEN:
1. [Action]
2. [Action]

FORMULAS USED:
| Cell | Formula | Purpose |
|------|---------|---------|
| B2 | =AVERAGE('00_RAW_DATA_LOCKED'!B2:B{state['n_rows']+1}) | Calculate mean |

SHEET CREATED: [Sheet Name]

READY FOR QC REVIEW"""

    messages = [
        SystemMessage(content=IMPLEMENTER_SYSTEM_PROMPT),
        HumanMessage(content=prompt)
    ]
    
    response = await llm.ainvoke(messages)
    task_output = response.content
    
    updated_task = dict(current_task)
    updated_task['status'] = 'in_review'
    
    updated_tasks = list(tasks)
    updated_tasks[current_idx] = updated_task
    
    log_entry = LogEntry(
        timestamp=datetime.now().isoformat(),
        agent="implementer",
        action=f"Executed task {current_task['id']}",
        details=f"Task: {current_task['name']}",
        task_id=current_task['id']
    )
    
    return {
        "current_task": updated_task,
        "current_task_output": task_output,
        "tasks": updated_tasks,
        "status": "reviewing",
        "execution_log": [log_entry],
        "messages": [{"role": "implementer", "content": f"Completed task {current_task['id']}: {current_task['name']}"}]
    }
