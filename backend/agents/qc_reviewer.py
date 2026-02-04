"""
Agent 3: QC Reviewer - DUAL REVIEW SYSTEM
Verifies every task meets PhD-level standards with VETO POWER.
Uses BOTH Claude Sonnet 4.5 AND OpenAI 5.2 for maximum accuracy.
Both must agree for approval - if either rejects, task is rejected.
"""

from typing import Dict, Any
from datetime import datetime

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


def build_review_prompt(current_task: Dict, task_output: str, revision_count: int, prev_feedback: str) -> str:
    """Build the review prompt for QC agents."""
    return f"""Review this task execution for PhD-level quality:

TASK SPECIFICATION (from Master Plan):
- Task ID: {current_task['id']}
- Phase: {current_task['phase']}
- Name: {current_task['name']}
- Objective: {current_task['objective']}
- Method: {current_task['method']}
- Expected Output: {current_task['output_sheet']}

IMPLEMENTER'S OUTPUT:
{task_output}

REVISION HISTORY:
This is revision attempt #{revision_count + 1}
{f"Previous QC feedback: {prev_feedback}" if revision_count > 0 else "First submission"}

Verify using the complete checklist:

A. METHODOLOGICAL VERIFICATION
☐ Task executed matches Master Plan specification EXACTLY
☐ Statistical method is appropriate for the data type
☐ Sample size considerations addressed
☐ Effect sizes included where appropriate

B. COMPUTATIONAL ACCURACY
☐ ALL cells contain formulas (starting with "=")
☐ ZERO hardcoded values exist
☐ Formulas reference correct data ranges
☐ Results are plausible

C. DOCUMENTATION QUALITY
☐ Formula documentation is complete
☐ Variable labels are clear
☐ Notes explain any decisions

D. PROFESSIONAL STANDARDS
☐ APA 7th edition formatting
☐ Proper decimal places
☐ Publication-ready appearance

Make your decision: APPROVE, REJECT, CONDITIONAL, or HALT
Provide specific feedback for any issues found."""


def parse_decision(review_text: str) -> str:
    """Parse decision from review text."""
    if "REJECT" in review_text.upper():
        return "REJECT"
    elif "HALT" in review_text.upper():
        return "HALT"
    elif "CONDITIONAL" in review_text.upper():
        return "CONDITIONAL"
    return "APPROVE"


async def qc_reviewer_node(state: SurveyAnalysisState) -> Dict[str, Any]:
    """
    DUAL QC Reviewer - uses BOTH Sonnet 4.5 AND OpenAI 5.2.
    Both reviewers must agree for approval.
    If EITHER rejects, the task is rejected.
    
    Args:
        state: Current workflow state
    
    Returns:
        State updates with QC decision
    """
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
    
    prompt = build_review_prompt(current_task, task_output, revision_count, prev_feedback)
    messages = [
        SystemMessage(content=QC_REVIEWER_SYSTEM_PROMPT),
        HumanMessage(content=prompt)
    ]
    
    # === REVIEW 1: Claude Sonnet 4.5 ===
    llm_sonnet = ChatAnthropic(
        model=QC_REVIEWER_MODEL_1,
        temperature=QC_REVIEWER_TEMP,
        max_tokens=QC_REVIEWER_MAX_TOKENS,
        api_key=ANTHROPIC_API_KEY
    )
    
    response_sonnet = await llm_sonnet.ainvoke(messages)
    review_sonnet = response_sonnet.content
    decision_sonnet = parse_decision(review_sonnet)
    
    # === REVIEW 2: OpenAI 5.2 ===
    llm_openai = ChatOpenAI(
        model=QC_REVIEWER_MODEL_2,
        temperature=QC_REVIEWER_TEMP,
        max_tokens=QC_REVIEWER_MAX_TOKENS,
        api_key=OPENAI_API_KEY
    )
    
    response_openai = await llm_openai.ainvoke(messages)
    review_openai = response_openai.content
    decision_openai = parse_decision(review_openai)
    
    # === DUAL REVIEW LOGIC ===
    # Both must approve for final approval
    # If either rejects, final decision is REJECT
    # If either halts, final decision is HALT
    
    if decision_sonnet == "HALT" or decision_openai == "HALT":
        final_decision = "HALT"
    elif decision_sonnet == "REJECT" or decision_openai == "REJECT":
        final_decision = "REJECT"
    elif decision_sonnet == "CONDITIONAL" or decision_openai == "CONDITIONAL":
        final_decision = "CONDITIONAL"
    else:
        final_decision = "APPROVE"
    
    # Combine feedback from both reviewers
    combined_feedback = f"""
═══════════════════════════════════════════════════════════
DUAL QC REVIEW RESULTS
═══════════════════════════════════════════════════════════

📋 REVIEW 1: Claude Sonnet 4.5
Decision: {decision_sonnet}
{review_sonnet}

───────────────────────────────────────────────────────────

📋 REVIEW 2: OpenAI 5.2
Decision: {decision_openai}
{review_openai}

───────────────────────────────────────────────────────────

🎯 FINAL DECISION: {final_decision}
(Both reviewers must agree for approval)
═══════════════════════════════════════════════════════════
"""
    
    qc_record = QCDecision(
        task_id=current_task['id'],
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
        action=f"Dual review of task {current_task['id']}",
        details=f"Sonnet: {decision_sonnet}, OpenAI: {decision_openai} → Final: {final_decision}",
        task_id=current_task['id']
    )
    
    new_revision_count = revision_count + 1 if final_decision == "REJECT" else 0
    
    return {
        "qc_decision": final_decision,
        "qc_feedback": combined_feedback,
        "qc_history": [qc_record],
        "task_revision_count": new_revision_count,
        "execution_log": [log_entry],
        "messages": [
            {"role": "qc_reviewer_sonnet", "content": f"Sonnet 4.5: {decision_sonnet}"},
            {"role": "qc_reviewer_openai", "content": f"OpenAI 5.2: {decision_openai}"},
            {"role": "qc_reviewer", "content": f"FINAL: {final_decision}"}
        ]
    }
