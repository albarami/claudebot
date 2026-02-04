"""
Main LangGraph workflow for PhD Survey Analyzer.
Assembles the complete multi-agent graph with nodes and edges.
"""

from langgraph.graph import StateGraph, END
from langgraph.checkpoint.memory import MemorySaver

from graph.state import SurveyAnalysisState
from graph.nodes import load_data_node, advance_task_node, generate_deliverables_node
from graph.edges import route_after_qc, route_after_audit, should_continue_tasks
from agents.strategist import strategist_node
from agents.implementer import implementer_node
from agents.qc_reviewer import qc_reviewer_node
from agents.auditor import auditor_node


def create_survey_analysis_workflow():
    """
    Create the complete LangGraph workflow for survey analysis.
    
    Workflow:
    1. load_data → strategist → implementer ↔ qc_reviewer → auditor → deliverables
    
    Returns:
        Compiled LangGraph application with checkpointing
    """
    workflow = StateGraph(SurveyAnalysisState)
    
    workflow.add_node("load_data", load_data_node)
    workflow.add_node("strategist", strategist_node)
    workflow.add_node("implementer", implementer_node)
    workflow.add_node("qc_reviewer", qc_reviewer_node)
    workflow.add_node("advance_task", advance_task_node)
    workflow.add_node("auditor", auditor_node)
    workflow.add_node("deliverables", generate_deliverables_node)
    
    async def error_node(state):
        return {
            "status": "failed",
            "errors": ["Workflow halted due to critical error or max revisions reached"],
            "messages": [{"role": "system", "content": "Workflow failed - check errors"}]
        }
    workflow.add_node("error", error_node)
    
    workflow.set_entry_point("load_data")
    
    workflow.add_edge("load_data", "strategist")
    
    workflow.add_edge("strategist", "implementer")
    
    workflow.add_edge("implementer", "qc_reviewer")
    
    workflow.add_conditional_edges(
        "qc_reviewer",
        route_after_qc,
        {
            "advance_task": "advance_task",
            "implementer": "implementer",
            "auditor": "auditor",
            "error": "error"
        }
    )
    
    workflow.add_conditional_edges(
        "advance_task",
        should_continue_tasks,
        {
            "implementer": "implementer",
            "auditor": "auditor"
        }
    )
    
    workflow.add_conditional_edges(
        "auditor",
        route_after_audit,
        {
            "deliverables": "deliverables",
            "implementer": "implementer"
        }
    )
    
    workflow.add_edge("deliverables", END)
    workflow.add_edge("error", END)
    
    memory = MemorySaver()
    
    app = workflow.compile(checkpointer=memory)
    
    return app


survey_workflow = create_survey_analysis_workflow()
