"""
FastAPI backend for PhD Survey Analyzer.
Provides REST API for the LangGraph multi-agent workflow.
"""

import asyncio
import uuid
from pathlib import Path
from datetime import datetime
from typing import Dict, Any, Optional

from fastapi import FastAPI, UploadFile, File, HTTPException, BackgroundTasks
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from pydantic import BaseModel
import aiofiles

from config import UPLOAD_DIR, OUTPUT_DIR, ANTHROPIC_API_KEY
from graph.state import create_initial_state
from graph.workflow import survey_workflow


app = FastAPI(
    title="PhD Survey Analyzer",
    description="LangGraph-based multi-agent system for PhD-level survey EDA",
    version="2.0.0"
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

sessions: Dict[str, Dict[str, Any]] = {}


class AnalysisRequest(BaseModel):
    """Request to start analysis."""
    session_id: str
    research_questions: Optional[list] = None


class StatusResponse(BaseModel):
    """Analysis status response."""
    session_id: str
    status: str
    progress: float
    current_agent: Optional[str]
    current_task: Optional[str]
    tasks_completed: int
    total_tasks: int
    certification: Optional[str]
    overall_score: Optional[float]
    logs: list
    errors: list
    verification_status: Optional[str] = None
    formula_coverage: Optional[float] = None
    output_type: Optional[str] = None


@app.on_event("startup")
async def startup():
    """Startup event."""
    print("🎓 PhD Survey Analyzer - LangGraph Multi-Agent System")
    print(f"📁 Upload dir: {UPLOAD_DIR}")
    print(f"📁 Output dir: {OUTPUT_DIR}")
    
    if not ANTHROPIC_API_KEY:
        print("⚠️  WARNING: ANTHROPIC_API_KEY not set in .env")
    else:
        print("✅ Anthropic API key configured")


@app.get("/")
async def root():
    """Root endpoint."""
    return {
        "name": "PhD Survey Analyzer",
        "version": "2.0.0",
        "framework": "LangGraph",
        "agents": [
            "Survey Strategist (Claude Opus 4.5)",
            "Survey Implementer (Claude Sonnet 4)",
            "Survey QC Reviewer (Claude Opus 4.5)",
            "Survey Auditor (Claude Opus 4.5)"
        ]
    }


@app.post("/api/upload")
async def upload_file(file: UploadFile = File(...)):
    """
    Upload a survey Excel file.
    
    Args:
        file: Excel file (.xlsx, .xls)
    
    Returns:
        Session ID and file info
    """
    if not file.filename.endswith(('.xlsx', '.xls')):
        raise HTTPException(status_code=400, detail="Only Excel files (.xlsx, .xls) are supported")
    
    session_id = f"session_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    safe_filename = f"{timestamp}_{file.filename}"
    file_path = UPLOAD_DIR / safe_filename
    
    async with aiofiles.open(file_path, 'wb') as f:
        content = await file.read()
        await f.write(content)
    
    sessions[session_id] = {
        "file_path": str(file_path),
        "file_name": file.filename,
        "status": "uploaded",
        "progress": 0,
        "current_agent": None,
        "current_task": None,
        "tasks_completed": 0,
        "total_tasks": 0,
        "logs": [],
        "errors": [],
        "state": None,
        "certification": None,
        "overall_score": None,
        "output_path": None,
        "research_questions": [],  # Will be populated by analyze endpoint
        "verification_status": "pending",
        "formula_coverage": None,
        "output_type": None
    }
    
    return {
        "session_id": session_id,
        "file_name": file.filename,
        "file_path": str(file_path),
        "status": "uploaded"
    }


async def run_analysis(session_id: str):
    """
    Run the full LangGraph analysis workflow.
    
    Args:
        session_id: Session identifier
    """
    session = sessions.get(session_id)
    if not session:
        return
    
    try:
        session["status"] = "running"
        session["logs"].append({
            "timestamp": datetime.now().isoformat(),
            "agent": "system",
            "message": "Starting LangGraph multi-agent workflow..."
        })
        
        initial_state = create_initial_state(
            session_id=session_id,
            file_path=session["file_path"]
        )
        
        # Pass research_questions to state
        if session.get("research_questions"):
            initial_state["research_questions"] = session["research_questions"]
        
        config = {"configurable": {"thread_id": session_id}}
        
        async for event in survey_workflow.astream(initial_state, config):
            for node_name, state_update in event.items():
                if node_name == "__end__":
                    continue
                
                messages = state_update.get("messages", [])
                for msg in messages:
                    session["logs"].append({
                        "timestamp": datetime.now().isoformat(),
                        "agent": msg.get("role", "system"),
                        "message": msg.get("content", "")
                    })
                
                status = state_update.get("status", session["status"])
                session["status"] = status
                
                if "current_task_idx" in state_update:
                    session["tasks_completed"] = state_update["current_task_idx"]
                
                if "total_tasks" in state_update:
                    session["total_tasks"] = state_update["total_tasks"]
                
                if session["total_tasks"] > 0:
                    session["progress"] = (session["tasks_completed"] / session["total_tasks"]) * 100
                
                if "current_task" in state_update and state_update["current_task"]:
                    session["current_task"] = state_update["current_task"].get("name", "")
                    session["current_agent"] = "implementer"
                
                if node_name == "strategist":
                    session["current_agent"] = "strategist"
                elif node_name == "qc_reviewer":
                    session["current_agent"] = "qc_reviewer"
                elif node_name == "auditor":
                    session["current_agent"] = "auditor"
                
                if "certification" in state_update:
                    session["certification"] = state_update["certification"]
                
                if "overall_score" in state_update:
                    session["overall_score"] = state_update["overall_score"]
                
                if "deliverables" in state_update:
                    session["deliverables"] = state_update["deliverables"]
                
                if "errors" in state_update:
                    session["errors"].extend(state_update["errors"])

                if "verification_status" in state_update:
                    session["verification_status"] = state_update["verification_status"]

                if "formula_coverage" in state_update:
                    session["formula_coverage"] = state_update["formula_coverage"]

                if "output_type" in state_update:
                    session["output_type"] = state_update["output_type"]
        
        session["status"] = "completed"
        session["progress"] = 100
        session["logs"].append({
            "timestamp": datetime.now().isoformat(),
            "agent": "system",
            "message": f"✅ Analysis complete! Score: {session.get('overall_score', 0):.1f}% - {session.get('certification', 'N/A')}"
        })
        
    except Exception as e:
        session["status"] = "error"
        session["errors"].append(str(e))
        session["logs"].append({
            "timestamp": datetime.now().isoformat(),
            "agent": "system",
            "message": f"❌ Error: {str(e)}"
        })


@app.post("/api/analyze")
async def start_analysis(request: AnalysisRequest, background_tasks: BackgroundTasks):
    """
    Start the analysis workflow.
    
    Args:
        request: Analysis request with session_id
    
    Returns:
        Confirmation that analysis started
    """
    session = sessions.get(request.session_id)
    if not session:
        raise HTTPException(status_code=404, detail="Session not found")
    
    if session["status"] not in ["uploaded", "error"]:
        raise HTTPException(status_code=400, detail="Analysis already running or completed")
    
    # Store research_questions in session for state initialization
    if request.research_questions:
        session["research_questions"] = request.research_questions
    
    background_tasks.add_task(run_analysis, request.session_id)
    
    return {
        "session_id": request.session_id,
        "status": "started",
        "message": "Analysis workflow started"
    }


@app.get("/api/status/{session_id}")
async def get_status(session_id: str) -> StatusResponse:
    """
    Get analysis status.
    
    Args:
        session_id: Session identifier
    
    Returns:
        Current status of the analysis
    """
    session = sessions.get(session_id)
    if not session:
        raise HTTPException(status_code=404, detail="Session not found")
    
    return StatusResponse(
        session_id=session_id,
        status=session["status"],
        progress=session["progress"],
        current_agent=session["current_agent"],
        current_task=session["current_task"],
        tasks_completed=session["tasks_completed"],
        total_tasks=session["total_tasks"],
        certification=session.get("certification"),
        overall_score=session.get("overall_score"),
        logs=session["logs"][-50:],
        errors=session["errors"],
        verification_status=session.get("verification_status"),
        formula_coverage=session.get("formula_coverage"),
        output_type=session.get("output_type")
    )


@app.get("/api/logs/{session_id}")
async def get_logs(session_id: str):
    """Get all logs for a session."""
    session = sessions.get(session_id)
    if not session:
        raise HTTPException(status_code=404, detail="Session not found")
    
    return {"logs": session["logs"]}


@app.get("/api/download/{session_id}")
async def download_results(session_id: str):
    """
    Download analysis results (audit certificate).
    
    Args:
        session_id: Session identifier
    
    Returns:
        Markdown file with audit certificate
    """
    session = sessions.get(session_id)
    if not session:
        raise HTTPException(status_code=404, detail="Session not found")
    
    if session["status"] != "completed":
        raise HTTPException(status_code=400, detail="Analysis not yet completed")
    
    deliverables = session.get("deliverables", [])
    if deliverables:
        for path in deliverables:
            if Path(path).exists() and path.endswith(".md"):
                return FileResponse(
                    path,
                    media_type="text/markdown",
                    filename=Path(path).name
                )
    
    audit_files = list(OUTPUT_DIR.glob(f"AUDIT_CERTIFICATE_{session_id}*"))
    if audit_files:
        return FileResponse(
            str(audit_files[0]),
            media_type="text/markdown",
            filename=audit_files[0].name
        )
    
    raise HTTPException(status_code=404, detail="No output files found")


@app.get("/api/download-excel/{session_id}")
async def download_excel(session_id: str):
    """
    Download Excel workbook with analysis results.
    Returns .xlsm (macro-enabled) if available, otherwise .xlsx.
    
    Args:
        session_id: Session identifier
    
    Returns:
        Excel workbook file
    """
    session = sessions.get(session_id)
    if not session:
        raise HTTPException(status_code=404, detail="Session not found")
    
    if session["status"] not in ["completed", "running"]:
        raise HTTPException(status_code=400, detail="Analysis not started or failed")
    
    xlsm_files = list(OUTPUT_DIR.glob(f"PhD_EDA_{session_id}*.xlsm"))
    if xlsm_files:
        return FileResponse(
            str(xlsm_files[0]),
            media_type="application/vnd.ms-excel.sheet.macroEnabled.12",
            filename=xlsm_files[0].name,
            headers={"X-Output-Type": "macro-enabled"}
        )
    
    xlsx_files = list(OUTPUT_DIR.glob(f"PhD_EDA_{session_id}*.xlsx"))
    if xlsx_files:
        return FileResponse(
            str(xlsx_files[0]),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename=xlsx_files[0].name,
            headers={"X-Output-Type": "standard"}
        )
    
    raise HTTPException(status_code=404, detail="Excel workbook not found")


@app.get("/api/verification/{session_id}")
async def get_verification_status(session_id: str):
    """
    Get verification status for the analysis.
    
    Args:
        session_id: Session identifier
    
    Returns:
        Verification metrics and status
    """
    session = sessions.get(session_id)
    if not session:
        raise HTTPException(status_code=404, detail="Session not found")
    
    xlsx_files = list(OUTPUT_DIR.glob(f"PhD_EDA_{session_id}*.xlsx"))
    xlsm_files = list(OUTPUT_DIR.glob(f"PhD_EDA_{session_id}*.xlsm"))
    
    workbook_exists = bool(xlsx_files or xlsm_files)
    output_type = "macro-enabled" if xlsm_files else ("standard" if xlsx_files else "none")
    
    return {
        "session_id": session_id,
        "workbook_exists": workbook_exists,
        "output_type": output_type,
        "verification_status": session.get("verification_status", "pending"),
        "formula_coverage": session.get("formula_coverage"),
        "qc_approvals": sum(1 for log in session.get("logs", []) 
                          if "APPROVE" in log.get("message", "")),
        "qc_rejections": sum(1 for log in session.get("logs", []) 
                           if "REJECT" in log.get("message", ""))
    }


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="127.0.0.1", port=8000)
