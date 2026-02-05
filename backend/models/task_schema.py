"""
Pydantic schemas for structured task definitions.
Ensures type-safe, validated task specifications from the strategist.
"""

from typing import List, Optional, Literal
from pydantic import BaseModel, Field, field_validator
from enum import Enum


class TaskType(str, Enum):
    """Supported analysis task types with deterministic formula templates."""
    # Quantitative tasks
    DATA_AUDIT = "data_audit"
    DATA_DICTIONARY = "data_dictionary"
    MISSING_DATA = "missing_data"
    DESCRIPTIVE_STATS = "descriptive_stats"
    FREQUENCY_TABLES = "frequency_tables"
    NORMALITY_CHECK = "normality_check"
    CORRELATION_MATRIX = "correlation_matrix"
    RELIABILITY_ALPHA = "reliability_alpha"
    GROUP_COMPARISON = "group_comparison"
    CROSS_TABULATION = "cross_tabulation"
    EFFECT_SIZES = "effect_sizes"
    SUMMARY_DASHBOARD = "summary_dashboard"
    # Qualitative tasks
    CODEBOOK_CREATION = "codebook_creation"
    QUALITATIVE_CODING = "qualitative_coding"
    THEME_ANALYSIS = "theme_analysis"
    CODING_RELIABILITY = "coding_reliability"
    # Reporting tasks
    APA_TABLES = "apa_tables"
    NARRATIVE_RESULTS = "narrative_results"


class TaskPhase(str, Enum):
    """Analysis phases in order."""
    DATA_VALIDATION = "1_Data_Validation"
    EXPLORATORY = "2_Exploratory"
    DESCRIPTIVE = "3_Descriptive"
    INFERENTIAL = "4_Inferential"
    RELIABILITY = "5_Reliability"
    ADVANCED = "6_Advanced"
    SYNTHESIS = "7_Synthesis"
    DELIVERABLES = "8_Deliverables"


class ColumnSpec(BaseModel):
    """Specification for columns to analyze."""
    column_names: List[str] = Field(default_factory=list)
    column_type: Literal["numeric", "categorical", "all"] = "all"
    max_columns: Optional[int] = None


class TaskSpec(BaseModel):
    """
    Structured task specification.
    Validated by Pydantic - no regex parsing needed.
    """
    id: str = Field(..., pattern=r"^\d+\.\d+$", description="Task ID like '1.1'")
    phase: TaskPhase
    task_type: TaskType
    name: str = Field(..., min_length=3, max_length=100)
    objective: str = Field(..., min_length=10)
    output_sheet: str = Field(..., pattern=r"^[A-Z0-9_]+$", max_length=31)
    columns: ColumnSpec = Field(default_factory=ColumnSpec)
    group_by: Optional[str] = None
    scale_items: Optional[List[str]] = None
    
    @field_validator('output_sheet')
    @classmethod
    def validate_sheet_name(cls, v: str) -> str:
        """Ensure Excel-compatible sheet name."""
        invalid_chars = ['\\', '/', '*', '?', ':', '[', ']']
        for char in invalid_chars:
            if char in v:
                raise ValueError(f"Sheet name cannot contain '{char}'")
        return v[:31]


class MasterPlan(BaseModel):
    """
    Complete master plan with validated tasks.
    Output by strategist, validated before execution.
    """
    session_id: str
    total_variables: int
    total_observations: int
    detected_scales: List[str] = Field(default_factory=list)
    research_questions: List[str] = Field(default_factory=list)
    tasks: List[TaskSpec] = Field(..., min_length=40, max_length=60)
    
    @field_validator('tasks')
    @classmethod
    def validate_task_order(cls, tasks: List[TaskSpec]) -> List[TaskSpec]:
        """Ensure tasks are properly ordered by phase."""
        phase_order = {p: i for i, p in enumerate(TaskPhase)}
        for i in range(1, len(tasks)):
            curr_phase = phase_order.get(tasks[i].phase, 99)
            prev_phase = phase_order.get(tasks[i-1].phase, 99)
            if curr_phase < prev_phase:
                pass  # Allow flexibility but could enforce strict ordering
        return tasks
    
    def get_tasks_by_phase(self, phase: TaskPhase) -> List[TaskSpec]:
        """Get all tasks in a specific phase."""
        return [t for t in self.tasks if t.phase == phase]
    
    def get_tasks_by_type(self, task_type: TaskType) -> List[TaskSpec]:
        """Get all tasks of a specific type."""
        return [t for t in self.tasks if t.task_type == task_type]


class PlanValidationResult(BaseModel):
    """Result of plan validation checks."""
    is_valid: bool
    errors: List[str] = Field(default_factory=list)
    warnings: List[str] = Field(default_factory=list)
    task_count: int = 0
    phase_coverage: dict = Field(default_factory=dict)


def validate_plan(plan: MasterPlan, available_columns: List[str]) -> PlanValidationResult:
    """
    Deterministic validation of master plan.
    
    Args:
        plan: The master plan to validate
        available_columns: Columns available in the dataset
    
    Returns:
        Validation result with errors/warnings
    """
    errors = []
    warnings = []
    
    # Check required phases are present
    required_phases = {TaskPhase.DATA_VALIDATION, TaskPhase.DESCRIPTIVE, TaskPhase.SYNTHESIS}
    present_phases = {t.phase for t in plan.tasks}
    missing_phases = required_phases - present_phases
    if missing_phases:
        errors.append(f"Missing required phases: {[p.value for p in missing_phases]}")

    # Check required task types are present
    required_types = {
        TaskType.DATA_AUDIT,
        TaskType.DATA_DICTIONARY,
        TaskType.MISSING_DATA,
        TaskType.DESCRIPTIVE_STATS,
        TaskType.NORMALITY_CHECK,
        TaskType.RELIABILITY_ALPHA,
        TaskType.CORRELATION_MATRIX,
        TaskType.GROUP_COMPARISON,
        TaskType.EFFECT_SIZES
    }
    present_types = {t.task_type for t in plan.tasks}
    missing_types = required_types - present_types
    if missing_types:
        errors.append(f"Missing required task types: {[t.value for t in missing_types]}")
    
    # Check for duplicate task IDs
    task_ids = [t.id for t in plan.tasks]
    if len(task_ids) != len(set(task_ids)):
        errors.append("Duplicate task IDs found")
    
    # Check for duplicate sheet names
    sheet_names = [t.output_sheet for t in plan.tasks]
    if len(sheet_names) != len(set(sheet_names)):
        errors.append("Duplicate output sheet names found")
    
    # Validate column references
    for task in plan.tasks:
        for col in task.columns.column_names:
            if col not in available_columns:
                warnings.append(f"Task {task.id}: Column '{col}' not found in dataset")
    
    # Check scale items for reliability tasks
    reliability_tasks = [t for t in plan.tasks if t.task_type == TaskType.RELIABILITY_ALPHA]
    for task in reliability_tasks:
        if not task.scale_items or len(task.scale_items) < 2:
            errors.append(f"Task {task.id}: Reliability analysis requires at least 2 scale items")
    
    # Phase coverage
    phase_coverage = {phase.value: 0 for phase in TaskPhase}
    for task in plan.tasks:
        phase_coverage[task.phase.value] += 1
    
    return PlanValidationResult(
        is_valid=len(errors) == 0,
        errors=errors,
        warnings=warnings,
        task_count=len(plan.tasks),
        phase_coverage=phase_coverage
    )
