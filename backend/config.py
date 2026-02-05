"""
Configuration for PhD Survey Analyzer.
Multi-model system using Claude (Anthropic) and OpenAI models.
"""

import os
from pathlib import Path
from dotenv import load_dotenv

load_dotenv()

# API Keys
ANTHROPIC_API_KEY = os.getenv("ANTHROPIC_API_KEY")
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")

# Model Configuration - Claude (Anthropic)
OPUS_MODEL = "claude-opus-4-20250514"
SONNET_MODEL = "claude-sonnet-4-20250514"

# Model Configuration - OpenAI
OPENAI_MODEL = "gpt-5.2"  # OpenAI 5.2

# Agent Model Assignments (as specified by user)
STRATEGIST_MODEL = SONNET_MODEL         # Sonnet 4.5 for planning
STRATEGIST_PROVIDER = "anthropic"

IMPLEMENTER_MODEL = OPUS_MODEL          # Opus 4.5 for execution (highest capability)
IMPLEMENTER_PROVIDER = "anthropic"

# QC uses DUAL review: first Sonnet 4.5, then OpenAI 5.2
QC_REVIEWER_MODEL_1 = SONNET_MODEL      # First review: Sonnet 4.5
QC_REVIEWER_PROVIDER_1 = "anthropic"
QC_REVIEWER_MODEL_2 = OPENAI_MODEL      # Second review: OpenAI 5.2
QC_REVIEWER_PROVIDER_2 = "openai"

AUDITOR_MODEL = OPENAI_MODEL            # OpenAI 5.2 for final certification
AUDITOR_PROVIDER = "openai"

# Temperature Settings
STRATEGIST_TEMP = 0.3    # Analytical
IMPLEMENTER_TEMP = 0.1   # Precision mode
QC_REVIEWER_TEMP = 0.2   # Strict
AUDITOR_TEMP = 0.1       # Maximum objectivity

# Token Limits
STRATEGIST_MAX_TOKENS = 8000
IMPLEMENTER_MAX_TOKENS = 8000  # Opus needs more for detailed execution
QC_REVIEWER_MAX_TOKENS = 4000
AUDITOR_MAX_TOKENS = 8000

# Quality Thresholds
PUBLICATION_READY_THRESHOLD = 97.0
THESIS_READY_THRESHOLD = 95.0
NEEDS_REVISION_THRESHOLD = 90.0

# Revision Limits
MAX_TASK_REVISIONS = 10  # Max revisions per task before escalation
MAX_PLAN_REVISIONS = 3   # Max plan revisions

# Excel recalculation (required for UDF verification)
REQUIRE_EXCEL_RECALC = os.getenv("REQUIRE_EXCEL_RECALC", "1") == "1"

# Template usage (disabled by default - dynamic workbooks only)
ALLOW_TEMPLATE = os.getenv("ALLOW_TEMPLATE", "0") == "1"

# Paths
BASE_DIR = Path(__file__).parent.parent
UPLOAD_DIR = BASE_DIR / "uploads"
OUTPUT_DIR = BASE_DIR / "output"

UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)
