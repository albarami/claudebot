# OpenClaw Quick Start Guide

## Step 5: Start OpenClaw

```powershell
cd C:\Projects\claudebot

# Load environment variables
$env:OPENCLAW_PASSWORD = "your-strong-password-123"
Get-Content .env | ForEach-Object {
    if ($_ -match '^([^=]+)=(.*)$') {
        Set-Item -Path "env:$($matches[1])" -Value $matches[2]
    }
}

# Start gateway with your config
openclaw gateway start --config openclaw-config.yaml

# You should see:
# ✓ OpenClaw Gateway started
# ✓ Listening on http://127.0.0.1:18789
# ✓ Authentication: enabled (password)
```

## 🔒 Step 6: Security Verification

```powershell
# Check it's localhost only (not exposed to internet)
netstat -an | findstr "18789"
# Should show: 127.0.0.1:18789 (NOT 0.0.0.0:18789)

# Run security audit
openclaw security audit --deep

# Expected output:
# ✓ Gateway bound to localhost
# ✓ Authentication enabled
# ✓ No public exposure detected
# ✓ Firewall: recommended
```

## 🎯 Step 7: Run Your First Analysis

### Option A: Web UI (Easiest)
1. Open browser: http://127.0.0.1:18789
2. Login with your password
3. Upload your Wellbeing Questionnaire.xlsx
4. Type: "Run phd_survey_eda skill on this file"
5. Watch the 4 agents work through 40-60 tasks
6. Download the final 16-sheet workbook

### Option B: Command Line (For automation)

```powershell
# Upload survey file
openclaw files upload "C:\path\to\your\survey.xlsx" --name "wellbeing_survey"

# Run the skill
openclaw run-skill phd_survey_eda --file wellbeing_survey.xlsx --output "C:\Projects\claudebot\output"

# Monitor progress
openclaw workflow status
```

### Option C: Windsurf IDE Integration

In Windsurf, create a script:

```javascript
// survey-analyzer.js
const openclaw = require('openclaw');

async function analyzesurvey(filePath) {
  const session = await openclaw.createSession({
    skill: 'phd_survey_eda',
    agents: ['survey_strategist', 'survey_implementer', 'survey_qc_reviewer', 'survey_auditor']
  });
  
  await session.upload(filePath);
  const result = await session.run();
  
  console.log(`Quality Score: ${result.audit.overall_quality}`);
  console.log(`Output: ${result.output_path}`);
}

analyzesurvey('C:\\Projects\\claudebot\\data\\wellbeing_survey.xlsx');
```

## 📊 What You'll Get (Example Output)

```
C:\Projects\claudebot\output\
└── SURVEY_EDA_COMPLETE_Wellbeing_PhD_20260204.xlsx
    ├── 00_RAW_DATA_LOCKED (175 rows × 69 cols - original data)
    ├── 01_CODEBOOK (variable definitions, types, scales)
    ├── 02_VALID_RESPONSES (171 rows after exclusions)
    ├── 03_DATA_QUALITY (missing patterns, outliers)
    ├── 04_CLEAN_NUMERIC (numeric conversions + recoding)
    ├── 05_MISSING_ANALYSIS (MCAR test results)
    ├── 06_DESCRIPTIVES (M, SD, skew, kurtosis by variable)
    ├── 07_SCALE_RELIABILITY (Cronbach's α for each scale)
    ├── 08_CORRELATIONS (r, p-values, significance stars)
    ├── 09_GROUP_COMPARISONS (gender, religion differences)
    ├── 10_EFFECT_SIZES (Cohen's d, eta-squared)
    ├── 11_VISUALIZATIONS (histograms, scatterplots, heatmaps)
    ├── 12_APA_RESULTS (publication-ready tables)
    ├── 13_METHODOLOGY (full methods section for your thesis)
    ├── 14_AUDIT_CERTIFICATE (Quality: 98.2% - Publication Ready)
    └── 15_EXECUTION_LOG (every step documented)
```

## Quality Metrics Example

```
═══════════════════════════════════════════
  ACADEMIC AUDIT CERTIFICATE
═══════════════════════════════════════════
Survey: Wellbeing Questionnaire
Date: 2026-02-04
Auditor: Agent 4 (Claude Opus 4.5)

QUALITY ASSESSMENT:
├─ Computational Accuracy:      100.0% ✓
├─ Methodological Soundness:     98.5% ✓
├─ Reproducibility:             100.0% ✓
├─ Academic Standards:           97.8% ✓
├─ Documentation Quality:        98.0% ✓
└─ OVERALL QUALITY SCORE:        98.2% ✓

CERTIFICATION: 🏆 PUBLICATION-READY

No critical issues detected.
Analysis suitable for:
- Doctoral dissertation
- Peer-reviewed journal submission
- Conference presentation

Signed: Academic Auditor Agent
═══════════════════════════════════════════
```
