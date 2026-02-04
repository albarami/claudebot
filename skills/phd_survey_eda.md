# PhD-Level Survey EDA Skill

## Objective
Perform comprehensive, academically rigorous Exploratory Data Analysis on survey data suitable for doctoral research and peer-reviewed publication.

## Input
- Excel file (.xlsx) with survey responses
- Research questions (optional - will infer if not provided)

## Process

### Phase 1: Strategic Planning (Agent 1)
1. Load and inspect survey structure
2. Identify variable types (demographic, scale items, outcomes)
3. Detect scales/subscales from column naming patterns
4. Create 40-60 task Master Plan covering all analysis phases
5. Output: MASTER_PLAN.md

### Phase 2: Iterative Execution (Agent 2 + Agent 3)
FOR EACH TASK in Master Plan:
  - Agent 2: Execute task using Excel formulas only
  - Agent 2: Document in EXECUTION_LOG.md
  - Agent 2: Submit for review
  - Agent 3: Audit against checklist
  - IF approved → proceed to next task
  - IF rejected → Agent 2 revises with feedback
  - REPEAT until all tasks approved

### Phase 3: Final Academic Audit (Agent 4)
1. Review entire workbook
2. Verify methodological soundness
3. Check computational accuracy
4. Assess publication readiness
5. Generate quality metrics
6. Issue certification or revision requests

## Output
Single Excel workbook with 14-16+ sheets:
- Raw data (locked)
- Codebook with variable definitions
- Cleaned numeric data
- Data quality assessment
- Descriptive statistics
- Reliability analysis (Cronbach's alpha)
- Correlation matrix with significance
- Group comparisons (t-tests, ANOVA, chi-square)
- Effect sizes
- Professional visualizations
- APA-formatted results
- Full methodology documentation
- Academic audit certificate
- Complete execution log

## Quality Standards
- 100% formula-based (no hardcoding)
- Publication-ready formatting
- APA 7th edition compliance
- Full reproducibility
- ≥97% overall quality score

## Estimated Time
2-6 hours depending on:
- Number of variables (50-150)
- Sample size (100-5000)
- Complexity of scales
- Number of group comparisons

## Estimated Cost
$10-25 per survey (using Opus 4.5 for quality)
