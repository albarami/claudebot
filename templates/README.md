# PhD Survey Analyzer - Output Format

## How It Works

Each survey analysis generates a **fresh Excel workbook** (`.xlsx`) tailored to that specific survey's data and structure. There is no static template - every workbook is dynamically created.

## Output Structure

Each generated workbook contains:

```
PhD_EDA_{session_id}.xlsx
├── 00_METADATA          ← Session info, generation timestamp
├── 00_RAW_DATA_LOCKED   ← Original survey data (read-only reference)
├── 01_DATA_AUDIT        ← Data quality checks
├── 02_DESCRIPTIVES      ← Summary statistics (all formulas)
├── 03_CORRELATIONS      ← Correlation matrix
├── ...                  ← Additional analysis sheets
```

## Computation Method

**Excel Formulas + Python Verification**

1. All numeric outputs use Excel formulas (e.g., `=AVERAGE()`, `=STDEV.S()`, `=CORREL()`)
2. Formulas reference the raw data sheet directly
3. Python computes ground-truth values for verification
4. Any mismatch triggers rejection and revision

## Why No Macro Template?

- Each survey has different columns, scales, and analysis needs
- A static template cannot accommodate dynamic survey structures
- Python provides reliable ground-truth for advanced statistics
- Standard `.xlsx` files work everywhere (no macro security issues)

## Advanced Statistics

For tests not native to Excel (Shapiro-Wilk, Cronbach's α, etc.):

- **Python computes the value** in `verification.py`
- Excel shows the result as a labeled value
- Verification report confirms accuracy

## VBA Reference (Optional)

The file `backend/tools/udf/analysis_udf.bas` contains VBA implementations of advanced statistics. These are provided as reference but are **not required** for the system to work. Python handles all verification.
