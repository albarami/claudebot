# Excel Macro Integration (No Template Required)

## Overview

The PhD Survey Analyzer generates macro-enabled Excel files (`.xlsm`) dynamically at runtime.
No fixed analysis template is used, because every survey produces different sheets.

Macros (UDFs) are injected automatically from the VBA source in:

```
backend/tools/udf/analysis_udf.bas
```

## Requirements
- Microsoft Excel installed on the host machine
- Trust access to the VBA project object model enabled in Excel
- `pywin32` installed in the Python environment

## UDF Functions Available

| Function | Purpose |
|----------|---------|
| `SHAPIRO_WILK(range)` | Normality test W statistic + p-value |
| `LEVENE_TEST(r1, r2)` | Homogeneity of variance F + p-value |
| `CRONBACH_ALPHA(range)` | Internal consistency alpha |
| `FISHER_Z(r)` | Fisher Z transformation |
| `P_VALUE_T(t, df)` | Two-tailed p for t-test |
| `P_VALUE_F(f, df1, df2)` | p-value for F statistic |
| `COHENS_D(mean1, sd1, n1, mean2, sd2, n2)` | Effect size d |
| `ETA_SQUARED(SSbetween, SStotal)` | ANOVA effect size eta^2 |
| `CRAMERS_V(chi2, n, k)` | Chi-square effect size |
| `CI_MEAN(mean, sd, n, conf)` | Confidence interval for mean |

## Troubleshooting

| Issue | Resolution |
|------|------------|
| `pywin32` missing | Install `pywin32` and retry |
| VBA access denied | Enable "Trust access to VBA project object model" in Excel |
| UDF shows `#NAME?` | Macros are disabled or VBA import failed |

## Notes
- If macro creation fails, the system fails closed by design.
- This preserves academic integrity by preventing partial or unverified output.

