"""
Statistics tools for PhD Survey Analyzer.
These calculate values to inform formula construction.
The actual Excel output uses FORMULAS referencing raw data.
"""

import re
from typing import Dict, List, Any, Tuple, Optional

import pandas as pd
import numpy as np
from scipy import stats


class SurveyDataAnalyzer:
    """Analyzes survey data structure and calculates statistics."""
    
    def __init__(self, df: pd.DataFrame):
        self.df = df
        self.n_rows = len(df)
        self.n_cols = len(df.columns)
        self.columns = list(df.columns)
    
    def get_column_types(self) -> Dict[str, str]:
        """Classify each column by type."""
        types = {}
        for col in self.columns:
            dtype = str(self.df[col].dtype)
            n_unique = self.df[col].nunique()
            
            if dtype in ['int64', 'float64']:
                if n_unique <= 7:
                    types[col] = "ordinal"
                else:
                    types[col] = "numeric"
            elif dtype == 'object':
                if n_unique <= 10:
                    types[col] = "categorical"
                else:
                    types[col] = "text"
            else:
                types[col] = "other"
        
        return types
    
    def get_numeric_columns(self) -> List[str]:
        """Get list of numeric columns."""
        return list(self.df.select_dtypes(include=[np.number]).columns)
    
    def get_categorical_columns(self) -> List[str]:
        """Get list of categorical columns."""
        return list(self.df.select_dtypes(include=['object', 'category']).columns)
    
    def detect_scales(self) -> Dict[str, List[str]]:
        """
        Detect scale patterns from column naming.
        E.g., Faith1, Faith2, Faith3 -> Faith scale
        """
        scales = {}
        pattern = re.compile(r'^([A-Za-z_]+)(\d+)$')
        
        for col in self.columns:
            match = pattern.match(str(col))
            if match:
                scale_name = match.group(1)
                if scale_name not in scales:
                    scales[scale_name] = []
                scales[scale_name].append(col)
        
        return {k: sorted(v) for k, v in scales.items() if len(v) >= 2}
    
    def create_data_summary(self) -> str:
        """Create comprehensive data summary for strategist."""
        lines = [
            "=" * 60,
            "SURVEY DATA SUMMARY",
            "=" * 60,
            "",
            f"DATASET DIMENSIONS:",
            f"  Rows: {self.n_rows}",
            f"  Columns: {self.n_cols}",
            f"  Total cells: {self.n_rows * self.n_cols}",
            "",
            f"MISSING DATA:",
            f"  Total missing: {self.df.isna().sum().sum()}",
            f"  Missing %: {self.df.isna().sum().sum() / (self.n_rows * self.n_cols) * 100:.1f}%",
            "",
            "COLUMN DETAILS:",
        ]
        
        col_types = self.get_column_types()
        for col in self.columns:
            dtype = col_types.get(col, "unknown")
            n_missing = self.df[col].isna().sum()
            n_unique = self.df[col].nunique()
            
            info = f"  - {col}: {dtype}, {n_unique} unique, {n_missing} missing"
            
            if self.df[col].dtype in ['int64', 'float64']:
                min_val = self.df[col].min()
                max_val = self.df[col].max()
                info += f", range [{min_val}-{max_val}]"
            
            lines.append(info)
        
        scales = self.detect_scales()
        if scales:
            lines.extend([
                "",
                "DETECTED SCALES:",
            ])
            for name, items in scales.items():
                lines.append(f"  - {name}: {len(items)} items ({', '.join(items[:3])}{'...' if len(items) > 3 else ''})")
        
        lines.extend([
            "",
            "NUMERIC COLUMNS FOR ANALYSIS:",
            f"  {', '.join(self.get_numeric_columns()[:10])}{'...' if len(self.get_numeric_columns()) > 10 else ''}",
            "",
            "CATEGORICAL COLUMNS:",
            f"  {', '.join(self.get_categorical_columns()[:5])}{'...' if len(self.get_categorical_columns()) > 5 else ''}",
        ])
        
        return "\n".join(lines)
    
    def calculate_descriptives(self, col: str) -> Dict[str, Any]:
        """Calculate descriptive statistics for a column."""
        data = self.df[col].dropna()
        n = len(data)
        
        if n == 0:
            return {"error": "No valid data"}
        
        mean = float(data.mean())
        std = float(data.std())
        se = std / np.sqrt(n) if n > 0 else 0
        
        return {
            "n": n,
            "mean": mean,
            "std": std,
            "se": se,
            "median": float(data.median()),
            "min": float(data.min()),
            "max": float(data.max()),
            "range": float(data.max() - data.min()),
            "q1": float(data.quantile(0.25)),
            "q3": float(data.quantile(0.75)),
            "iqr": float(data.quantile(0.75) - data.quantile(0.25)),
            "skewness": float(stats.skew(data)) if n > 2 else 0,
            "kurtosis": float(stats.kurtosis(data)) if n > 2 else 0,
            "ci_lower": mean - 1.96 * se,
            "ci_upper": mean + 1.96 * se
        }
    
    def test_normality(self, col: str) -> Dict[str, Any]:
        """Run Shapiro-Wilk normality test."""
        data = self.df[col].dropna()
        n = len(data)
        
        if n < 3 or n > 5000:
            return {"error": f"Sample size {n} outside valid range (3-5000)"}
        
        try:
            w, p = stats.shapiro(data)
            return {
                "n": n,
                "w": float(w),
                "p": float(p),
                "is_normal": p > 0.05,
                "interpretation": "Normal (p > .05)" if p > 0.05 else "Non-normal (p ≤ .05)"
            }
        except Exception as e:
            return {"error": str(e)}
    
    def calculate_reliability(self, items: List[str]) -> Dict[str, Any]:
        """Calculate Cronbach's alpha for scale items."""
        valid_items = [i for i in items if i in self.df.columns]
        if len(valid_items) < 2:
            return {"error": "Need at least 2 items"}
        
        data = self.df[valid_items].dropna()
        if len(data) < 10:
            return {"error": "Insufficient cases for reliability"}
        
        k = len(valid_items)
        item_vars = data.var()
        total_var = data.sum(axis=1).var()
        
        if total_var == 0:
            return {"error": "Zero variance in total"}
        
        alpha = (k / (k - 1)) * (1 - item_vars.sum() / total_var)
        
        if alpha >= 0.9:
            interp = "Excellent"
        elif alpha >= 0.8:
            interp = "Good"
        elif alpha >= 0.7:
            interp = "Acceptable"
        elif alpha >= 0.6:
            interp = "Questionable"
        else:
            interp = "Poor"
        
        return {
            "items": valid_items,
            "n_items": k,
            "n_cases": len(data),
            "alpha": float(alpha),
            "interpretation": interp
        }
    
    def calculate_correlation(self, col1: str, col2: str) -> Dict[str, Any]:
        """Calculate Pearson correlation with p-value."""
        data = self.df[[col1, col2]].dropna()
        if len(data) < 3:
            return {"error": "Insufficient data"}
        
        r, p = stats.pearsonr(data[col1], data[col2])
        
        if abs(r) >= 0.5:
            strength = "strong"
        elif abs(r) >= 0.3:
            strength = "moderate"
        else:
            strength = "weak"
        
        direction = "positive" if r > 0 else "negative"
        
        return {
            "r": float(r),
            "p": float(p),
            "n": len(data),
            "significant": p < 0.05,
            "interpretation": f"{strength} {direction}"
        }
    
    def run_ttest(self, dv: str, group_var: str) -> Dict[str, Any]:
        """Run independent samples t-test."""
        groups = self.df[group_var].dropna().unique()
        if len(groups) != 2:
            return {"error": f"Need exactly 2 groups, found {len(groups)}"}
        
        g1 = self.df[self.df[group_var] == groups[0]][dv].dropna()
        g2 = self.df[self.df[group_var] == groups[1]][dv].dropna()
        
        if len(g1) < 5 or len(g2) < 5:
            return {"error": "Groups too small"}
        
        t, p = stats.ttest_ind(g1, g2)
        
        pooled_std = np.sqrt(
            ((len(g1) - 1) * g1.var() + (len(g2) - 1) * g2.var()) /
            (len(g1) + len(g2) - 2)
        )
        d = (g1.mean() - g2.mean()) / pooled_std if pooled_std > 0 else 0
        
        if abs(d) >= 0.8:
            effect_interp = "large"
        elif abs(d) >= 0.5:
            effect_interp = "medium"
        else:
            effect_interp = "small"
        
        return {
            "t": float(t),
            "df": len(g1) + len(g2) - 2,
            "p": float(p),
            "cohens_d": float(d),
            "effect_interpretation": effect_interp,
            "significant": p < 0.05,
            "group1": {"name": str(groups[0]), "n": len(g1), "mean": float(g1.mean()), "sd": float(g1.std())},
            "group2": {"name": str(groups[1]), "n": len(g2), "mean": float(g2.mean()), "sd": float(g2.std())}
        }
    
    def run_anova(self, dv: str, group_var: str) -> Dict[str, Any]:
        """Run one-way ANOVA."""
        groups = self.df.groupby(group_var)[dv].apply(lambda x: x.dropna().tolist())
        groups = [g for g in groups if len(g) >= 5]
        
        if len(groups) < 2:
            return {"error": "Need at least 2 groups with n≥5"}
        
        f, p = stats.f_oneway(*groups)
        
        all_data = np.concatenate(groups)
        grand_mean = all_data.mean()
        ss_between = sum(len(g) * (np.mean(g) - grand_mean) ** 2 for g in groups)
        ss_total = sum((all_data - grand_mean) ** 2)
        eta_sq = ss_between / ss_total if ss_total > 0 else 0
        
        if eta_sq >= 0.14:
            effect_interp = "large"
        elif eta_sq >= 0.06:
            effect_interp = "medium"
        else:
            effect_interp = "small"
        
        return {
            "f": float(f),
            "df_between": len(groups) - 1,
            "df_within": sum(len(g) for g in groups) - len(groups),
            "p": float(p),
            "eta_squared": float(eta_sq),
            "effect_interpretation": effect_interp,
            "significant": p < 0.05,
            "n_groups": len(groups)
        }
    
    def get_missing_analysis(self) -> Dict[str, Any]:
        """Analyze missing data patterns."""
        missing_by_col = self.df.isna().sum().to_dict()
        missing_pct = {col: (n / self.n_rows * 100) for col, n in missing_by_col.items()}
        missing_per_row = self.df.isna().sum(axis=1)
        
        return {
            "total_missing": int(self.df.isna().sum().sum()),
            "total_pct": self.df.isna().sum().sum() / (self.n_rows * self.n_cols) * 100,
            "by_column": missing_by_col,
            "pct_by_column": missing_pct,
            "vars_with_missing": sum(1 for v in missing_by_col.values() if v > 0),
            "complete_cases": int((missing_per_row == 0).sum()),
            "complete_pct": (missing_per_row == 0).sum() / self.n_rows * 100,
            "high_missing_rows": int((missing_per_row > self.n_cols * 0.3).sum())
        }
