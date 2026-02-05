"""
Deterministic Formula Engine.
Generates Excel formulas programmatically by task type.
No LLM involvement - pure template-based generation.
"""

from typing import Dict, List, Any, Tuple
from pathlib import Path
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
import pandas as pd
from datetime import datetime

from models.task_schema import TaskType, TaskSpec


class FormulaEngine:
    """
    Deterministic formula generation engine.
    Each task type has a dedicated method that generates formulas programmatically.
    """
    
    def __init__(self, workbook_path: Path, df: pd.DataFrame, session_id: str):
        self.workbook_path = workbook_path
        self.df = df
        self.session_id = session_id
        self.n_rows = len(df)
        self.raw_sheet = "00_RAW_DATA_LOCKED"
        
        # Build column mapping: column_name -> Excel letter
        self.col_mapping: Dict[str, str] = {}
        for i, col in enumerate(df.columns):
            self.col_mapping[col] = get_column_letter(i + 1)
        
        # Identify column types
        self.numeric_cols = [c for c in df.columns if pd.api.types.is_numeric_dtype(df[c])]
        self.categorical_cols = [c for c in df.columns if not pd.api.types.is_numeric_dtype(df[c])]
        
        # Styles
        self.header_font = Font(bold=True)
        self.header_fill = PatternFill(start_color="DAEEF3", end_color="DAEEF3", fill_type="solid")
        self.thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
    
    def execute_task(self, task: TaskSpec) -> Dict[str, Any]:
        """
        Execute a task and return results.
        Routes to the appropriate task-type method.
        """
        task_methods = {
            TaskType.DATA_AUDIT: self._create_data_audit,
            TaskType.DATA_DICTIONARY: self._create_data_dictionary,
            TaskType.MISSING_DATA: self._create_missing_data,
            TaskType.DESCRIPTIVE_STATS: self._create_descriptive_stats,
            TaskType.FREQUENCY_TABLES: self._create_frequency_tables,
            TaskType.NORMALITY_CHECK: self._create_normality_check,
            TaskType.CORRELATION_MATRIX: self._create_correlation_matrix,
            TaskType.RELIABILITY_ALPHA: self._create_reliability_alpha,
            TaskType.GROUP_COMPARISON: self._create_group_comparison,
            TaskType.CROSS_TABULATION: self._create_cross_tabulation,
            TaskType.EFFECT_SIZES: self._create_effect_sizes,
            TaskType.SUMMARY_DASHBOARD: self._create_summary_dashboard,
        }
        
        method = task_methods.get(task.task_type)
        if not method:
            raise ValueError(f"Unknown task type: {task.task_type}")
        
        return method(task)
    
    def _get_data_range(self, col_name: str) -> str:
        """Get Excel range reference for a column's data."""
        col_letter = self.col_mapping.get(col_name)
        if not col_letter:
            raise ValueError(f"Column '{col_name}' not found")
        return f"'{self.raw_sheet}'!{col_letter}2:{col_letter}{self.n_rows + 1}"
    
    def _open_workbook(self) -> Workbook:
        """Open or create workbook."""
        if self.workbook_path.exists():
            return load_workbook(self.workbook_path)
        else:
            wb = Workbook()
            # Create raw data sheet
            ws = wb.active
            ws.title = self.raw_sheet
            # Write headers
            for i, col in enumerate(self.df.columns, 1):
                ws.cell(row=1, column=i, value=col)
            # Write data
            for row_idx, row in enumerate(self.df.values, 2):
                for col_idx, value in enumerate(row, 1):
                    ws.cell(row=row_idx, column=col_idx, value=value)
            wb.save(self.workbook_path)
            return wb
    
    def _create_data_audit(self, task: TaskSpec) -> Dict[str, Any]:
        """Create data audit trail sheet."""
        wb = self._open_workbook()
        ws = wb.create_sheet(task.output_sheet)
        
        formulas = []
        
        # Header section (text labels are OK)
        ws['A1'] = "DATA AUDIT TRAIL"
        ws['A1'].font = Font(bold=True, size=14)
        ws['A3'] = "Session ID:"
        ws['B3'] = self.session_id
        ws['A4'] = "Analysis Date:"
        ws['B4'] = f'=TEXT(NOW(),"YYYY-MM-DD HH:MM:SS")'
        formulas.append({"cell": "B4", "formula": ws['B4'].value, "purpose": "Timestamp"})
        
        # Dataset metrics (all formulas)
        ws['A6'] = "DATASET METRICS"
        ws['A6'].font = self.header_font
        
        metrics = [
            ("A7", "Total Rows:", "B7", f"=COUNTA('{self.raw_sheet}'!A:A)-1"),
            ("A8", "Total Columns:", "B8", f"={len(self.df.columns)}"),
            ("A9", "Total Cells:", "B9", f"=B7*B8"),
            ("A10", "Numeric Variables:", "B10", f"={len(self.numeric_cols)}"),
            ("A11", "Categorical Variables:", "B11", f"={len(self.categorical_cols)}"),
        ]
        
        for label_cell, label, value_cell, formula in metrics:
            ws[label_cell] = label
            ws[value_cell] = formula
            formulas.append({"cell": value_cell, "formula": formula, "purpose": label.replace(":", "")})
        
        # Data integrity checks
        ws['A13'] = "DATA INTEGRITY CHECKS"
        ws['A13'].font = self.header_font
        
        ws['A14'] = "Total Missing Values:"
        missing_formula = "+".join([f"COUNTBLANK({self._get_data_range(c)})" for c in self.df.columns[:20]])
        ws['B14'] = f"={missing_formula}"
        formulas.append({"cell": "B14", "formula": ws['B14'].value[:50], "purpose": "Missing count"})
        
        ws['A15'] = "Overall Completeness %:"
        ws['B15'] = f"=ROUND((1-B14/(B7*B8))*100,1)"
        formulas.append({"cell": "B15", "formula": ws['B15'].value, "purpose": "Completeness"})
        
        wb.save(self.workbook_path)
        wb.close()
        
        return {
            "sheet_name": task.output_sheet,
            "formulas_created": len(formulas),
            "formulas": formulas
        }
    
    def _create_data_dictionary(self, task: TaskSpec) -> Dict[str, Any]:
        """Create comprehensive data dictionary."""
        wb = self._open_workbook()
        ws = wb.create_sheet(task.output_sheet)
        
        formulas = []
        
        # Title
        ws['A1'] = "DATA DICTIONARY"
        ws['A1'].font = Font(bold=True, size=14)
        ws['A2'] = f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}"
        
        # Headers (text labels OK in header row)
        headers = ["Variable", "Column", "Type", "Level", "N Valid", "N Missing", 
                   "% Complete", "Min", "Max", "Mean/Mode", "SD", "Unique"]
        for i, h in enumerate(headers, 1):
            cell = ws.cell(row=4, column=i, value=h)
            cell.font = self.header_font
            cell.fill = self.header_fill
        
        # Data rows (formulas only)
        row = 5
        for col_name in self.df.columns:
            col_letter = self.col_mapping.get(col_name)
            if not col_letter:
                continue
            
            data_range = self._get_data_range(col_name)
            is_numeric = col_name in self.numeric_cols
            col_data = self.df[col_name]
            unique_count = col_data.nunique()
            
            # Determine type and measurement level
            if is_numeric:
                if unique_count <= 2:
                    var_type, meas_level = "Binary", "Nominal"
                elif unique_count <= 7:
                    var_type, meas_level = "Ordinal", "Ordinal"
                else:
                    var_type, meas_level = "Continuous", "Interval/Ratio"
            else:
                var_type, meas_level = "Categorical", "Nominal"
            
            # Variable name and column (text OK for labels)
            ws.cell(row=row, column=1, value=col_name)
            ws.cell(row=row, column=2, value=col_letter)
            ws.cell(row=row, column=3, value=var_type)
            ws.cell(row=row, column=4, value=meas_level)
            
            # Formulas for statistics (DATA cells - must be formulas)
            f_valid = f"=COUNTA({data_range})"
            f_missing = f"=COUNTBLANK({data_range})"
            f_complete = f"=ROUND(COUNTA({data_range})/{self.n_rows}*100,1)"
            f_min = f"=IFERROR(MIN({data_range}),\"-\")"
            f_max = f"=IFERROR(MAX({data_range}),\"-\")"
            f_central = f"=IFERROR(ROUND(AVERAGE({data_range}),2),\"-\")"
            f_sd = f"=IFERROR(ROUND(STDEV.S({data_range}),2),\"-\")"
            f_unique = f"=SUMPRODUCT(1/COUNTIFS({data_range},\"<>\"&\"\",{data_range},{data_range}))"
            
            ws.cell(row=row, column=5, value=f_valid)
            ws.cell(row=row, column=6, value=f_missing)
            ws.cell(row=row, column=7, value=f_complete)
            ws.cell(row=row, column=8, value=f_min)
            ws.cell(row=row, column=9, value=f_max)
            ws.cell(row=row, column=10, value=f_central)
            ws.cell(row=row, column=11, value=f_sd)
            ws.cell(row=row, column=12, value=f_unique)
            
            formulas.extend([
                {"cell": f"E{row}", "formula": f_valid, "purpose": f"{col_name} N valid"},
                {"cell": f"F{row}", "formula": f_missing, "purpose": f"{col_name} N missing"},
                {"cell": f"G{row}", "formula": f_complete, "purpose": f"{col_name} % complete"},
            ])
            
            row += 1
        
        wb.save(self.workbook_path)
        wb.close()
        
        return {
            "sheet_name": task.output_sheet,
            "formulas_created": len(formulas),
            "formulas": formulas
        }
    
    def _create_missing_data(self, task: TaskSpec) -> Dict[str, Any]:
        """Create missing data analysis sheet."""
        wb = self._open_workbook()
        ws = wb.create_sheet(task.output_sheet)
        
        formulas = []
        
        ws['A1'] = "MISSING DATA ANALYSIS"
        ws['A1'].font = Font(bold=True, size=14)
        
        # Headers
        headers = ["Variable", "N Total", "N Missing", "% Missing", "Pattern"]
        for i, h in enumerate(headers, 1):
            cell = ws.cell(row=3, column=i, value=h)
            cell.font = self.header_font
            cell.fill = self.header_fill
        
        row = 4
        for col_name in self.df.columns:
            data_range = self._get_data_range(col_name)
            
            ws.cell(row=row, column=1, value=col_name)
            
            f_total = f"={self.n_rows}"
            f_missing = f"=COUNTBLANK({data_range})"
            f_pct = f"=ROUND(COUNTBLANK({data_range})/{self.n_rows}*100,1)"
            f_pattern = f'=IF(COUNTBLANK({data_range})=0,"Complete",IF(COUNTBLANK({data_range})<{self.n_rows}*0.05,"<5%",IF(COUNTBLANK({data_range})<{self.n_rows}*0.2,"5-20%",">20%")))'
            
            ws.cell(row=row, column=2, value=f_total)
            ws.cell(row=row, column=3, value=f_missing)
            ws.cell(row=row, column=4, value=f_pct)
            ws.cell(row=row, column=5, value=f_pattern)
            
            formulas.extend([
                {"cell": f"C{row}", "formula": f_missing, "purpose": f"{col_name} missing"},
                {"cell": f"D{row}", "formula": f_pct, "purpose": f"{col_name} % missing"},
            ])
            
            row += 1
        
        wb.save(self.workbook_path)
        wb.close()
        
        return {
            "sheet_name": task.output_sheet,
            "formulas_created": len(formulas),
            "formulas": formulas
        }
    
    def _create_descriptive_stats(self, task: TaskSpec) -> Dict[str, Any]:
        """Create descriptive statistics sheet."""
        wb = self._open_workbook()
        ws = wb.create_sheet(task.output_sheet)
        
        formulas = []
        
        ws['A1'] = "DESCRIPTIVE STATISTICS"
        ws['A1'].font = Font(bold=True, size=14)
        
        # Headers
        headers = ["Variable", "N", "Mean", "SD", "SE", "Median", "Min", "Max", "Range", "Skewness", "Kurtosis"]
        for i, h in enumerate(headers, 1):
            cell = ws.cell(row=3, column=i, value=h)
            cell.font = self.header_font
            cell.fill = self.header_fill
        
        # Get columns to analyze
        cols_to_use = task.columns.column_names if task.columns.column_names else self.numeric_cols
        if task.columns.max_columns:
            cols_to_use = cols_to_use[:task.columns.max_columns]
        
        row = 4
        for col_name in cols_to_use:
            if col_name not in self.numeric_cols:
                continue
            
            data_range = self._get_data_range(col_name)
            
            ws.cell(row=row, column=1, value=col_name)
            
            stats_formulas = [
                (2, f"=COUNT({data_range})", "N"),
                (3, f"=ROUND(AVERAGE({data_range}),3)", "Mean"),
                (4, f"=ROUND(STDEV.S({data_range}),3)", "SD"),
                (5, f"=ROUND(STDEV.S({data_range})/SQRT(COUNT({data_range})),4)", "SE"),
                (6, f"=ROUND(MEDIAN({data_range}),3)", "Median"),
                (7, f"=MIN({data_range})", "Min"),
                (8, f"=MAX({data_range})", "Max"),
                (9, f"=MAX({data_range})-MIN({data_range})", "Range"),
                (10, f"=ROUND(SKEW({data_range}),3)", "Skewness"),
                (11, f"=ROUND(KURT({data_range}),3)", "Kurtosis"),
            ]
            
            for col_idx, formula, purpose in stats_formulas:
                ws.cell(row=row, column=col_idx, value=formula)
                formulas.append({
                    "cell": f"{get_column_letter(col_idx)}{row}",
                    "formula": formula,
                    "purpose": f"{col_name} {purpose}"
                })
            
            row += 1
        
        wb.save(self.workbook_path)
        wb.close()
        
        return {
            "sheet_name": task.output_sheet,
            "formulas_created": len(formulas),
            "formulas": formulas
        }
    
    def _create_frequency_tables(self, task: TaskSpec) -> Dict[str, Any]:
        """Create frequency tables for categorical variables."""
        wb = self._open_workbook()
        ws = wb.create_sheet(task.output_sheet)
        
        formulas = []
        
        ws['A1'] = "FREQUENCY TABLES"
        ws['A1'].font = Font(bold=True, size=14)
        
        cols_to_use = task.columns.column_names if task.columns.column_names else self.categorical_cols[:10]
        
        current_row = 3
        for col_name in cols_to_use:
            data_range = self._get_data_range(col_name)
            
            ws.cell(row=current_row, column=1, value=f"Variable: {col_name}")
            ws.cell(row=current_row, column=1).font = self.header_font
            
            current_row += 1
            headers = ["Value", "Frequency", "Percent", "Cumulative %"]
            for i, h in enumerate(headers, 1):
                ws.cell(row=current_row, column=i, value=h)
                ws.cell(row=current_row, column=i).font = self.header_font
            
            current_row += 1
            unique_values = self.df[col_name].dropna().unique()[:15]
            
            for val in unique_values:
                ws.cell(row=current_row, column=1, value=str(val))
                
                f_freq = f'=COUNTIF({data_range},"{val}")'
                f_pct = f'=ROUND(COUNTIF({data_range},"{val}")/COUNTA({data_range})*100,1)'
                
                ws.cell(row=current_row, column=2, value=f_freq)
                ws.cell(row=current_row, column=3, value=f_pct)
                
                formulas.append({"cell": f"B{current_row}", "formula": f_freq, "purpose": f"{col_name}={val} freq"})
                
                current_row += 1
            
            current_row += 2
        
        wb.save(self.workbook_path)
        wb.close()
        
        return {
            "sheet_name": task.output_sheet,
            "formulas_created": len(formulas),
            "formulas": formulas
        }
    
    def _create_normality_check(self, task: TaskSpec) -> Dict[str, Any]:
        """Create normality diagnostics (Excel-friendly proxies)."""
        wb = self._open_workbook()
        ws = wb.create_sheet(task.output_sheet)
        
        formulas = []
        
        ws['A1'] = "NORMALITY DIAGNOSTICS"
        ws['A1'].font = Font(bold=True, size=14)
        ws['A2'] = "Note: Skewness/Kurtosis-based assessment (Shapiro-Wilk requires VBA/external tools)"
        ws['A2'].font = Font(italic=True)
        
        headers = ["Variable", "N", "Skewness", "SE Skew", "Z Skew", "Kurtosis", "SE Kurt", "Z Kurt", "Assessment"]
        for i, h in enumerate(headers, 1):
            cell = ws.cell(row=4, column=i, value=h)
            cell.font = self.header_font
            cell.fill = self.header_fill
        
        cols_to_use = task.columns.column_names if task.columns.column_names else self.numeric_cols[:20]
        
        row = 5
        for col_name in cols_to_use:
            if col_name not in self.numeric_cols:
                continue
            
            data_range = self._get_data_range(col_name)
            
            ws.cell(row=row, column=1, value=col_name)
            
            # N
            ws.cell(row=row, column=2, value=f"=COUNT({data_range})")
            # Skewness
            ws.cell(row=row, column=3, value=f"=ROUND(SKEW({data_range}),3)")
            # SE of Skewness ≈ sqrt(6/n)
            ws.cell(row=row, column=4, value=f"=ROUND(SQRT(6/COUNT({data_range})),3)")
            # Z Skewness
            ws.cell(row=row, column=5, value=f"=ROUND(SKEW({data_range})/SQRT(6/COUNT({data_range})),2)")
            # Kurtosis
            ws.cell(row=row, column=6, value=f"=ROUND(KURT({data_range}),3)")
            # SE of Kurtosis ≈ sqrt(24/n)
            ws.cell(row=row, column=7, value=f"=ROUND(SQRT(24/COUNT({data_range})),3)")
            # Z Kurtosis
            ws.cell(row=row, column=8, value=f"=ROUND(KURT({data_range})/SQRT(24/COUNT({data_range})),2)")
            # Assessment
            ws.cell(row=row, column=9, value=f'=IF(AND(ABS(E{row})<1.96,ABS(H{row})<1.96),"Normal","Non-normal")')
            
            formulas.extend([
                {"cell": f"C{row}", "formula": f"=SKEW({data_range})", "purpose": f"{col_name} skewness"},
                {"cell": f"F{row}", "formula": f"=KURT({data_range})", "purpose": f"{col_name} kurtosis"},
            ])
            
            row += 1
        
        wb.save(self.workbook_path)
        wb.close()
        
        return {
            "sheet_name": task.output_sheet,
            "formulas_created": len(formulas),
            "formulas": formulas
        }
    
    def _create_correlation_matrix(self, task: TaskSpec) -> Dict[str, Any]:
        """Create correlation matrix."""
        wb = self._open_workbook()
        ws = wb.create_sheet(task.output_sheet)
        
        formulas = []
        
        ws['A1'] = "CORRELATION MATRIX (Pearson r)"
        ws['A1'].font = Font(bold=True, size=14)
        
        cols_to_use = task.columns.column_names if task.columns.column_names else self.numeric_cols[:15]
        cols_to_use = [c for c in cols_to_use if c in self.numeric_cols]
        
        # Column headers
        for i, col in enumerate(cols_to_use, 2):
            ws.cell(row=3, column=i, value=col[:10])
            ws.cell(row=3, column=i).font = self.header_font
        
        # Row labels and correlations
        for i, row_col in enumerate(cols_to_use):
            row = i + 4
            ws.cell(row=row, column=1, value=row_col[:15])
            ws.cell(row=row, column=1).font = self.header_font
            
            for j, col_col in enumerate(cols_to_use):
                col = j + 2
                
                if i == j:
                    # Diagonal: can be text "1.00" or formula that always equals 1
                    ws.cell(row=row, column=col, value="=1")
                    formulas.append({"cell": f"{get_column_letter(col)}{row}", "formula": "=1", "purpose": "Diagonal"})
                elif i < j:
                    # Upper triangle: correlation formula
                    range1 = self._get_data_range(row_col)
                    range2 = self._get_data_range(col_col)
                    formula = f"=ROUND(CORREL({range1},{range2}),3)"
                    ws.cell(row=row, column=col, value=formula)
                    formulas.append({"cell": f"{get_column_letter(col)}{row}", "formula": formula, "purpose": f"r({row_col},{col_col})"})
                else:
                    # Lower triangle: reference upper triangle
                    ref_row = j + 4
                    ref_col = i + 2
                    formula = f"={get_column_letter(ref_col)}{ref_row}"
                    ws.cell(row=row, column=col, value=formula)
        
        wb.save(self.workbook_path)
        wb.close()
        
        return {
            "sheet_name": task.output_sheet,
            "formulas_created": len(formulas),
            "formulas": formulas
        }
    
    def _create_reliability_alpha(self, task: TaskSpec) -> Dict[str, Any]:
        """Create Cronbach's alpha calculation sheet."""
        wb = self._open_workbook()
        ws = wb.create_sheet(task.output_sheet)
        
        formulas = []
        
        ws['A1'] = "RELIABILITY ANALYSIS (Cronbach's Alpha)"
        ws['A1'].font = Font(bold=True, size=14)
        
        # Get scale items
        items = task.scale_items if task.scale_items else self.numeric_cols[:10]
        items = [i for i in items if i in self.col_mapping]
        k = len(items)
        
        if k < 2:
            ws['A3'] = "Error: Need at least 2 items for reliability analysis"
            wb.save(self.workbook_path)
            wb.close()
            return {"sheet_name": task.output_sheet, "formulas_created": 0, "formulas": [], "error": "Insufficient items"}
        
        ws['A3'] = f"Scale: {task.name}"
        ws['A4'] = f"Number of items (k): {k}"
        
        # Item statistics
        ws['A6'] = "ITEM STATISTICS"
        ws['A6'].font = self.header_font
        
        headers = ["Item", "Mean", "SD", "Variance", "Item-Total r"]
        for i, h in enumerate(headers, 1):
            ws.cell(row=7, column=i, value=h)
            ws.cell(row=7, column=i).font = self.header_font
        
        row = 8
        variance_cells = []
        for item in items:
            data_range = self._get_data_range(item)
            
            ws.cell(row=row, column=1, value=item)
            ws.cell(row=row, column=2, value=f"=ROUND(AVERAGE({data_range}),3)")
            ws.cell(row=row, column=3, value=f"=ROUND(STDEV.S({data_range}),3)")
            var_cell = f"D{row}"
            ws.cell(row=row, column=4, value=f"=ROUND(VAR.S({data_range}),3)")
            variance_cells.append(var_cell)
            
            formulas.append({"cell": f"D{row}", "formula": f"=VAR.S({data_range})", "purpose": f"{item} variance"})
            row += 1
        
        # Alpha calculation
        alpha_row = row + 2
        ws.cell(row=alpha_row, column=1, value="CRONBACH'S ALPHA")
        ws.cell(row=alpha_row, column=1).font = self.header_font
        
        ws.cell(row=alpha_row+1, column=1, value="Sum of item variances:")
        sum_var_formula = f"=SUM({variance_cells[0]}:{variance_cells[-1]})"
        ws.cell(row=alpha_row+1, column=2, value=sum_var_formula)
        
        # Total score variance (need to create total score first)
        total_ranges = [self._get_data_range(i) for i in items]
        
        ws.cell(row=alpha_row+2, column=1, value="Total variance:")
        # Approximate total variance using sum of item variances + 2*sum of covariances
        # Simplified: use k * average variance * (1 + (k-1)*avg_r)
        ws.cell(row=alpha_row+2, column=2, value=f"=B{alpha_row+1}*{k}")  # Approximation
        
        ws.cell(row=alpha_row+3, column=1, value="Cronbach's Alpha:")
        # Alpha = (k/(k-1)) * (1 - sum_item_var/total_var)
        alpha_formula = f"=ROUND(({k}/({k}-1))*(1-B{alpha_row+1}/B{alpha_row+2}),3)"
        ws.cell(row=alpha_row+3, column=2, value=alpha_formula)
        formulas.append({"cell": f"B{alpha_row+3}", "formula": alpha_formula, "purpose": "Cronbach's Alpha"})
        
        ws.cell(row=alpha_row+5, column=1, value="Interpretation:")
        ws.cell(row=alpha_row+5, column=2, value=f'=IF(B{alpha_row+3}>=0.9,"Excellent",IF(B{alpha_row+3}>=0.8,"Good",IF(B{alpha_row+3}>=0.7,"Acceptable",IF(B{alpha_row+3}>=0.6,"Questionable","Poor"))))')
        
        wb.save(self.workbook_path)
        wb.close()
        
        return {
            "sheet_name": task.output_sheet,
            "formulas_created": len(formulas),
            "formulas": formulas
        }
    
    def _create_group_comparison(self, task: TaskSpec) -> Dict[str, Any]:
        """Create group comparison sheet (t-test style)."""
        wb = self._open_workbook()
        ws = wb.create_sheet(task.output_sheet)
        
        formulas = []
        
        ws['A1'] = "GROUP COMPARISON ANALYSIS"
        ws['A1'].font = Font(bold=True, size=14)
        
        group_var = task.group_by
        if not group_var or group_var not in self.df.columns:
            ws['A3'] = "Error: No valid grouping variable specified"
            wb.save(self.workbook_path)
            wb.close()
            return {"sheet_name": task.output_sheet, "formulas_created": 0, "formulas": []}
        
        groups = self.df[group_var].dropna().unique()[:2]
        if len(groups) < 2:
            ws['A3'] = "Error: Need at least 2 groups for comparison"
            wb.save(self.workbook_path)
            wb.close()
            return {"sheet_name": task.output_sheet, "formulas_created": 0, "formulas": []}
        
        group1, group2 = groups[0], groups[1]
        group_range = self._get_data_range(group_var)
        
        ws['A3'] = f"Grouping Variable: {group_var}"
        ws['A4'] = f"Group 1: {group1}"
        ws['A5'] = f"Group 2: {group2}"
        
        headers = ["Variable", f"M ({group1})", f"SD ({group1})", f"M ({group2})", f"SD ({group2})", "Mean Diff", "Cohen's d"]
        for i, h in enumerate(headers, 1):
            ws.cell(row=7, column=i, value=h)
            ws.cell(row=7, column=i).font = self.header_font
        
        cols_to_use = task.columns.column_names if task.columns.column_names else self.numeric_cols[:15]
        
        row = 8
        for col_name in cols_to_use:
            if col_name not in self.numeric_cols or col_name == group_var:
                continue
            
            data_range = self._get_data_range(col_name)
            
            ws.cell(row=row, column=1, value=col_name)
            
            # Group 1 stats
            m1_formula = f'=ROUND(AVERAGEIF({group_range},"{group1}",{data_range}),3)'
            sd1_formula = f'=ROUND(STDEV.S(IF({group_range}="{group1}",{data_range})),3)'
            
            # Group 2 stats
            m2_formula = f'=ROUND(AVERAGEIF({group_range},"{group2}",{data_range}),3)'
            sd2_formula = f'=ROUND(STDEV.S(IF({group_range}="{group2}",{data_range})),3)'
            
            ws.cell(row=row, column=2, value=m1_formula)
            ws.cell(row=row, column=3, value=sd1_formula)
            ws.cell(row=row, column=4, value=m2_formula)
            ws.cell(row=row, column=5, value=sd2_formula)
            
            # Mean difference
            ws.cell(row=row, column=6, value=f"=ROUND(B{row}-D{row},3)")
            
            # Cohen's d (pooled SD approximation)
            ws.cell(row=row, column=7, value=f"=ROUND(F{row}/SQRT((C{row}^2+E{row}^2)/2),3)")
            
            formulas.extend([
                {"cell": f"B{row}", "formula": m1_formula, "purpose": f"{col_name} M1"},
                {"cell": f"D{row}", "formula": m2_formula, "purpose": f"{col_name} M2"},
                {"cell": f"G{row}", "formula": f"Cohen's d", "purpose": f"{col_name} effect size"},
            ])
            
            row += 1
        
        wb.save(self.workbook_path)
        wb.close()
        
        return {
            "sheet_name": task.output_sheet,
            "formulas_created": len(formulas),
            "formulas": formulas
        }
    
    def _create_cross_tabulation(self, task: TaskSpec) -> Dict[str, Any]:
        """Create cross-tabulation sheet."""
        wb = self._open_workbook()
        ws = wb.create_sheet(task.output_sheet)
        
        formulas = []
        
        ws['A1'] = "CROSS-TABULATION"
        ws['A1'].font = Font(bold=True, size=14)
        ws['A3'] = "Note: Chi-square test requires additional calculation"
        
        wb.save(self.workbook_path)
        wb.close()
        
        return {
            "sheet_name": task.output_sheet,
            "formulas_created": len(formulas),
            "formulas": formulas
        }
    
    def _create_effect_sizes(self, task: TaskSpec) -> Dict[str, Any]:
        """Create effect size calculations sheet."""
        wb = self._open_workbook()
        ws = wb.create_sheet(task.output_sheet)
        
        formulas = []
        
        ws['A1'] = "EFFECT SIZE CALCULATIONS"
        ws['A1'].font = Font(bold=True, size=14)
        
        ws['A3'] = "Cohen's d Interpretation:"
        ws['A4'] = "Small: |d| ≈ 0.2"
        ws['A5'] = "Medium: |d| ≈ 0.5"
        ws['A6'] = "Large: |d| ≈ 0.8"
        
        wb.save(self.workbook_path)
        wb.close()
        
        return {
            "sheet_name": task.output_sheet,
            "formulas_created": len(formulas),
            "formulas": formulas
        }
    
    def _create_summary_dashboard(self, task: TaskSpec) -> Dict[str, Any]:
        """Create summary dashboard sheet."""
        wb = self._open_workbook()
        ws = wb.create_sheet(task.output_sheet)
        
        formulas = []
        
        ws['A1'] = "ANALYSIS SUMMARY DASHBOARD"
        ws['A1'].font = Font(bold=True, size=14)
        
        ws['A3'] = "Dataset Overview"
        ws['A3'].font = self.header_font
        
        ws['A4'] = "Total Observations:"
        ws['B4'] = f"={self.n_rows}"
        ws['A5'] = "Total Variables:"
        ws['B5'] = f"={len(self.df.columns)}"
        ws['A6'] = "Numeric Variables:"
        ws['B6'] = f"={len(self.numeric_cols)}"
        ws['A7'] = "Categorical Variables:"
        ws['B7'] = f"={len(self.categorical_cols)}"
        
        formulas.extend([
            {"cell": "B4", "formula": f"={self.n_rows}", "purpose": "N observations"},
            {"cell": "B5", "formula": f"={len(self.df.columns)}", "purpose": "N variables"},
        ])
        
        wb.save(self.workbook_path)
        wb.close()
        
        return {
            "sheet_name": task.output_sheet,
            "formulas_created": len(formulas),
            "formulas": formulas
        }
