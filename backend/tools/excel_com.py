"""
Excel COM automation helpers.
Used to recalculate workbooks and evaluate UDFs.
Requires Excel + pywin32.
"""

from pathlib import Path
from typing import Optional

try:
    import win32com.client as win32  # type: ignore
except Exception:
    win32 = None


def recalculate_workbook(workbook_path: Path) -> None:
    """
    Force Excel to recalculate all formulas (including UDFs) and save.
    """
    if win32 is None:
        raise RuntimeError("pywin32/Excel COM is required for recalculation")

    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    wb = None
    try:
        wb = excel.Workbooks.Open(str(workbook_path))
        excel.CalculateFullRebuild()
        wb.Save()
    finally:
        if wb is not None:
            wb.Close(SaveChanges=True)
        excel.Quit()

