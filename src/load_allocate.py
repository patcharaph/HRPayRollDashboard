from __future__ import annotations

from io import BytesIO
from pathlib import Path
from typing import Dict

import pandas as pd


def load_allocation_workbook(file_path: str | Path | bytes) -> Dict[str, pd.DataFrame]:
    """
    Load all sheets from allocation workbook as raw DataFrames.
    """
    if isinstance(file_path, (str, Path)):
        workbook_path = Path(file_path)
        sheets = pd.read_excel(workbook_path, sheet_name=None, engine="openpyxl")
    else:
        sheets = pd.read_excel(BytesIO(file_path), sheet_name=None, engine="openpyxl")
    return {str(k): v for k, v in sheets.items()}


def find_mapping_sheet(sheets: Dict[str, pd.DataFrame], mapping_sheet_hint: str = "3-69") -> pd.DataFrame:
    """
    Pick employee mapping sheet. Prefer exact/contains `3-69`.
    """
    for sheet_name, df in sheets.items():
        if sheet_name.strip() == mapping_sheet_hint:
            return df.copy()
    for sheet_name, df in sheets.items():
        if mapping_sheet_hint in sheet_name:
            return df.copy()
    raise ValueError(f"Mapping sheet '{mapping_sheet_hint}' not found in workbook.")
