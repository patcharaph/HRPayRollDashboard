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
    Pick employee mapping sheet.
    Priority:
    1) exact/contains `mapping_sheet_hint` (default: 3-69)
    2) best-score sheet by expected mapping columns (Code/Name/Department/Cost Center/Type/Front-Back)
    """
    for sheet_name, df in sheets.items():
        if sheet_name.strip() == mapping_sheet_hint:
            return df.copy()
    for sheet_name, df in sheets.items():
        if mapping_sheet_hint in sheet_name:
            return df.copy()

    def norm_col(col: object) -> str:
        return str(col).strip().lower()

    candidate_keywords = [
        "code",
        "employee",
        "name",
        "department",
        "dept",
        "cost center",
        "cost_center",
        "costcenter",
        "type",
        "front/back",
        "front",
        "back",
    ]

    best_name = None
    best_df = None
    best_score = -1

    for sheet_name, df in sheets.items():
        cols = [norm_col(c) for c in df.columns]
        score = 0
        for kw in candidate_keywords:
            if any(kw in c for c in cols):
                score += 1
        # Prefer sheets with at least some data rows.
        if not df.empty:
            score += 1
        if score > best_score:
            best_score = score
            best_name = sheet_name
            best_df = df

    # Require minimum confidence to avoid selecting random numeric sheets.
    if best_df is not None and best_score >= 4:
        return best_df.copy()

    raise ValueError(
        f"Mapping sheet '{mapping_sheet_hint}' not found in workbook, "
        "and no fallback sheet matched mapping structure."
    )
