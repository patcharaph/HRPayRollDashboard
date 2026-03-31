from __future__ import annotations

from pathlib import Path
from typing import List

import numpy as np
import pandas as pd


def _normalize_cols(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    out.columns = [str(c).strip() for c in out.columns]
    return out


def _find_first_column(columns: List[str], keywords: List[str]) -> str | None:
    lowered = [c.lower() for c in columns]
    for kw in keywords:
        for i, c in enumerate(lowered):
            if kw in c:
                return columns[i]
    return None


def _as_string_id(series: pd.Series) -> pd.Series:
    """
    Normalize id-like values to string and remove trailing '.0' from Excel numeric parsing.
    """
    out = series.astype(str).str.strip()
    out = out.str.replace("\u00A0", "", regex=False)  # non-breaking space
    out = out.str.replace("\u200b", "", regex=False)  # zero-width space
    out = out.str.replace(r"^'+", "", regex=True)     # leading apostrophe from Excel text cells
    out = out.str.replace(r"\s+", "", regex=True)     # remove embedded spaces in id
    out = out.str.replace(r"\.0+$", "", regex=True)
    out = out.replace(
        {
            "nan": "",
            "NaN": "",
            "None": "",
            "NONE": "",
            "null": "",
            "NULL": "",
            "<NA>": "",
        }
    )
    return out


def transform_payroll_to_fact(
    payroll_raw: pd.DataFrame,
    month_key: str = "2026-03",
    source_file: str = "3. ข้อมูลการจ่ายแยกรายคน lhsec 0326.xls",
) -> pd.DataFrame:
    """
    Convert payroll raw data (wide) to long fact:
    month_key, employee_id, employee_name, pay_item, amount, source_file
    """
    raw = _normalize_cols(payroll_raw)
    cols = list(raw.columns)
    if "Code" in cols:
        emp_id_col = "Code"
    else:
        emp_id_col = _find_first_column(cols, ["employee_id", "emp_id", "code", "รหัสพนักงาน", "รหัส"])
    emp_name_col = _find_first_column(cols, ["employee_name", "name", "ชื่อพนักงาน", "ชื่อ"])

    if emp_id_col is None:
        raise ValueError("Cannot find employee_id column in payroll file.")
    if emp_name_col is None:
        raw["employee_name"] = ""
        emp_name_col = "employee_name"

    id_columns = [emp_id_col, emp_name_col]
    numeric_candidates = [c for c in cols if c not in id_columns]

    # Keep columns that have at least one numeric value.
    value_columns = []
    for c in numeric_candidates:
        as_num = pd.to_numeric(raw[c], errors="coerce")
        if as_num.notna().any():
            value_columns.append(c)

    if not value_columns:
        raise ValueError("Cannot identify payroll amount columns in source file.")

    long_df = raw.melt(
        id_vars=id_columns,
        value_vars=value_columns,
        var_name="pay_item",
        value_name="amount",
    )
    long_df = long_df.rename(
        columns={
            emp_id_col: "employee_id",
            emp_name_col: "employee_name",
        }
    )
    long_df["amount"] = pd.to_numeric(long_df["amount"], errors="coerce").fillna(0.0)
    long_df = long_df[long_df["amount"] != 0].copy()
    long_df["employee_id"] = _as_string_id(long_df["employee_id"]).astype("string")
    long_df["employee_name"] = long_df["employee_name"].astype(str).str.strip()
    long_df["month_key"] = month_key
    long_df["source_file"] = Path(source_file).name
    long_df = long_df[
        ["month_key", "employee_id", "employee_name", "pay_item", "amount", "source_file"]
    ].reset_index(drop=True)
    return long_df


def summarize_payroll_by_employee(payroll_fact: pd.DataFrame) -> pd.DataFrame:
    summary = (
        payroll_fact.groupby(["month_key", "employee_id", "employee_name"], as_index=False)["amount"]
        .sum()
        .rename(columns={"amount": "direct_payroll_cost"})
    )
    return summary
