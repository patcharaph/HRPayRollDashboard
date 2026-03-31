from __future__ import annotations

import re
from typing import Dict, List

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


def _clean_text(value: object) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def _as_string_id(series: pd.Series) -> pd.Series:
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


def _prepare_mapping_table(mapping_raw: pd.DataFrame) -> pd.DataFrame:
    """
    Try to normalize mapping sheet and auto-detect header row when header is not on first row.
    """
    raw = mapping_raw.copy()
    raw = raw.dropna(axis=1, how="all").dropna(axis=0, how="all")
    raw = _normalize_cols(raw)

    id_keywords = ["employee_id", "emp_id", "employee code", "personnel", "รหัสพนักงาน", "รหัส"]
    existing_cols = [str(c).strip() for c in raw.columns]
    if _find_first_column(existing_cols, id_keywords):
        return raw

    # Detect a header row from first 20 rows by keyword hits.
    scan_rows = min(20, len(raw))
    best_row = None
    best_score = -1
    for idx in range(scan_rows):
        row_values = [_clean_text(v) for v in raw.iloc[idx].tolist()]
        row_lower = [v.lower() for v in row_values]

        score = 0
        if any(any(k in v for k in id_keywords) for v in row_lower):
            score += 3
        if any(("name" in v) or ("ชื่อ" in v) for v in row_lower):
            score += 1
        if any(("department" in v) or ("แผนก" in v) or ("dept" in v) for v in row_lower):
            score += 1
        if any(("cost center" in v) or ("cost_center" in v) or ("ศูนย์ต้นทุน" in v) or ("cc" == v) for v in row_lower):
            score += 1

        if score > best_score:
            best_score = score
            best_row = idx

    if best_row is not None and best_score >= 3:
        header = [_clean_text(v) if _clean_text(v) else f"unnamed_{i}" for i, v in enumerate(raw.iloc[best_row].tolist())]
        data = raw.iloc[best_row + 1 :].copy()
        data.columns = header
        data = data.dropna(axis=1, how="all").dropna(axis=0, how="all")
        return _normalize_cols(data)

    return raw


def transform_employee_master(mapping_raw: pd.DataFrame, month_key: str = "2026-03") -> pd.DataFrame:
    df = _prepare_mapping_table(mapping_raw)
    cols = list(df.columns)

    # Explicit business mapping: "Code" = employee_id (รหัสพนักงาน)
    if "Code" in cols:
        employee_id_col = "Code"
    else:
        employee_id_col = None

    if employee_id_col is None:
        employee_id_col = _find_first_column(
            cols,
            [
                "employee_id",
                "emp_id",
                "employee code",
                "personnel",
                "staff id",
                "code",
                "no.",
                "no",
                "รหัสพนักงาน",
                "รหัส",
            ],
        )
    employee_name_col = _find_first_column(cols, ["employee_name", "emp_name", "name", "ชื่อพนักงาน", "ชื่อ"])
    department_col = _find_first_column(cols, ["department", "dept", "ฝ่าย", "แผนก"])
    cost_center_col = _find_first_column(cols, ["cost_center", "cost center", "costcenter", "cc", "ศูนย์ต้นทุน"])
    employee_type_col = _find_first_column(cols, ["employee_type", "emp_type", "type", "ประเภทพนักงาน"])
    front_back_col = _find_first_column(cols, ["front_back", "front/back", "front", "back"])

    if employee_id_col is None:
        # Heuristic fallback: pick column that looks like id code (mostly alnum, many unique values).
        best_col = None
        best_score = -1.0
        for c in cols:
            series = df[c].astype(str).str.strip()
            non_empty = series[series != ""]
            if non_empty.empty:
                continue
            unique_ratio = non_empty.nunique() / len(non_empty)
            looks_like_code_ratio = non_empty.str.match(r"^[A-Za-z0-9\-_]+$", na=False).mean()
            score = (0.6 * unique_ratio) + (0.4 * looks_like_code_ratio)
            if score > best_score:
                best_score = score
                best_col = c
        if best_col is not None and best_score >= 0.55:
            employee_id_col = best_col

    if employee_id_col is None:
        preview_cols = ", ".join([str(c) for c in cols[:20]])
        raise ValueError(
            "Cannot find employee_id in mapping sheet 3-69. "
            f"Detected columns: {preview_cols}"
        )

    master = pd.DataFrame(
        {
            "month_key": month_key,
            "employee_id": df[employee_id_col],
            "employee_name": df[employee_name_col] if employee_name_col else "",
            "department": df[department_col] if department_col else "",
            "cost_center": df[cost_center_col] if cost_center_col else "",
            "employee_type": df[employee_type_col] if employee_type_col else "",
            "front_back": df[front_back_col] if front_back_col else "",
        }
    )
    master["employee_id"] = _as_string_id(master["employee_id"]).astype("string")
    master["employee_name"] = master["employee_name"].astype(str).str.strip()
    master["department"] = master["department"].astype(str).str.strip()
    master["cost_center"] = _as_string_id(master["cost_center"]).astype("string")
    master["employee_type"] = master["employee_type"].astype(str).str.strip()
    master["front_back"] = master["front_back"].astype(str).str.strip()

    master = master[master["employee_id"] != ""].copy()

    # If same employee_id appears multiple rows, keep the row with the most complete mapping fields.
    completeness_cols = ["employee_name", "department", "cost_center", "employee_type", "front_back"]
    for c in completeness_cols:
        master[c] = master[c].fillna("").astype(str).str.strip()
    master["_completeness_score"] = master[completeness_cols].apply(
        lambda r: sum(1 for v in r.tolist() if v not in ["", "nan", "None", "<NA>"]),
        axis=1,
    )
    master = (
        master.sort_values(["employee_id", "_completeness_score"], ascending=[True, False])
        .drop_duplicates(subset=["employee_id"], keep="first")
        .drop(columns=["_completeness_score"])
        .reset_index(drop=True)
    )
    return master


def transform_allocation_fact(
    workbook_sheets: Dict[str, pd.DataFrame],
    month_key: str = "2026-03",
    mapping_sheet_hint: str = "3-69",
) -> pd.DataFrame:
    """
    Build allocation fact from all non-mapping sheets.
    """
    frames = []
    for sheet_name, raw in workbook_sheets.items():
        if mapping_sheet_hint in sheet_name or sheet_name.strip() == mapping_sheet_hint:
            continue
        df = _normalize_cols(raw)
        cols = list(df.columns)
        vendor_col = _find_first_column(cols, ["vendor", "supplier", "ผู้ขาย"])
        expense_col = _find_first_column(cols, ["expense_type", "expense", "รายการ", "ประเภทค่าใช้จ่าย"])
        cc_col = _find_first_column(cols, ["cost_center", "cost center", "cc", "ศูนย์ต้นทุน"])
        amount_col = _find_first_column(cols, ["allocated_amount", "allocated", "amount", "ยอด", "ค่าใช้จ่าย"])

        if cc_col is None or amount_col is None:
            continue

        frame = pd.DataFrame(
            {
                "month_key": month_key,
                "vendor_name": df[vendor_col] if vendor_col else "Unknown Vendor",
                "expense_type": df[expense_col] if expense_col else "Unknown Expense",
                "cost_center": df[cc_col],
                "allocated_amount": pd.to_numeric(df[amount_col], errors="coerce"),
                "source_sheet": sheet_name,
            }
        )
        frame["vendor_name"] = frame["vendor_name"].astype(str).str.strip()
        frame["expense_type"] = frame["expense_type"].astype(str).str.strip()
        frame["cost_center"] = _as_string_id(frame["cost_center"]).astype("string")
        frame = frame[frame["allocated_amount"].notna()].copy()
        frames.append(frame)

    if not frames:
        return pd.DataFrame(
            columns=[
                "month_key",
                "vendor_name",
                "expense_type",
                "cost_center",
                "allocated_amount",
                "source_sheet",
            ]
        )

    allocation_fact = pd.concat(frames, ignore_index=True)
    return allocation_fact.reset_index(drop=True)


def summarize_allocation_by_cost_center(allocation_fact: pd.DataFrame) -> pd.DataFrame:
    summary = (
        allocation_fact.groupby(["month_key", "cost_center"], as_index=False)["allocated_amount"]
        .sum()
        .rename(columns={"allocated_amount": "total_allocated_amount"})
    )
    return summary


def extract_total_cost_from_mapping_sheet(mapping_raw: pd.DataFrame) -> float:
    """
    Estimate payroll total from mapping sheet (3-69).
    Priority:
    1) Sum salary-like column (e.g. เงินเดือน / salary) if found.
    2) Fallback to sum of numeric pay-item columns.
    """
    df = _prepare_mapping_table(mapping_raw)
    if df.empty:
        return 0.0

    def to_num(series: pd.Series) -> pd.Series:
        s = series.astype(str).str.strip()
        # Remove thousand separators and keep negative with parenthesis support.
        s = s.str.replace(",", "", regex=False)
        s = s.str.replace(r"^\((.*)\)$", r"-\1", regex=True)
        return pd.to_numeric(s, errors="coerce")

    # 1) Salary-like column first.
    salary_keywords = ["เงินเดือน", "salary", "wage", "basic"]
    for c in df.columns:
        c_lower = str(c).strip().lower()
        if any(k in c_lower for k in salary_keywords):
            s = to_num(df[c]).fillna(0)
            if s.abs().sum() > 0:
                return float(s.sum())

    excluded_keywords = [
        "employee",
        "emp_",
        "code",
        "no",
        "name",
        "department",
        "dept",
        "cost",
        "center",
        "type",
        "front",
        "back",
        "remark",
        "source",
    ]
    numeric_cols = []
    for c in df.columns:
        c_lower = str(c).strip().lower()
        if any(k in c_lower for k in excluded_keywords):
            continue
        s = to_num(df[c])
        if s.notna().any():
            numeric_cols.append(c)

    if not numeric_cols:
        return 0.0

    total = 0.0
    for c in numeric_cols:
        total += float(to_num(df[c]).fillna(0).sum())
    return total
