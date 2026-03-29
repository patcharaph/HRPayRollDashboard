from __future__ import annotations

from io import BytesIO
from pathlib import Path

import pandas as pd


def to_excel_bytes(sheets: dict[str, pd.DataFrame]) -> bytes:
    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        for sheet_name, df in sheets.items():
            (df if df is not None else pd.DataFrame()).to_excel(writer, sheet_name=sheet_name, index=False)
    bio.seek(0)
    return bio.getvalue()


def to_mkt_excel_bytes(mkt_df: pd.DataFrame) -> bytes:
    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        mkt_df.to_excel(writer, sheet_name="Value", index=False)
        pd.DataFrame().to_excel(writer, sheet_name="Sheet1", index=False)
    bio.seek(0)
    return bio.getvalue()


def to_accounting_excel_bytes(
    accounting_df: pd.DataFrame,
    company_name: str = "บมจ.หลักทรัพย์ แลนด์ แอนด์ เฮ้าส์",
    period_text: str = "01/03/2026-31/03/2026",
) -> bytes:
    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        sheet_name = "Allocate salary for Accounting"
        # Keep same report-like header layout as sample file:
        # Row 1: Allocation Salary 2025
        # Row 2: Company + period label/value
        # Row 4: table header (startrow=3)
        accounting_df.to_excel(writer, sheet_name=sheet_name, index=False, startrow=3)
        pd.DataFrame().to_excel(writer, sheet_name="Sheet1", index=False)

        ws = writer.book[sheet_name]
        ws["A1"] = "Allocation Salary 2025"
        ws["A2"] = company_name
        ws["R2"] = "ประจำงวด"
        ws["S2"] = period_text
    bio.seek(0)
    return bio.getvalue()


def export_example_outputs(
    mkt_df: pd.DataFrame,
    accounting_df: pd.DataFrame,
    month_suffix: str = "03 26",
    period_text: str = "01/03/2026-31/03/2026",
    output_dir: str | Path = "output",
) -> dict:
    out_dir = Path(output_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    mkt_path = out_dir / f"Allocate MKT {month_suffix}.xlsx"
    acc_path = out_dir / f"Allocate salary {month_suffix} for Accounting.xlsx"

    mkt_path.write_bytes(to_mkt_excel_bytes(mkt_df))
    acc_path.write_bytes(to_accounting_excel_bytes(accounting_df, period_text=period_text))

    return {
        "allocate_mkt": str(mkt_path),
        "allocate_salary_accounting": str(acc_path),
    }
