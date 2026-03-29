from __future__ import annotations

from pathlib import Path

import pandas as pd


def export_outputs(
    summary_payroll_by_employee: pd.DataFrame,
    summary_allocation_by_cost_center: pd.DataFrame,
    output_dir: str | Path = "output",
) -> dict:
    out_dir = Path(output_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    payroll_path = out_dir / "summary_payroll_by_employee.csv"
    allocation_path = out_dir / "summary_allocation_by_cost_center.csv"

    summary_payroll_by_employee.to_csv(payroll_path, index=False, encoding="utf-8-sig")
    summary_allocation_by_cost_center.to_csv(allocation_path, index=False, encoding="utf-8-sig")

    return {
        "summary_payroll_by_employee": str(payroll_path),
        "summary_allocation_by_cost_center": str(allocation_path),
    }
