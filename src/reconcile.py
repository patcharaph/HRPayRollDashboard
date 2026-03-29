from __future__ import annotations

import pandas as pd


def build_reconciliation_checks(
    payroll_fact: pd.DataFrame,
    allocation_fact: pd.DataFrame,
    employee_master: pd.DataFrame,
) -> pd.DataFrame:
    """
    Build reconciliation table for dashboard view.
    """
    rows = []
    month = payroll_fact["month_key"].iloc[0] if not payroll_fact.empty else "2026-03"

    payroll_total = float(payroll_fact["amount"].sum()) if not payroll_fact.empty else 0.0
    payroll_summary_total = (
        float(payroll_fact.groupby(["month_key", "employee_id"])["amount"].sum().sum())
        if not payroll_fact.empty
        else 0.0
    )
    rows.append(
        {
            "check_name": "payroll_total_vs_payroll_summary",
            "month_key": month,
            "left_value": payroll_total,
            "right_value": payroll_summary_total,
            "difference": payroll_total - payroll_summary_total,
            "status": "ok" if abs(payroll_total - payroll_summary_total) <= 0.01 else "mismatch",
        }
    )

    alloc_total = float(allocation_fact["allocated_amount"].sum()) if not allocation_fact.empty else 0.0
    by_cc_total = (
        float(allocation_fact.groupby(["month_key", "cost_center"])["allocated_amount"].sum().sum())
        if not allocation_fact.empty
        else 0.0
    )
    rows.append(
        {
            "check_name": "allocation_total_vs_cost_center_total",
            "month_key": month,
            "left_value": alloc_total,
            "right_value": by_cc_total,
            "difference": alloc_total - by_cc_total,
            "status": "ok" if abs(alloc_total - by_cc_total) <= 0.01 else "mismatch",
        }
    )

    mapped = payroll_fact.merge(employee_master[["employee_id", "cost_center"]], on="employee_id", how="left")
    mapped_total = float(mapped[mapped["cost_center"].notna()]["amount"].sum()) if not mapped.empty else 0.0
    rows.append(
        {
            "check_name": "payroll_mapped_to_cost_center",
            "month_key": month,
            "left_value": payroll_total,
            "right_value": mapped_total,
            "difference": payroll_total - mapped_total,
            "status": "ok" if abs(payroll_total - mapped_total) <= 0.01 else "mismatch",
        }
    )

    return pd.DataFrame(rows)

