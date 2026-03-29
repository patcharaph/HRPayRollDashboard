from __future__ import annotations

import numpy as np
import pandas as pd


def apply_employee_mapping(payroll_fact: pd.DataFrame, employee_master: pd.DataFrame) -> pd.DataFrame:
    payroll = payroll_fact.copy()
    master = employee_master.copy()
    payroll["employee_id"] = payroll["employee_id"].astype(str).str.strip().str.replace(r"\.0+$", "", regex=True)
    master["employee_id"] = master["employee_id"].astype(str).str.strip().str.replace(r"\.0+$", "", regex=True)
    master["cost_center"] = master["cost_center"].astype(str).str.strip().str.replace(r"\.0+$", "", regex=True)
    return payroll.merge(
        master[
            ["employee_id", "employee_name", "department", "cost_center", "employee_type", "front_back"]
        ],
        on="employee_id",
        how="left",
        suffixes=("", "_master"),
    )


def build_employee_cost_summary(
    payroll_with_mapping: pd.DataFrame,
    allocation_fact: pd.DataFrame,
) -> pd.DataFrame:
    direct = (
        payroll_with_mapping.groupby(
            ["month_key", "employee_id", "employee_name", "department", "cost_center"], as_index=False
        )["amount"]
        .sum()
        .rename(columns={"amount": "direct_payroll_cost"})
    )

    alloc_by_cc = (
        allocation_fact.groupby(["month_key", "cost_center"], as_index=False)["allocated_amount"]
        .sum()
        .rename(columns={"allocated_amount": "allocated_overhead_cost_center"})
    )
    base = direct.merge(alloc_by_cc, on=["month_key", "cost_center"], how="left")
    base["allocated_overhead_cost_center"] = base["allocated_overhead_cost_center"].fillna(0.0)

    # Distribute overhead to employees in the same cost center by payroll share.
    group_total = base.groupby(["month_key", "cost_center"])["direct_payroll_cost"].transform("sum")
    employee_count = base.groupby(["month_key", "cost_center"])["employee_id"].transform("count")
    payroll_share = np.where(group_total > 0, base["direct_payroll_cost"] / group_total, 1 / employee_count.clip(1))
    base["allocated_overhead_employee"] = base["allocated_overhead_cost_center"] * payroll_share
    base["fully_allocated_cost"] = base["direct_payroll_cost"] + base["allocated_overhead_employee"]
    base["employee_id"] = base["employee_id"].astype(str).str.strip().str.replace(r"\.0+$", "", regex=True)
    base["cost_center"] = base["cost_center"].astype(str).str.strip().str.replace(r"\.0+$", "", regex=True)
    return base


def compute_executive_kpis(
    payroll_fact: pd.DataFrame,
    allocation_fact: pd.DataFrame,
    employee_master: pd.DataFrame,
) -> dict:
    return {
        "total_cost": float(payroll_fact["amount"].sum()) if not payroll_fact.empty else 0.0,
        "total_allocated_cost": float(allocation_fact["allocated_amount"].sum()) if not allocation_fact.empty else 0.0,
        "employee_count": int(employee_master["employee_id"].nunique()) if not employee_master.empty else 0,
        "cost_center_count": int(employee_master["cost_center"].replace("", np.nan).dropna().nunique())
        if not employee_master.empty
        else 0,
    }


def monthly_cost_trend(payroll_fact: pd.DataFrame, allocation_fact: pd.DataFrame) -> pd.DataFrame:
    p = payroll_fact.groupby("month_key", as_index=False)["amount"].sum().rename(columns={"amount": "payroll_cost"})
    a = (
        allocation_fact.groupby("month_key", as_index=False)["allocated_amount"]
        .sum()
        .rename(columns={"allocated_amount": "allocated_cost"})
    )
    trend = p.merge(a, on="month_key", how="outer").fillna(0)
    trend["total_combined"] = trend["payroll_cost"] + trend["allocated_cost"]
    return trend.sort_values("month_key")


def top_cost_centers(payroll_with_mapping: pd.DataFrame, n: int = 10) -> pd.DataFrame:
    out = (
        payroll_with_mapping.groupby("cost_center", as_index=False)["amount"]
        .sum()
        .rename(columns={"amount": "total_cost"})
        .sort_values("total_cost", ascending=False)
        .head(n)
    )
    return out


def top_vendors(allocation_fact: pd.DataFrame, n: int = 10) -> pd.DataFrame:
    out = (
        allocation_fact.groupby("vendor_name", as_index=False)["allocated_amount"]
        .sum()
        .sort_values("allocated_amount", ascending=False)
        .head(n)
    )
    return out
