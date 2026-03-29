from __future__ import annotations

from typing import List

import pandas as pd


def _issue_row(
    issue_type: str,
    issue_level: str,
    month_key: str,
    employee_id: str = "",
    employee_name: str = "",
    cost_center: str = "",
    amount: float | int | None = None,
    source_sheet: str = "",
    note: str = "",
) -> dict:
    return {
        "issue_type": issue_type,
        "issue_level": issue_level,
        "month_key": month_key,
        "employee_id": employee_id,
        "employee_name": employee_name,
        "cost_center": cost_center,
        "amount": amount if amount is not None else 0.0,
        "source_sheet": source_sheet,
        "note": note,
    }


def run_dq_checks(
    payroll_fact: pd.DataFrame,
    employee_master: pd.DataFrame,
    allocation_fact: pd.DataFrame,
) -> pd.DataFrame:
    issues: List[dict] = []

    # Payroll rules
    missing_emp_id = payroll_fact[payroll_fact["employee_id"].astype(str).str.strip() == ""]
    for _, r in missing_emp_id.iterrows():
        issues.append(
            _issue_row(
                issue_type="employee_id_missing",
                issue_level="error",
                month_key=str(r["month_key"]),
                employee_name=str(r.get("employee_name", "")),
                amount=float(r.get("amount", 0.0)),
                source_sheet=str(r.get("source_file", "")),
                note="Payroll record has missing employee_id.",
            )
        )

    zero_amount = payroll_fact[pd.to_numeric(payroll_fact["amount"], errors="coerce").fillna(0) == 0]
    for _, r in zero_amount.iterrows():
        issues.append(
            _issue_row(
                issue_type="amount_zero",
                issue_level="warning",
                month_key=str(r["month_key"]),
                employee_id=str(r.get("employee_id", "")),
                employee_name=str(r.get("employee_name", "")),
                amount=float(r.get("amount", 0.0)),
                source_sheet=str(r.get("source_file", "")),
                note="Payroll amount is zero.",
            )
        )

    dup_name = payroll_fact[payroll_fact.duplicated(subset=["month_key", "employee_name"], keep=False)]
    for _, r in dup_name.iterrows():
        issues.append(
            _issue_row(
                issue_type="employee_name_duplicate",
                issue_level="review",
                month_key=str(r["month_key"]),
                employee_id=str(r.get("employee_id", "")),
                employee_name=str(r.get("employee_name", "")),
                amount=float(r.get("amount", 0.0)),
                source_sheet=str(r.get("source_file", "")),
                note="Duplicate employee_name found in payroll month.",
            )
        )

    dup_keys = payroll_fact[
        payroll_fact.duplicated(subset=["month_key", "employee_id", "pay_item"], keep=False)
    ]
    for _, r in dup_keys.iterrows():
        issues.append(
            _issue_row(
                issue_type="duplicate_record",
                issue_level="error",
                month_key=str(r["month_key"]),
                employee_id=str(r.get("employee_id", "")),
                employee_name=str(r.get("employee_name", "")),
                amount=float(r.get("amount", 0.0)),
                source_sheet=str(r.get("source_file", "")),
                note="Duplicate by month_key + employee_id + pay_item.",
            )
        )

    # Mapping rules
    mapped = payroll_fact.merge(
        employee_master[["employee_id", "cost_center"]],
        on="employee_id",
        how="left",
    )
    map_not_found = mapped[mapped["cost_center"].isna()]
    for _, r in map_not_found.iterrows():
        issues.append(
            _issue_row(
                issue_type="mapping_not_found",
                issue_level="error",
                month_key=str(r["month_key"]),
                employee_id=str(r.get("employee_id", "")),
                employee_name=str(r.get("employee_name", "")),
                amount=float(r.get("amount", 0.0)),
                source_sheet=str(r.get("source_file", "")),
                note="employee_id from payroll not found in 3-69 mapping.",
            )
        )

    map_cc_empty = mapped[mapped["cost_center"].astype(str).str.strip().eq("")]
    for _, r in map_cc_empty.iterrows():
        issues.append(
            _issue_row(
                issue_type="mapping_cost_center_empty",
                issue_level="error",
                month_key=str(r["month_key"]),
                employee_id=str(r.get("employee_id", "")),
                employee_name=str(r.get("employee_name", "")),
                amount=float(r.get("amount", 0.0)),
                note="Mapped employee has empty cost_center.",
            )
        )

    # Allocation rules
    alloc_cc_empty = allocation_fact[allocation_fact["cost_center"].astype(str).str.strip().eq("")]
    for _, r in alloc_cc_empty.iterrows():
        issues.append(
            _issue_row(
                issue_type="allocation_cost_center_empty",
                issue_level="error",
                month_key=str(r["month_key"]),
                cost_center=str(r.get("cost_center", "")),
                amount=float(r.get("allocated_amount", 0.0)),
                source_sheet=str(r.get("source_sheet", "")),
                note="Allocation record has no cost_center.",
            )
        )

    # Optional allocation completeness rule when source_total exists.
    if "source_total" in allocation_fact.columns:
        cmp_df = (
            allocation_fact.groupby(["month_key", "source_sheet"], as_index=False)
            .agg(allocated_total=("allocated_amount", "sum"), source_total=("source_total", "max"))
            .fillna(0)
        )
        mismatch = cmp_df[(cmp_df["allocated_total"] - cmp_df["source_total"]).abs() > 0.01]
        for _, r in mismatch.iterrows():
            issues.append(
                _issue_row(
                    issue_type="allocation_totals_mismatch",
                    issue_level="error",
                    month_key=str(r["month_key"]),
                    amount=float(r["allocated_total"]),
                    source_sheet=str(r["source_sheet"]),
                    note=f"Allocated {r['allocated_total']:.2f} != source {r['source_total']:.2f}",
                )
            )
        incomplete = cmp_df[cmp_df["allocated_total"] + 0.01 < cmp_df["source_total"]]
        for _, r in incomplete.iterrows():
            issues.append(
                _issue_row(
                    issue_type="incomplete_allocation",
                    issue_level="error",
                    month_key=str(r["month_key"]),
                    amount=float(r["allocated_total"]),
                    source_sheet=str(r["source_sheet"]),
                    note="Allocated amount is lower than source total.",
                )
            )

    dq = pd.DataFrame(issues)
    if dq.empty:
        return pd.DataFrame(
            columns=[
                "issue_type",
                "issue_level",
                "month_key",
                "employee_id",
                "employee_name",
                "cost_center",
                "amount",
                "source_sheet",
                "note",
            ]
        )
    return dq

