from __future__ import annotations

from pathlib import Path
import re

import pandas as pd
import plotly.express as px
import streamlit as st

from src.dq_checks import run_dq_checks
from src.export_excel import (
    to_accounting_excel_bytes,
    to_mkt_excel_bytes,
)
from src.load_allocate import find_mapping_sheet, load_allocation_workbook
from src.load_payroll import (
    load_payroll_xls,
    load_payroll_xls_from_bytes,
    payroll_bytes_is_encrypted,
    payroll_file_is_encrypted,
)
from src.metrics import (
    apply_employee_mapping,
    build_employee_cost_summary,
)
from src.reconcile import build_reconciliation_checks
from src.transform_allocate import (
    extract_total_cost_from_mapping_sheet,
    summarize_allocation_by_cost_center,
    transform_allocation_fact,
    transform_employee_master,
)
from src.transform_payroll import transform_payroll_to_fact


st.set_page_config(page_title="HR Payroll Dashboard", page_icon=":bar_chart:", layout="wide")
st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Manrope:wght@400;500;600;700;800&family=Plus+Jakarta+Sans:wght@500;700;800&display=swap');

    :root {
        --brand-900: #0b2a62;
        --brand-800: #113b86;
        --brand-700: #1d5ed8;
        --brand-600: #2f6feb;
        --ink-900: #0f172a;
        --ink-600: #475569;
        --surface-0: #ffffff;
        --surface-50: #f8fbff;
        --surface-100: #eef4ff;
        --line-200: #dbe7ff;
    }

    .stApp {
        font-family: 'Manrope', sans-serif;
        font-size: 14px;
        background:
          radial-gradient(1200px 520px at 102% -8%, rgba(47, 111, 235, 0.16), transparent 58%),
          radial-gradient(900px 420px at -2% 2%, rgba(29, 94, 216, 0.10), transparent 55%),
          linear-gradient(180deg, #ffffff 0%, #f7faff 100%);
    }
    section[data-testid="stSidebar"] {
        background: linear-gradient(180deg, #f9fbff 0%, #edf4ff 100%);
        border-right: 1px solid var(--line-200);
    }
    section[data-testid="stSidebar"] * {
        color: var(--brand-800) !important;
    }
    section[data-testid="stSidebar"] h1,
    section[data-testid="stSidebar"] h2,
    section[data-testid="stSidebar"] h3,
    section[data-testid="stSidebar"] label,
    section[data-testid="stSidebar"] p,
    section[data-testid="stSidebar"] span,
    section[data-testid="stSidebar"] div {
        color: var(--brand-800) !important;
    }
    section[data-testid="stSidebar"] [data-testid="stFileUploader"] label p,
    section[data-testid="stSidebar"] [data-testid="stFileUploader"] label span {
        font-weight: 800 !important;
        font-size: 0.95rem !important;
        line-height: 1.35 !important;
        color: var(--brand-900) !important;
    }
    h1, h2, h3 {
        color: var(--brand-900);
    }
    h1 {
        font-family: 'Plus Jakarta Sans', sans-serif;
        font-weight: 800;
        font-size: 2.8rem;
        letter-spacing: -0.4px;
    }
    .block-container {
        padding-top: 1.25rem;
        max-width: 1320px;
    }
    div[data-baseweb="tab-list"] {
        gap: 8px;
        padding: 6px;
        background: rgba(17, 59, 134, 0.06);
        border-radius: 14px;
        border: 1px solid rgba(17, 59, 134, 0.10);
        backdrop-filter: blur(4px);
    }
    button[role="tab"] {
        background: var(--surface-0);
        border-radius: 10px;
        color: var(--brand-900);
        border: 1px solid var(--line-200);
        padding: 8px 14px;
        font-weight: 600;
        transition: all .18s ease;
    }
    button[role="tab"][aria-selected="true"] {
        background: linear-gradient(180deg, var(--brand-700) 0%, var(--brand-800) 100%);
        color: #ffffff;
        border: 1px solid var(--brand-800);
        box-shadow: 0 8px 18px rgba(17, 59, 134, 0.30);
        transform: translateY(-1px);
    }
    div[data-testid="stMetric"] {
        background: linear-gradient(180deg, #ffffff 0%, #f2f7ff 100%);
        border: 1px solid #cfe0ff;
        border-radius: 16px;
        padding: 14px 16px;
        box-shadow: 0 12px 24px rgba(16, 56, 128, 0.10);
        backdrop-filter: blur(3px);
    }
    div[data-testid="stMetricLabel"] {
        font-weight: 700;
        color: var(--brand-800);
    }
    div[data-testid="stMetricValue"] {
        font-family: 'Plus Jakarta Sans', sans-serif;
        letter-spacing: -0.2px;
    }
    .stButton > button, .stDownloadButton > button {
        background: linear-gradient(180deg, var(--brand-600) 0%, var(--brand-800) 100%);
        color: #ffffff !important;
        border: 1px solid var(--brand-800);
        border-radius: 12px;
        font-weight: 600;
        box-shadow: 0 10px 20px rgba(16, 56, 128, 0.24);
        transition: all .18s ease;
    }
    .stButton > button p, .stButton > button span,
    .stDownloadButton > button p, .stDownloadButton > button span {
        color: #ffffff !important;
    }
    .stButton > button:hover, .stDownloadButton > button:hover {
        background: linear-gradient(180deg, #1b57c6 0%, #0c388d 100%);
        border-color: #0c388d;
        transform: translateY(-1px);
    }
    [data-testid="stDataFrame"], [data-testid="stTable"] {
        border: 1px solid #d8e5ff;
        border-radius: 14px;
        box-shadow: 0 10px 22px rgba(14, 60, 146, 0.08);
        background: #ffffff;
    }
    /* Prevent nested scrollbars in Streamlit dataframe viewport */
    [data-testid="stDataFrame"] > div {
        overflow: visible !important;
    }
    .stAlert {
        border-radius: 12px;
        border: 1px solid #d2e3ff;
        background: linear-gradient(180deg, #ffffff 0%, #f7fbff 100%);
    }
    .stCaption {
        color: var(--ink-600);
    }
    /* Uploaded file row style (the exact row under each uploader) */
    section[data-testid="stSidebar"] [data-testid="stFileUploader"] [data-testid="stFileUploaderFile"] {
        background: #dcfce7 !important;
        border: 1px solid #86efac !important;
        border-radius: 10px !important;
        padding: 6px 8px !important;
        color: #166534 !important;
    }
    section[data-testid="stSidebar"] [data-testid="stFileUploader"] [data-testid="stFileUploaderFileName"]::before {
        content: "✅ ";
        font-weight: 800;
        color: #166534 !important;
    }
    .payroll-password-label {
        color: #dc2626 !important;
        font-weight: 800 !important;
        margin: 6px 0 2px 0;
        font-size: 1rem;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

DATA_DIR = Path("data")
PAYROLL_FILE = DATA_DIR / "payroll.xls"
ALLOCATE_FILE = DATA_DIR / "allocation.xlsx"
PAYROLL_PASSWORD = ""
DEFAULT_MONTH_KEY = "2026-03"


def find_local_file(patterns: list[str]) -> Path | None:
    for p in patterns:
        files = sorted(DATA_DIR.glob(p))
        if files:
            return files[0]
    return None


def infer_month_key_from_filename(file_name: str | None) -> str | None:
    """
    Examples:
    - '...0326...' -> 2026-03
    - '...03 26...' -> 2026-03
    - '...03-26...' -> 2026-03
    """
    if not file_name:
        return None
    name = str(file_name)
    candidates = []
    candidates.extend(re.findall(r"(?<!\d)(0[1-9]|1[0-2])[\s\-_]?(\d{2})(?!\d)", name))
    candidates.extend(re.findall(r"(?<!\d)(0[1-9]|1[0-2])(\d{2})(?!\d)", name))
    for mm, yy in candidates:
        month = int(mm)
        year = 2000 + int(yy)
        return f"{year:04d}-{month:02d}"
    return None


def validate_uploaded_filename(upload_kind: str, file_name: str) -> str | None:
    name = str(file_name or "").strip()
    lower_name = name.lower()
    label = "Payroll" if upload_kind == "payroll" else "Previous Payroll"

    if upload_kind in {"payroll", "prev_payroll"}:
        # Payroll file name must include MMYY token such as 0326 or 03-26.
        if infer_month_key_from_filename(name) is None:
            return (
                f"Invalid {label} file name: '{name}'. "
                "Expected month token MMYY in the name (example: 0326)."
            )

    if upload_kind == "allocation":
        # Allocation naming convention: include allocat* and 4-digit year.
        if "allocat" not in lower_name or re.search(r"(?<!\d)20\d{2}(?!\d)", name) is None:
            return (
                f"Invalid allocation file name: '{name}'. "
                "Expected name to contain 'Allocation/Allocate' and year YYYY (example: Allocation Cost 2026.xlsx)."
            )

    return None


def previous_month_key(month_key: str) -> str:
    try:
        d = pd.to_datetime(f"{month_key}-01") - pd.offsets.MonthBegin(1)
        return d.strftime("%Y-%m")
    except Exception:
        return "2026-02"


@st.cache_data(show_spinner=True)
def run_pipeline(
    payroll_upload_bytes: bytes | None,
    allocation_upload_bytes: bytes | None,
    payroll_password: str,
    month_key: str,
    payroll_file_label: str,
) -> dict:
    local_payroll_file = find_local_file(["*.xls", "*.xlsx"])
    local_allocate_file = find_local_file(["*Allocate*.xlsx", "*allocate*.xlsx", "*.xlsx"])

    payroll_available = (payroll_upload_bytes is not None) or (local_payroll_file is not None)
    allocation_available = (allocation_upload_bytes is not None) or (local_allocate_file is not None)

    if not payroll_available or not allocation_available:
        return {
            "error": (
                "Missing source files. Upload both files from sidebar or place in data/:\n"
                "- Payroll file (.xls)\n"
                "- Allocation file (.xlsx)"
            )
        }

    payroll_encrypted = False
    if payroll_upload_bytes is not None:
        payroll_encrypted = payroll_bytes_is_encrypted(payroll_upload_bytes)
    else:
        payroll_encrypted = payroll_file_is_encrypted(local_payroll_file)

    if payroll_encrypted and not payroll_password.strip():
        return {"error": "Payroll file is password-protected. Please enter payroll password in sidebar first."}

    try:
        if payroll_upload_bytes is not None:
            payroll_raw = load_payroll_xls_from_bytes(payroll_upload_bytes, password=payroll_password)
        else:
            payroll_raw = load_payroll_xls(local_payroll_file, password=payroll_password)
    except Exception:
        return {"error": "Cannot open payroll file. Please verify payroll password and try again."}

    if allocation_upload_bytes is not None:
        allocate_sheets = load_allocation_workbook(allocation_upload_bytes)
    else:
        allocate_sheets = load_allocation_workbook(local_allocate_file)

    # Dynamic mapping sheet hint by month, e.g. 2026-02 -> 2-69, 2026-03 -> 3-69
    month_num_hint = "3"
    try:
        month_num_hint = str(int(str(month_key).split("-")[1]))
    except Exception:
        month_num_hint = "3"
    mapping_sheet_hint = f"{month_num_hint}-69"

    def _norm_id(series: pd.Series) -> pd.Series:
        s = series.astype(str).str.strip()
        s = s.str.replace("\u00A0", "", regex=False)
        s = s.str.replace("\u200b", "", regex=False)
        s = s.str.replace(r"^'+", "", regex=True)
        s = s.str.replace(r"\s+", "", regex=True)
        s = s.str.replace(r"\.0+$", "", regex=True)
        numeric_mask = s.str.match(r"^\d+$", na=False)
        s.loc[numeric_mask] = s.loc[numeric_mask].str.lstrip("0")
        s.loc[numeric_mask & (s == "")] = "0"
        return s

    try:
        payroll_fact = transform_payroll_to_fact(
            payroll_raw,
            month_key=month_key,
            source_file=payroll_file_label,
        )

        # Choose mapping sheet by actual id overlap with payroll to avoid wrong n-69 sheet.
        payroll_ids = set(_norm_id(payroll_fact["employee_id"]).tolist())
        mapping_sheet_name = mapping_sheet_hint
        mapping_raw = None
        best_overlap = -1
        best_name = None

        candidate_names = []
        for sheet_name in allocate_sheets.keys():
            if re.match(r"^\s*\d{1,2}\s*-\s*69\s*$", str(sheet_name)):
                candidate_names.append(sheet_name)
        if mapping_sheet_hint in allocate_sheets:
            candidate_names = [mapping_sheet_hint] + [n for n in candidate_names if n != mapping_sheet_hint]
        if not candidate_names:
            candidate_names = list(allocate_sheets.keys())

        for sheet_name in candidate_names:
            try:
                candidate_master = transform_employee_master(allocate_sheets[sheet_name], month_key=month_key)
                candidate_ids = set(_norm_id(candidate_master["employee_id"]).tolist())
                overlap = len(payroll_ids.intersection(candidate_ids))
                if overlap > best_overlap:
                    best_overlap = overlap
                    best_name = sheet_name
            except Exception:
                continue

        if best_name is not None and best_overlap >= 1:
            mapping_sheet_name = str(best_name)
            mapping_raw = allocate_sheets[best_name].copy()
        else:
            mapping_raw = find_mapping_sheet(allocate_sheets, mapping_sheet_hint=mapping_sheet_hint)
            mapping_sheet_name = mapping_sheet_hint

        employee_master = transform_employee_master(mapping_raw, month_key=month_key)

        # Enrich mapping by other *-69 sheets for ids missing in selected month sheet.
        all_master_frames = []
        for sheet_name in candidate_names:
            try:
                cm = transform_employee_master(allocate_sheets[sheet_name], month_key=month_key)
                cm["_source_sheet"] = str(sheet_name)
                all_master_frames.append(cm)
            except Exception:
                continue
        if all_master_frames:
            all_master = pd.concat(all_master_frames, ignore_index=True)
            all_master["employee_id"] = _norm_id(all_master["employee_id"])
            all_master["employee_name"] = all_master["employee_name"].astype(str).fillna("").str.strip()
            all_master["department"] = all_master["department"].astype(str).fillna("").str.strip()
            all_master["cost_center"] = all_master["cost_center"].astype(str).fillna("").str.strip()
            all_master["employee_type"] = all_master["employee_type"].astype(str).fillna("").str.strip()
            all_master["front_back"] = all_master["front_back"].astype(str).fillna("").str.strip()
            all_master = all_master[all_master["employee_id"] != ""].copy()

            all_master["_score"] = (
                (all_master["employee_name"] != "").astype(int)
                + (all_master["department"] != "").astype(int)
                + (all_master["cost_center"] != "").astype(int)
                + (all_master["employee_type"] != "").astype(int)
                + (all_master["front_back"] != "").astype(int)
            )
            all_master["_prefer_sheet"] = (all_master["_source_sheet"] == str(mapping_sheet_name)).astype(int)
            all_master = (
                all_master.sort_values(
                    ["employee_id", "_score", "_prefer_sheet"],
                    ascending=[True, False, False],
                )
                .drop_duplicates(subset=["employee_id"], keep="first")
                .drop(columns=["_score", "_prefer_sheet", "_source_sheet"])
                .reset_index(drop=True)
            )
            employee_master = all_master

        allocation_mapping_total_cost = extract_total_cost_from_mapping_sheet(mapping_raw)
        allocation_fact = transform_allocation_fact(
            allocate_sheets,
            month_key=month_key,
            mapping_sheet_hint=mapping_sheet_name,
        )
    except Exception as e:
        return {"error": f"Data transform failed: {e}"}

    payroll_with_mapping = apply_employee_mapping(payroll_fact, employee_master)
    employee_summary = build_employee_cost_summary(payroll_with_mapping, allocation_fact)
    allocation_summary = summarize_allocation_by_cost_center(allocation_fact)
    dq_issues = run_dq_checks(payroll_fact, employee_master, allocation_fact)
    recon_table = build_reconciliation_checks(payroll_fact, allocation_fact, employee_master)

    return {
        "payroll_fact": payroll_fact,
        "employee_master": employee_master,
        "allocation_fact": allocation_fact,
        "payroll_with_mapping": payroll_with_mapping,
        "employee_summary": employee_summary,
        "allocation_summary": allocation_summary,
        "dq_issues": dq_issues,
        "recon_table": recon_table,
        "allocation_mapping_total_cost": allocation_mapping_total_cost,
        "mapping_sheet_name": mapping_sheet_name,
    }


def normalize_cost_center(series: pd.Series) -> pd.Series:
    """
    Keep cost center as string. Numeric-like values are zero-padded to 6 chars.
    """
    out = series.astype(str).str.strip().str.replace(r"\.0+$", "", regex=True)
    numeric_mask = out.str.match(r"^\d+$", na=False)
    out.loc[numeric_mask] = out.loc[numeric_mask].str.zfill(6)
    return out


def normalize_employee_code(series: pd.Series) -> pd.Series:
    out = series.astype(str).str.strip().str.replace(r"\.0+$", "", regex=True)
    return out


def is_valid_code(series: pd.Series) -> pd.Series:
    s = series.astype(str).str.strip().str.lower()
    return ~s.isin(["", "nan", "<na>", "none", "null"])


def is_valid_cost_center(series: pd.Series) -> pd.Series:
    s = series.astype(str).str.strip().str.lower()
    return ~s.isin(["", "nan", "<na>", "none", "null"])


def build_thai_header_table(df: pd.DataFrame) -> pd.DataFrame:
    col_map = {
        "month_key": "เดือน",
        "cost_center": "Cost Center",
        "department": "แผนก",
        "direct_payroll_cost": "เงินเดือน",
        "employee_count": "จำนวนพนักงาน",
    }
    out = df.copy()
    new_cols = []
    for c in out.columns:
        if c in col_map:
            new_cols.append(col_map[c])
        elif c.startswith("pay_total_"):
            item_name = c.replace("pay_total_", "").strip()
            new_cols.append(item_name)
        else:
            new_cols.append(c)

    # Streamlit/pyarrow requires unique column names.
    seen = {}
    unique_cols = []
    for name in new_cols:
        key = str(name).strip()
        if key not in seen:
            seen[key] = 0
            unique_cols.append(key)
        else:
            seen[key] += 1
            if key == "เงินเดือน":
                unique_cols.append(f"{key} (รายการ)")
            else:
                unique_cols.append(f"{key} ({seen[key] + 1})")

    out.columns = unique_cols
    return out


def _norm_item_name(name: str) -> str:
    return str(name).strip().lower().replace(" ", "")


def order_pay_item_columns(pay_item_cols: list[str]) -> list[str]:
    # Fixed business order requested by user.
    # Each row allows Thai/English aliases found in source files.
    order_alias_groups = [
        ["เงินเดือน", "salary"],
        ["โบนัส", "bonus"],
        ["ชดเชย+บอกกล่าว"],
        ["วันหยุดคงเหลือ"],
        ["ตกเบิก/ปรับย้อน", "ตกเบิก", "ปรับย้อน", "incentive", "adjust"],
        ["ค่าล่วงเวลา", "ot"],
        ["เงินช่วยเหลือ", "allowance", "all"],
        ["ค่าโทรศัพท์", "tel"],
        ["ค่าครองชีพ"],
        ["ประกันกลุ่ม flex point (tax)", "flex point"],
        ["ค่าตำแหน่ง"],
        ["ค่าพาหนะ", "car allowance", "car"],
        ["ประกันสังคมลูกจ้าง", "soc.", "soc"],
        ["ประกันสังคมนายจ้าง"],
        ["เงินสะสมกองทุนพนักงาน", "pf พนง.", "pf พนง"],
        ["เงินสมทบกองทุนบริษัท", "pf บริษัท"],
        ["esp"],
        ["กยศ"],
        ["บังคับคดี"],
    ]

    base_cols = {c: _norm_item_name(c.replace("pay_total_", "")) for c in pay_item_cols}
    ordered = []
    used = set()

    for aliases in order_alias_groups:
        alias_norms = [_norm_item_name(a) for a in aliases]
        for col, base_norm in base_cols.items():
            if col in used:
                continue
            if any((a in base_norm) or (base_norm in a) for a in alias_norms):
                ordered.append(col)
                used.add(col)

    # Append remaining pay items at the end, keeping original appearance order.
    for c in pay_item_cols:
        if c not in used:
            ordered.append(c)
    return ordered


def order_employee_item_columns(employee_item_cols: list[str]) -> list[str]:
    """
    Order employee pay-item columns using the same business order as cost-center summary.
    Input format: 'รายการค่าใช้จ่าย: <item>'
    """
    to_pay_total = {
        c: f"pay_total_{str(c).replace('รายการค่าใช้จ่าย: ', '').strip()}"
        for c in employee_item_cols
    }
    ordered_pay_total = order_pay_item_columns(list(to_pay_total.values()))
    back_map = {v: k for k, v in to_pay_total.items()}
    ordered = [back_map[c] for c in ordered_pay_total if c in back_map]
    for c in employee_item_cols:
        if c not in ordered:
            ordered.append(c)
    return ordered


def is_cost_pay_item_name(name: str) -> bool:
    n = _norm_item_name(str(name))
    non_cost_tokens = [
        "no.",
        "no",
        "ลำดับ",
        "count",
        "code",
        "รหัส",
        "costcenter",
        "center",
        "department",
        "dept",
        "type",
        "front/back",
        "front",
        "back",
        "remark",
    ]
    return not any(tok in n for tok in non_cost_tokens)


def find_pay_col(columns: list[str], aliases: list[str]) -> str | None:
    alias_norm = [_norm_item_name(a) for a in aliases]
    for c in columns:
        base = str(c).replace("pay_total_", "").replace("รายการค่าใช้จ่าย: ", "").strip()
        c_norm = _norm_item_name(base)
        if any(a in c_norm or c_norm in a for a in alias_norm):
            return c
    return None


def month_key_to_period_text(month_key: str) -> str:
    try:
        d = pd.to_datetime(f"{month_key}-01")
        end = d + pd.offsets.MonthEnd(1)
        return f"{d:%d/%m/%Y}-{end:%d/%m/%Y}"
    except Exception:
        return "01/03/2026-31/03/2026"


def month_key_to_suffix(month_key: str) -> str:
    try:
        d = pd.to_datetime(f"{month_key}-01")
        return f"{d:%m %y}"
    except Exception:
        return "03 26"


def count_csv_tokens(value: object) -> int:
    if value is None:
        return 0
    try:
        if pd.isna(value):
            return 0
    except Exception:
        pass

    if isinstance(value, (list, tuple, set)):
        return len([v for v in value if str(v).strip() != ""])

    text = str(value).strip()
    if text in {"", "<NA>", "nan", "None"}:
        return 0
    return len([v for v in text.split(",") if v.strip() != ""])


st.title("HR Payroll & Allocation Dashboard")
st.caption("Executive Summary | Employee/Payroll | Allocation | Data Quality/Reconciliation")

st.sidebar.header("Data Source")
if "uploader_nonce" not in st.session_state:
    st.session_state["uploader_nonce"] = 0
uploader_nonce = st.session_state["uploader_nonce"]

if st.sidebar.button("Clear / Start Over"):
    st.cache_data.clear()
    st.session_state["uploader_nonce"] = st.session_state.get("uploader_nonce", 0) + 1
    for k in list(st.session_state.keys()):
        if k.startswith("payroll_upload_file_") or k.startswith("allocation_upload_file_") or k.startswith(
            "prev_payroll_upload_file_"
        ):
            del st.session_state[k]
    for k in [
        "payroll_password_input",
        "payroll_password_submitted",
        "filter_month",
        "filter_cost_center",
        "filter_department",
    ]:
        if k in st.session_state:
            del st.session_state[k]
    st.rerun()

st.sidebar.warning(
    "คำเตือน: อัปโหลดไฟล์ให้ถูกประเภทและถูกเดือน "
    "(Payroll .xls / Allocation .xlsx) เพื่อป้องกันตัวเลขคลาดเคลื่อน"
)
payroll_upload = st.sidebar.file_uploader(
    "Upload Payroll (.xls)",
    type=["xls"],
    accept_multiple_files=False,
    key=f"payroll_upload_file_{uploader_nonce}",
)
allocation_upload = st.sidebar.file_uploader(
    "Upload Allocation (.xlsx)",
    type=["xlsx"],
    accept_multiple_files=False,
    key=f"allocation_upload_file_{uploader_nonce}",
)
prev_payroll_upload = st.sidebar.file_uploader(
    "Upload Previous Payroll (.xls, optional)",
    type=["xls"],
    accept_multiple_files=False,
    key=f"prev_payroll_upload_file_{uploader_nonce}",
)

upload_name_errors: list[str] = []
if payroll_upload is not None:
    payroll_name_error = validate_uploaded_filename("payroll", payroll_upload.name)
    if payroll_name_error:
        upload_name_errors.append(payroll_name_error)
if allocation_upload is not None:
    allocation_name_error = validate_uploaded_filename("allocation", allocation_upload.name)
    if allocation_name_error:
        upload_name_errors.append(allocation_name_error)
if prev_payroll_upload is not None:
    prev_payroll_name_error = validate_uploaded_filename("prev_payroll", prev_payroll_upload.name)
    if prev_payroll_name_error:
        upload_name_errors.append(prev_payroll_name_error)

if upload_name_errors:
    for msg in upload_name_errors:
        st.sidebar.error(msg)
    st.error("Invalid uploaded file name format. Please rename file(s) and upload again.")
    st.stop()

local_payroll_file_for_label = find_local_file(["*.xls", "*.xlsx"])
payroll_file_label = (
    payroll_upload.name
    if payroll_upload is not None
    else (local_payroll_file_for_label.name if local_payroll_file_for_label is not None else "payroll.xls")
)
auto_month_key = infer_month_key_from_filename(payroll_file_label) or DEFAULT_MONTH_KEY
auto_prev_month_key = previous_month_key(auto_month_key)

st.sidebar.markdown("<div class='payroll-password-label'>Payroll Password</div>", unsafe_allow_html=True)
payroll_password = st.sidebar.text_input(
    "Payroll Password",
    value=PAYROLL_PASSWORD,
    type="password",
    key="payroll_password_input",
    label_visibility="collapsed",
)
if st.sidebar.button("Submit Password", key="payroll_password_submit_btn"):
    if str(payroll_password).strip():
        st.session_state["payroll_password_submitted"] = True
    else:
        st.session_state["payroll_password_submitted"] = False
        st.sidebar.warning("Please enter password before submit.")
if st.session_state.get("payroll_password_submitted", False):
    st.sidebar.success("Password submitted")
# Force month keys from payroll filename every run (non-editable).
month_key_input = auto_month_key
prev_month_key_input = auto_prev_month_key
st.sidebar.text_input("Month Key (Auto)", value=month_key_input, disabled=True)
st.sidebar.text_input("Previous Month Key (Auto)", value=prev_month_key_input, disabled=True)
if infer_month_key_from_filename(payroll_file_label):
    st.sidebar.caption(f"Auto month from payroll file name: {month_key_input}")

payroll_upload_bytes = payroll_upload.getvalue() if payroll_upload is not None else None
allocation_upload_bytes = allocation_upload.getvalue() if allocation_upload is not None else None
prev_payroll_upload_bytes = prev_payroll_upload.getvalue() if prev_payroll_upload is not None else None

data = run_pipeline(
    payroll_upload_bytes=payroll_upload_bytes,
    allocation_upload_bytes=allocation_upload_bytes,
    payroll_password=payroll_password,
    month_key=month_key_input,
    payroll_file_label=payroll_file_label,
)
if "error" in data:
    st.error(data["error"])
    st.stop()

payroll_fact = data["payroll_fact"]
employee_master = data["employee_master"]
allocation_fact = data["allocation_fact"]
payroll_with_mapping = data["payroll_with_mapping"]
employee_summary = data["employee_summary"]
allocation_summary = data["allocation_summary"]
dq_issues = data["dq_issues"]
recon_table = data["recon_table"]
allocation_mapping_total_cost = data["allocation_mapping_total_cost"]
mapping_sheet_name = data.get("mapping_sheet_name", "3-69")

employee_master["cost_center"] = normalize_cost_center(employee_master["cost_center"])
allocation_fact["cost_center"] = normalize_cost_center(allocation_fact["cost_center"])
payroll_with_mapping["cost_center"] = normalize_cost_center(payroll_with_mapping["cost_center"])
employee_summary["cost_center"] = normalize_cost_center(employee_summary["cost_center"])

# Sidebar Filters
st.sidebar.header("Filters")
months = sorted(payroll_fact["month_key"].astype(str).dropna().unique().tolist())
selected_month = st.sidebar.selectbox(
    "Month",
    months,
    index=0 if months else None,
    key="filter_month",
)

cost_centers = sorted(
    {
        cc
        for cc in pd.concat(
            [employee_master["cost_center"], allocation_fact["cost_center"]],
            ignore_index=True,
        )
        .astype(str)
        .fillna("")
        .tolist()
        if cc.strip() != ""
    }
)
selected_cc = st.sidebar.multiselect("Cost Center", cost_centers, default=[], key="filter_cost_center")

departments = sorted([d for d in employee_master["department"].astype(str).unique().tolist() if d.strip() != ""])
selected_dept = st.sidebar.multiselect("Department", departments, default=[], key="filter_department")

employee_summary_f = employee_summary[employee_summary["month_key"] == selected_month].copy()
allocation_summary_f = allocation_summary[allocation_summary["month_key"] == selected_month].copy()
payroll_with_mapping_f = payroll_with_mapping[payroll_with_mapping["month_key"] == selected_month].copy()
allocation_fact_f = allocation_fact[allocation_fact["month_key"] == selected_month].copy()
dq_issues_f = dq_issues[dq_issues["month_key"] == selected_month].copy()

if selected_cc:
    employee_summary_f = employee_summary_f[employee_summary_f["cost_center"].isin(selected_cc)]
    allocation_summary_f = allocation_summary_f[allocation_summary_f["cost_center"].isin(selected_cc)]
    payroll_with_mapping_f = payroll_with_mapping_f[payroll_with_mapping_f["cost_center"].isin(selected_cc)]
    allocation_fact_f = allocation_fact_f[allocation_fact_f["cost_center"].isin(selected_cc)]
    dq_issues_f = dq_issues_f[dq_issues_f["cost_center"].isin(selected_cc)]

if selected_dept:
    employee_summary_f = employee_summary_f[employee_summary_f["department"].isin(selected_dept)]
    payroll_with_mapping_f = payroll_with_mapping_f[payroll_with_mapping_f["department"].isin(selected_dept)]
    dq_issues_f = dq_issues_f[dq_issues_f["employee_id"].isin(employee_summary_f["employee_id"])]

employee_summary_f["employee_id"] = (
    normalize_employee_code(employee_summary_f["employee_id"])
)
employee_summary_f["cost_center"] = normalize_cost_center(employee_summary_f["cost_center"])
allocation_fact_f["cost_center"] = normalize_cost_center(allocation_fact_f["cost_center"])
payroll_with_mapping_f["cost_center"] = normalize_cost_center(payroll_with_mapping_f["cost_center"])
payroll_with_mapping_f["employee_id"] = normalize_employee_code(payroll_with_mapping_f["employee_id"])

valid_employee_summary_f = employee_summary_f[is_valid_code(employee_summary_f["employee_id"])].copy()
valid_payroll_with_mapping_f = payroll_with_mapping_f[is_valid_code(payroll_with_mapping_f["employee_id"])].copy()
valid_payroll_cost_f = valid_payroll_with_mapping_f[
    valid_payroll_with_mapping_f["pay_item"].astype(str).map(is_cost_pay_item_name)
].copy()
employee_payitem_f = (
    valid_payroll_cost_f.pivot_table(
        index=["month_key", "employee_id"],
        columns="pay_item",
        values="amount",
        aggfunc="sum",
        fill_value=0.0,
    )
    .reset_index()
)
employee_payitem_f["month_key"] = employee_payitem_f["month_key"].astype(str).str.strip()
employee_payitem_f["employee_id"] = normalize_employee_code(employee_payitem_f["employee_id"])
employee_payitem_f = employee_payitem_f.rename(
    columns={
        c: f"รายการค่าใช้จ่าย: {c}"
        for c in employee_payitem_f.columns
        if c not in ["month_key", "employee_id"]
    }
)

# Use salary-only for direct payroll column in summary.
salary_mask = valid_payroll_cost_f["pay_item"].astype(str).str.contains(
    r"เงินเดือน|salary|wage|basic",
    case=False,
    regex=True,
    na=False,
)
salary_payroll_f = valid_payroll_cost_f[salary_mask].copy()
base_for_direct_salary = salary_payroll_f if not salary_payroll_f.empty else valid_payroll_cost_f

direct_cc_summary = (
    base_for_direct_salary.groupby("cost_center", as_index=False)["amount"]
    .sum()
    .rename(columns={"amount": "direct_payroll_cost"})
)
alloc_cc_summary = (
    allocation_fact_f.groupby("cost_center", as_index=False)["allocated_amount"]
    .sum()
    .rename(columns={"allocated_amount": "allocated_overhead"})
)
employee_cc_count = (
    valid_employee_summary_f.groupby("cost_center", as_index=False)["employee_id"]
    .nunique()
    .rename(columns={"employee_id": "employee_count"})
)
cc_department_map = (
    valid_payroll_with_mapping_f[["cost_center", "department"]]
    .dropna()
    .assign(department=lambda d: d["department"].astype(str).str.strip())
)
cc_department_map = cc_department_map[cc_department_map["department"] != ""]
cc_department_map = (
    cc_department_map.groupby("cost_center", as_index=False)["department"]
    .agg(lambda x: ", ".join(sorted(set(x))))
)
pay_item_cc_summary = (
    valid_payroll_cost_f.pivot_table(
        index="cost_center",
        columns="pay_item",
        values="amount",
        aggfunc="sum",
        fill_value=0.0,
    )
    .reset_index()
)
pay_item_cc_summary.columns = [
    "cost_center" if c == "cost_center" else f"pay_total_{str(c).strip()}"
    for c in pay_item_cc_summary.columns
]

cc_summary_f = direct_cc_summary.merge(alloc_cc_summary, on="cost_center", how="outer").fillna(0.0)
cc_summary_f = cc_summary_f.merge(employee_cc_count, on="cost_center", how="left")
cc_summary_f = cc_summary_f.merge(cc_department_map, on="cost_center", how="left")
cc_summary_f = cc_summary_f.merge(pay_item_cc_summary, on="cost_center", how="left")
cc_summary_f["employee_count"] = cc_summary_f["employee_count"].fillna(0).astype(int)
cc_summary_f["department"] = cc_summary_f["department"].fillna("")
base_cols = [
    "cost_center",
    "department",
    "employee_count",
    "direct_payroll_cost",
]
pay_item_cols = [c for c in cc_summary_f.columns if c.startswith("pay_total_")]
if pay_item_cols:
    cc_summary_f[pay_item_cols] = cc_summary_f[pay_item_cols].fillna(0.0)
    # Remove duplicated salary pay-item column; keep main "เงินเดือน" column only.
    salary_like_cols = [
        c
        for c in pay_item_cols
        if _norm_item_name(c.replace("pay_total_", "")) in ["salary", "เงินเดือน"]
    ]
    if salary_like_cols:
        cc_summary_f = cc_summary_f.drop(columns=salary_like_cols, errors="ignore")
        pay_item_cols = [c for c in pay_item_cols if c not in salary_like_cols]
    cc_summary_f["total_employee_cost"] = cc_summary_f[pay_item_cols].sum(axis=1)
else:
    cc_summary_f["total_employee_cost"] = cc_summary_f["direct_payroll_cost"]
ordered_pay_item_cols = order_pay_item_columns(pay_item_cols)
cc_summary_f = cc_summary_f[base_cols + ordered_pay_item_cols]
cc_summary_f.insert(0, "month_key", selected_month)
cc_summary_f = cc_summary_f[
    is_valid_cost_center(cc_summary_f["cost_center"])
].copy()

# KPI on current filtered view only
if not cc_summary_f.empty and ordered_pay_item_cols:
    # Total Cost = direct salary + all remaining pay-item totals
    kpi_total_cost = float(cc_summary_f["direct_payroll_cost"].sum() + cc_summary_f[ordered_pay_item_cols].sum().sum())
else:
    kpi_total_cost = float(cc_summary_f["direct_payroll_cost"].sum()) if not cc_summary_f.empty else 0.0
kpi_total_allocated = float(allocation_fact_f["allocated_amount"].sum()) if not allocation_fact_f.empty else 0.0
kpi_employee_count = int(valid_employee_summary_f["employee_id"].nunique()) if not valid_employee_summary_f.empty else 0
kpi_cost_center_count = int(cc_summary_f["cost_center"].nunique()) if not cc_summary_f.empty else 0

# Reconciliation check: Summary by Cost Center vs Upload Payroll (.xls)
valid_payroll_cc_f = valid_payroll_cost_f[is_valid_cost_center(valid_payroll_cost_f["cost_center"])].copy()
payroll_xls_total = float(valid_payroll_cc_f["amount"].sum()) if not valid_payroll_cc_f.empty else 0.0
summary_cc_total = (
    float(cc_summary_f[ordered_pay_item_cols].sum().sum())
    if (not cc_summary_f.empty and ordered_pay_item_cols)
    else float(cc_summary_f["direct_payroll_cost"].sum()) if not cc_summary_f.empty else 0.0
)
recon_summary_vs_payroll = pd.DataFrame(
    columns=["check_name", "month_key", "left_value", "right_value", "difference", "status"]
)

# Reconciliation checks: duplicate employee name / duplicate employee code
recon_source = valid_payroll_with_mapping_f.copy()
if selected_cc:
    recon_source = recon_source[recon_source["cost_center"].isin(selected_cc)]
if selected_dept:
    recon_source = recon_source[recon_source["department"].isin(selected_dept)]

recon_unique_emp = (
    recon_source[["month_key", "employee_id", "employee_name"]]
    .drop_duplicates()
    .copy()
)
recon_unique_emp["employee_id"] = normalize_employee_code(recon_unique_emp["employee_id"])
recon_unique_emp["employee_name"] = recon_unique_emp["employee_name"].astype(str).str.strip()
recon_unique_emp = recon_unique_emp[is_valid_code(recon_unique_emp["employee_id"])].copy()

dup_name_count = int(
    recon_unique_emp[recon_unique_emp["employee_name"].str.strip() != ""]
    .duplicated(subset=["month_key", "employee_name"], keep=False)
    .sum()
)
dup_code_count = int(
    recon_unique_emp.duplicated(subset=["month_key", "employee_id"], keep=False).sum()
)

recon_duplicate_checks = pd.DataFrame(
    [
        {
            "check_name": "duplicate_employee_name",
            "month_key": selected_month,
            "left_value": float(dup_name_count),
            "right_value": 0.0,
            "difference": float(dup_name_count),
            "status": "ok" if dup_name_count == 0 else "mismatch",
        },
        {
            "check_name": "duplicate_employee_code",
            "month_key": selected_month,
            "left_value": float(dup_code_count),
            "right_value": 0.0,
            "difference": float(dup_code_count),
            "status": "ok" if dup_code_count == 0 else "mismatch",
        },
    ]
)
recon_summary_vs_payroll = pd.concat([recon_summary_vs_payroll, recon_duplicate_checks], ignore_index=True)

# Additional reconciliation checks requested
summary_total_from_rows = (
    float(cc_summary_f["direct_payroll_cost"].sum() + cc_summary_f[ordered_pay_item_cols].sum().sum())
    if (not cc_summary_f.empty and ordered_pay_item_cols)
    else float(cc_summary_f["direct_payroll_cost"].sum()) if not cc_summary_f.empty else 0.0
)

employee_total_from_payroll = float(valid_payroll_cost_f["amount"].sum()) if not valid_payroll_cost_f.empty else 0.0
cost_center_total_from_summary = float(summary_total_from_rows)
recon_check_employee_vs_cc = pd.DataFrame(
    [
        {
            "check_name": "payroll_employee_total_vs_cost_center_total",
            "month_key": selected_month,
            "left_value": employee_total_from_payroll,
            "right_value": cost_center_total_from_summary,
            "difference": employee_total_from_payroll - cost_center_total_from_summary,
            "status": "ok"
            if abs(employee_total_from_payroll - cost_center_total_from_summary) <= 0.01
            else "mismatch",
        }
    ]
)

employee_master_f = employee_master[employee_master["month_key"] == selected_month].copy()
if selected_cc:
    employee_master_f = employee_master_f[employee_master_f["cost_center"].isin(selected_cc)]
if selected_dept:
    employee_master_f = employee_master_f[employee_master_f["department"].isin(selected_dept)]
employee_master_f["cost_center"] = normalize_cost_center(employee_master_f["cost_center"])

payroll_cc_set = set(cc_summary_f["cost_center"].astype(str).str.strip().tolist())
allocation_cc_set = set(
    allocation_fact_f["cost_center"]
    .astype(str)
    .str.strip()
    .loc[is_valid_cost_center(allocation_fact_f["cost_center"])]
    .tolist()
)
mapping_cc_set = set(
    employee_master_f["cost_center"]
    .astype(str)
    .str.strip()
    .loc[is_valid_cost_center(employee_master_f["cost_center"])]
    .tolist()
)
alloc_map_cc_set = allocation_cc_set.union(mapping_cc_set)
missing_in_alloc_map = payroll_cc_set - alloc_map_cc_set
extra_in_alloc_map = alloc_map_cc_set - payroll_cc_set

recon_check_cc_set = pd.DataFrame(
    [
        {
            "check_name": "cost_center_set_match",
            "month_key": selected_month,
            "left_value": float(len(payroll_cc_set)),
            "right_value": float(len(alloc_map_cc_set)),
            "difference": float(len(missing_in_alloc_map) + len(extra_in_alloc_map)),
            "status": "ok" if (len(missing_in_alloc_map) == 0 and len(extra_in_alloc_map) == 0) else "mismatch",
            "missing_in_alloc_map": ", ".join(sorted(missing_in_alloc_map)),
            "extra_in_alloc_map": ", ".join(sorted(extra_in_alloc_map)),
        }
    ]
)

recon_prev_salary = pd.DataFrame(
    columns=["check_name", "month_key", "left_value", "right_value", "difference", "status"]
)
prev_month_salary_diff_detail = pd.DataFrame(
    columns=[
        "employee_id",
        "employee_name",
        "cost_center",
        "department",
        "current_salary",
        "previous_salary",
        "difference",
        "change_type",
    ]
)
if prev_payroll_upload_bytes is not None:
    try:
        prev_payroll_raw = load_payroll_xls_from_bytes(prev_payroll_upload_bytes, password=payroll_password)
        prev_payroll_fact = transform_payroll_to_fact(
            prev_payroll_raw,
            month_key=prev_month_key_input,
            source_file=prev_payroll_upload.name if prev_payroll_upload is not None else "previous_payroll.xls",
        )
        prev_payroll_mapped = apply_employee_mapping(prev_payroll_fact, employee_master)
        prev_payroll_mapped["employee_id"] = normalize_employee_code(prev_payroll_mapped["employee_id"])
        prev_payroll_mapped["cost_center"] = normalize_cost_center(prev_payroll_mapped["cost_center"])
        if selected_cc:
            prev_payroll_mapped = prev_payroll_mapped[prev_payroll_mapped["cost_center"].isin(selected_cc)]
        if selected_dept:
            prev_payroll_mapped = prev_payroll_mapped[prev_payroll_mapped["department"].isin(selected_dept)]
        prev_valid = prev_payroll_mapped[is_valid_code(prev_payroll_mapped["employee_id"])].copy()
        prev_valid_cost = prev_valid[
            prev_valid["pay_item"].astype(str).map(is_cost_pay_item_name)
        ].copy()

        current_salary_mask = valid_payroll_cost_f["pay_item"].astype(str).str.contains(
            r"salary|wage|basic|เงินเดือน",
            case=False,
            regex=True,
            na=False,
        )
        prev_salary_mask = prev_valid_cost["pay_item"].astype(str).str.contains(
            r"salary|wage|basic|เงินเดือน",
            case=False,
            regex=True,
            na=False,
        )
        current_salary_total = float(valid_payroll_cost_f[current_salary_mask]["amount"].sum())
        prev_salary_total = float(prev_valid_cost[prev_salary_mask]["amount"].sum())
        if current_salary_total == 0.0:
            current_salary_total = float(cc_summary_f["direct_payroll_cost"].sum()) if not cc_summary_f.empty else 0.0
        if prev_salary_total == 0.0:
            prev_salary_total = float(prev_valid_cost["amount"].sum()) if not prev_valid_cost.empty else 0.0

        curr_salary_by_emp = (
            valid_payroll_cost_f[current_salary_mask]
            .groupby(["employee_id", "employee_name"], as_index=False)["amount"]
            .sum()
            .rename(columns={"amount": "current_salary"})
        )
        prev_salary_by_emp = (
            prev_valid_cost[prev_salary_mask]
            .groupby(["employee_id", "employee_name"], as_index=False)["amount"]
            .sum()
            .rename(columns={"amount": "previous_salary"})
        )
        curr_emp_info = (
            valid_payroll_with_mapping_f[["employee_id", "cost_center", "department"]]
            .drop_duplicates(subset=["employee_id"])
            .rename(columns={"cost_center": "cost_center_curr", "department": "department_curr"})
        )
        prev_emp_info = (
            prev_valid[["employee_id", "cost_center", "department"]]
            .drop_duplicates(subset=["employee_id"])
            .rename(columns={"cost_center": "cost_center_prev", "department": "department_prev"})
        )
        prev_month_salary_diff_detail = curr_salary_by_emp.merge(
            prev_salary_by_emp,
            on="employee_id",
            how="outer",
            suffixes=("_curr", "_prev"),
        )
        prev_month_salary_diff_detail = prev_month_salary_diff_detail.merge(
            curr_emp_info,
            on="employee_id",
            how="left",
        ).merge(
            prev_emp_info,
            on="employee_id",
            how="left",
        )
        prev_month_salary_diff_detail["employee_name"] = (
            prev_month_salary_diff_detail["employee_name_curr"]
            .fillna(prev_month_salary_diff_detail["employee_name_prev"])
            .fillna("")
        )
        prev_month_salary_diff_detail["cost_center"] = (
            prev_month_salary_diff_detail["cost_center_curr"]
            .fillna(prev_month_salary_diff_detail["cost_center_prev"])
            .fillna("")
        )
        prev_month_salary_diff_detail["department"] = (
            prev_month_salary_diff_detail["department_curr"]
            .fillna(prev_month_salary_diff_detail["department_prev"])
            .fillna("")
        )
        prev_month_salary_diff_detail = prev_month_salary_diff_detail.drop(
            columns=[
                "employee_name_curr",
                "employee_name_prev",
                "cost_center_curr",
                "cost_center_prev",
                "department_curr",
                "department_prev",
            ],
            errors="ignore",
        )
        prev_month_salary_diff_detail["employee_id"] = normalize_employee_code(
            prev_month_salary_diff_detail["employee_id"]
        )
        prev_month_salary_diff_detail["cost_center"] = normalize_cost_center(
            prev_month_salary_diff_detail["cost_center"]
        )
        prev_month_salary_diff_detail["current_salary"] = pd.to_numeric(
            prev_month_salary_diff_detail["current_salary"], errors="coerce"
        ).fillna(0.0)
        prev_month_salary_diff_detail["previous_salary"] = pd.to_numeric(
            prev_month_salary_diff_detail["previous_salary"], errors="coerce"
        ).fillna(0.0)
        prev_month_salary_diff_detail["difference"] = (
            prev_month_salary_diff_detail["current_salary"] - prev_month_salary_diff_detail["previous_salary"]
        )
        prev_month_salary_diff_detail["change_type"] = "changed"
        prev_month_salary_diff_detail.loc[
            (prev_month_salary_diff_detail["previous_salary"] == 0)
            & (prev_month_salary_diff_detail["current_salary"] != 0),
            "change_type",
        ] = "new_in_current"
        prev_month_salary_diff_detail.loc[
            (prev_month_salary_diff_detail["current_salary"] == 0)
            & (prev_month_salary_diff_detail["previous_salary"] != 0),
            "change_type",
        ] = "missing_in_current"
        prev_month_salary_diff_detail = prev_month_salary_diff_detail[
            prev_month_salary_diff_detail["difference"].abs() > 0.01
        ].sort_values("difference", key=lambda s: s.abs(), ascending=False).reset_index(drop=True)

        recon_prev_salary = pd.DataFrame(
            [
                {
                    "check_name": "current_salary_total_vs_previous_month_salary",
                    "month_key": selected_month,
                    "left_value": current_salary_total,
                    "right_value": prev_salary_total,
                    "difference": current_salary_total - prev_salary_total,
                    "status": "same" if abs(current_salary_total - prev_salary_total) <= 0.01 else "review",
                }
            ]
        )
    except Exception as e:
        recon_prev_salary = pd.DataFrame(
            [
                {
                    "check_name": "previous_month_file_parse_error",
                    "month_key": selected_month,
                    "left_value": 0.0,
                    "right_value": 0.0,
                    "difference": 0.0,
                    "status": f"error: {e}",
                }
            ]
        )

recon_summary_vs_payroll = pd.concat(
    [
        recon_summary_vs_payroll,
        recon_check_employee_vs_cc,
        recon_check_cc_set,
        recon_prev_salary,
    ],
    ignore_index=True,
)
if "missing_in_alloc_map" not in recon_summary_vs_payroll.columns:
    recon_summary_vs_payroll["missing_in_alloc_map"] = ""
if "extra_in_alloc_map" not in recon_summary_vs_payroll.columns:
    recon_summary_vs_payroll["extra_in_alloc_map"] = ""

# Build output files in the same structure as provided examples
emp_meta = (
    valid_payroll_with_mapping_f[
        ["employee_id", "employee_name", "department", "cost_center", "employee_type", "front_back"]
    ]
    .drop_duplicates(subset=["employee_id"])
    .copy()
)
emp_meta["employee_id"] = normalize_employee_code(emp_meta["employee_id"])
employee_export = emp_meta.merge(employee_payitem_f, on="employee_id", how="left")
employee_export = employee_export.fillna(0)

pay_cols_emp = [c for c in employee_export.columns if c.startswith("รายการค่าใช้จ่าย: ")]
salary_col_emp = find_pay_col(pay_cols_emp, ["salary", "เงินเดือน"])
bonus_col_emp = find_pay_col(pay_cols_emp, ["bonus", "โบนัส"])
incentive_col_emp = find_pay_col(pay_cols_emp, ["incentive"])
severance_col_emp = find_pay_col(pay_cols_emp, ["ชดเชย+บอกกล่าว"])
leave_col_emp = find_pay_col(pay_cols_emp, ["วันหยุดคงเหลือ"])
ot_col_emp = find_pay_col(pay_cols_emp, ["ot", "ค่าล่วงเวลา"])
all_col_emp = find_pay_col(pay_cols_emp, ["allowance", "เงินช่วยเหลือ", "all"])
car_col_emp = find_pay_col(pay_cols_emp, ["car allowance", "car", "ค่าพาหนะ"])
living_col_emp = find_pay_col(pay_cols_emp, ["ค่าครองชีพ"])
flex_col_emp = find_pay_col(pay_cols_emp, ["flex point", "ประกันกลุ่ม"])
position_col_emp = find_pay_col(pay_cols_emp, ["ค่าตำแหน่ง"])
tel_col_emp = find_pay_col(pay_cols_emp, ["tel", "ค่าโทรศัพท์"])
soc_col_emp = find_pay_col(pay_cols_emp, ["soc", "ประกันสังคมลูกจ้าง"])
pf_emp_col_emp = find_pay_col(pay_cols_emp, ["pf พนง", "เงินสะสมกองทุนพนักงาน"])
pf_comp_col_emp = find_pay_col(pay_cols_emp, ["pf บริษัท", "เงินสมทบกองทุนบริษัท"])
accrued_col_emp = find_pay_col(pay_cols_emp, ["accrued bonus"])

def _v(df: pd.DataFrame, col: str | None) -> pd.Series:
    if col is None or col not in df.columns:
        return pd.Series([0.0] * len(df))
    return pd.to_numeric(df[col], errors="coerce").fillna(0.0)

mkt_output_df = pd.DataFrame(
    {
        "Code": employee_export["employee_id"],
        "No.": range(1, len(employee_export) + 1),
        "Name": employee_export["employee_name"],
        "Department": employee_export["department"],
        "Cost Center": employee_export["cost_center"],
        "Type": employee_export["employee_type"],
        "Front/Back": employee_export["front_back"],
        "Remark": "",
        "COUNT": 1,
        "Salary": _v(employee_export, salary_col_emp),
        "Bonus": _v(employee_export, bonus_col_emp),
        "Incentive": _v(employee_export, incentive_col_emp),
        "ชดเชย+บอกกล่าว": _v(employee_export, severance_col_emp),
        "วันหยุดคงเหลือ": _v(employee_export, leave_col_emp),
        "OT": _v(employee_export, ot_col_emp),
        "Allowance": _v(employee_export, all_col_emp),
        "Car Allowance ": _v(employee_export, car_col_emp),
        "ค่าครองชีพ": _v(employee_export, living_col_emp),
        "ประกันกลุ่ม Flex Point (Tax)": _v(employee_export, flex_col_emp),
        "ค่าตำแหน่ง": _v(employee_export, position_col_emp),
        "Tel": _v(employee_export, tel_col_emp),
        "Soc.": _v(employee_export, soc_col_emp),
        "PF พนง.": _v(employee_export, pf_emp_col_emp),
        "PF บริษัท": _v(employee_export, pf_comp_col_emp),
        "Accrued Bonus": _v(employee_export, accrued_col_emp),
    }
)
# Keep only Type = F-Fix 1 and F-Inc as requested.
type_norm = mkt_output_df["Type"].astype(str).str.strip().str.lower()
mkt_output_df = mkt_output_df[
    type_norm.isin(["f-fix 1", "f-inc"])
].copy()
# Sort Allocate MKT output by Cost Center and re-number running No.
mkt_output_df["Cost Center"] = normalize_cost_center(mkt_output_df["Cost Center"])
mkt_output_df = mkt_output_df.sort_values("Cost Center", ascending=True).reset_index(drop=True)
mkt_output_df["No."] = range(1, len(mkt_output_df) + 1)

cc_type_map = (
    valid_payroll_with_mapping_f[["cost_center", "front_back"]]
    .dropna()
    .groupby("cost_center", as_index=False)["front_back"]
    .first()
    .rename(columns={"front_back": "Type"})
)
accounting_base = cc_summary_f.merge(cc_type_map, on="cost_center", how="left")
accounting_base["Type"] = accounting_base["Type"].fillna("")
accounting_base = accounting_base.sort_values("cost_center").reset_index(drop=True)

pay_cols_cc = [c for c in accounting_base.columns if c.startswith("pay_total_")]
salary_col_cc = find_pay_col(pay_cols_cc, ["salary", "เงินเดือน"])
bonus_col_cc = find_pay_col(pay_cols_cc, ["bonus", "โบนัส"])
incentive_col_cc = find_pay_col(pay_cols_cc, ["incentive"])
severance_col_cc = find_pay_col(pay_cols_cc, ["ชดเชย+บอกกล่าว"])
leave_col_cc = find_pay_col(pay_cols_cc, ["วันหยุดคงเหลือ"])
ot_col_cc = find_pay_col(pay_cols_cc, ["ot", "ค่าล่วงเวลา"])
all_col_cc = find_pay_col(pay_cols_cc, ["allowance", "เงินช่วยเหลือ", "all"])
car_col_cc = find_pay_col(pay_cols_cc, ["car allowance", "car", "ค่าพาหนะ"])
living_col_cc = find_pay_col(pay_cols_cc, ["ค่าครองชีพ"])
flex_col_cc = find_pay_col(pay_cols_cc, ["flex point", "ประกันกลุ่ม"])
position_col_cc = find_pay_col(pay_cols_cc, ["ค่าตำแหน่ง"])
tel_col_cc = find_pay_col(pay_cols_cc, ["tel", "ค่าโทรศัพท์"])
soc_col_cc = find_pay_col(pay_cols_cc, ["soc", "ประกันสังคมลูกจ้าง"])
pf_emp_col_cc = find_pay_col(pay_cols_cc, ["pf พนง", "เงินสะสมกองทุนพนักงาน"])
pf_comp_col_cc = find_pay_col(pay_cols_cc, ["pf บริษัท", "เงินสมทบกองทุนบริษัท"])
accrued_col_cc = find_pay_col(pay_cols_cc, ["accrued bonus"])

# Salary column should always come from direct payroll (เงินเดือนหลัก) per cost center.
salary_series_cc = pd.to_numeric(accounting_base["direct_payroll_cost"], errors="coerce").fillna(0.0)

accounting_output_df = pd.DataFrame(
    {
        "No.": range(1, len(accounting_base) + 1),
        "Center": accounting_base["cost_center"],
        "Dept.": accounting_base["department"],
        "Type": accounting_base["Type"],
        "Salary": salary_series_cc,
        "Bonus": _v(accounting_base, bonus_col_cc),
        "Incentive": _v(accounting_base, incentive_col_cc),
        "ชดเชย+บอกกล่าว": _v(accounting_base, severance_col_cc),
        "วันหยุดคงเหลือ": _v(accounting_base, leave_col_cc),
        "OT": _v(accounting_base, ot_col_cc),
        "ALL": _v(accounting_base, all_col_cc),
        "Car": _v(accounting_base, car_col_cc),
        "ค่าครองชีพ": _v(accounting_base, living_col_cc),
        "ประกันกลุ่ม Flex Point (Tax)": _v(accounting_base, flex_col_cc),
        "ค่าตำแหน่ง": _v(accounting_base, position_col_cc),
        "Tel": _v(accounting_base, tel_col_cc),
        "Soc.": _v(accounting_base, soc_col_cc),
        "PF พนง.": _v(accounting_base, pf_emp_col_cc),
        "PF บริษัท": _v(accounting_base, pf_comp_col_cc),
        "Accrued Bonus": _v(accounting_base, accrued_col_cc),
    }
)

st.sidebar.subheader("Downloads")
period_text = month_key_to_period_text(selected_month)
month_suffix = month_key_to_suffix(selected_month)
mkt_output_filename = f"Allocate MKT {month_suffix}.xlsx"
acc_output_filename = f"Allocate salary {month_suffix} for Accounting.xlsx"
st.sidebar.download_button(
    "Download Allocate MKT",
    data=to_mkt_excel_bytes(mkt_output_df),
    file_name=mkt_output_filename,
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
st.sidebar.download_button(
    "Download Allocate Accounting",
    data=to_accounting_excel_bytes(accounting_output_df, period_text=period_text),
    file_name=acc_output_filename,
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
# Main tabs
tab1, tab2, tab3 = st.tabs(
    ["Executive Summary", "Employee / Payroll", "Reconciliation"]
)

with tab1:
    c1, c3, c4 = st.columns(3)
    c1.metric("Total Cost", f"{kpi_total_cost:,.2f}")
    c3.metric("Employee Count", f"{kpi_employee_count:,}")
    c4.metric("Cost Center Count", f"{kpi_cost_center_count:,}")

    st.subheader("Summary by Cost Center")
    if cc_summary_f.empty:
        st.info("No cost center summary for selected filters.")
    else:
        cc_summary_display = build_thai_header_table(
            cc_summary_f.sort_values("direct_payroll_cost", ascending=False)
        )
        number_cols = cc_summary_display.select_dtypes(include="number").columns.tolist()
        format_map = {c: "{:,.2f}" for c in number_cols}
        table_height = 36 * (10 + 1)
        st.dataframe(
            cc_summary_display.style.format(format_map),
            column_order=list(cc_summary_display.columns),
            use_container_width=False,
            hide_index=True,
            height=table_height,
        )
        dept_summary = employee_summary_f.copy()
        dept_summary["department"] = dept_summary["department"].astype(str).str.strip()
        dept_summary.loc[dept_summary["department"] == "", "department"] = "Unknown"
        dept_summary = (
            dept_summary.groupby("department", as_index=False)[
                ["direct_payroll_cost", "allocated_overhead_employee", "fully_allocated_cost"]
            ]
            .sum()
            .sort_values("fully_allocated_cost", ascending=False)
            .head(15)
        )
        dept_chart_df = dept_summary.melt(
            id_vars=["department"],
            value_vars=["direct_payroll_cost", "allocated_overhead_employee"],
            var_name="cost_type",
            value_name="amount",
        )
        fig_dept_stack = px.bar(
            dept_chart_df,
            x="department",
            y="amount",
            color="cost_type",
            barmode="stack",
            title="Cost Breakdown by Department (Direct vs Allocated)",
            color_discrete_sequence=["#60a5fa", "#1d4ed8"],
        )
        fig_dept_stack.update_layout(showlegend=False)
        fig_dept_stack.add_scatter(
            x=dept_summary["department"],
            y=dept_summary["fully_allocated_cost"],
            mode="text",
            text=[f"{v:,.2f}" for v in dept_summary["fully_allocated_cost"]],
            textposition="top center",
            showlegend=False,
            hoverinfo="skip",
        )
        st.plotly_chart(fig_dept_stack, use_container_width=True)

with tab2:
    st.subheader("Employee / Payroll")
    if employee_summary_f.empty:
        st.info("No data for selected filters.")
    else:
        employee_display = employee_summary_f[
            [
                "month_key",
                "employee_id",
                "employee_name",
                "department",
                "cost_center",
                "direct_payroll_cost",
            ]
        ].copy()
        employee_display = employee_display.rename(
            columns={
                "month_key": "เดือน",
                "employee_id": "รหัสพนักงาน",
                "employee_name": "ชื่อพนักงาน",
                "department": "แผนก",
                "cost_center": "Cost Center",
                "direct_payroll_cost": "เงินเดือน",
            }
        )
        employee_display["เดือน"] = employee_display["เดือน"].astype(str).str.strip()
        employee_display["รหัสพนักงาน"] = normalize_employee_code(employee_display["รหัสพนักงาน"])
        # Add full pay-item columns (เงินเดือน/โบนัส/ฯลฯ) per employee.
        employee_display = employee_display.merge(
            employee_payitem_f,
            left_on=["เดือน", "รหัสพนักงาน"],
            right_on=["month_key", "employee_id"],
            how="left",
        )
        employee_display = employee_display.drop(columns=["month_key", "employee_id"], errors="ignore")
        fixed_cols = [
            "เดือน",
            "รหัสพนักงาน",
            "ชื่อพนักงาน",
            "แผนก",
            "Cost Center",
        ]
        fixed_cols_existing = [c for c in fixed_cols if c in employee_display.columns]
        pay_item_cols_tab2 = [
            c
            for c in employee_payitem_f.columns
            if c not in ["month_key", "employee_id"]
        ]
        pay_item_cols_tab2 = order_employee_item_columns(pay_item_cols_tab2)
        for c in pay_item_cols_tab2:
            if c not in employee_display.columns:
                employee_display[c] = 0.0
        employee_display = employee_display[fixed_cols_existing + pay_item_cols_tab2]
        number_cols_tab2 = employee_display.select_dtypes(include="number").columns.tolist()
        format_map_tab2 = {c: "{:,.2f}" for c in number_cols_tab2}
        sort_col = pay_item_cols_tab2[0] if pay_item_cols_tab2 else (
            "รหัสพนักงาน" if "รหัสพนักงาน" in employee_display.columns else employee_display.columns[0]
        )
        st.dataframe(
            employee_display.sort_values(sort_col, ascending=False).style.format(format_map_tab2),
            column_order=list(employee_display.columns),
            use_container_width=False,
            hide_index=True,
            height=36 * (10 + 1),
        )

with tab3:
    st.subheader("Reconciliation Checks")
    # Debug summary for employee coverage and mapping status
    payroll_month_all = payroll_fact[payroll_fact["month_key"] == selected_month].copy()
    payroll_month_all["employee_id"] = normalize_employee_code(payroll_month_all["employee_id"])
    payroll_month_all = payroll_month_all[is_valid_code(payroll_month_all["employee_id"])].copy()

    payroll_after_filter = valid_payroll_with_mapping_f.copy()
    payroll_after_filter["employee_id"] = normalize_employee_code(payroll_after_filter["employee_id"])
    payroll_after_filter = payroll_after_filter[is_valid_code(payroll_after_filter["employee_id"])].copy()

    mapped_after_filter = payroll_after_filter[
        is_valid_cost_center(payroll_after_filter["cost_center"])
    ].copy()
    unmapped_after_filter = payroll_after_filter[
        ~is_valid_cost_center(payroll_after_filter["cost_center"])
    ].copy()

    total_emp_month = int(payroll_month_all["employee_id"].nunique()) if not payroll_month_all.empty else 0
    total_emp_after_filter = int(payroll_after_filter["employee_id"].nunique()) if not payroll_after_filter.empty else 0
    mapped_emp_count = int(mapped_after_filter["employee_id"].nunique()) if not mapped_after_filter.empty else 0
    unmapped_emp_count = int(unmapped_after_filter["employee_id"].nunique()) if not unmapped_after_filter.empty else 0

    with st.expander("Debug: Employee Coverage", expanded=False):
        d1, d2, d3, d4 = st.columns(4)
        d1.metric("Employees in payroll (month)", f"{total_emp_month:,}")
        d2.metric("Employees after filters", f"{total_emp_after_filter:,}")
        d3.metric("Mapped employees", f"{mapped_emp_count:,}")
        d4.metric("Unmapped employees", f"{unmapped_emp_count:,}")

        if unmapped_emp_count > 0:
            unmapped_list = (
                unmapped_after_filter[["employee_id", "employee_name"]]
                .drop_duplicates()
                .sort_values("employee_id")
                .head(30)
            )
            st.caption("Sample unmapped employees (max 30):")
            st.dataframe(unmapped_list, use_container_width=True, hide_index=True)

    recon_desc = {
        "duplicate_employee_name": "ตรวจชื่อพนักงานซ้ำในเดือนเดียวกัน",
        "duplicate_employee_code": "ตรวจรหัสพนักงานซ้ำในเดือนเดียวกัน",
        "summary_total_vs_cost_center_sum": "ตรวจว่า Total Cost KPI เท่ากับผลรวมจาก Summary by Cost Center",
        "payroll_employee_total_vs_cost_center_total": "ตรวจยอดรวม Payroll เทียบกับยอดรวมตาม Cost Center",
        "cost_center_set_match": "ตรวจชุดรหัส Cost Center ระหว่าง Payroll กับ Allocation/Mapping",
        "current_salary_total_vs_previous_month_salary": "Compare current month salary total with previous month salary total",
        "previous_month_file_parse_error": "Cannot parse previous payroll file",
    }
    recon_display = recon_summary_vs_payroll.copy()
    recon_display["description"] = recon_display["check_name"].map(recon_desc).fillna("")
    recon_display["missing_count"] = recon_display["missing_in_alloc_map"].apply(count_csv_tokens)
    recon_display["extra_count"] = recon_display["extra_in_alloc_map"].apply(count_csv_tokens)
    recon_display = recon_display.rename(
        columns={
            "check_name": "check",
            "month_key": "month",
            "left_value": "left_value",
            "right_value": "right_value",
            "difference": "difference",
            "status": "status",
            "missing_in_alloc_map": "missing_in_alloc_map",
            "extra_in_alloc_map": "extra_in_alloc_map",
        }
    )
    recon_display = recon_display[
        [
            "check",
            "description",
            "month",
            "left_value",
            "right_value",
            "difference",
            "status",
            "extra_in_alloc_map",
        ]
    ]
    st.dataframe(recon_display, use_container_width=True, hide_index=True)

    st.caption("พนักงานที่เงินเดือนต่างจากเดือนก่อนหน้า")
    if prev_payroll_upload_bytes is None:
        st.info("Upload Previous Payroll (.xls) to see employee-level differences.")
    elif prev_month_salary_diff_detail.empty:
        st.success("No employee-level salary differences found.")
    else:
        diff_count = int(len(prev_month_salary_diff_detail))
        diff_net = float(prev_month_salary_diff_detail["difference"].sum())
        m1, m2 = st.columns(2)
        m1.metric("Employees Changed", f"{diff_count:,}")
        m2.metric("Net Difference", f"{diff_net:,.2f}")
        st.caption(
            f"Changed employees (unique): {prev_month_salary_diff_detail['employee_id'].astype(str).nunique():,}"
        )

        prev_diff_display = prev_month_salary_diff_detail[
            [
                "employee_id",
                "employee_name",
                "cost_center",
                "department",
                "current_salary",
                "previous_salary",
                "difference",
                "change_type",
            ]
        ].rename(
            columns={
                "employee_id": "รหัสพนักงาน",
                "employee_name": "ชื่อพนักงาน",
                "cost_center": "Cost Center",
                "department": "แผนก",
                "current_salary": "เงินเดือนเดือนปัจจุบัน",
                "previous_salary": "เงินเดือนเดือนก่อน",
                "difference": "ผลต่าง",
                "change_type": "สถานะ",
            }
        )
        st.download_button(
            "Download Salary Difference CSV",
            data=prev_diff_display.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig"),
            file_name=f"salary_diff_{selected_month}_vs_{prev_month_key_input}.csv",
            mime="text/csv",
        )

        st.dataframe(
            prev_diff_display.style.format(
                {
                    "เงินเดือนเดือนปัจจุบัน": "{:,.2f}",
                    "เงินเดือนเดือนก่อน": "{:,.2f}",
                    "ผลต่าง": "{:,.2f}",
                }
            ),
            use_container_width=True,
            hide_index=True,
            height=36 * (10 + 1),
        )
