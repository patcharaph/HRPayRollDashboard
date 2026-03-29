# HR Payroll & Allocation Dashboard

Interactive dashboard ด้วย Streamlit สำหรับวิเคราะห์ Payroll, สรุปตาม Cost Center, ตรวจ Reconciliation และ export ไฟล์ตามรูปแบบที่กำหนด

## Current Scope

- Tabs ในแอป
  - Executive Summary
  - Employee / Payroll
  - Reconciliation
- Filters
  - Month
  - Cost Center
  - Department
- Upload
  - Payroll `.xls` (รองรับไฟล์มีรหัสผ่าน)
  - Allocation `.xlsx`
- Export
  - `Allocate MKT <MM YY>.xlsx`
  - `Allocate salary <MM YY> for Accounting.xlsx`

## Tech Stack

- `streamlit`
- `pandas`
- `numpy`
- `plotly`
- `openpyxl`
- `xlrd`
- `msoffcrypto-tool`

## Project Structure

```text
HRPayRollDaskboard/
├── app.py
├── requirements.txt
├── README.md
├── .gitignore
├── data/                      # optional local input files
├── output/                    # exported output files
└── src/
    ├── __init__.py
    ├── load_payroll.py
    ├── load_allocate.py
    ├── transform_payroll.py
    ├── transform_allocate.py
    ├── metrics.py
    ├── reconcile.py
    ├── dq_checks.py
    ├── export_csv.py
    └── export_excel.py
```

## Installation

```bash
pip install -r requirements.txt
```

## Run

```bash
streamlit run app.py
```

## How To Use

1. เปิดแอปด้วย `streamlit run app.py`
2. อัปโหลดไฟล์จาก Sidebar
   - Payroll `.xls`
   - Allocation `.xlsx`
3. ถ้าไฟล์ payroll ถูกล็อก ให้กรอก `Payroll Password`
4. เลือก `Month Key (YYYY-MM)`
5. ใช้ filters ตามต้องการ
6. ดาวน์โหลดไฟล์ output หรือกด `Export XLSX to output/`

## KPI Definition (Current)

- **Total Cost**: ยอดรวมค่าใช้จ่ายทั้งหมดจาก payroll ตาม filter ปัจจุบัน
  - เงินเดือนหลัก + รายการค่าใช้จ่ายอื่นที่จัดเป็น cost item
- **Total Allocated Cost**: ยอดรวมจาก allocation fact ตาม filter
- **Employee Count**: จำนวนรหัสพนักงาน (Code) ที่ valid และไม่ซ้ำ
- **Cost Center Count**: จำนวน cost center ที่ valid ตามมุมมองปัจจุบัน

## Reconciliation Checks (Current)

- `duplicate_employee_name`
- `duplicate_employee_code`
- `payroll_employee_total_vs_cost_center_total`
- `cost_center_set_match`
  - มีรายละเอียด `extra_in_alloc_map` เพื่อบอกรหัสที่เกิน

## Output Files

ระบบ export เป็น `.xlsx` ตามรูปแบบตัวอย่าง:

1. `Allocate MKT <MM YY>.xlsx`
   - Sheet: `Value`, `Sheet1`
   - กรอง Type ให้เหลือเฉพาะ `F-Fix 1` และ `F-Inc`

2. `Allocate salary <MM YY> for Accounting.xlsx`
   - Sheet: `Allocate salary for Accounting`, `Sheet1`
   - มี report header ด้านบนก่อนตาราง

> `<MM YY>` มาจาก `Month Key` ที่เลือก เช่น `2026-03` -> `03 26`

## Data Privacy

- ไฟล์ที่ upload ผ่าน UI ประมวลผลในหน่วยความจำ (RAM)
- ไม่บันทึกไฟล์ input ลงดิสก์อัตโนมัติ
- หากปิด/รีสตาร์ทแอป ต้อง upload ใหม่

## Git / Deployment Notes

- แนะนำใช้ GitHub Private Repo
- `.gitignore` ตั้งค่าไม่ให้ push `data/` และ `output/`
- Deploy แนะนำเริ่มที่ Streamlit Community Cloud

## Troubleshooting

- ถ้าเจอปัญหาอ่าน `.xls` ไม่ได้ ให้ตรวจรหัสผ่าน
- ถ้าคอลัมน์ไม่ตรง template จริง ให้ส่งชื่อคอลัมน์ตัวอย่างมาเพื่อ lock mapping เพิ่ม
- ถ้า output ยังไม่ตรงตัวอย่าง ให้แนบตัวอย่างไฟล์ล่าสุดจาก `output/` เพื่อปรับ format 1:1