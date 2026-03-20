#!/usr/bin/env python3
import pandas as pd
from openpyxl import load_workbook

src = pd.read_excel("DataCheck_DemoCode.xlsx")
out = pd.read_excel("DataCheck_DemoCode_processed.xlsx")

wb = load_workbook("DataCheck_DemoCode_processed.xlsx")
ws = wb[wb.sheetnames[0]]

assert len(src) == len(out), f"Row count mismatch: {len(src)} vs {len(out)}"
assert "Result" in out.columns, "Missing Result column"
assert "Comment" in out.columns, "Missing Comment column"

assert ws.cell(row=1, column=14).value == "Result", "Result header is not in column N"
assert ws.cell(row=1, column=15).value == "Comment", "Comment header is not in column O"

allowed = {"VALID", "INVALID", "RECHECK"}
bad = out[~out["Result"].astype(str).isin(allowed)]
assert bad.empty, f'Unexpected Result values: {bad["Result"].dropna().unique()[:10]}'

empty_comments = out[out["Comment"].fillna("").astype(str).str.strip() == ""]
assert empty_comments.empty, f"Empty comments found: {len(empty_comments)}"

for row_idx in range(2, len(src) + 2):
    r_visible = str(ws.cell(row=row_idx, column=14).value or "").strip()
    c_visible = str(ws.cell(row=row_idx, column=15).value or "").strip()
    assert r_visible in {
        "VALID",
        "INVALID",
        "RECHECK",
    }, f"Invalid N cell value at row {row_idx}: {r_visible}"
    assert c_visible != "", f"Empty O cell comment at row {row_idx}"

n1 = out[out["sub status"].astype(str).str.strip().str.lower() == "n1: nwc"]
for i, row in n1.iterrows():
    st = str(row.get("status", "")).strip().lower()
    res = str(row.get("Result", "")).strip()
    if st in ("", "valid"):
        assert res == "VALID", f"N1 expected VALID at row {i + 2} status={st}"
    elif st == "a":
        assert res == "INVALID", f"N1 expected INVALID at row {i + 2} status=a"
    elif st == "!":
        assert res == "INVALID", f"N1 expected INVALID at row {i + 2} status=!"
    elif st in ("r", "no info", "no company match"):
        assert res == "RECHECK", f"N1 expected RECHECK at row {i + 2} status={st}"

print("QA OK")
print("Rows:", len(out))
print("Result distribution:", out["Result"].value_counts(dropna=False).to_dict())
print("N1 rows checked:", len(n1))
print("Column checks: N=Result, O=Comment")
