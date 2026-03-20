#!/usr/bin/env python3
import re
from typing import Dict, List, Mapping, Optional, Tuple
from urllib.parse import urlparse

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter


PUBLIC_EMAIL_DOMAINS = {
    "gmail.com",
    "yahoo.com",
    "hotmail.com",
    "outlook.com",
    "aol.com",
    "icloud.com",
    "proton.me",
    "protonmail.com",
    "mail.com",
    "gmx.com",
}

EMAIL_REGEX = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")
WORD_REGEX = re.compile(r"\w+")

INPUT_FILE = "DataCheck_DemoCode.xlsx"
OUTPUT_XLSX = "Lead_Data_Validation_Results.xlsx"
OUTPUT_CSV = "Lead_Data_Validation_Results.csv"


# ---------------- BASIC HELPERS ----------------
def norm(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    return str(v).strip()


def parse_req(req: str) -> Dict[str, str]:
    out = {}
    for part in req.split("|"):
        if ":" in part:
            k, v = part.split(":", 1)
            out[k.strip().lower()] = v.strip()
    return out


def email_domain(e):
    return e.lower().split("@")[-1] if "@" in e else ""


def base_domain(d):
    d = d.replace("www.", "")
    parts = d.split(".")
    return ".".join(parts[-2:]) if len(parts) >= 2 else d


def normalize_url_for_hyperlink(raw_link: str) -> Optional[str]:
    link = norm(raw_link)
    if not link:
        return None

    # Keep valid absolute URLs as-is.
    parsed = urlparse(link)
    if parsed.scheme in {"http", "https"} and parsed.netloc:
        return link

    # Support values like "www.company.com/profile" or "linkedin.com/in/...".
    candidate = f"https://{link}"
    parsed_candidate = urlparse(candidate)
    if parsed_candidate.netloc and "." in parsed_candidate.netloc:
        return candidate
    return None


# ---------------- CHECKS ----------------
def check_prooflink(link, email):
    link = norm(link).lower()
    email = norm(email)

    if not link:
        return "INVALID", "Missing prooflink"

    if "linkedin.com/in/" in link or "zoominfo.com/p/" in link:
        return "VALID", "Valid prooflink: trusted source (LinkedIn/ZoomInfo)"

    host = urlparse(link).netloc.replace("www.", "")
    if host and base_domain(host) == base_domain(email_domain(email)):
        return (
            "VALID",
            "Valid prooflink: official website matches corporate email domain",
        )

    return (
        "INVALID",
        "Invalid prooflink source: expected linkedin.com/in/, zoominfo.com/p/, or company domain match",
    )


def check_title(row, req_map):
    title = norm(row.get("title")).lower()
    if not title:
        return "INVALID", "Missing title"

    keywords = req_map.get("keywords", "").lower()
    level = req_map.get("job_level", "").lower()

    issues = []
    matched_keyword = None

    if keywords and "any" not in keywords:
        title_tokens = set(WORD_REGEX.findall(title))
        for keyword in [k.strip() for k in re.split(r",|;", keywords) if k.strip()]:
            keyword_tokens = set(WORD_REGEX.findall(keyword))
            if keyword_tokens and keyword_tokens.issubset(title_tokens):
                matched_keyword = keyword
                break
        if not matched_keyword:
            issues.append("Missing keywords")

    level_map = {"c": 6, "vp": 5, "director": 4, "manager": 3, "lead": 2}
    req_level = 0
    for k, v in level_map.items():
        if k in level:
            req_level = v

    title_level = 1
    for k, v in level_map.items():
        if k in title:
            title_level = v

    if req_level and title_level < req_level:
        issues.append("Low title level")

    if issues:
        detail = []
        if "Missing keywords" in issues:
            detail.append("Title missing required keywords")
        if "Low title level" in issues:
            detail.append("Title level below requirement")
        return "INVALID", "; ".join(detail)

    if keywords and "any" not in keywords:
        return "VALID", f"Title contains keywords: {matched_keyword}"
    return "VALID", "No keyword requirements"


def check_nwc(status):
    s = norm(status).lower()
    if not s or s == "valid":
        return "VALID", "NWC status clear"
    if s == "a":
        return "INVALID", "Retrieved lead"
    if s == "!":
        return "INVALID", "Suspicious lead"
    if s in {"r", "no info", "no company match"}:
        return "RECHECK", "Manual recheck needed"
    return "INVALID", f"Unknown NWC status: {s}"


def check_other(row):
    needed = ["first_name", "last_name", "company", "email"]
    missing = [c for c in needed if not norm(row.get(c))]
    if missing:
        return "INVALID", f"Missing required data: {', '.join(missing)}"

    email = norm(row.get("email")).lower()
    if not EMAIL_REGEX.match(email):
        return "INVALID", "Invalid email format"
    if email_domain(email) in PUBLIC_EMAIL_DOMAINS:
        return "INVALID", "Email must be corporate (public mailbox detected)"

    return "VALID", "Lead data matches requirements"


def validate(row):
    sub = norm(row.get("sub status")).lower()
    req_map = parse_req(norm(row.get("req")))

    if "nwc" in sub:
        return check_nwc(row.get("status"))
    if "prooflink" in sub:
        return check_prooflink(row.get("prooflink"), row.get("email"))
    if "title" in sub:
        return check_title(row, req_map)
    if "other" in sub:
        return check_other(row)

    return check_other(row)


# ---------------- MAIN ----------------
def main():
    df = pd.read_excel(INPUT_FILE)

    results, comments = [], []

    for row in df.to_dict(orient="records"):
        r, c = validate(row)
        results.append(r)
        comments.append(c)

    df["Result"] = results
    df["Comment"] = comments

    out_xlsx = OUTPUT_XLSX
    out_csv = OUTPUT_CSV

    df.to_excel(out_xlsx, index=False)
    df.to_csv(out_csv, index=False)

    # -------- FIX EXCEL FORMAT --------
    wb = load_workbook(out_xlsx)
    ws = wb.active

    hyperlink_cols = set()
    for idx, cell in enumerate(ws[1], start=1):
        header = norm(cell.value).lower()
        if header in {"prooflink", "employees_prooflink"}:
            hyperlink_cols.add(idx)

    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)

        for cell in col:
            val = str(cell.value) if cell.value else ""
            max_len = max(max_len, len(val))

            # wrap long text columns
            if col_letter in ["N", "O", "F"]:  # Result, Comment, req
                cell.alignment = Alignment(wrap_text=True, vertical="top")
            else:
                cell.alignment = Alignment(vertical="center")

            # Force prooflink columns to be clickable URLs when possible.
            if cell.column in hyperlink_cols and cell.row > 1 and cell.value:
                raw = str(cell.value).strip()
                target = normalize_url_for_hyperlink(raw)
                if not target:
                    # Fallback keeps cells clickable even for imperfect source formatting.
                    target = f"https://{raw}"
                cell.hyperlink = target
                cell.style = "Hyperlink"

        # limit width so text wraps instead of overflowing
        ws.column_dimensions[col_letter].width = min(max_len + 2, 45)

    wb.save(out_xlsx)

    print("Done:", out_xlsx, out_csv)


if __name__ == "__main__":
    main()
