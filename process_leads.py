#!/usr/bin/env python3
import argparse
import re
from typing import Callable, Dict, List, Mapping, Optional, Tuple
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
    "yandex.com",
    "qq.com",
}

RESULT_COL_IDX = 14  # Column N
COMMENT_COL_IDX = 15  # Column O

EMAIL_REGEX = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")
TAG_REGEX = re.compile(r"<[^>]+>")
SPACE_REGEX = re.compile(r"\s+")

COUNTRY_ALIASES = {
    "usa": ["usa", "united states", "us", "u.s.", "united states of america"],
    "uk": ["uk", "united kingdom", "britain", "great britain"],
    "uae": ["uae", "united arab emirates", "u.a.e."],
}

COMMENT_NORMALIZATION = {
    "Title does not contain required keywords": "Title missing required keywords",
    "Title level does not meet requirement": "Title level below requirement",
    "Industry: Industry does not match requirements": "Industry does not match requirements",
    "Status requires manual recheck": "Manual recheck needed",
    "Lead data matches requirements": "Lead data matches requirements",
    "Title contains required keyword(s)": "Title matches required keywords",
    "No keyword requirements": "No keyword requirements",
    "Valid prooflink source": "Valid prooflink",
    "Prooflink matches corporate domain": "Prooflink matches company domain",
    "Prooflink does not match allowed sources": "Invalid prooflink source",
}


def norm_text(value: object) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and pd.isna(value):
        return ""
    return str(value).strip()


def clean_text(value: object) -> str:
    text = norm_text(value)
    text = TAG_REGEX.sub(" ", text)
    text = SPACE_REGEX.sub(" ", text)
    return text.strip()


def parse_req_map(req: str) -> Dict[str, str]:
    result: Dict[str, str] = {}
    for part in req.split("|"):
        if ":" not in part:
            continue
        key, value = part.split(":", 1)
        result[key.strip().lower()] = value.strip()
    return result


def to_slug_tokens(text: str) -> List[str]:
    text = text.lower()
    text = re.sub(r"[^a-z0-9\s]", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    if not text:
        return []
    return [t for t in text.split(" ") if len(t) > 1]


def normalize_tokens(text: str) -> set:
    return set(to_slug_tokens(text))


def contains_word(text: str, word: str) -> bool:
    normalized_word = SPACE_REGEX.sub(" ", word.strip().lower())
    if not normalized_word:
        return False
    return re.search(rf"\b{re.escape(normalized_word)}\b", text.lower()) is not None


def location_matches_candidate(location: str, candidate: str) -> bool:
    c = candidate.strip().lower()
    aliases = COUNTRY_ALIASES.get(c, [c])
    return any(contains_word(location, alias) for alias in aliases)


def extract_keywords(req_map: Dict[str, str], req_raw: str) -> Optional[List[str]]:
    raw = req_map.get("keywords", "")
    raw_l = raw.lower()
    if not raw:
        return None
    if "any" in raw_l:
        return None

    if "see comment" not in raw_l:
        candidates = [k.strip() for k in re.split(r",|;", raw) if k.strip()]
        return candidates or None

    plain = clean_text(req_raw)
    matches = re.findall(r"keywords?\s*:\s*([^|]+)", plain, flags=re.IGNORECASE)
    bucket: List[str] = []
    for m in matches:
        segment = re.split(
            r"job\s+area|job\s+levels?|industry|comment", m, flags=re.IGNORECASE
        )[0]
        for item in re.split(r",|;", segment):
            item = item.strip(" .")
            if item:
                bucket.append(item)

    dedup = []
    seen = set()
    for item in bucket:
        key = item.lower()
        if key not in seen:
            seen.add(key)
            dedup.append(item)
    return dedup or None


def infer_min_level(level_text: str, req_raw: str) -> int:
    level_text = (level_text or "").lower().strip()
    if "see comment" in level_text:
        plain = clean_text(req_raw)
        m = re.search(r"job\s+levels?\s*:\s*([^|]+)", plain, flags=re.IGNORECASE)
        if m:
            level_text = m.group(1).lower().strip()

    if not level_text or "any" in level_text:
        return 0
    if "c-level" in level_text or "chief" in level_text:
        return 6
    if (
        "vp" in level_text
        or "vice president" in level_text
        or "evp" in level_text
        or "svp" in level_text
    ):
        return 5
    if "director" in level_text:
        return 4
    if "manager" in level_text:
        return 3
    if "lead" in level_text or "head" in level_text:
        return 2
    return 0


def infer_title_level(title: str) -> float:
    t = title.lower()

    base_level = 1.0
    if any(
        k in t
        for k in [
            "chief ",
            " cto",
            " cfo",
            " ceo",
            " cmo",
            " cio",
            "co-founder",
            "founder",
        ]
    ):
        base_level = 6.0
    elif any(k in t for k in ["vice president", " vp", "evp", "svp"]):
        base_level = 5.0
    elif "director" in t:
        base_level = 4.0
    elif "manager" in t:
        base_level = 3.0
    elif any(k in t for k in ["head", "lead", "principal"]):
        base_level = 2.0

    if "assistant" in t:
        base_level -= 1.0
    if "senior" in t:
        base_level += 0.5

    return max(base_level, 1.0)


def email_domain(email: str) -> str:
    email = email.lower().strip()
    if "@" not in email:
        return ""
    return email.rsplit("@", 1)[1]


def base_domain(domain: str) -> str:
    domain = domain.lower().strip()
    domain = re.sub(r"^www\.", "", domain)
    parts = [p for p in domain.split(".") if p]
    if len(parts) <= 2:
        return domain
    # Handles common ccTLD shape like company.co.uk
    if len(parts[-1]) == 2 and parts[-2] in {"co", "com", "org", "net"}:
        return ".".join(parts[-3:])
    return ".".join(parts[-2:])


def prooflink_matches_email_domain(prooflink: str, email: str) -> bool:
    if not prooflink or not email:
        return False
    parsed = urlparse(prooflink)
    host = parsed.netloc.lower().replace("www.", "")
    if not host:
        return False
    domain = email_domain(email)
    if not domain:
        return False
    return base_domain(host) == base_domain(domain)


def parse_employee_floor(employees: str) -> Optional[int]:
    text = employees.replace(",", "")
    m = re.search(r"(\d+)\s*-", text)
    if m:
        return int(m.group(1))
    m = re.search(r"(\d+)\s*\+", text)
    if m:
        return int(m.group(1))
    m = re.search(r"(\d+)", text)
    if m:
        return int(m.group(1))
    return None


def parse_required_company_min(size_text: str) -> Optional[int]:
    s = size_text.lower().replace(",", "").strip()
    if not s or "any" in s:
        return None
    m = re.search(r"from\s*(\d+)", s)
    if m:
        return int(m.group(1))
    m = re.search(r"(\d+)\s*\+", s)
    if m:
        return int(m.group(1))
    m = re.search(r"(\d+)\s*-", s)
    if m:
        return int(m.group(1))
    m = re.search(r"(\d+)", s)
    if m:
        return int(m.group(1))
    return None


def extract_industry_requirements(
    req_map: Dict[str, str], req_raw: str
) -> Optional[List[str]]:
    raw = req_map.get("industry", "")
    raw_l = raw.lower().strip()
    if not raw or "any" in raw_l:
        return None

    if "see comment" not in raw_l:
        items = [x.strip() for x in re.split(r",|;", raw) if x.strip()]
        return items or None

    plain = clean_text(req_raw)
    matches = re.findall(r"industry\s*:\s*([^|]+)", plain, flags=re.IGNORECASE)
    out: List[str] = []
    for m in matches:
        for token in re.split(r",|;", m):
            token = token.strip(" .")
            if (
                token
                and "all subindustries" not in token.lower()
                and "see comment" not in token.lower()
            ):
                out.append(token)
    return out or None


def first_matching_keyword(title: str, keywords: List[str]) -> Optional[str]:
    title_tokens = set(to_slug_tokens(title))
    if not title_tokens:
        return None
    for kw in keywords:
        kt = to_slug_tokens(kw)
        if not kt:
            continue
        if all(token in title_tokens for token in kt):
            return kw
    return None


def check_title_pl_summary(row: Mapping[str, object]) -> Tuple[str, str]:
    title = norm_text(row.get("title"))
    req_raw = norm_text(row.get("req"))
    req_map = parse_req_map(req_raw)

    if not title:
        return "INVALID", "Missing title"

    keywords = extract_keywords(req_map, req_raw)
    min_level = infer_min_level(req_map.get("job_level", ""), req_raw)
    title_level = infer_title_level(title)
    matched_keyword = first_matching_keyword(title, keywords) if keywords else None

    issues: List[str] = []
    if keywords and not matched_keyword:
        issues.append("Title does not contain required keywords")

    if min_level and title_level < min_level:
        issues.append("Title level does not meet requirement")

    if issues:
        return "INVALID", "; ".join(issues[:2])

    if matched_keyword:
        return "VALID", f"Title contains keywords: {matched_keyword}"
    if keywords:
        return "VALID", "Title contains required keyword(s)"
    return "VALID", "No keyword requirements"


def check_prooflink(row: Mapping[str, object]) -> Tuple[str, str]:
    prooflink = norm_text(row.get("prooflink")).lower()
    email = norm_text(row.get("email"))

    if not prooflink:
        return "INVALID", "Missing prooflink"

    if "linkedin.com/in/" in prooflink or "zoominfo.com/p/" in prooflink:
        return "VALID", "Valid prooflink source"

    if prooflink_matches_email_domain(prooflink, email):
        return "VALID", "Prooflink matches corporate domain"

    return "INVALID", "Prooflink does not match allowed sources"


def check_geo(row: Mapping[str, object]) -> Tuple[str, str]:
    req_map = parse_req_map(norm_text(row.get("req")))
    geo_req = req_map.get("geo", "")
    location = clean_text(row.get("location")).lower()

    if not geo_req:
        return "VALID", "No GEO requirements"

    if "any" in geo_req.lower():
        return "VALID", "No GEO requirements"

    if not location:
        return "INVALID", "Missing location data"

    # Extract country-like tokens from requirement text.
    geo_text = geo_req.lower().replace("only", " ")
    raw_tokens = re.split(r",|/|&|\\(|\\)|-", geo_text)
    candidates = []
    for token in raw_tokens:
        token = token.strip()
        if len(token) < 3:
            continue
        if token in {
            "geo",
            "focus on companies in europe with a hq in emea",
            "emea",
            "apac",
            "n/a",
        }:
            continue
        if "city" in token or "region" in token:
            continue
        candidates.append(token)

    if not candidates:
        return "VALID", "GEO requirements could not be parsed"

    for c in candidates:
        if location_matches_candidate(location, c):
            return "VALID", "Location matches GEO requirements"

    return "INVALID", "Location does not match GEO requirements"


def check_nwc(row: Mapping[str, object]) -> Tuple[str, str]:
    status = norm_text(row.get("status")).lower()
    if not status or status == "valid":
        return "VALID", "NWC status clear"
    if status == "a":
        return "INVALID", "Retrieved lead"
    if status == "!":
        return "INVALID", "Suspicious lead"
    if status in {"r", "no info", "no company match"}:
        return "RECHECK", "Status requires manual recheck"
    return "INVALID", f"Unknown status value: {status}"


def validate_required_profile_fields(row: Mapping[str, object]) -> Optional[str]:
    mandatory = ["first_name", "last_name", "company", "title", "email", "industry"]
    missing = [c for c in mandatory if not norm_text(row.get(c))]
    if missing:
        return "Missing profile data"
    return None


def validate_email_field(email: str) -> Optional[str]:
    if not EMAIL_REGEX.match(email):
        return "Invalid email format"
    domain = email_domain(email)
    if domain in PUBLIC_EMAIL_DOMAINS:
        return "Email should be corporate"
    return None


def validate_industry_field(
    row_industry: str, req_map: Dict[str, str], req_raw: str
) -> Optional[str]:
    industries = extract_industry_requirements(req_map, req_raw)
    req_tokens = normalize_tokens(clean_text(req_raw))

    if not row_industry:
        return None

    row_tokens = normalize_tokens(row_industry)
    if not row_tokens:
        return None

    if industries:
        for i in industries:
            if normalize_tokens(i) & row_tokens:
                return None
        return "Industry: Industry does not match requirements"

    raw_ind = req_map.get("industry", "").lower()
    if "see comment" in raw_ind and not (row_tokens & req_tokens):
        return "Industry: Industry does not match requirements"
    return None


def validate_company_size_field(
    employees: str, req_map: Dict[str, str]
) -> Optional[str]:
    required_min = parse_required_company_min(req_map.get("company_size", ""))
    row_floor = parse_employee_floor(employees)
    if required_min is not None and row_floor is not None and row_floor < required_min:
        return "Company size below requirement"
    return None


def check_other(row: Mapping[str, object]) -> Tuple[str, str]:
    req_raw = norm_text(row.get("req"))
    req_map = parse_req_map(req_raw)
    issues: List[str] = []

    # Keep status shortcut for explicit retired/suspicious markers.
    status = norm_text(row.get("status")).lower()
    if status == "a":
        return "INVALID", "Retrieved lead"

    email = norm_text(row.get("email")).lower()
    row_industry = norm_text(row.get("industry")).lower()

    validators: List[Callable[[Mapping[str, object]], Optional[str]]] = [
        lambda r: validate_required_profile_fields(r),
        lambda r: validate_email_field(email),
        lambda r: validate_industry_field(row_industry, req_map, req_raw),
        lambda r: validate_company_size_field(norm_text(r.get("employees")), req_map),
    ]

    for validator in validators:
        issue = validator(row)
        if issue:
            issues.append(issue)

    # Apply title/keyword logic for mixed "Other" checks when company/industry checks passed.
    title_result, title_comment = check_title_pl_summary(row)
    if title_result == "INVALID":
        issues.append(title_comment)

    if issues:
        return "INVALID", "; ".join(issues[:2])

    if title_comment:
        return "VALID", title_comment
    return "VALID", "Lead data matches requirements"


def validate_row(row: Mapping[str, object]) -> Tuple[str, str]:
    sub_status = norm_text(row.get("sub status")).lower()

    if sub_status == "n/a: title/pl summary":
        return check_title_pl_summary(row)
    if sub_status == "n/a: prooflink":
        return check_prooflink(row)
    if sub_status == "n1: nwc":
        return check_nwc(row)
    if sub_status == "n/a: country/geo":
        return check_geo(row)

    if sub_status in {
        "n/a: other",
        "n/a: other (auto)",
        "n/a: other (company)",
        "n2: out of business/bad data",
        "n2: out of business/bad data (company)",
    }:
        return check_other(row)

    # Fallback for unexpected categories.
    return check_other(row)


def compact_comment(comment: str) -> str:
    text = norm_text(comment)
    if not text:
        return text

    parts = [p.strip() for p in text.split(";") if p.strip()]
    compact_parts = [COMMENT_NORMALIZATION.get(p, p) for p in parts]
    compact = "; ".join(compact_parts)
    return re.sub(r"\s+", " ", compact).strip()


def process_file(input_path: str, output_base: str) -> None:
    df = pd.read_excel(input_path)

    results: List[str] = []
    comments: List[str] = []
    records = df.to_dict(orient="records")
    for row in records:
        result, comment = validate_row(row)
        results.append(result)
        comments.append(compact_comment(comment))

    xlsx_path = f"{output_base}.xlsx"
    csv_path = f"{output_base}.csv"

    # Keep the original worksheet layout and write outputs next to source columns.
    wb = load_workbook(input_path)
    ws = wb[wb.sheetnames[0]]

    result_col_letter = get_column_letter(RESULT_COL_IDX)
    comment_col_letter = get_column_letter(COMMENT_COL_IDX)
    ws.column_dimensions[result_col_letter].width = 12
    ws.column_dimensions[comment_col_letter].width = 48

    ws.cell(row=1, column=RESULT_COL_IDX, value="Result")
    ws.cell(row=1, column=COMMENT_COL_IDX, value="Comment")
    ws.cell(row=1, column=RESULT_COL_IDX).alignment = Alignment(
        horizontal="center", vertical="center", wrap_text=True
    )
    ws.cell(row=1, column=COMMENT_COL_IDX).alignment = Alignment(
        horizontal="left", vertical="center", wrap_text=True
    )

    for i, (result, comment) in enumerate(zip(results, comments), start=2):
        ws.cell(row=i, column=RESULT_COL_IDX, value=result)
        ws.cell(row=i, column=COMMENT_COL_IDX, value=comment)
        ws.cell(row=i, column=RESULT_COL_IDX).alignment = Alignment(
            horizontal="center", vertical="center", wrap_text=True
        )
        ws.cell(row=i, column=COMMENT_COL_IDX).alignment = Alignment(
            horizontal="left", vertical="top", wrap_text=True
        )

    wb.save(xlsx_path)

    # CSV keeps a compact tabular version with explicit output fields.
    df["Result"] = results
    df["Comment"] = comments
    df.to_csv(csv_path, index=False)

    summary = df["Result"].value_counts(dropna=False).to_dict()
    print("Saved:", xlsx_path)
    print("Saved:", csv_path)
    print("Result distribution:", summary)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Validate lead rows by sub-status rules."
    )
    parser.add_argument(
        "--input",
        default="DataCheck_DemoCode.xlsx",
        help="Input XLSX path (default: DataCheck_DemoCode.xlsx)",
    )
    parser.add_argument(
        "--output-base",
        default="DataCheck_DemoCode_processed",
        help="Output file base name without extension",
    )
    return parser


def main() -> None:
    args = build_parser().parse_args()
    process_file(args.input, args.output_base)


if __name__ == "__main__":
    main()
