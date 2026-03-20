# Logic Flow

## 1) Objective
This document explains the complete logic implemented in lead_data_validator.py, from low-level helper choices to full system behavior.

Primary goals:
- Validate each lead row according to sub status and req definitions.
- Return one final decision per row: VALID, INVALID, or RECHECK.
- Return a clear, human-readable Comment explaining the decision.
- Export both XLSX and CSV outputs.
- Make prooflink fields clickable in Excel for fast QA review.

Outputs:
- Lead_Data_Validation_Results.xlsx
- Lead_Data_Validation_Results.csv

## 2) Why these libraries were selected

### pandas
Why used:
- Reliable reading and writing of tabular datasets (XLSX and CSV).
- Efficient row iteration for rule-based validation.
- Natural way to append Result and Comment columns.

Why it is the right fit:
- Input is spreadsheet-shaped data with mixed field types.
- Validation pipeline is row-centric and pandas handles this cleanly.

### openpyxl
Why used:
- Needed for post-processing Excel formatting after data export.
- Supports column widths, cell alignment, and hyperlinks.

Why it is the right fit:
- pandas writes values well, but workbook styling and clickable links need worksheet-level control.

### re
Why used:
- Email format checks.
- Keyword token extraction for title keyword matching.

Why it is the right fit:
- Lightweight and precise for string patterns required by this validator.

### urllib.parse.urlparse
Why used:
- Safe URL decomposition to extract host/domain from proof links.

Why it is the right fit:
- Domain matching logic is central for official-website prooflink validation.

### typing
Why used:
- Improves readability and maintainability with explicit intent for function inputs/outputs.

Why it is the right fit:
- Reduces ambiguity and improves robustness for future updates.

## 3) Global constants and why they exist

### PUBLIC_EMAIL_DOMAINS
Purpose:
- Blocks free mailbox domains for corporate lead checks.

Reason:
- Corporate validation requires business email quality, not personal inboxes.

### EMAIL_REGEX
Purpose:
- Baseline syntax validation before any domain-level checks.

Reason:
- Prevent false acceptance of malformed addresses.

### WORD_REGEX
Purpose:
- Tokenization for keyword-in-title matching.

Reason:
- Improves precision for multi-word keyword checks.

## 4) Helper functions and design rationale

### norm
What it does:
- Converts null-like values into safe empty strings.

Why needed:
- Prevents crashes and repeated null-handling logic.
- Standardizes all downstream validators.

### parse_req
What it does:
- Converts req text into a dictionary using key:value pairs split by |.

Why needed:
- Makes requirements machine-readable and easy to query by validator.

### email_domain
What it does:
- Extracts domain portion from an email.

Why needed:
- Required for corporate domain checks and prooflink-domain comparison.

### base_domain
What it does:
- Normalizes domains by removing www and comparing root-like suffixes.

Why needed:
- Avoids false mismatches from subdomains.

### normalize_url_for_hyperlink
What it does:
- Accepts full URLs as-is.
- Converts bare domains/paths into https-prefixed links when possible.

Why needed:
- Ensures prooflink text becomes clickable links in Excel even if source lacks scheme.

## 5) Validation blocks and why each exists

### check_prooflink
Checks:
- Missing link -> INVALID.
- linkedin.com/in or zoominfo.com/p -> VALID.
- Official website accepted when prooflink host domain matches corporate email domain -> VALID.
- Otherwise INVALID.

Why:
- Matches requested acceptance policy for proof evidence.

Comment style:
- Specific reason for pass/fail, not generic OK/Bad.

### check_title
Checks:
- Missing title -> INVALID.
- Keyword requirement from req keywords.
- Seniority requirement from req job_level compared against title seniority.

Why:
- Sub status Title/PL Summary requires both content relevance and level fit.

Detail behavior:
- Can return combined reasons, for example keyword miss plus low level.
- Returns positive detail with matched keyword when applicable.

### check_nwc
Checks:
- Empty or valid status -> VALID.
- a -> INVALID (retrieved/retired lead path).
- ! -> INVALID (suspicious).
- r, no info, no company match -> RECHECK.
- Other unknown status -> INVALID with explicit status value.

Why:
- Implements required N1 NWC decision matrix exactly.

### check_other
Checks:
- Required identity/company fields.
- Email syntax validity.
- Corporate mailbox requirement.

Why:
- Sub status Other should guarantee baseline lead quality and completeness.

Comment style:
- Lists missing fields explicitly.
- Distinguishes malformed email from public mailbox usage.

## 6) Decision router and system control flow

Function:
- validate

Routing:
- If sub status contains nwc -> check_nwc.
- If contains prooflink -> check_prooflink.
- If contains title -> check_title.
- If contains other -> check_other.
- Fallback -> check_other.

Why this order:
- Most specific sub-status logic first.
- Ensures one deterministic path per row.
- Fallback protects system against unknown labels.

## 7) Main pipeline behavior

Function:
- main

Pipeline:
1. Read DataCheck_DemoCode.xlsx.
2. Validate each row and collect Result + Comment.
3. Append Result and Comment columns.
4. Save both XLSX and CSV outputs.
5. Re-open XLSX with openpyxl for formatting and hyperlink pass.
6. Save finalized workbook.

Why this two-step write process:
- pandas handles data export efficiently.
- openpyxl handles presentation features pandas does not fully manage.

## 8) Why comments are detailed

Design target:
- Reviewer should know exactly why a row passed or failed without reading code.

Benefits:
- Faster QA and analyst review cycles.
- Better traceability for disputes and rechecks.
- Easier model tuning when rules evolve.

Comment design principles used:
- Actionable language.
- Single-source reason or compact multi-reason list.
- Consistent phrasing per rule block.

## 9) Why links are visible and clickable

What is implemented:
- Hyperlink conversion for both prooflink and employees_prooflink columns.
- Non-empty values are converted to clickable links.
- URL normalization includes https fallback.

Why this matters:
- QA reviewers can open sources instantly.
- Reduces copy/paste friction and manual errors.
- Speeds up validation at scale.

## 10) Why column widths and wrapping are set this way

Current strategy:
- Width is calculated from content length and capped at 45.
- Wrap text on Result, Comment, and req columns.

Why this is used:
- Prevents very long values from exploding worksheet width.
- Keeps detailed comments readable in-cell.
- Preserves compact horizontal layout for reviewers.

Vertical alignment choices:
- Wrapped columns use top alignment for readability.
- Other columns use center alignment for cleaner scanning.

## 11) Robustness features

- Null-safe text normalization everywhere via norm.
- Defensive URL normalization for malformed proof links.
- Deterministic row routing with fallback.
- Explicit handling for unknown NWC status values.
- Field-level comments to reduce silent failures.

## 12) Optimization choices

Current optimizations:
- Single read of source workbook with pandas.
- Simple list accumulation for Result/Comment then single assignment.
- One formatting pass in openpyxl after write.

Why this is efficient enough:
- Complexity is linear by row count.
- Suitable for medium-to-large sheets while keeping code maintainable.

Future scale upgrades if needed for 50k to 70k rows:
- Cache parsed req values by unique req text to avoid repeated parsing.
- Batch title keyword checks by grouped patterns.
- Reduce per-cell styling to targeted columns only.
- Optional chunked processing for memory control.

## 13) Clean-up policy applied in this logic file

This document intentionally excludes:
- Redundant prose that repeats function names without adding intent.
- Generic statements with no operational impact.
- Historical alternatives not used in current implementation.

This document intentionally includes:
- Why each function exists.
- Why each library is used.
- Why visual formatting and link behavior are required.
- Why each decision path returns its specific Result and Comment.

## 14) End-to-end summary

The system is a deterministic row validator with explicit rule routing by sub status, requirement parsing from req, detailed decision comments, and reviewer-oriented XLSX output enhancements. It is built to be readable, auditable, and practical for real QA workflows where explanation quality and link accessibility matter as much as raw pass/fail outcomes.
