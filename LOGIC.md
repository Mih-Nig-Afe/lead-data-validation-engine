# Validation Logic

## Objective

`lead_data_validator.py` validates each source row and appends two explicit outputs:
- `Result`
- `Comment`

Allowed result values:
- `VALID`
- `INVALID`
- `RECHECK`

The engine also exports both workbook and CSV outputs:
- `Lead_Data_Validation_Results.xlsx`
- `Lead_Data_Validation_Results.csv`

## Execution Model

Run the validator:

```bash
python3 lead_data_validator.py
```

Run the QA checks against the generated workbook:

```bash
python3 qa_check.py
```

Both scripts also accept explicit path overrides through CLI flags.

## Core Helpers

### `norm`
- normalizes null-like values into safe strings
- prevents repetitive null handling across rule functions

### `parse_req`
- converts the pipe-delimited `req` text into a dictionary
- provides structured access to fields like `keywords` and `job_level`

### `normalize_url_for_hyperlink`
- preserves valid absolute URLs
- adds `https://` when a value looks like a valid bare domain/path
- keeps proof links clickable in output workbooks

### requirement placeholder detection
- treats values like `-`, `any`, and `see comment` as non-deterministic
- avoids pretending that every free-form requirement can be enforced safely

## Rule Families

### 1. NWC

`N1: NWC` rows map status codes to final decisions:
- empty or `valid` -> `VALID`
- `a` -> `INVALID`
- `!` -> `INVALID`
- `r`, `no info`, `no company match` -> `RECHECK`
- unknown codes -> `INVALID`

Why:
- these statuses already represent a compact business decision table

### 2. Prooflink

Accepted evidence:
- `linkedin.com/in/...`
- `zoominfo.com/p/...`
- company-hosted URLs whose base domain matches the corporate email domain

Why:
- proof must come either from a trusted external profile source or the company itself

### 3. Title

Title validation performs two checks when the requirement is deterministic:
- keyword fit
- seniority threshold

Important implementation details:
- title ranking is phrase-based, not single-letter substring matching
- placeholder requirement values do not trigger hard rejections
- mixed level strings such as `ANY level ... Manager+` are treated as ambiguous and skipped
- positive matches return explainable comments, not just a status code

### 4. Country/GEO

Current GEO behavior is intentionally conservative:
- explicit location-to-country matches are validated automatically
- ambiguous location text returns `RECHECK`
- the large-scale region/county strategy is documented separately in `COUNTRY_CHECK_SCALING_NOTES.md`

Why:
- the task asks for design notes around high-volume GEO processing, and false-positive country matching is worse than an explicit review handoff

### 5. Out of Business / Bad Data

Rows marked as `Out of Business/Bad data` are immediately set to `INVALID`.

Why:
- this status already encodes a business decision and should not fall through generic data-hygiene checks

### 6. Other

Baseline data hygiene checks:
- required identity fields exist
- email format is valid
- public mailbox domains are rejected

Why:
- this is the minimum quality bar for a lead record even when richer evidence is unavailable

## Routing

Validation is dispatched by `sub status`:
- contains `nwc` -> NWC logic
- contains `country` or `geo` -> GEO logic
- contains `prooflink` -> prooflink logic
- contains `title` -> title logic
- contains `out of business` or `bad data` -> explicit invalidation
- contains `other` -> baseline data checks
- unknown values -> baseline data checks as a safe fallback

## Output Formatting

After validation:
1. data is written to XLSX and CSV
2. the workbook is reopened with `openpyxl`
3. prooflink columns are converted to hyperlinks when possible
4. `Result`, `Comment`, and `req` columns are wrapped for readability
5. column widths are capped to avoid unusable sheet layouts

Design choice:
- formatting is header-driven instead of tied to hard-coded column letters

## Reliability Signals

- unit tests cover core rule behavior and output generation
- `qa_check.py` verifies output structure and hyperlink presence
- CI runs the tests on push and pull request
