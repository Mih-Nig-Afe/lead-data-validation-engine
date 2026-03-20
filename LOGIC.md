# Lead Data Validation Logic

## Deliverables Produced

- `DataCheck_DemoCode_processed.xlsx`: processed test sheet with `Result` and `Comment` columns.
- `DataCheck_DemoCode_processed.csv`: CSV version of the same processed results.
- `process_leads.py`: code used to process and validate the test table.
- This document (`LOGIC.md`): logic explanation and scaling notes.

## What the script does

- Reads the test sheet from `DataCheck_DemoCode.xlsx`.
- Routes each row to validation logic based on `sub status`.
- Compares row data against parsed rules from `req`.
- Uses dictionary iteration (`to_dict(orient="records")`) for better performance on large datasets.
- Uses regex-based normalization and matching to reduce false positives.
- Writes `Result` and `Comment` columns.
- Exports both XLSX and CSV outputs.

## Rule routing by sub status

- `N/A: Title/PL Summary`
  - Validates title against requirement keywords.
  - Validates title seniority against required job level (`Director+`, `Manager+`, etc.).
- `N/A: Prooflink`
  - Accepts `linkedin.com/in/` and `zoominfo.com/p/` links.
  - Also accepts official website links when website domain matches corporate email domain.
- `N1: NWC`
  - `status` empty or `valid` => `VALID`
  - `status = a` => `INVALID` (`Retrieved lead`)
  - `status = !` => `INVALID` (`Suspicious lead`)
  - `status in {r, no info, no company match}` => `RECHECK`
- `N/A: Country/GEO`
  - Parses `Geo:` requirement and checks word-boundary country matches to avoid substring errors.
- `N/A: Other`, `N/A: Other (auto)`, `N/A: Other (company)`, `N2: ...`
  - Checks required profile fields.
  - Checks email format and corporate domain usage.
  - Checks industry match where explicit industry requirements exist.
  - Checks minimum company size when requirement includes threshold (example: `500+`, `from 500`).

## Output semantics

- `VALID`: row passed active checks.
- `INVALID`: one or more rule violations found.
- `RECHECK`: status indicates manual verification is needed.

The system is designed to be easily extendable by adding new validators without modifying core validation flow.

## Notes for high-volume GEO checks (50k-70k rows)

If `location` contains only region/county text (without clean country values) and volume is large:

1. Normalize location text once

- Lowercase, remove punctuation, normalize spaces.
- Precompute normalized location column once and reuse in vectorized matching.

1. Pre-parse GEO requirements once per unique `req`

- Build a lookup dictionary: `req_string -> parsed_allowed_countries`.
- Reuse this parsed result for all rows sharing the same requirement text.

1. Use vectorized operations instead of row-by-row regex

- Use `pandas.Series.str.contains` with precompiled OR patterns per requirement group.
- Group rows by identical GEO requirement, then evaluate each group in batch.

1. Add country and county synonym map

- Map variants like `uae -> united arab emirates`, `uk -> united kingdom`.
- Map county/region aliases to country (example: `Bavaria -> Germany`, `Scotland -> United Kingdom`) before matching.
- Keep this map in memory for fast replacement.

1. Optional scaling upgrade

- For very large recurring datasets, process in chunks and/or move GEO normalization into a lightweight SQL/Polars pipeline for faster throughput.

## Run command

```bash
/usr/bin/python3 process_leads.py --input DataCheck_DemoCode.xlsx --output-base DataCheck_DemoCode_processed
```
