# Country Check Scaling Notes (50k-70k Rows)

## Purpose
This document explains how to process country requirements when the location column contains only region/county text, at high volume.

## Problem Context
- Requirement asks for country-level validation.
- Source location may contain county, region, city, abbreviations, or inconsistent spelling.
- Dataset size is large (50k-70k rows), so naive row-by-row regex checks can become slow and noisy.

## Recommended Processing Strategy

### 1) Normalize Once, Reuse Everywhere
Create normalized helper fields at load time:
- lowercased text
- punctuation removed
- spaces normalized
- common separators unified

Why:
- Reduces repeated text cleanup work.
- Improves matching consistency.

### 2) Parse GEO Requirements Once per Unique req
- Group by unique req text.
- Parse allowed countries once.
- Cache in a map: req_text -> parsed_geo_rules.

Why:
- Many rows share the same req value.
- Avoids redundant parsing and speeds up processing.

### 3) Region/County to Country Mapping Layer
Maintain lookup dictionaries:
- county -> country
- region -> country
- city aliases -> country (if needed)
- abbreviation and typo variants

Examples:
- bavaria -> germany
- scotland -> united kingdom
- ny, new york state -> united states

Why:
- Converts indirect location references into country-level signals.

### 4) Multi-Stage Match Pipeline
Use deterministic match order:
1. Exact country name match in location.
2. Alias/abbreviation country match.
3. Region/county inferred country match.
4. No confident match -> RECHECK or INVALID (policy-based).

Why:
- Maximizes precision first.
- Keeps uncertain cases transparent.

### 5) Confidence-Based Decisioning
Assign confidence levels to matches:
- High confidence: explicit country mention.
- Medium confidence: trusted region/county mapping.
- Low confidence: ambiguous token overlap.

Policy suggestion:
- High confidence + allowed country -> VALID.
- Medium confidence + allowed country -> VALID or RECHECK (depending on strictness).
- Low confidence -> RECHECK.
- Confident mismatch -> INVALID.

### 6) Vectorized Evaluation Instead of Per-Row Heavy Logic
- Use grouped/vectorized string operations where possible.
- Apply checks by req group, not row-by-row Python loops.

Why:
- Better throughput on 50k-70k rows.
- Lower CPU overhead.

### 7) Keep Explainable Comments
Write comments that explain exactly what happened:
- Location inferred as Germany from region Bavaria
- Location ambiguous: county token matched multiple countries
- Location does not meet GEO requirement

Why:
- Improves QA and auditability.
- Helps manual recheck teams work faster.

### 8) Performance and Reliability Controls
- Cache parsed requirements and mapping lookups.
- Keep reference dictionaries in memory.
- Avoid per-cell Excel styling passes during core validation.
- Apply workbook formatting only once, after data decisions are complete.

## Suggested Output Semantics for GEO
- VALID: location confidently satisfies GEO requirement.
- INVALID: location confidently contradicts GEO requirement.
- RECHECK: location cannot be confidently resolved.

## Data Quality Enhancements (Optional)
- Maintain versioned geo alias dictionaries.
- Track unmatched tokens to continuously improve mappings.
- Add unit tests for known edge-case regions and ambiguous names.

## Practical Summary
For large sheets with county/region-only locations, the best approach is:
- normalize once,
- parse requirements once,
- map region/county to country through curated dictionaries,
- run vectorized grouped checks,
- and preserve explainable comments with confidence-based outcomes.
