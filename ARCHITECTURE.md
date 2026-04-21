# Architecture

## 1. Current System Boundary

The current implementation is a batch validation engine for CRM lead exports. It is not pretending to be a full CRM. That is deliberate.

Current responsibilities:
- ingest a workbook export
- route each row to the correct rule set
- produce a deterministic decision and explanation
- export analyst-friendly XLSX and CSV outputs

Current non-goals:
- user authentication
- multi-tenant persistence
- distributed job scheduling
- interactive review UI

This keeps the present code honest while still showing how the domain should evolve in a real backend system.

## 2. Logical Components

### Workbook Adapter
- Reads input rows from Excel.
- Converts the batch into a tabular structure for rule execution.

### Validation Router
- Uses `sub status` to decide which rule family applies.
- Keeps one deterministic path per row.

### Rule Modules
- `NWC` rules evaluate manual status codes into `VALID`, `INVALID`, or `RECHECK`.
- `GEO` rules validate only explicit country-resolution cases and defer ambiguous locations to review.
- `Prooflink` rules validate trusted profile sources and company-domain evidence.
- `Title` rules check keyword fit and seniority requirements when the requirement text is machine-readable.
- `Bad Data` rules immediately invalidate rows already marked as out of business or bad data.
- `Other` rules enforce baseline data hygiene such as required fields and corporate email checks.

### Decision Layer
- Produces the final result and an audit-friendly comment.
- Comments are treated as product output, not debug strings.

### Output Formatter
- Writes CSV for downstream systems.
- Formats Excel for human QA with clickable proof links and readable columns.

## 3. Domain Entities

### Present Batch Model

| Entity | Key fields | Notes |
| --- | --- | --- |
| `Lead` | name, company, title, email, location | One spreadsheet row |
| `RequirementProfile` | req text, parsed keys | Current source of business constraints |
| `Evidence` | prooflink, employees prooflink, status | Inputs used by validators |
| `ValidationDecision` | result, comment | Final decision payload |
| `ValidationRun` | input file, timestamps, output files | Batch execution metadata |

### CRM-Oriented Target Model

For a production CRM workflow I would model:

| Entity | Purpose |
| --- | --- |
| `User` | Analyst, reviewer, or admin actor |
| `Workspace` | Customer or internal business unit boundary |
| `Lead` | Canonical person/company identity |
| `LeadEvidence` | URLs, status codes, enrichment artifacts |
| `RequirementTemplate` | Reusable targeting rules per campaign or workflow |
| `ValidationJob` | Async batch request and lifecycle state |
| `ValidationDecision` | Versioned decisions over time |
| `ManualReviewTask` | Operational queue for `RECHECK` outcomes |
| `RuleVersion` | Auditability for decision reproducibility |

Relationship expectations:
- one `Workspace` has many `Users`, `Leads`, and `ValidationJobs`
- one `ValidationJob` evaluates many `Leads`
- one `Lead` has many `LeadEvidence` records and many `ValidationDecision` records over time
- one `RequirementTemplate` can be reused across many jobs
- one `ManualReviewTask` belongs to one decision but may be reassigned across users

## 4. Decision Flow

```text
ingest workbook
-> normalize row values
-> parse requirement string
-> route by sub status
-> run domain validator
-> emit result + comment
-> write XLSX/CSV
-> post-process workbook formatting
```

Important architectural choice:
- formatting happens after decisioning, so validation logic stays independent from presentation concerns.

## 5. Why These Choices

### Why batch-first instead of database-first
- The source system is clearly spreadsheet-driven.
- A batch adapter lets the engine solve the actual workflow quickly.
- It avoids fake infrastructure that adds noise without improving the core validation behavior.

### Why deterministic comments
- In ops-heavy systems, explainability matters more than model cleverness.
- Reviewers need to know exactly why a record failed.
- This also makes disputes and rule tuning materially easier.

### Why conservative handling of ambiguous requirements
- Free-form requirement text often includes instructions such as `see comment`.
- Treating those as hard machine-readable rules creates false confidence.
- The engine now skips non-deterministic title checks instead of inventing precision.

## 6. Scaling To 10k+ Users And Larger Batches

There are two different scaling dimensions here:
- end-user scale: many analysts or customers submitting validation jobs
- data scale: large files or many concurrent jobs

### Near-Term Scale

For tens of thousands of leads per job:
- cache parsed `req` strings by unique value
- vectorize low-complexity checks where possible
- isolate workbook styling into a single final pass
- keep outputs append-only and immutable per run

### Service Scale

For 10k+ users or many concurrent workspaces:
- move ingestion to an API plus object storage
- create async `ValidationJob` workers behind a queue
- persist decisions and artifacts in PostgreSQL
- store raw uploads and generated outputs in blob storage
- add per-workspace rate limits, quotas, and job status tracking

Suggested production topology:

```text
API -> job queue -> validation workers -> relational DB + blob storage
                                 -> manual review queue -> reviewer UI
```

### Data Model Scale

To support CRM evolution:
- separate canonical lead identity from one-off import rows
- version requirement templates
- version rule sets
- track reviewer overrides and reason codes
- emit metrics on invalidation causes to improve campaign targeting upstream

## 7. Failure Modes And Controls

Primary risks:
- ambiguous requirement text
- ambiguous geography in free-form location fields
- malformed or scheme-less proof links
- workbook layout drift
- false positives from naive title parsing

Mitigations in this repo:
- placeholder requirement detection
- conservative GEO handling that only auto-validates explicit single-country matches
- URL normalization for hyperlink and domain validation paths
- header-driven workbook formatting instead of fixed column letters
- unit tests around title and prooflink logic
- separate QA script for generated output validation

## 8. What I Would Build Next

If this were the foundation of a real CRM validation subsystem, the next implementation steps would be:
- convert rule families into explicit strategy classes
- persist `ValidationRun`, `ValidationDecision`, and `ManualReviewTask` tables
- expose a REST or async job API
- add rule-versioning and replayability
- add observability around throughput, invalidation reasons, and recheck rates
