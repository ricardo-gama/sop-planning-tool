# Track 2 — Publish Pipeline: schema & decisions (design only, no code yet)

Status: design doc per NEXT_STEPS.md Part B. No Office Script, PA flow, or Planning
Tool code has been written — this resolves the open questions from B5 and finalises
the B1 schema so a future session can build against a settled contract.

---

## Decisions resolved

### B5.1 — LeftSync deals in publish
**Decided: include until removed.** A deal that leaves the CEP sync feed (`CepStatus
= "LeftSync"`) still publishes to the Planning Tool. The planner is the one who
decides it's done — via the existing "Remove from pipeline" action in the Demand
Tool (which now also unflags `Promote` in OppImport, so it won't resurface via
SyncPipeline). No separate exclusion logic needed at publish time beyond the
existing `DealType ≠ "Lost"` filter.

### B5.2 — Where PipelinePublished lives
**Deferred.** Not decided yet — revisit before B2/B4 are built. Doesn't block
finalising the schema below, since the column list is the same regardless of which
workbook hosts the table.

### B5.4 — How the Planning Tool detects a new publish
**Revised from the original doc.** The original assumption was a single
`PublishedDate` the Planning Tool polls on load. That doesn't hold up once publish
runs multiple times a day (nightly baseline + manual intra-day publishes, per B4
Option C): a single timestamp only tells the Planning Tool *that something changed
run-to-run*, not *which rows*.

**New approach: per-row change tracking**, computed by the publish script at each
run by diffing against the previous `PipelinePublished` snapshot (keyed on `OppId`):

- `RowStatus`: `"New"` | `"Updated"` | `"Unchanged"` | `"Removed"` — computed by
  comparing this run's row to the same `OppId` in the table before it's replaced.
  "Updated" fires when any of `DealType`, `EffectiveOiDate`, `SelectedOfferID`,
  `Risk`, `Value` differ from the prior snapshot.
- `FirstPublishedDate`: set once, when the `OppId` first appears in
  `PipelinePublished`; carried forward unchanged on every later run.
- `LastChangedDate`: updated only on runs where `RowStatus` is `"New"` or
  `"Updated"`; otherwise carried forward from the prior snapshot. This is the field
  the Planning Tool actually watches — it can filter `LastChangedDate > <last time
  I looked>` to find exactly the rows that need review, independent of how many
  publish runs happened in between.
- `PublishedDate`: still written every run (this run's timestamp), kept for audit /
  "how fresh is this table" purposes — but the Planning Tool should *not* rely on it
  to detect row-level change.

**Removed rows:** because the publish is a wholesale replace, a deal that no longer
qualifies (`DealType = "Lost"`, or the DemandWork row was deleted) simply disappears
from the table — the Planning Tool would never see a `"Removed"` status. To make
removal visible instead of silent, the publish script keeps a removed row for
exactly one additional run with `RowStatus = "Removed"` before dropping it for good
(i.e. diff the previous snapshot's `OppId` set against this run's; anything missing
gets re-inserted with `RowStatus = "Removed"` and today's `LastChangedDate`, then is
actually omitted on the run after that).

This still needs a real trial run against production-shaped data before being
treated as final — flagged here as the next thing to validate once B2 is built.

---

## B1 — PipelinePublished schema (finalised)

| # | Column | Type | Source | Notes |
|---|--------|------|--------|-------|
| 0 | OppId | Text | DemandWork.OppId | CEP# join key |
| 1 | OpportunityName | Text | DW snapshot | |
| 2 | Customer | Text | DW snapshot | |
| 3 | Region | Text | DW snapshot | |
| 4 | SubRegion | Text | DW snapshot | |
| 5 | MarketPortfolio | Text | DW snapshot | |
| 6 | Value | Number | DW snapshot | Raw € value |
| 7 | ForecastStatus | Text | DW snapshot | |
| 8 | DealType | Text | DemandWork | The planner's deal classification |
| 9 | SelectedOfferID | Text | DemandWork | Committed offer |
| 10 | EffectiveOiDate | Date | Computed | AdjustedOiDate if set, else CrmOiDate |
| 11 | Risk | Boolean | DemandWork | |
| 12 | CepStatus | Text | DemandWork | "Active" or "LeftSync" |
| 13 | ScoreBand | Text | DW snapshot | |
| 14 | LastModifiedBy | Text | DemandWork | Planner who last saved |
| 15 | LastModifiedDate | Date | DemandWork | |
| 16 | RowStatus | Text | Computed (publish) | "New" \| "Updated" \| "Unchanged" \| "Removed" |
| 17 | FirstPublishedDate | Date | Computed (publish) | Set once, carried forward |
| 18 | LastChangedDate | Date | Computed (publish) | Updated only on New/Updated runs — this is what the Planning Tool polls |
| 19 | PublishedBy | Text | Script | User/system that triggered this run |
| 20 | PublishedDate | Date | Script | Timestamp of this run — audit only, not for change detection |

Changes from the original NEXT_STEPS.md draft:
- Reordered `ForecastStatus` next to the other DW-snapshot columns (mirrors the A4
  Pipeline-list column change already shipped).
- Added `RowStatus`, `FirstPublishedDate`, `LastChangedDate` per the B5.4 revision
  above — these replace reliance on a single `PublishedDate` poll.

---

## Still open (unchanged from NEXT_STEPS.md, not addressed here)

- **B5.2** — Demand Tool workbook vs Planning Tool workbook vs PA-mediated. Deferred.
- **B5.3** — Publish history: last-timestamp only vs full PublishLog table. Not
  discussed this round; original doc's "start with A" stands unless revisited.
- **B5.5** — `InPipeline` flag's final role. Still parked.
- **B2–B7** — Office Script implementation, PA flow, add-in publish button, Planning
  Tool read/promote logic. None of this is built. Do not assume `RowStatus` /
  `LastChangedDate` are live columns anywhere yet — this document defines the target
  schema for when B2 is actually built, not current state.
