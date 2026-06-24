# S&OP Capacity Planning Tool — Planning Logic Reference

**Version:** v0.2.0  
**Last updated:** June 2026  
**Scope:** Covers both planning scripts (Constrained and Unconstrained), the Sync Demand script, data model, scheduling mechanics, deal alignment feature, and all deliberate design decisions.

---

## 1. Overview

The planning tool produces two independent views of the same demand book:

| Script | Table outputs | Purpose |
|---|---|---|
| **Rebalance Constrained** | `Results`, `Schedule` | Finite-capacity scheduling — respects plant bay limits, carries overflow forward |
| **Rebalance Unconstrained** | `ResultsUnconstrained`, `ScheduleUnconstrained` | Infinite-capacity scheduling — shows demand shape independent of bottlenecks |

Both scripts share the same input tables, routing logic, milestone chain, and deal alignment feature. The only structural difference is whether Assembly and Testing placement respects capacity constraints.

---

## 2. Data Model

### Input tables

| Table | Purpose |
|---|---|
| `Demand` | One row per equipment line. Source of truth for what needs to be planned. |
| `DemandImport` | Raw feed from the source system. Sync Demand merges this into Demand. |
| `Routings` | Per-equipment-family load and duration parameters. |
| `CapacityBase` | Default weekly capacity per Plant × Stage. |
| `CapacityOverride` | Week-specific capacity exceptions (temporary adjustments). |
| `Backlog` | Already-committed load per Plant × Stage × Week from in-progress orders. |
| `Adjustments` | Manual per-line overrides: load values and forced start weeks. |
| `Calendar` | Maps week indices (1–156) to ISO week labels and Excel date serials. |
| `Settings` | Global planning parameters (tolerance, max passes). |

### Output tables

| Table | Columns |
|---|---|
| `Results` / `ResultsUnconstrained` | Plant, Stage, Week, Status, Backlog, Planned, Total, Cap, Util%, Overload |
| `Schedule` / `ScheduleUnconstrained` | OppID, LineID, Equipment, Plant, OI, Status, KOM, iKOM, 1stBOM, FrozenDesign, FinalBOM, AsmStart, AsmFinish, TstStart, TstFinish, PDI, FAT, EOP, FCA, LT |

### Planning horizon

156 weeks (3 years) from the week marked `WeekIndex = 1` in the Calendar table. All week indices are 1-based; all date arithmetic uses Excel serial numbers.

---

## 3. Demand Ordering

Both scripts process demand in the same priority order before any load is placed. This order determines which lines get first access to finite capacity in the constrained script.

**Sort key (ascending):**

1. `Priority` — explicit numeric priority set by the planner (blank = lowest, treated as +∞)
2. `Status rank` — fixed hierarchy:

| Status | Rank |
|---|---|
| FROM OI TO KO | 1 |
| ANTICIPATION - ENGINEERING ONLY + PROCUREMENT | 2 |
| ANTICIPATION - ENGINEERING ONLY | 3 |
| ON HAND | 4 |
| FORECAST | 5 |
| BACKUP | 6 |
| OUTLOOK | 7 |
| OTHER | 8 |

3. `OI date` (earliest first)
4. Original row index (stable tiebreaker)

Lines with `ChangeFlag = "Removed"` are excluded from planning entirely.

---

## 4. Routing Parameters

Each equipment line references a `RoutingKey` which maps to a row in the Routings table. The routing defines all load values and working-day offsets used to derive milestones and durations.

| Parameter | Description |
|---|---|
| `Assembly Load` | Total assembly bay-hours for this equipment |
| `Testing Load` | Total testing bay-hours |
| `ESL Load` | Engineering support load (unconstrained, spread across full lifecycle) |
| `CTO Load` | Engineering configuration load (unconstrained, spread across design window) |
| `FAT Load` | Factory Acceptance Test load (single-week placement) |
| `Assembly Duration (days)` | Working days for assembly stage |
| `Testing Duration (days)` | Working days for testing stage |
| `Procurement (days)` | Working days from Final BOM to material ready |
| `Procurement Buffer (days)` | Additional buffer added to procurement lead time |
| `KOM (days)` | Working days from OI to Kick-Off Meeting |
| `1st BOM (days)` | Working days from OI to first BOM release |
| `Frozen Packing Design (days)` | Working days from OI to frozen design milestone |
| `Final BOM (days)` | Working days from OI to Final BOM release |
| `PDI (days)` | Working days from Testing finish to Pre-Delivery Inspection |
| `FAT (days)` | Working days from Testing finish to FAT |
| `EOP (days)` | Working days from Testing finish to End of Production |
| `FCA (days)` | Working days from EOP to Final Customer Acceptance |

All offsets are working-day counts (Mon–Fri, no calendar exceptions beyond weekends).

---

## 5. Milestone Chain

Every line's milestone dates are derived from a fixed chain anchored at the OI (Order Intake) date. Engineering milestones are always OI-anchored and are never moved by capacity constraints or alignment logic.

```
OI
├── + KOM days              → KOM date
├── + KOM + 5               → iKOM date
├── + KOM + 1stBOM          → 1st BOM date
├── + KOM + FrozenDesign    → Frozen Packing Design date
└── + KOM + FinalBOM        → Final BOM date
                                    │
                                    └── + Procurement + ProcBuffer → earliest Assembly start date
                                                                            │
                                                                    [Assembly: constrained or spread]
                                                                            │
                                                                    [Testing: follows Assembly]
                                                                            │
                                                                    ├── + PDI days  → PDI date
                                                                    ├── + FAT days  → FAT date
                                                                    ├── + EOP days  → EOP date
                                                                    └── EOP + FCA   → FCA date
```

The **lead time (LT)** reported in the Schedule table is the working-week count from OI to the realized FCA date: `workdaysBetween(OI, FCA) / 5`.

---

## 6. Stage Load Model

### Constrained stages (Assembly, Testing)

These are the only stages that compete for finite bay capacity. Load is allocated by `allocateStage()` — see Section 7.

### Unconstrained stages

| Stage | Placement method | Window |
|---|---|---|
| CTO | `spread()` — evenly distributed | Frozen Packing Design → Final BOM |
| ESL | `spread()` — evenly distributed | OI → FCA (full lifecycle) |
| FAT | `place()` — single week | FAT date week |

**Design decision:** CTO and ESL intentionally bypass `allocateStage()` and use the unconstrained `spread()` primitive. This is not a bug.

- **ESL** represents engineering support load across the whole project lifecycle. It is informational — no capacity ceiling applies to it.
- **CTO** is tied to the engineering team's configuration window between Frozen Design and Final BOM. It is not subject to bay capacity.

These decisions were confirmed during design and documented here to prevent them being misread as gaps during reviews.

---

## 7. Constrained Scheduling — `allocateStage()`

The constrained script's core mechanic. Places a line's load respecting the plant's remaining weekly capacity.

**Inputs:** plant, stage, start week, total load, duration in weeks, status  
**Output:** first and last week where load was actually placed

**Algorithm:**

1. Divide total load evenly across `durationWeeks` → `perWeek = totalLoad / durationWeeks`
2. For each week in the planned window:
   - Accumulate `carry += perWeek`
   - Calculate `room = cap(plant, stage, week) − usedAt(plant, stage, week)`
   - Place `put = min(carry, room)`; subtract from carry
3. If carry remains after the planned window, continue week by week beyond it until carry is exhausted or the horizon is reached
4. If carry still remains at horizon end, dump into the last feasible week (ensures load is never silently lost)

`usedAt()` sums both backlog load (fixed, pre-committed) and planned load already placed by earlier-priority lines in this pass.

**Testing follows Assembly:** Testing start defaults to the next working day after Assembly's realized finish date (`addWorkdays(asmScheduleFinishDate, 1)`). This ensures Testing never begins while Assembly is still running.

---

## 8. Unconstrained Scheduling — `spread()`

The unconstrained script places all stage load using `spread()`: load is divided evenly across the window and placed without checking or consuming capacity. No carry-forward occurs. The result represents demand shape, not a feasible production plan.

Assembly and Testing start weeks follow the same logic as the constrained script (natural start from OI + routing offsets, or manual Adjustments override), but `spread()` distributes load across the duration window without respecting limits.

---

## 9. Manual Overrides — Adjustments Table

The `Adjustments` table allows planners to override two things per line per stage:

| Column | Effect |
|---|---|
| `LoadOverride` | Replaces the routing's standard load value for this line/stage |
| `StartWeek` | Forces Assembly (or Testing) to begin at a specific week instead of the computed start |

**Priority of Assembly start resolution (both scripts):**

1. Manual pin (`Adjustments.StartWeek`) — always wins; alignment never overrides this
2. Alignment-deferred start — computed by the alignment pass (see Section 11)
3. Natural start — derived from OI + routing offsets

Manual pins for Testing work the same way.

---

## 10. Backlog

The `Backlog` table holds already-committed load from in-progress orders that are not in the Demand table. It is loaded once at the start of each script run and never modified.

In the constrained script, backlog occupies capacity before any planned load is placed — `usedAt()` always includes backlog. In the Results output, backlog rows appear as `Status = "Backlog"` with load in the `Backlog` column (not the `Planned` column).

In the constrained script's iteration loop (alignment passes), backlog is preserved across all passes — only `plannedByStatus` and `plannedTotal` are reset between passes.

---

## 11. Deal Alignment Feature

### Purpose

A deal (`OppID`) can contain multiple equipment lines with different lead times. Without alignment, shorter lines finish weeks or months before the slowest line, which is wasteful and operationally misleading. The alignment feature defers shorter lines so all equipment in a flagged deal finishes within a configurable tolerance window of the slowest line.

### Activation

Opt-in per deal via the `AlignFlag` column on the Demand table. **Any line of an OppID with a non-blank, non-zero AlignFlag activates alignment for the entire deal.** Unflagged deals run byte-for-byte as before this feature existed.

The flag survives Sync Demand refreshes: it is treated identically to RoutingKey, Plant, and Priority — preserved from the existing Demand row by LineID and re-emitted into the rebuilt row.

### Anchor

The deal anchor is the **maximum FCA week across all lines of the deal**. All shorter lines are deferred toward this anchor. The anchor is never pulled earlier — alignment is delay-only.

### Tolerance band

A line is only deferred if its gap to the anchor exceeds `AlignToleranceWeeks` (from the Settings table, default 2). Lines already within N weeks of the anchor are left untouched. Deferred lines target `anchor − N` weeks, landing at the edge of the band rather than exactly on the anchor. This softens load concentration.

### Manual pins

Lines with a manual `StartWeek` in the Adjustments table are **never moved by alignment**. A manual pin always wins. A manually-pinned line may still be the deal anchor (and drag other lines later) — this is intentional: if a line was pushed to smooth a plant overload, the deal genuinely finishes when that line finishes.

### Build-only deferral

Alignment defers only Assembly onward. Engineering milestones (KOM, 1st BOM, Frozen Design, Final BOM, CTO load) remain OI-anchored. Procurement offsets remain OI-anchored (procurement carries no capacity load in the current model). ESL re-stretches automatically because its window is always OI → realized FCA.

### Unconstrained script — exact two-pass

1. **Phase 1:** Loop all demand lines, compute each line's natural FCA week via pure date arithmetic (no load placement). Derive `dealAnchor[OppID] = max(naturalFCA)` for each aligned deal.
2. **Compute defers:** For each aligned, non-pinned line where `gap > N`: `deferWeeks = gap − N`; new Assembly start = `naturalAsmStartWeek + deferWeeks`.
3. **Phase 2:** Placement loop, identical to the standard loop, with deferred starts injected. Result is exact — no capacity contention, so deferred lines land precisely at `anchor − N`.

### Constrained script — bounded ratcheting iteration

Because finite capacity can cause carry-forward that shifts a line's realized FCA, exact two-pass alignment is not possible. The constrained script uses bounded iteration:

1. **Pass 1 (baseline):** Full scheduling pass with no overrides. Produces realized FCA week per line. Identical to the pre-alignment behavior.
2. **Ratchet anchor initialization:** For each aligned deal, `anchorByDeal[OppID] = max(realizedFCA)`. The anchor only ever moves later — never earlier. This prevents oscillation.
3. **Compute defers:** Same logic as unconstrained: for non-pinned lines where `gap > N`, defer Assembly start by `gap − N` weeks from the natural start.
4. **Pass 2..K:** Re-run the full scheduling loop (resetting `plannedByStatus` and `plannedTotal`; backlog is untouched). After each pass, update the ratchet anchor if any line's realized FCA moved later. Compute new defers. Stop early if no overrides changed (alignment stabilized). Cap at `MaxAlignPasses` (Settings table, default 3).
5. **Single I/O:** Excel reads happen once before the loop; `writeFast` writes happen once after the final pass. Only the in-memory allocation loop repeats — no performance impact for typical demand sizes.

**Residual imperfection (documented by design):** When deferred lines pile into the anchor's window, new contention can push one or two lines slightly past the frozen anchor. The tolerance band N absorbs most of this. Remaining overshoot is bounded and predictable. This is accepted behavior, not a bug — equivalent to the documented ESL/CTO bypass of `allocateStage()`.

### Settings parameters

| Key | Default | Description |
|---|---|---|
| `AlignToleranceWeeks` | 2 | Lines finishing within N weeks of the anchor are not deferred. Increase for looser alignment; decrease for tighter. |
| `MaxAlignPasses` | 3 | Maximum scheduling iterations for the constrained script. Minimum effective value is 2. Ignored by unconstrained script. |

Both scripts fall back to defaults silently if the Settings table does not exist — safe for workbooks created before this feature.

---

## 12. Sync Demand Script

Merges `DemandImport` (source system feed) into `Demand` (planning master). Runs as a separate Office Script before either planning script.

**Per-line logic:**

- Match by `SourceLineID` (= `LineID` in Demand)
- If matched and signature unchanged → `ChangeFlag = "Unchanged"`
- If matched and signature changed → `ChangeFlag = "Changed"`, diff written to `ChangedFields`
- If not matched in import → `ChangeFlag = "Removed"` (line stays in Demand, excluded from planning)
- If new in import → `ChangeFlag = "New"`, `ChangedFields = "assign RoutingKey + Plant"`

**Preserved columns (survive every sync by LineID match):**

`RoutingKey`, `Plant`, `Priority`, `AlignFlag`

These are planner-set values not present in the source system. They are read from the existing Demand row and re-emitted into the rebuilt row. A line newly added to the import arrives with all four blank.

**Signature** is the concatenation of `OppID|MachineDescription|OIDate|Status`. A change to any of these four fields triggers `ChangeFlag = "Changed"`.

---

## 13. Results Output Format

Both scripts produce Results rows in the same format. **Each row represents the load contribution of a single status in a single week** — there is no pre-aggregated total row across all statuses. Aggregation is done in the UI or via Excel pivot/filter on the Results table.

| Column | Notes |
|---|---|
| Plant | Plant code |
| Stage | ESL / CTO / Assembly / Testing / FAT |
| Week | ISO week label (YYYY-WW) |
| Status | Backlog, or one of the 8 demand statuses |
| Backlog | Load from the Backlog table (0 for planned rows) |
| Planned | Load from this planning run (0 for backlog rows) |
| Total | = Backlog + Planned for this row only |
| Cap | Weekly capacity for this Plant × Stage × Week |
| Util% | `round(Total / Cap × 100)` — 0 if Cap = 0 |
| Overload | `max(0, Total − Cap)` |

Cap is sourced from `CapacityOverride` if a week-specific override exists, otherwise from `CapacityBase`. In the Results output, Cap reflects only the first occurrence per plant per week (to avoid double-counting when multiple status rows exist for the same plant/stage/week).

---

## 14. Known Limitations and Deferred Items

| Item | Status | Notes |
|---|---|---|
| Triggering Office Scripts from the add-in | Deferred post-MVP | Microsoft platform gap; scripts must be run manually from the Automate tab |
| FAT `place()` behavior | Deferred decision | FAT always places unconditionally regardless of capacity; not yet flagged as a constraint |
| Weekly snapshot history to Fabric lakehouse | Deferred | Licensing exists; deferred until the operational tool is stable |
| Procurement load stage | Not modelled | Procurement is a timing offset only; no capacity load. JIT-shift of procurement is trivial to add if a Procurement stage is introduced later |
| Alignment convergence iteration | Single bounded pass accepted | For very congested plants, a line may land slightly past the anchor after K passes. Tolerance band N absorbs most of this |
| ESL capacity ceiling | Not enforced | ESL is load visibility only; adding a capacity ceiling would require promoting ESL to `allocateStage()` |

---

## 15. Deployment Notes

- **Add-in hosting:** GitHub Pages at `https://ricardo-gama.github.io/sop-planning-tool/`
- **Manifest distribution:** Network-share trusted catalog via synced SharePoint folder; UNC-style path required in Excel Trust Center
- **HTTP:** Add-in runs on HTTP with `AllowHTTP` registry DWORD set; manifest uses `http://` not `https://`
- **Deployment sequence:** Always commit source to `main` first (`git add . && git commit && git push`), then `npm run build` → `npm run deploy`. Omitting the commit step risks losing source changes
- **Build currency:** `APP_VERSION` constant in the add-in header confirms which build is live. Current version: `v0.2.0`
- **Co-authoring:** The workbook uses AutoSave with no Power Pivot (Power Pivot disables co-authoring). All named tables are standard Excel tables safe for concurrent editing
