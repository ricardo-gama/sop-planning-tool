/// <reference types="office-js" />
import * as React from "react";
import { useState, useEffect, useCallback, useMemo, useRef } from "react";

/* ─── version ─────────────────────────────────────────────────────────────── */
const APP_VERSION = "v0.2.1";

/* ─── types ───────────────────────────────────────────────────────────────── */
type View = "tabs" | "pipeline-detail" | "opp-detail";
type TabId = "pipeline" | "opp";

interface OppImportRow {
  rowIdx: number;
  opportunityId: string;
  cep: string;
  opportunityName: string;
  customer: string;
  accountId: string;
  region: string;
  subRegion: string;
  country: string;
  stage: string;
  value: number;
  marketProbability: string;
  groupProbability: string;
  expectedOiDate: string;
  forecastStatus: string;
  activeOfferId: string;
  activeOfferName: string;
  offerBusinessId: string;
  offerQuoteGuid: string;
  marketCategoryL1: string;
  portfolioL2: string;
  marketPortfolio: string;
  scoreBand: string;
  modifiedOn: string;
  promote: boolean;
}

interface DemandWorkRow {
  rowIdx: number;
  oppId: string;
  status: string;
  inPipeline: boolean;
  dealType: string;
  selectedOfferId: string;
  crmOiDate: string;
  adjustedOiDate: string;
  risk: boolean;
  lastModifiedBy: string;
  lastModifiedDate: string;
  // ── anticipation / OI-Header fields — required only for certain DealTypes ─
  anticipationStartDate: string;
  anticipationExpirationDate: string;
  oiHeader: string;
  // ── snapshot cols 10-18 — blank on pre-migration rows ────────────────────
  opportunityName: string;
  customer: string;
  region: string;
  subRegion: string;
  marketPortfolio: string;
  value: number;
  forecastStatus: string;
  scoreBand: string;
  cepStatus: string;  // "Active" | "LeftSync" | "" (blank = pre-migration)
}

interface CommentRow {
  rowIdx: number;
  commentId: string;
  oppId: string;
  commentText: string;
  author: string;
  timestamp: string;
}

interface ChangeLogEntry {
  rowIdx: number;
  changeLogId: string;
  oppId: string;
  opportunityName: string;
  eventType: string;
  eventDetail: string;
  offerId: string;
  firstDetectedDate: string;
  lastSeenDate: string;
  acknowledgedBy: string;
  acknowledgedDate: string;
}

/**
 * EquipmentResolved — machine-derived, materialized effective equipment list.
 * Written by SyncEquipment.ts (nightly) and SyncPipeline.ts (on-demand, step 5).
 * Read-only from the add-in's point of view — never written here.
 */
interface EquipmentResolvedRow {
  rowIdx: number;
  lineKey: string;
  offerBusinessId: string;
  oppId: string; // CEP — join key back to OppImport/DemandWork
  origin: string; // "Import" | "Estimate" | "Added"
  l5Code: string;
  qty: number;
  resolvedArchetype: string;
  resolvedPlant: string;
  resolutionStatus: string; // "Resolved" | "NoMapping" | "NoPlantRule" | "NoSubRegion" | "Ambiguous"
  resolutionDetail: string;
  advisory: string; // "" | "PhasedOutArchetype" | "L5NotActive" | "L5StatusUnknown" | "InactiveArchetype"
  masterVersion: string;
}

/**
 * EquipmentAdjust — human-owned overlay table. The add-in writes Exclude /
 * OverrideQty / Estimate / Added rows here; SyncEquipment.ts / SyncPipeline.ts
 * read this to rebuild EquipmentResolved. ProductArchetype is the Ch.4
 * handoff redesign column — planner picks a family, plant is resolved by the
 * script, not picked here. ArchetypeCode/Plant remain as a legacy fallback
 * read path only; the add-in no longer writes to them.
 */
interface EquipmentAdjustRow {
  rowIdx: number;
  adjustmentId: string;
  offerBusinessId: string;
  adjustType: string; // "Exclude" | "OverrideQty" | "Estimate" | "Added"
  targetLineKey: string;
  productArchetype: string;
  archetypeCode: string; // legacy, read-only from the add-in's point of view
  plant: string;          // legacy, read-only from the add-in's point of view
  qty: number;
  superseded: boolean;
  supersededDate: string;
  createdDate: string;
}

/** ArchetypeMaster — only the columns the Equipment tab needs (62 cols total in the mirror). */
interface ArchetypeMasterRow {
  rowIdx: number;
  archetype: string;
  archetypeName: string;
  productArchetype: string;
  family: string;
  active: boolean;
}

/** EquipmentImport — only what the Equipment tab needs for display + Estimate/Added origin detection. */
interface EquipmentImportLineInfo {
  lineKey: string;
  offerBusinessId: string;
  offerItemL5: string;
  l5Code: string;
}

interface PipelineRow {
  opp: OppImportRow;   // live from OppImport, or synthesised from DW snapshot
  dw: DemandWorkRow;
  effectiveOiDate: string;
  offerDiverged: boolean;
  isLeftSync: boolean; // true = deal no longer in CEP sync feed
  anticipationOverdue: boolean; // true = anticipation DealType past its expiration date
}

/* ─── constants ───────────────────────────────────────────────────────────── */
// Display label changed to "Deal status" — column name in DemandWork stays "DealType"
const DEAL_STATUSES = [
  "Anticipation — Eng. Only",
  "Anticipation — In Full",
  "Backup",
  "Forecast",
  "From OI to KOM",
  "Lost",
  "On Hand",
  "Other",
];

// DealTypes that need an anticipation start/expiration date
const ANTICIPATION_TYPES = ["Anticipation — Eng. Only", "Anticipation — In Full"];
// DealTypes that need an OI Header (execution deal number)
const OI_HEADER_TYPES = [...ANTICIPATION_TYPES, "From OI to KOM"];
// Letter + "1" + 6 digits, e.g. H1000000 — leading letter can vary
const OI_HEADER_PATTERN = /^[A-Za-z]1\d{6}$/;

const BAND_COLOR: Record<string, string> = {
  A: "#16a34a", B: "#ca8a04", C: "#dc2626", D: "#dc2626",
};

/** Stable empty Set reference for the read-only EquipmentTable usage (Opportunity detail has no pending-removal concept). */
const EMPTY_KEY_SET: Set<string> = new Set();

/* ─── helpers ─────────────────────────────────────────────────────────────── */
/**
 * Safe boolean parser — handles Excel TRUE/FALSE, 1/0, and JS booleans.
 * Boolean("FALSE") === true in JS, so never use Boolean() on Excel cell values.
 */
function parseBool(v: unknown): boolean {
  if (v === true  || v === 1)   return true;
  if (v === false || v === 0)   return false;
  const s = String(v).trim().toUpperCase();
  return s === "TRUE" || s === "1" || s === "YES";
}

/**
 * Format a date for display as dd/mm/yyyy.
 * Handles Excel serial numbers (stored as number or numeric string)
 * and ISO date strings from Dataverse / CRM.
 */
function formatDate(val: string | number | null | undefined): string {
  if (val === null || val === undefined || val === "") return "—";
  // Excel serial date: numbers between 40000–70000 cover 2009–2091
  const n = Number(val);
  if (!isNaN(n) && n > 40000 && n < 70000) {
    const d = new Date((n - 25569) * 86400000);
    return `${String(d.getUTCDate()).padStart(2,"0")}/${String(d.getUTCMonth()+1).padStart(2,"0")}/${d.getUTCFullYear()}`;
  }
  // ISO date string ("2026-09-15" or "2026-09-15T00:00:00Z")
  if (typeof val === "string" && val.length >= 8) {
    const d = new Date(val);
    if (!isNaN(d.getTime())) {
      return `${String(d.getUTCDate()).padStart(2,"0")}/${String(d.getUTCMonth()+1).padStart(2,"0")}/${d.getUTCFullYear()}`;
    }
  }
  return String(val) || "—";
}

/**
 * Converts an Excel serial date or ISO date string into the "YYYY-MM-DD" format
 * required by <input type="date">. Native date inputs silently show blank for
 * any other format, so this must run before a raw workbook value is bound to one.
 */
function toDateInputValue(val: string | number | null | undefined): string {
  if (val === null || val === undefined || val === "") return "";
  const n = Number(val);
  if (!isNaN(n) && n > 40000 && n < 70000) {
    const d = new Date((n - 25569) * 86400000);
    return `${d.getUTCFullYear()}-${String(d.getUTCMonth()+1).padStart(2,"0")}-${String(d.getUTCDate()).padStart(2,"0")}`;
  }
  if (typeof val === "string" && val.length >= 8) {
    const d = new Date(val);
    if (!isNaN(d.getTime())) {
      return `${d.getUTCFullYear()}-${String(d.getUTCMonth()+1).padStart(2,"0")}-${String(d.getUTCDate()).padStart(2,"0")}`;
    }
  }
  return "";
}

/**
 * True when a deal is one of the anticipation DealTypes and its expiration
 * date has passed. Compares by day only (no time-of-day drift).
 */
function isAnticipationOverdue(dealType: string, anticipationExpirationDate: string): boolean {
  if (!ANTICIPATION_TYPES.includes(dealType)) return false;
  const iso = toDateInputValue(anticipationExpirationDate);
  if (!iso) return false;
  const today = new Date().toISOString().slice(0, 10);
  return iso < today;
}

/**
 * Synthesises an OppImportRow from DemandWork snapshot columns when the live
 * OppImport join fails (deal has left the CEP sync feed). Fields not stored
 * in the snapshot are left blank — the result is display-only.
 */
function makeSnapshotOpp(dw: DemandWorkRow): OppImportRow {
  return {
    rowIdx: -1,
    opportunityId: "",
    cep: dw.oppId,
    opportunityName: dw.opportunityName,
    customer: dw.customer,
    accountId: "",
    region: dw.region,
    subRegion: dw.subRegion,
    country: "",
    stage: "",
    value: dw.value,
    //value: Number(dw.value ?? 0),
    marketProbability: "",
    groupProbability: "",
    expectedOiDate: dw.crmOiDate,
    forecastStatus: dw.forecastStatus,
    activeOfferId: "",
    activeOfferName: "",
    offerBusinessId: dw.selectedOfferId,
    offerQuoteGuid: "",
    marketCategoryL1: "",
    portfolioL2: "",
    marketPortfolio: dw.marketPortfolio,
    scoreBand: dw.scoreBand,
    modifiedOn: "",
    promote: false,
  };
}

/* ─── Office.js helpers ───────────────────────────────────────────────────── */
async function tryReadTable(
  name: string
): Promise<{ headers: string[]; rows: (string | number | boolean)[][] }> {
  try {
    return await Excel.run(async (ctx) => {
      const tbl = ctx.workbook.tables.getItem(name);
      const hdrRange  = tbl.getHeaderRowRange();
      const bodyRange = tbl.getDataBodyRange();
      hdrRange.load("values");
      bodyRange.load("values");
      await ctx.sync();
      const headers = (hdrRange.values[0] as string[]).map((h) => String(h ?? "").trim());
      const rows    = (bodyRange.values as (string | number | boolean)[][])
                        .filter((r) => r[0] !== "" && r[0] !== null);
      return { headers, rows };
    });
  } catch {
    return { headers: [], rows: [] };
  }
}

async function writeCellInTable(
  tableName: string, rowIdx: number, colName: string,
  value: string | number | boolean, headers: string[]
): Promise<void> {
  const colIdx = headers.indexOf(colName);
  if (colIdx < 0) throw new Error(`Column "${colName}" not found in ${tableName}`);
  await Excel.run(async (ctx) => {
    ctx.workbook.tables.getItem(tableName).getDataBodyRange().getCell(rowIdx, colIdx).values = [[value]];
    await ctx.sync();
  });
}

async function writeMultipleCells(
  tableName: string, rowIdx: number,
  updates: { col: string; value: string | number | boolean }[],
  headers: string[]
): Promise<void> {
  await Excel.run(async (ctx) => {
    const body = ctx.workbook.tables.getItem(tableName).getDataBodyRange();
    for (const u of updates) {
      const ci = headers.indexOf(u.col);
      if (ci >= 0) body.getCell(rowIdx, ci).values = [[u.value]];
    }
    await ctx.sync();
  });
}

async function appendRowToTable(tableName: string, row: (string | number | boolean)[]): Promise<void> {
  await Excel.run(async (ctx) => {
    ctx.workbook.tables.getItem(tableName).rows.add(-1, [row]);
    await ctx.sync();
  });
}

/**
 * Builds a positional row array from a set of named values, placed at
 * whatever index each column actually occupies in `headers` — instead of
 * assuming a fixed column order. Used for EquipmentAdjust writes because
 * that table's exact column order (especially the new ProductArchetype
 * column) isn't guaranteed to match any particular position; any column
 * present in `headers` but not in `valuesByCol` is written as "".
 */
function buildRowFromHeaders(
  headers: string[],
  valuesByCol: Record<string, string | number | boolean>
): (string | number | boolean)[] {
  return headers.map((h) => (Object.prototype.hasOwnProperty.call(valuesByCol, h) ? valuesByCol[h] : ""));
}

async function deleteRowFromTable(tableName: string, rowIdx: number): Promise<void> {
  await Excel.run(async (ctx) => {
    const table = ctx.workbook.tables.getItem(tableName);
    // An active AutoFilter on the table can also block a row delete/shift —
    // clear it first so the delete always succeeds. The filter dropdowns stay
    // on the table; only the active criteria are cleared.
    const filter = table.autoFilter;
    filter.load("isDataFiltered");
    await ctx.sync();
    if (filter.isDataFiltered) filter.clearCriteria();

    // Delete via the Table's row API, not a raw Range.delete() — a Range delete
    // shifts sheet cells and Excel refuses (InsertDeleteConflict) when another
    // table sits below and would need to shift too. TableRow.delete() shifts
    // only within the table.
    table.rows.getItemAt(rowIdx).delete();
    await ctx.sync();
  });
}

/* ─── CollapsibleSection ──────────────────────────────────────────────────── */
function CollapsibleSection({
  title, badge, subtitle, defaultOpen = true, dimTitle = false, children,
}: React.PropsWithChildren<{
  title: React.ReactNode;
  badge?: React.ReactNode;
  subtitle?: string;
  defaultOpen?: boolean;
  dimTitle?: boolean;
}>): React.ReactElement {
  const [open, setOpen] = useState(defaultOpen);
  return (
    <div style={s.section}>
      <button style={s.sectionBtn} onClick={() => setOpen(!open)} aria-expanded={open}>
        <span style={{ ...s.sectionTitle, color: dimTitle ? "#94a3b8" : "#0f2942" }}>
          {title}{badge && <span style={{ marginLeft: 6 }}>{badge}</span>}
        </span>
        <div style={{ display: "flex" as const, alignItems: "center" as const, gap: 6 }}>
          {subtitle && <span style={s.sectionSub}>{subtitle}</span>}
          <span style={{
            fontSize: 16, color: "#94a3b8", display: "inline-block" as const,
            transform: open ? "rotate(0deg)" : "rotate(-90deg)", transition: "transform 0.15s", lineHeight: 1,
          }}>⌄</span>
        </div>
      </button>
      {open && <div style={s.sectionBody}>{children}</div>}
    </div>
  );
}

/* ─── CrmGrid — 2-column layout (label | value | label | value per row) ──── */
function CrmGrid({ pairs }: { pairs: [string, string][] }) {
  return (
    <div style={s.crmGrid}>
      {pairs.map(([lbl, val], i) => (
        <React.Fragment key={i}>
          <span style={s.crmLabel}>{lbl}</span>
          <span style={s.crmValue}>{val || "—"}</span>
        </React.Fragment>
      ))}
    </div>
  );
}

/* ─── MultiSelect — checkbox dropdown filter ─────────────────────────────── */
interface MultiSelectOption { label: string; value: string; }
interface MultiSelectProps {
  label: string;
  options: string[];
  selected: string[];
  onChange: (vals: string[]) => void;
  extraOptions?: MultiSelectOption[]; // prepended before dynamic options
}
function MultiSelect({ label, options, selected, onChange, extraOptions }: MultiSelectProps) {
  const [open, setOpen] = useState(false);
  const ref = useRef<HTMLDivElement>(null);

  useEffect(() => {
    function onDocClick(e: MouseEvent) {
      if (ref.current && !ref.current.contains(e.target as Node)) setOpen(false);
    }
    document.addEventListener("mousedown", onDocClick);
    return () => document.removeEventListener("mousedown", onDocClick);
  }, []);

  const allOptions: MultiSelectOption[] = [
    ...(extraOptions ?? []),
    ...options.map((o) => ({ label: o, value: o })),
  ];
  const allSelected = allOptions.length > 0 && selected.length === allOptions.length;

  const toggle = (val: string) => {
    onChange(selected.includes(val) ? selected.filter((v) => v !== val) : [...selected, val]);
  };

  return (
    <div style={s.multiSelWrap} ref={ref}>
      <button style={selected.length > 0 ? s.multiSelBtnOn : s.multiSelBtn} onClick={() => setOpen(!open)}>
        {label}{selected.length > 0 ? ` (${selected.length})` : ""}
      </button>
      {open && (
        <div style={s.multiSelPanel}>
          <div style={s.multiSelToggle} onClick={() => onChange(allSelected ? [] : allOptions.map((o) => o.value))}>
            {allSelected ? "Clear" : "Select all"}
          </div>
          {allOptions.map((o) => (
            <label key={o.value} style={s.multiSelRow}>
              <input type="checkbox" checked={selected.includes(o.value)} onChange={() => toggle(o.value)} />
              {o.label}
            </label>
          ))}
        </div>
      )}
    </div>
  );
}

/**
 * Description shown for an equipment line. Import-origin lines prefer
 * EquipmentImport's own OfferItemL5 (the category-level label carried on the
 * imported line itself); if that's blank, fall back to L5Mapping.L5Description
 * via the line's L5Code — same underlying information, just from the other
 * source. Doesn't apply to Estimate/Added lines: those never carry an
 * L5Code, so there's nothing to look up — the Product Archetype column
 * carries the meaning instead for those rows.
 */
function equipmentDescription(
  line: EquipmentResolvedRow,
  importLineByKey: Map<string, EquipmentImportLineInfo>,
  l5DescByCode: Map<string, string>
): string {
  if (line.origin !== "Import") return "";
  const imp = importLineByKey.get(line.lineKey);
  if (imp?.offerItemL5) return imp.offerItemL5;
  return l5DescByCode.get(line.l5Code) || "";
}

/* ─── EquipmentTable — shared between Opportunity detail (view-only) and ──
   Pipeline detail (editable) ────────────────────────────────────────────── */
interface EquipmentTableProps {
  lines: EquipmentResolvedRow[];
  excluded: EquipmentAdjustRow[];
  pendingManual: EquipmentAdjustRow[];
  pendingRestore: EquipmentImportLineInfo[];
  importLineByKey: Map<string, EquipmentImportLineInfo>;
  l5DescByCode: Map<string, string>;
  archetypeMasterByCode: Map<string, ArchetypeMasterRow>;
  adjustByAdjustmentId: Map<string, EquipmentAdjustRow>;
  pendingRemovalKeys: Set<string>;
  editable: boolean;
  saving: boolean;
  onExclude: (line: EquipmentResolvedRow) => void;
  onRestore: (adj: EquipmentAdjustRow) => void;
  onRemoveManual: (line: EquipmentResolvedRow) => void;
  onRemovePending: (adj: EquipmentAdjustRow) => void;
  onKeepLeftCep: (line: EquipmentResolvedRow) => void;
  onRemoveLeftCep: (line: EquipmentResolvedRow) => void;
}

function EquipmentTable({
  lines, excluded, pendingManual, pendingRestore, importLineByKey, l5DescByCode, archetypeMasterByCode, adjustByAdjustmentId,
  pendingRemovalKeys, editable, saving, onExclude, onRestore, onRemoveManual, onRemovePending,
  onKeepLeftCep, onRemoveLeftCep,
}: EquipmentTableProps): React.ReactElement {
  const grid = editable ? s.eqGridEdit : s.eqGridView;
  const cols = editable
    ? ["Description", "Product archetype", "Origin", "Plant", "Archetype", "Status", "Advisory", ""]
    : ["Description", "Product archetype", "Origin", "Plant", "Archetype", "Status", "Advisory"];

  const hasAnyRows = lines.length > 0 || pendingManual.length > 0 || pendingRestore.length > 0;

  if (!hasAnyRows && excluded.length === 0) {
    return <div style={s.emptySmall}>No equipment lines for this offer yet.</div>;
  }

  return (
    <>
      {hasAnyRows && (
        <div style={s.eqTableWrap}>
          <div style={{ minWidth: editable ? 980 : 860 }}>
            <div style={{ ...grid, ...s.eqTableHead }}>
              {cols.map((h, i) => <span key={i} style={s.th}>{h}</span>)}
            </div>
            {editable && pendingManual.map(adj => (
              <div key={`pending-${adj.rowIdx}`} style={{ ...grid, background: "#eff6ff", padding: "6px 14px" }}>
                <span style={{ ...s.td, ...s.ell }}>—</span>
                <span style={{ ...s.td, ...s.ell }}>{adj.productArchetype || adj.archetypeCode || "—"}</span>
                <span style={s.td}><span style={s.eqOriginBadge}>{adj.adjustType || "—"}</span></span>
                <span style={{ ...s.td, ...s.ell }}>—</span>
                <span style={s.td}>—</span>
                <span style={s.td}><span style={s.eqStatusPending}>Pending sync</span></span>
                <span style={s.td} />
                <span style={s.eqActionRow}>
                  <button style={s.deleteLink} onClick={() => onRemovePending(adj)} disabled={saving}>Remove</button>
                </span>
              </div>
            ))}
            {editable && pendingRestore.map(imp => (
              <div key={`restore-${imp.lineKey}`} style={{ ...grid, background: "#eff6ff", padding: "6px 14px" }}>
                <span style={{ ...s.td, ...s.ell }} title={imp.offerItemL5 || l5DescByCode.get(imp.l5Code) || ""}>
                  {imp.offerItemL5 || l5DescByCode.get(imp.l5Code) || "—"}
                </span>
                <span style={{ ...s.td, ...s.ell }}>—</span>
                <span style={s.td}><span style={s.eqOriginBadge}>Import</span></span>
                <span style={{ ...s.td, ...s.ell }}>—</span>
                <span style={s.td}>—</span>
                <span style={s.td}><span style={s.eqStatusPending}>Pending sync</span></span>
                <span style={s.td} />
                <span style={s.td} />
              </div>
            ))}
            {lines.map((line, i) => {
              const isManual = line.origin !== "Import";
              const desc = equipmentDescription(line, importLineByKey, l5DescByCode);
              // EquipmentResolved has no ProductArchetype column of its own —
              // for a resolved manual line, derive it from the resolved
              // archetype code; if resolution FAILED (NoMapping etc.) that
              // lookup is empty, so fall back to what the planner originally
              // picked, still sitting on the underlying EquipmentAdjust row.
              // This is what keeps a failed/unresolved manual line from
              // showing up completely blank.
              const productArchetype =
                archetypeMasterByCode.get(line.resolvedArchetype)?.productArchetype
                || (isManual ? adjustByAdjustmentId.get(line.lineKey)?.productArchetype : "")
                || "";
              const statusOk = line.resolutionStatus === "Resolved";
              const isPendingRemoval = editable && pendingRemovalKeys.has(line.lineKey);
              const isLeftCep = line.advisory === "LeftCep";
              const bg = isPendingRemoval ? "#fef2f2" : isLeftCep ? "#fffbeb" : i % 2 === 0 ? "#fff" : "#f8fafc";

              return (
                <div key={line.lineKey} style={{ ...grid, background: bg, padding: "6px 14px" }}>
                  <span style={{ ...s.td, ...s.ell }} title={desc}>{desc || "—"}</span>
                  <span style={{ ...s.td, ...s.ell }}>{productArchetype || "—"}</span>
                  <span style={s.td}><span style={s.eqOriginBadge}>{line.origin || "—"}</span></span>
                  <span style={{ ...s.td, ...s.ell }}>{line.resolvedPlant || "—"}</span>
                  <span style={{ ...s.td, ...s.eqArchetypeCode }}>{line.resolvedArchetype || "—"}</span>
                  <span style={s.td}>
                    <span style={statusOk ? s.eqStatusOk : s.eqStatusWarn} title={line.resolutionDetail}>
                      {line.resolutionStatus || "Unresolved"}
                    </span>
                  </span>
                  <span style={s.td}>
                    {line.advisory && (
                      <span style={isLeftCep ? s.eqAdvisoryLeftCep : s.eqAdvisoryBadge} title={isLeftCep ? "No longer present in this offer per CEP sync — Keep to make it a permanent line, or Remove to let it drop." : undefined}>
                        {isLeftCep ? "left CEP" : line.advisory}
                      </span>
                    )}
                  </span>
                  {editable && (
                    isPendingRemoval ? (
                      <span style={s.eqStatusRemoving} title="Removed from EquipmentAdjust — run SyncPipeline in Excel (or wait for the nightly sync) to clear it from this list">
                        Removing…
                      </span>
                    ) : isLeftCep ? (
                      <span style={s.eqActionRow}>
                        <button style={s.deleteLink} onClick={() => onKeepLeftCep(line)} disabled={saving}>Keep</button>
                        <button style={s.deleteLink} onClick={() => onRemoveLeftCep(line)} disabled={saving}>Remove</button>
                      </span>
                    ) : (
                      <span style={s.eqActionRow}>
                        {!isManual && (
                          <button style={s.deleteLink} onClick={() => onExclude(line)} disabled={saving}>Exclude</button>
                        )}
                        {isManual && <button style={s.deleteLink} onClick={() => onRemoveManual(line)} disabled={saving}>Remove</button>}
                      </span>
                    )
                  )}
                </div>
              );
            })}
          </div>
        </div>
      )}

      {editable && excluded.length > 0 && (
        <CollapsibleSection title="Excluded lines" defaultOpen={false} dimTitle subtitle={`${excluded.length}`}>
          {excluded.map(adj => {
            const imp = importLineByKey.get(adj.targetLineKey);
            const desc = imp?.offerItemL5 || (imp ? l5DescByCode.get(imp.l5Code) : "") || adj.targetLineKey;
            return (
              <div key={adj.rowIdx} style={s.eqExcludedRow}>
                <span style={s.eqLabel}>{desc}</span>
                <button style={s.deleteLink} onClick={() => onRestore(adj)} disabled={saving}>Restore</button>
              </div>
            );
          })}
        </CollapsibleSection>
      )}
    </>
  );
}

/* ─── Main App ────────────────────────────────────────────────────────────── */
export default function App() {
  /* ── view ── */
  const [view, setView] = useState<View>("tabs");
  const [tab, setTab]   = useState<TabId>("pipeline");

  /* ── raw data ── */
  const [oppImport,   setOppImport]   = useState<OppImportRow[]>([]);
  const [oppHeaders,  setOppHeaders]  = useState<string[]>([]);
  const [demandWork,  setDemandWork]  = useState<DemandWorkRow[]>([]);
  const [dwHeaders,   setDwHeaders]   = useState<string[]>([]);
  const [comments,    setComments]    = useState<CommentRow[]>([]);
  const [commentHeaders, setCommentHeaders] = useState<string[]>([]);
  const [changeLogs,  setChangeLogs]  = useState<ChangeLogEntry[]>([]);
  const [clHeaders,   setClHeaders]   = useState<string[]>([]);
  const [equipmentResolved, setEquipmentResolved] = useState<EquipmentResolvedRow[]>([]);
  const [equipmentAdjust,   setEquipmentAdjust]   = useState<EquipmentAdjustRow[]>([]);
  const [eaHeaders,         setEaHeaders]         = useState<string[]>([]);
  const [archetypeMaster,   setArchetypeMaster]   = useState<ArchetypeMasterRow[]>([]);
  const [equipmentImportLines, setEquipmentImportLines] = useState<EquipmentImportLineInfo[]>([]);
  const [l5Descriptions, setL5Descriptions] = useState<{ l5Code: string; l5Description: string }[]>([]);

  /* ── ui ── */
  const [loading,    setLoading]    = useState(false);
  const [loadMsg,    setLoadMsg]    = useState("");
  const [errMsg,     setErrMsg]     = useState("");
  const [actionNote, setActionNote] = useState("");
  const [lastLoaded, setLastLoaded] = useState("");

  /* ── navigation ── */
  const [selectedPipeline, setSelectedPipeline] = useState<PipelineRow | null>(null);
  const [selectedOpp,      setSelectedOpp]      = useState<OppImportRow | null>(null);

  /* ── pipeline detail edit state ── */
  const [pdDealType,       setPdDealType]       = useState("");
  const [pdSelectedOffer,  setPdSelectedOffer]  = useState("");
  const [pdAdjustedOi,     setPdAdjustedOi]     = useState("");
  const [pdRisk,           setPdRisk]           = useState(false);
  const [pdAnticipationStart,      setPdAnticipationStart]      = useState("");
  const [pdAnticipationExpiration, setPdAnticipationExpiration] = useState("");
  const [pdOiHeader,               setPdOiHeader]               = useState("");
  const [pdSaving,         setPdSaving]         = useState(false);
  const [pdNewComment,     setPdNewComment]      = useState("");
  const [pdAddingComment,  setPdAddingComment]  = useState(false);

  /* ── pipeline filter state ── */
  const [pipeSearch,             setPipeSearch]             = useState("");
  const [pipeFilterRegion,       setPipeFilterRegion]       = useState<string[]>([]);
  const [pipeFilterSubRegion,    setPipeFilterSubRegion]    = useState<string[]>([]);
  const [pipeFilterPortfolio,    setPipeFilterPortfolio]    = useState<string[]>([]);
  const [pipeFilterForecastStatus, setPipeFilterForecastStatus] = useState<string[]>([]);
  const [pipeFilterDealStatus,   setPipeFilterDealStatus]   = useState<string[]>([]);
  const [pipeFilterChangeLog,    setPipeFilterChangeLog]    = useState("");

  /* ── opportunities filter state ── */
  const [oppSearch,                setOppSearch]                = useState("");
  const [oppFilterForecastStatus,  setOppFilterForecastStatus]  = useState<string[]>([]);
  const [oppFilterRegion,          setOppFilterRegion]          = useState<string[]>([]);
  const [oppFilterPortfolio,       setOppFilterPortfolio]       = useState<string[]>([]);
  const [oppFilterBand,            setOppFilterBand]            = useState<string[]>([]);
  const [oppFlaggedOnly,           setOppFlaggedOnly]           = useState(false);

  /* ── opp detail ── */
  const [odPromoting, setOdPromoting] = useState(false);

  /* ── equipment (opp detail) ── */
  const [eqSaving,          setEqSaving]          = useState(false);
  const [eqAddingLine,      setEqAddingLine]      = useState(false);
  const [eqNewFamily,           setEqNewFamily]           = useState("");
  const [eqNewProductArchetype, setEqNewProductArchetype] = useState("");
  const [eqActionNote,      setEqActionNote]      = useState("");
  // LineKeys just Excluded/Removed by the planner but not yet actually gone
  // from EquipmentResolved (that only happens once SyncPipeline/SyncEquipment
  // rebuilds it) — tinted light red in the table as a "this is on its way
  // out" signal, purely client-side bookkeeping.
  const [pendingRemovalKeys, setPendingRemovalKeys] = useState<Set<string>>(new Set());
  // EquipmentImport LineKeys whose Exclude was just undone (Restore) but
  // haven't reappeared in EquipmentResolved yet — tinted light blue like a
  // pending manual add, same "sync hasn't caught up" signal in reverse.
  const [pendingRestoreKeys, setPendingRestoreKeys] = useState<Set<string>>(new Set());

  /* ── pipeline row delete ── */
  const [pendingDeleteCep, setPendingDeleteCep] = useState<string | null>(null);
  const [deleting,         setDeleting]         = useState(false);

  /* ── load ─────────────────────────────────────────────────────────────── */
  const load = useCallback(async () => {
    setLoading(true); setErrMsg(""); setLoadMsg("Reading workbook…");
    try {
      const [oData, dwData, cmData, clData, erData, eaData, amData, eiData, l5mData] = await Promise.all([
        tryReadTable("OppImport"),
        tryReadTable("DemandWork"),
        tryReadTable("DemandWorkComments"),
        tryReadTable("ChangeLog"),
        tryReadTable("EquipmentResolved"),
        tryReadTable("EquipmentAdjust"),
        tryReadTable("ArchetypeMaster"),
        tryReadTable("EquipmentImport"),
        tryReadTable("L5Mapping"),
      ]);

      // OppImport — 24 cols, positional
      setOppHeaders(oData.headers);
      setOppImport(oData.rows.map((r, i) => ({
        rowIdx: i,
        opportunityId:     String(r[0]  ?? ""),
        cep:               String(r[1]  ?? ""),
        opportunityName:   String(r[2]  ?? ""),
        customer:          String(r[3]  ?? ""),
        accountId:         String(r[4]  ?? ""),
        region:            String(r[5]  ?? ""),
        subRegion:         String(r[6]  ?? ""),
        country:           String(r[7]  ?? ""),
        stage:             String(r[8]  ?? ""),
        value:             Number(r[9]  ?? 0),
        marketProbability: String(r[10] ?? ""),
        groupProbability:  String(r[11] ?? ""),
        expectedOiDate:    String(r[12] ?? ""),
        forecastStatus:    String(r[13] ?? ""),
        activeOfferId:     String(r[14] ?? ""),
        activeOfferName:   String(r[15] ?? ""),
        offerBusinessId:   String(r[16] ?? ""),
        offerQuoteGuid:    String(r[17] ?? ""),
        marketCategoryL1:  String(r[18] ?? ""),
        portfolioL2:       String(r[19] ?? ""),
        marketPortfolio:   String(r[20] ?? ""),
        scoreBand:         String(r[21] ?? ""),
        modifiedOn:        String(r[22] ?? ""),
        promote:           parseBool(r[23]),   // safe boolean parse
      })));

      // DemandWork — 10 cols
      const dwh = dwData.headers;
      setDwHeaders(dwh);
      const dwIdx = (col: string) => dwh.indexOf(col);
      setDemandWork(dwData.rows.map((r, i) => ({
        rowIdx: i,
        oppId:           String(r[dwIdx("OppId")]            ?? ""),
        status:          String(r[dwIdx("Status")]           ?? ""),
        inPipeline:      parseBool(r[dwIdx("InPipeline")]),
        dealType:        String(r[dwIdx("DealType")]         ?? ""),
        selectedOfferId: String(r[dwIdx("SelectedOfferID")]  ?? ""),
        crmOiDate:       String(r[dwIdx("CrmOiDate")]        ?? ""),
        adjustedOiDate:  String(r[dwIdx("AdjustedOiDate")]   ?? ""),
        risk:            parseBool(r[dwIdx("Risk")]),
        lastModifiedBy:  String(r[dwIdx("LastModifiedBy")]   ?? ""),
        lastModifiedDate:String(r[dwIdx("LastModifiedDate")] ?? ""),
        anticipationStartDate:      String(r[dwIdx("AnticipationStartDate")]      ?? ""),
        anticipationExpirationDate: String(r[dwIdx("AnticipationExpirationDate")] ?? ""),
        oiHeader:                   String(r[dwIdx("OIHeader")]                   ?? ""),
        // ── snapshot cols (10-18) — blank on pre-migration rows ──────────
        opportunityName: String(r[dwIdx("OpportunityName")] ?? ""),
        customer:        String(r[dwIdx("Customer")]        ?? ""),
        region:          String(r[dwIdx("Region")]          ?? ""),
        subRegion:       String(r[dwIdx("SubRegion")]       ?? ""),
        marketPortfolio: String(r[dwIdx("MarketPortfolio")] ?? ""),
        value:           Number(r[dwIdx("Value")]           ?? 0),
        forecastStatus:  String(r[dwIdx("ForecastStatus")]  ?? ""),
        scoreBand:       String(r[dwIdx("ScoreBand")]       ?? ""),
        cepStatus:       String(r[dwIdx("CepStatus")]       ?? ""),
      })));

      // DemandWorkComments — 5 cols
      const cmh = cmData.headers;
      setCommentHeaders(cmh);
      const cmIdx = (col: string) => cmh.indexOf(col);
      setComments(cmData.rows.map((r, i) => ({
        rowIdx:      i,
        commentId:   String(r[cmIdx("CommentId")]   ?? ""),
        oppId:       String(r[cmIdx("OppId")]       ?? ""),
        commentText: String(r[cmIdx("CommentText")] ?? ""),
        author:      String(r[cmIdx("Author")]      ?? ""),
        timestamp:   String(r[cmIdx("Timestamp")]   ?? ""),
      })));

      // ChangeLog — 10 cols
      const clh = clData.headers;
      setClHeaders(clh);
      const clIdx = (col: string) => clh.indexOf(col);
      setChangeLogs(clData.rows.map((r, i) => ({
        rowIdx:            i,
        changeLogId:       String(r[clIdx("ChangeLogId")]       ?? ""),
        oppId:             String(r[clIdx("OppId")]             ?? ""),
        opportunityName:   String(r[clIdx("OpportunityName")]   ?? ""),
        eventType:         String(r[clIdx("EventType")]         ?? ""),
        eventDetail:       String(r[clIdx("EventDetail")]       ?? ""),
        offerId:           String(r[clIdx("OfferId")]           ?? ""),
        firstDetectedDate: String(r[clIdx("FirstDetectedDate")] ?? ""),
        lastSeenDate:      String(r[clIdx("LastSeenDate")]      ?? ""),
        acknowledgedBy:    String(r[clIdx("AcknowledgedBy")]    ?? ""),
        acknowledgedDate:  String(r[clIdx("AcknowledgedDate")]  ?? ""),
      })));

      // EquipmentResolved — read by header name (evolving table, don't assume position)
      const erh = erData.headers;
      const erIdx = (col: string) => erh.indexOf(col);
      const resolvedRows = erData.rows.map((r, i) => ({
        rowIdx:            i,
        lineKey:           String(r[erIdx("LineKey")]           ?? ""),
        offerBusinessId:   String(r[erIdx("OfferBusinessId")]   ?? ""),
        oppId:             String(r[erIdx("OppId")]             ?? ""),
        origin:            String(r[erIdx("Origin")]            ?? ""),
        l5Code:            String(r[erIdx("L5Code")]            ?? ""),
        qty:               Number(r[erIdx("Qty")]               ?? 0),
        resolvedArchetype: String(r[erIdx("ResolvedArchetype")] ?? ""),
        resolvedPlant:     String(r[erIdx("ResolvedPlant")]     ?? ""),
        resolutionStatus:  String(r[erIdx("ResolutionStatus")]  ?? ""),
        resolutionDetail:  String(r[erIdx("ResolutionDetail")]  ?? ""),
        advisory:          String(r[erIdx("Advisory")]          ?? ""),
        masterVersion:     String(r[erIdx("MasterVersion")]     ?? ""),
      }));
      setEquipmentResolved(resolvedRows);

      // A LineKey marked pending-removal drops out of tracking once it's
      // actually gone from EquipmentResolved — the sync caught up, nothing
      // left to flag.
      const stillResolvedKeys = new Set(resolvedRows.map(r => r.lineKey));
      setPendingRemovalKeys(prev => {
        const next = new Set<string>();
        prev.forEach(k => { if (stillResolvedKeys.has(k)) next.add(k); });
        return next.size === prev.size ? prev : next;
      });
      // Mirror image: a pending-restore key drops out once it's actually
      // back in EquipmentResolved — the sync caught up, nothing left to flag.
      setPendingRestoreKeys(prev => {
        const next = new Set<string>();
        prev.forEach(k => { if (!stillResolvedKeys.has(k)) next.add(k); });
        return next.size === prev.size ? prev : next;
      });

      // EquipmentAdjust — ProductArchetype is the Ch.4 redesign column; on a
      // sheet where it hasn't been added yet, eaIdx("ProductArchetype") is -1
      // and every row just reads back "" — safe, falls into the legacy path.
      const eah = eaData.headers;
      setEaHeaders(eah);
      const eaIdx = (col: string) => eah.indexOf(col);
      setEquipmentAdjust(eaData.rows.map((r, i) => ({
        rowIdx:           i,
        adjustmentId:     String(r[eaIdx("AdjustmentId")]     ?? ""),
        offerBusinessId:  String(r[eaIdx("OfferBusinessId")]  ?? ""),
        adjustType:       String(r[eaIdx("AdjustType")]       ?? ""),
        targetLineKey:    String(r[eaIdx("TargetLineKey")]    ?? ""),
        productArchetype: String(r[eaIdx("ProductArchetype")] ?? ""),
        archetypeCode:    String(r[eaIdx("ArchetypeCode")]    ?? ""),
        plant:            String(r[eaIdx("Plant")]            ?? ""),
        qty:              Number(r[eaIdx("Qty")]              ?? 0),
        superseded:       parseBool(r[eaIdx("Superseded")]),
        supersededDate:   String(r[eaIdx("SupersededDate")]   ?? ""),
        createdDate:      String(r[eaIdx("CreatedDate")]      ?? ""),
      })));

      // ArchetypeMaster — only the 5 columns the Equipment tab needs (62 cols total)
      const amh = amData.headers;
      const amIdx = (col: string) => amh.indexOf(col);
      setArchetypeMaster(amData.rows.map((r, i) => ({
        rowIdx:           i,
        archetype:        String(r[amIdx("Archetype")]        ?? ""),
        archetypeName:    String(r[amIdx("ArchetypeName")]    ?? ""),
        productArchetype: String(r[amIdx("ProductArchetype")] ?? ""),
        family:           String(r[amIdx("Family")]           ?? ""),
        active:           parseBool(r[amIdx("Active")]),
      })));

      // EquipmentImport — display info + Estimate/Added origin detection.
      // OfferItemL5 (not Description) is what the Equipment table shows —
      // Description is SmartQuote's fully personalized per-line text, too
      // noisy for a table column; OfferItemL5 is the shorter category-level
      // label, consistent with what L5Mapping.L5Description also represents.
      const eih = eiData.headers;
      const eiIdx = (col: string) => eih.indexOf(col);
      setEquipmentImportLines(eiData.rows.map((r) => ({
        lineKey:         String(r[eiIdx("LineKey")]         ?? ""),
        offerBusinessId: String(r[eiIdx("OfferBusinessId")] ?? ""),
        offerItemL5:     String(r[eiIdx("OfferItemL5")]     ?? ""),
        l5Code:          String(r[eiIdx("L5Code")]          ?? ""),
      })));

      // L5Mapping — only L5Code + L5Description, used as the fallback source
      // for the Equipment table's Description column when EquipmentImport's
      // own (line-specific) Description is blank. Doesn't apply to
      // Estimate/Added lines — those never carry an L5Code.
      const l5mh = l5mData.headers;
      const l5mIdx = (col: string) => l5mh.indexOf(col);
      setL5Descriptions(l5mData.rows.map((r) => ({
        l5Code:        String(r[l5mIdx("L5Code")]        ?? ""),
        l5Description: String(r[l5mIdx("L5Description")] ?? ""),
      })));

      const now = new Date();
      setLastLoaded(`${String(now.getHours()).padStart(2,"0")}:${String(now.getMinutes()).padStart(2,"0")}`);
      setLoadMsg(`Loaded ${oData.rows.length} opportunities`);
    } catch (e: unknown) {
      setErrMsg(String(e)); setLoadMsg("");
    }
    setLoading(false);
  }, []);

  useEffect(() => { load(); }, [load]);

  /* ── computed ─────────────────────────────────────────────────────────── */
  const dwByOppId = useMemo(() => {
    const m = new Map<string, DemandWorkRow>();
    for (const dw of demandWork) m.set(dw.oppId, dw);
    return m;
  }, [demandWork]);

  const pipelineRows: PipelineRow[] = useMemo(() => {
    const rows: PipelineRow[] = [];
    for (const dw of demandWork) {
      // All DemandWork rows appear in the Pipeline tab.
      // LeftSync / legacy Inactive rows get an amber indicator.
      const isLeftSync = dw.cepStatus === "LeftSync" || dw.status === "Inactive";
      // Prefer live OppImport data; fall back to DW snapshot for LeftSync deals.
      const liveOpp = oppImport.find((o) => o.cep === dw.oppId);
      const opp     = liveOpp ?? makeSnapshotOpp(dw);
      const effectiveOiDate = dw.adjustedOiDate || dw.crmOiDate || opp.expectedOiDate;
      // Offer divergence is only meaningful for live deals
      const offerDiverged = !isLeftSync
        && !!dw.selectedOfferId
        && !!opp.offerBusinessId
        && dw.selectedOfferId !== opp.offerBusinessId;
      const anticipationOverdue = isAnticipationOverdue(dw.dealType, dw.anticipationExpirationDate);
      rows.push({ opp, dw, effectiveOiDate, offerDiverged, isLeftSync, anticipationOverdue });
    }
    return rows;
  }, [demandWork, oppImport]);

  // Pipeline filter options (from actual pipeline data)
  const pipeRegions     = useMemo(() => Array.from(new Set(pipelineRows.map(pr => pr.opp.region).filter(v => !!v))).sort(),          [pipelineRows]);
  const pipeSubRegions  = useMemo(() => Array.from(new Set(pipelineRows.map(pr => pr.opp.subRegion).filter(v => !!v))).sort(),       [pipelineRows]);
  const pipePortfolios  = useMemo(() => Array.from(new Set(pipelineRows.map(pr => pr.opp.marketPortfolio).filter(v => !!v))).sort(), [pipelineRows]);
  const pipeForecastStatuses = useMemo(() => Array.from(new Set(pipelineRows.map(pr => pr.opp.forecastStatus).filter(v => !!v))).sort(), [pipelineRows]);
  const openClByOpp     = useMemo(() => {const m = new Map<string, number>();
    for (const cl of changeLogs) if (!cl.acknowledgedBy) m.set(cl.oppId, (m.get(cl.oppId) ?? 0) + 1);
    return m;
  }, [changeLogs]);

  // Pipeline filtered rows
  const filteredPipeline: PipelineRow[] = useMemo(() => {
    return pipelineRows.filter(pr => {
      if (pipeSearch) {
        const q = pipeSearch.toLowerCase();
        if (!pr.opp.cep.toLowerCase().includes(q) &&
            !pr.opp.opportunityName.toLowerCase().includes(q) &&
            !pr.opp.customer.toLowerCase().includes(q)) return false;
      }
      if (pipeFilterRegion.length     > 0 && !pipeFilterRegion.includes(pr.opp.region))          return false;
      if (pipeFilterSubRegion.length  > 0 && !pipeFilterSubRegion.includes(pr.opp.subRegion))     return false;
      if (pipeFilterPortfolio.length  > 0 && !pipeFilterPortfolio.includes(pr.opp.marketPortfolio)) return false;
      if (pipeFilterForecastStatus.length > 0 && !pipeFilterForecastStatus.includes(pr.opp.forecastStatus)) return false;
      if (pipeFilterDealStatus.length > 0 && !pipeFilterDealStatus.includes(pr.dw.dealType))      return false;

      const clCount = openClByOpp.get(pr.opp.cep) ?? 0;
      if (pipeFilterChangeLog === "open" && clCount === 0) return false;
      if (pipeFilterChangeLog === "none" && clCount > 0) return false;
      return true;
    });
  }, [pipelineRows, pipeSearch, pipeFilterRegion, pipeFilterSubRegion, pipeFilterPortfolio, pipeFilterForecastStatus, pipeFilterDealStatus, pipeFilterChangeLog, openClByOpp]);

  // Opportunities filter options
  const allForecastStatuses = useMemo(() => Array.from(new Set(oppImport.map(o => o.forecastStatus).filter(v => !!v))).sort(), [oppImport]);
  const allRegions          = useMemo(() => Array.from(new Set(oppImport.map(o => o.region).filter(v => !!v))).sort(),          [oppImport]);
  const allPortfolios       = useMemo(() => Array.from(new Set(oppImport.map(o => o.marketPortfolio).filter(v => !!v))).sort(), [oppImport]);
  const allBands            = useMemo(() => Array.from(new Set(oppImport.map(o => o.scoreBand).filter(v => !!v))).sort(),       [oppImport]);

  const filteredOpps: OppImportRow[] = useMemo(() => {
    return oppImport.filter(o => {
      if (oppFlaggedOnly                    && !o.promote)                                  return false;
      if (oppFilterForecastStatus.length > 0 && !oppFilterForecastStatus.includes(o.forecastStatus)) return false;
      if (oppFilterRegion.length         > 0 && !oppFilterRegion.includes(o.region))         return false;
      if (oppFilterPortfolio.length      > 0 && !oppFilterPortfolio.includes(o.marketPortfolio)) return false;
      if (oppFilterBand.length           > 0 && !oppFilterBand.includes(o.scoreBand))        return false;
      if (oppSearch) {
        const q = oppSearch.toLowerCase();
        if (!o.cep.toLowerCase().includes(q) &&
            !o.opportunityName.toLowerCase().includes(q) &&
            !o.customer.toLowerCase().includes(q)) return false;
      }
      return true;
    });
  }, [oppImport, oppSearch, oppFilterForecastStatus, oppFilterRegion, oppFilterPortfolio, oppFilterBand, oppFlaggedOnly]);

  const totalOpenCl = useMemo(() => changeLogs.filter(c => !c.acknowledgedBy).length, [changeLogs]);

  /* ── equipment (opp detail) computed ─────────────────────────────────── */
  const archetypeMasterByCode = useMemo(() => {
    const m = new Map<string, ArchetypeMasterRow>();
    for (const a of archetypeMaster) if (a.archetype) m.set(a.archetype, a);
    return m;
  }, [archetypeMaster]);

  const familyOptions = useMemo(() =>
    Array.from(new Set(archetypeMaster.map(a => a.family).filter(v => !!v))).sort(),
    [archetypeMaster]);

  // ProductArchetype options, grouped by Family — the add-line form filters
  // ProductArchetype choices to whichever Family the planner picks first.
  const productArchetypesByFamily = useMemo(() => {
    const m = new Map<string, string[]>();
    for (const a of archetypeMaster) {
      if (!a.family || !a.productArchetype) continue;
      const list = m.get(a.family);
      if (list) { if (!list.includes(a.productArchetype)) list.push(a.productArchetype); }
      else m.set(a.family, [a.productArchetype]);
    }
    for (const list of m.values()) list.sort();
    return m;
  }, [archetypeMaster]);

  const importLineByKey = useMemo(() => {
    const m = new Map<string, EquipmentImportLineInfo>();
    for (const l of equipmentImportLines) if (l.lineKey) m.set(l.lineKey, l);
    return m;
  }, [equipmentImportLines]);

  const l5DescByCode = useMemo(() => {
    const m = new Map<string, string>();
    for (const l of l5Descriptions) if (l.l5Code && l.l5Description && !m.has(l.l5Code)) m.set(l.l5Code, l.l5Description);
    return m;
  }, [l5Descriptions]);

  const equipmentByOppId = useMemo(() => {
    const m = new Map<string, EquipmentResolvedRow[]>();
    for (const r of equipmentResolved) {
      const list = m.get(r.oppId);
      if (list) list.push(r); else m.set(r.oppId, [r]);
    }
    return m;
  }, [equipmentResolved]);

  /**
   * EquipmentAdjust rows keyed by AdjustmentId — for Estimate/Added lines,
   * AdjustmentId IS the resulting EquipmentResolved.LineKey (see applyOverlay
   * in the scripts), so this recovers the planner's original ProductArchetype
   * for a manual line even when resolution failed and ResolvedArchetype is
   * blank. EquipmentResolved itself has no ProductArchetype column — this is
   * the only place that information lives.
   */
  const adjustByAdjustmentId = useMemo(() => {
    const m = new Map<string, EquipmentAdjustRow>();
    for (const a of equipmentAdjust) if (a.adjustmentId) m.set(a.adjustmentId, a);
    return m;
  }, [equipmentAdjust]);

  const excludedAdjustByOffer = useMemo(() => {
    const m = new Map<string, EquipmentAdjustRow[]>();
    for (const a of equipmentAdjust) {
      if (a.adjustType !== "Exclude") continue;
      const list = m.get(a.offerBusinessId);
      if (list) list.push(a); else m.set(a.offerBusinessId, [a]);
    }
    return m;
  }, [equipmentAdjust]);

  /**
   * Estimate/Added EquipmentAdjust rows that don't have a matching
   * EquipmentResolved row yet — i.e. added since the last SyncPipeline/
   * SyncEquipment run. Shown as distinctly-colored "pending sync" rows so a
   * planner can see the line was actually added, without waiting for a
   * reload after the next sync.
   */
  const pendingManualByOffer = useMemo(() => {
    const resolvedKeys = new Set(equipmentResolved.map(r => r.lineKey));
    const m = new Map<string, EquipmentAdjustRow[]>();
    for (const a of equipmentAdjust) {
      if (a.adjustType !== "Estimate" && a.adjustType !== "Added") continue;
      if (a.superseded) continue;
      if (resolvedKeys.has(a.adjustmentId)) continue; // already resolved — show via equipmentByOppId instead
      const list = m.get(a.offerBusinessId);
      if (list) list.push(a); else m.set(a.offerBusinessId, [a]);
    }
    return m;
  }, [equipmentAdjust, equipmentResolved]);

  /**
   * EquipmentImport lines whose Exclude was just undone (Restore) but that
   * haven't reappeared in EquipmentResolved yet. Sourced from
   * importLineByKey rather than EquipmentAdjust, since the Exclude row that
   * used to reference them has already been deleted.
   */
  const pendingRestoreByOffer = useMemo(() => {
    const m = new Map<string, EquipmentImportLineInfo[]>();
    pendingRestoreKeys.forEach((key) => {
      const imp = importLineByKey.get(key);
      if (!imp) return;
      const list = m.get(imp.offerBusinessId);
      if (list) list.push(imp); else m.set(imp.offerBusinessId, [imp]);
    });
    return m;
  }, [pendingRestoreKeys, importLineByKey]);

  /* ── actions ──────────────────────────────────────────────────────────── */
  const openPipelineDetail = (pr: PipelineRow) => {
    setSelectedPipeline(pr);
    setPdDealType(pr.dw.dealType);
    setPdSelectedOffer(pr.dw.selectedOfferId);
    setPdAdjustedOi(toDateInputValue(pr.dw.adjustedOiDate));
    setPdRisk(pr.dw.risk);
    setPdAnticipationStart(toDateInputValue(pr.dw.anticipationStartDate));
    setPdAnticipationExpiration(toDateInputValue(pr.dw.anticipationExpirationDate));
    setPdOiHeader(pr.dw.oiHeader);
    setPdNewComment("");
    setActionNote("");
    setEqActionNote("");
    setEqAddingLine(false);
    setEqNewFamily("");
    setEqNewProductArchetype("");
    setPendingRemovalKeys(new Set());
    setPendingRestoreKeys(new Set());
    setView("pipeline-detail");
  };

  const openOppDetail = (opp: OppImportRow) => {
    setSelectedOpp(opp);
    setActionNote("");
    setView("opp-detail");
  };

  const goBack = () => setView("tabs");

  const savePlanningFields = async () => {
    if (!selectedPipeline) return;
    setPdSaving(true); setActionNote("");
    try {
      const today = new Date().toISOString().slice(0, 10);
      await writeMultipleCells("DemandWork", selectedPipeline.dw.rowIdx, [
        { col: "DealType",         value: pdDealType },
        { col: "SelectedOfferID",  value: pdSelectedOffer },
        { col: "AdjustedOiDate",   value: pdAdjustedOi },
        { col: "Risk",             value: pdRisk },
        { col: "AnticipationStartDate",      value: pdAnticipationStart },
        { col: "AnticipationExpirationDate", value: pdAnticipationExpiration },
        { col: "OIHeader",                   value: pdOiHeader },
        { col: "LastModifiedBy",   value: "Planner" }, // TODO: SSO display name
        { col: "LastModifiedDate", value: today },
      ], dwHeaders);
      setActionNote(`Saved ${selectedPipeline.opp.cep}`);
      await load();
    } catch (e: unknown) { setActionNote(`Error: ${String(e)}`); }
    setPdSaving(false);
  };

  const acceptActiveOffer = async () => {
    if (!selectedPipeline) return;
    setPdSaving(true);
    try {
      const newOffer = selectedPipeline.opp.offerBusinessId;
      await writeCellInTable("DemandWork", selectedPipeline.dw.rowIdx, "SelectedOfferID", newOffer, dwHeaders);
      // Optimistic update — clear the divergence banner immediately
      setPdSelectedOffer(newOffer);
      setSelectedPipeline(prev => prev ? {
        ...prev,
        dw: { ...prev.dw, selectedOfferId: newOffer },
        offerDiverged: false,
      } : null);
      setActionNote("Offer updated to match CEP.");
      await load();
    } catch (e: unknown) { setActionNote(`Error: ${String(e)}`); }
    setPdSaving(false);
  };

  const addComment = async () => {
    if (!selectedPipeline || !pdNewComment.trim()) return;
    setPdAddingComment(true);
    try {
      const ts = new Date().toISOString();
      await appendRowToTable("DemandWorkComments", [
        `${selectedPipeline.opp.cep}-${Date.now()}`,
        selectedPipeline.opp.cep,
        pdNewComment.trim(),
        "Planner", // TODO: SSO display name
        ts,
      ]);
      setPdNewComment("");
      setActionNote("Comment added.");
      await load();
    } catch (e: unknown) { setActionNote(`Error: ${String(e)}`); }
    setPdAddingComment(false);
  };

  const deleteComment = async (c: CommentRow) => {
    try {
      await deleteRowFromTable("DemandWorkComments", c.rowIdx);
      setActionNote("Comment deleted.");
      await load();
    } catch (e: unknown) { setActionNote(`Error: ${String(e)}`); }
  };

  const acknowledgeChangeLog = async (cl: ChangeLogEntry) => {
    try {
      const today = new Date().toISOString().slice(0, 10);
      await writeMultipleCells("ChangeLog", cl.rowIdx, [
        { col: "AcknowledgedBy",   value: "Planner" }, // TODO: SSO display name
        { col: "AcknowledgedDate", value: today },
      ], clHeaders);
      setActionNote("Event acknowledged.");
      await load();
    } catch (e: unknown) { setActionNote(`Error: ${String(e)}`); }
  };

  const togglePromote = async (opp: OppImportRow) => {
    setOdPromoting(true);
    try {
      const newVal = !opp.promote;
      await writeCellInTable("OppImport", opp.rowIdx, "Promote", newVal, oppHeaders);
      setActionNote(newVal ? "Flagged for promotion." : "Promotion flag cleared.");
      // Optimistic update for detail page if open
      if (selectedOpp?.rowIdx === opp.rowIdx) setSelectedOpp({ ...opp, promote: newVal });
      await load();
    } catch (e: unknown) { setActionNote(`Error: ${String(e)}`); }
    setOdPromoting(false);
  };

  const deleteFromPipeline = async (pr: PipelineRow) => {
    setDeleting(true);
    try {
      await deleteRowFromTable("DemandWork", pr.dw.rowIdx);
      // Unflag Promotion in OppImport — otherwise the deal reappears on the next SyncPipeline.
      if (pr.opp.rowIdx >= 0 && pr.opp.promote) {
        await writeCellInTable("OppImport", pr.opp.rowIdx, "Promote", false, oppHeaders);
      }

      // Auto-acknowledge any still-open ChangeLog events for this opportunity —
      // they're moot now that the deal is gone from the pipeline.
      const today = new Date().toISOString().slice(0, 10);
      const openForOpp = changeLogs.filter((cl) => cl.oppId === pr.opp.cep && !cl.acknowledgedBy);
      for (const cl of openForOpp) {
        await writeMultipleCells("ChangeLog", cl.rowIdx, [
          { col: "AcknowledgedBy",   value: "System (auto)" },
          { col: "AcknowledgedDate", value: today },
        ], clHeaders);
      }

      // Log the removal itself as a (pre-acknowledged) ChangeLog entry, so
      // there's a record of when/why the deal left the pipeline.
      if (clHeaders.length > 0) {
        await appendRowToTable("ChangeLog", buildRowFromHeaders(clHeaders, {
          ChangeLogId:        `${pr.opp.cep}-Removed-${Date.now()}`,
          OppId:               pr.opp.cep,
          OpportunityName:     pr.dw.opportunityName,
          EventType:           "RemovedFromPipeline",
          EventDetail:         "Opportunity removed from pipeline by planner.",
          OfferId:             pr.dw.selectedOfferId,
          FirstDetectedDate:   today,
          LastSeenDate:        today,
          AcknowledgedBy:      "System (auto)",
          AcknowledgedDate:    today,
        }));
      }

      setActionNote(`${pr.dw.oppId} removed from pipeline.`);
      setPendingDeleteCep(null);
      await load();
    } catch (e: unknown) { setActionNote(`Error: ${String(e)}`); }
    setDeleting(false);
  };

  /* ── equipment (opp detail) actions ──────────────────────────────────── */
  /**
   * Exclude an imported line. Appends an "Exclude" EquipmentAdjust row —
   * doesn't touch EquipmentImport or EquipmentResolved directly. The line
   * disappears from EquipmentResolved once SyncPipeline (or the nightly
   * SyncEquipment) re-runs the overlay + resolution.
   */
  const excludeEquipmentLine = async (offerBusinessId: string, line: EquipmentResolvedRow) => {
    setEqSaving(true); setEqActionNote("");
    try {
      const today = new Date().toISOString().slice(0, 10);
      await appendRowToTable("EquipmentAdjust", buildRowFromHeaders(eaHeaders, {
        AdjustmentId: `${offerBusinessId}-ADJ-${Date.now()}`,
        OfferBusinessId: line.offerBusinessId || offerBusinessId,
        AdjustType: "Exclude",
        TargetLineKey: line.lineKey,
        Superseded: false,
        CreatedDate: today,
      }));
      setPendingRemovalKeys(prev => new Set(prev).add(line.lineKey));
      setEqActionNote("Line excluded — run SyncPipeline in Excel (or wait for the nightly sync) to refresh the list.");
      await load();
    } catch (e: unknown) { setEqActionNote(`Error: ${String(e)}`); }
    setEqSaving(false);
  };

  /**
   * Deletes an EquipmentAdjust row outright — used for "Remove" on a
   * still-pending Estimate/Added row (nothing to resolve yet, so there's no
   * EquipmentResolved line to clean up — just delete the adjustment itself).
   */
  const deleteAdjustRow = async (adj: EquipmentAdjustRow, doneMessage: string) => {
    setEqSaving(true); setEqActionNote("");
    try {
      await deleteRowFromTable("EquipmentAdjust", adj.rowIdx);
      setEqActionNote(doneMessage);
      await load();
    } catch (e: unknown) { setEqActionNote(`Error: ${String(e)}`); }
    setEqSaving(false);
  };

  /**
   * Undoes an Exclude by deleting its EquipmentAdjust row. The underlying
   * line doesn't reappear in EquipmentResolved until SyncPipeline/
   * SyncEquipment re-runs, so it's tracked client-side as "pending restore"
   * and shown tinted blue in the meantime — same treatment as a freshly
   * added manual line.
   */
  const restoreExcludedLine = async (adj: EquipmentAdjustRow) => {
    setEqSaving(true); setEqActionNote("");
    try {
      await deleteRowFromTable("EquipmentAdjust", adj.rowIdx);
      // The excluded line may still be showing pending-removal red if this
      // undoes an Exclude click from the same session — clear that first.
      setPendingRemovalKeys(prev => {
        if (!prev.has(adj.targetLineKey)) return prev;
        const next = new Set(prev);
        next.delete(adj.targetLineKey);
        return next;
      });
      setPendingRestoreKeys(prev => new Set(prev).add(adj.targetLineKey));
      setEqActionNote("Exclusion removed — run SyncPipeline in Excel (or wait for the nightly sync) to refresh the list.");
      await load();
    } catch (e: unknown) { setEqActionNote(`Error: ${String(e)}`); }
    setEqSaving(false);
  };

  const removePendingLine = (adj: EquipmentAdjustRow) =>
    deleteAdjustRow(adj, "Line removed.");

  /**
   * Adds a planner-picked ProductArchetype line. Plant is NOT picked here —
   * it's resolved by SyncEquipment.ts/SyncPipeline.ts from the same
   * ArchetypeMaster-candidate + PlantRules tier-walk an imported line goes
   * through (Ch.4 handoff redesign). Qty is always 1 for a new line, per
   * the one-row-per-physical-unit convention.
   *
   * Estimate vs Added is not a planner choice — it's set here based on
   * whether this offer currently has any EquipmentImport rows at all,
   * matching the rule documented in SyncEquipment.ts.
   */
  const addManualEquipmentLine = async (offerBusinessId: string) => {
    if (!eqNewProductArchetype) return;
    setEqSaving(true); setEqActionNote("");
    try {
      const today = new Date().toISOString().slice(0, 10);
      const hasImportLines = equipmentImportLines.some(l => l.offerBusinessId === offerBusinessId);
      await appendRowToTable("EquipmentAdjust", buildRowFromHeaders(eaHeaders, {
        AdjustmentId: `${offerBusinessId}-ADJ-${Date.now()}`,
        OfferBusinessId: offerBusinessId,
        AdjustType: hasImportLines ? "Estimate" : "Added",
        ProductArchetype: eqNewProductArchetype,
        Qty: 1,
        Superseded: false,
        CreatedDate: today,
      }));
      setEqNewFamily("");
      setEqNewProductArchetype("");
      setEqAddingLine(false);
      setEqActionNote("Line added — run SyncPipeline in Excel (or wait for the nightly sync) to resolve it.");
      await load();
    } catch (e: unknown) { setEqActionNote(`Error: ${String(e)}`); }
    setEqSaving(false);
  };

  /** Removes a planner-added Estimate/Added line entirely (not an exclude — there's no underlying import row to keep). */
  const removeManualEquipmentLine = async (line: EquipmentResolvedRow) => {
    const adj = equipmentAdjust.find(a => a.adjustmentId === line.lineKey);
    if (!adj) { setEqActionNote("Couldn't find the underlying adjustment row — try reloading."); return; }
    setEqSaving(true); setEqActionNote("");
    try {
      await deleteRowFromTable("EquipmentAdjust", adj.rowIdx);
      setPendingRemovalKeys(prev => new Set(prev).add(line.lineKey));
      setEqActionNote("Line removed.");
      await load();
    } catch (e: unknown) { setEqActionNote(`Error: ${String(e)}`); }
    setEqSaving(false);
  };

  /**
   * "Keep" on a LeftCep line — the planner wants this equipment to stay on
   * the offer even though CEP no longer serves it. Clones it into a real
   * planner-owned Added/Estimate line (same ProductArchetype, or the legacy
   * ArchetypeCode+Plant fallback if that lookup comes up empty) and
   * acknowledges the ChangeLog event. Once SyncPipeline/SyncEquipment next
   * runs, this becomes a genuine Added-origin line and the script stops
   * carrying the orphaned Import-origin one forward.
   */
  const keepLeftCepLine = async (line: EquipmentResolvedRow) => {
    setEqSaving(true); setEqActionNote("");
    try {
      const today = new Date().toISOString().slice(0, 10);
      const productArchetype = archetypeMasterByCode.get(line.resolvedArchetype)?.productArchetype ?? "";
      const hasImportLines = equipmentImportLines.some(l => l.offerBusinessId === line.offerBusinessId);
      await appendRowToTable("EquipmentAdjust", buildRowFromHeaders(eaHeaders, {
        AdjustmentId: `${line.offerBusinessId}-ADJ-${Date.now()}`,
        OfferBusinessId: line.offerBusinessId,
        AdjustType: hasImportLines ? "Estimate" : "Added",
        ProductArchetype: productArchetype,
        ArchetypeCode: productArchetype ? "" : line.resolvedArchetype,
        Plant: productArchetype ? "" : line.resolvedPlant,
        Qty: line.qty || 1,
        Superseded: false,
        CreatedDate: today,
      }));
      const cl = changeLogs.find(c => c.changeLogId === `${line.lineKey}-LeftCep` && !c.acknowledgedBy);
      if (cl) {
        await writeMultipleCells("ChangeLog", cl.rowIdx, [
          { col: "AcknowledgedBy",   value: "Planner (kept)" },
          { col: "AcknowledgedDate", value: today },
        ], clHeaders);
      }
      setEqActionNote("Kept — run SyncPipeline in Excel (or wait for the nightly sync) to make it permanent.");
      await load();
    } catch (e: unknown) { setEqActionNote(`Error: ${String(e)}`); }
    setEqSaving(false);
  };

  /**
   * "Remove" on a LeftCep line — the planner agrees it's gone. Just
   * acknowledges the ChangeLog event; no adjust row is written, so the
   * script stops carrying the line forward on its next run.
   */
  const removeLeftCepLine = async (line: EquipmentResolvedRow) => {
    setEqSaving(true); setEqActionNote("");
    try {
      const today = new Date().toISOString().slice(0, 10);
      const cl = changeLogs.find(c => c.changeLogId === `${line.lineKey}-LeftCep` && !c.acknowledgedBy);
      if (cl) {
        await writeMultipleCells("ChangeLog", cl.rowIdx, [
          { col: "AcknowledgedBy",   value: "Planner (removed)" },
          { col: "AcknowledgedDate", value: today },
        ], clHeaders);
      }
      setEqActionNote("Acknowledged — it'll drop off the list once SyncPipeline (or the nightly sync) next runs.");
      await load();
    } catch (e: unknown) { setEqActionNote(`Error: ${String(e)}`); }
    setEqSaving(false);
  };

  /* ── CEP snapshot pairs (renamed from "CRM snapshot") ────────────────── */
  const oppCepPairs = (opp: OppImportRow): [string, string][] => [
    ["Customer",         opp.customer],
    ["Region",           opp.region],
    ["Sub-region",       opp.subRegion],
    ["Country",          opp.country],
    ["Mkt portfolio",    opp.marketPortfolio],
    ["Category L1",      opp.marketCategoryL1],
    ["Portfolio L2",     opp.portfolioL2],
    ["Value",            opp.value ? `€ ${(Number(opp.value) / 1_000_000).toFixed(1)}M` : "—"],
    ["CEP stage",        opp.stage],
    ["Forecast status",  opp.forecastStatus],
    ["Score band",       opp.scoreBand],
    ["Mkt probability",  opp.marketProbability],
    ["Grp probability",  opp.groupProbability],
    ["Active offer",     opp.offerBusinessId || opp.activeOfferName],
    ["Exp. OI date",     formatDate(opp.expectedOiDate)],
    ["Last modified",    formatDate(opp.modifiedOn)],
  ];

  /* ─── RENDER ────────────────────────────────────────────────────────────── */
  return (
    <div style={s.shell}>

      {/* ── Header ── */}
      <div style={s.header}>
        {view !== "tabs" && (
          <button style={s.backBtn} onClick={goBack} aria-label="Back">← Back</button>
        )}
        <div style={s.headerTitle}>
          Demand Tool <span style={s.version}>{APP_VERSION}</span>
        </div>
        <div style={s.headerRight}>
          {lastLoaded && <span style={s.syncTime}>Synced {lastLoaded}</span>}
          <button style={s.reloadBtn} onClick={load} disabled={loading} title="Reload" aria-label="Reload">
            {loading ? "…" : "↻"}
          </button>
        </div>
      </div>

      {/* ── Tab bar ── */}
      {view === "tabs" && (
        <div style={s.tabBar}>
          {(["pipeline", "opp"] as TabId[]).map(t => (
            <button key={t} style={tab === t ? s.tabActive : s.tab} onClick={() => setTab(t)}>
              {t === "pipeline" ? "Pipeline" : "Opportunities"}
              {t === "pipeline" && totalOpenCl > 0 && <span style={s.clBadge}>{totalOpenCl}</span>}
            </button>
          ))}
        </div>
      )}

      {/* ── Status bars ── */}
      {loadMsg && !errMsg && <div style={s.statusBar}>{loadMsg}</div>}
      {errMsg             && <div style={s.errorBar}>{errMsg}</div>}
      {actionNote         && <div style={s.actionBar}>{actionNote}</div>}

      {/* ══════════════════════════════════════════════
          PIPELINE LIST
      ══════════════════════════════════════════════ */}
      {view === "tabs" && tab === "pipeline" && (
        <div style={s.content}>
          {/* Filter bar */}
          <div style={s.filterBar}>
            <div style={s.searchBox}>
              <span style={{ fontSize: 15, color: "#94a3b8", flexShrink: 0 }}>⌕</span>
              <input style={s.searchInput} placeholder="Search CEP, name, customer…"
                value={pipeSearch} onChange={e => setPipeSearch(e.target.value)} />
              {pipeSearch && <button style={s.clearBtn} onClick={() => setPipeSearch("")}>✕</button>}
            </div>
            <div style={s.chips}>
              <MultiSelect label="Region" options={pipeRegions} selected={pipeFilterRegion} onChange={setPipeFilterRegion} />
              <MultiSelect label="Sub-region" options={pipeSubRegions} selected={pipeFilterSubRegion} onChange={setPipeFilterSubRegion} />
              <MultiSelect label="Portfolio" options={pipePortfolios} selected={pipeFilterPortfolio} onChange={setPipeFilterPortfolio} />
              <MultiSelect label="Forecast status" options={pipeForecastStatuses} selected={pipeFilterForecastStatus} onChange={setPipeFilterForecastStatus} />
              <MultiSelect label="Deal status" options={DEAL_STATUSES} selected={pipeFilterDealStatus} onChange={setPipeFilterDealStatus}
                extraOptions={[{ label: "— No status —", value: "" }]} />
              <select style={s.filterSel} value={pipeFilterChangeLog} onChange={e => setPipeFilterChangeLog(e.target.value)}>
                <option value="">Change Log</option>
                <option value="open">Open Events</option>
                <option value="none">No Events</option>
              </select>
            </div>
          </div>

          {/* Summary */}
          <div style={s.summaryBar}>
            <span style={s.summaryLabel}>
              {filteredPipeline.length} of {pipelineRows.length} deals
            </span>
            {filteredPipeline.length > 0 && <>
              <span style={s.dot}>·</span>
              <span style={s.summaryValue}>
                € {(filteredPipeline.reduce((a,pr) => a + (pr.opp.value ?? 0), 0) / 1_000_000).toFixed(1)}M
              </span>
            </>}
            {totalOpenCl > 0 && <>
              <span style={s.dot}>·</span>
              <span style={s.clWarning}>{totalOpenCl} change log event{totalOpenCl !== 1 ? "s" : ""}</span>
            </>}
          </div>

          {filteredPipeline.length === 0 ? (
            <div style={s.empty}>
              {pipelineRows.length === 0
                ? "No deals in pipeline yet.\nFlag opportunities and run SyncPipeline in Excel."
                : "No deals match the current filters."}
            </div>
          ) : (
            <div style={s.tableArea}>
              <div style={{ minWidth: 1200 }}>
                <div style={{ ...s.pipelineGrid, ...s.tableHead }}>
                  {["CEP","Opp. name","Customer","Mkt portfolio","Region","Sub-region","Forecast status","Deal status","Offer","Eff. OI",""].map((h,i) => (
                    <span key={i} style={s.th}>{h}</span>
                  ))}
                </div>
                {filteredPipeline.map((pr, i) => {
                  const bg       = (pr.isLeftSync || pr.anticipationOverdue) ? "#fffbeb" : i % 2 === 0 ? "#fff" : "#f8fafc";
                  const clCount  = openClByOpp.get(pr.dw.oppId) ?? 0;
                  const shortType = pr.dw.dealType.replace("Anticipation — ","Antic. ").replace("From OI to KOM","OI→KOM");

                  if (pendingDeleteCep === pr.dw.oppId) {
                    return (
                      <div key={pr.opp.cep} style={{ ...s.pipelineGrid, background: "#fef2f2", padding: "6px 14px" }}>
                        <span style={s.deleteConfirmRow}>
                          Remove <b>{pr.dw.oppId}</b> from pipeline?
                          <button style={s.confirmDeleteBtn} onClick={() => deleteFromPipeline(pr)} disabled={deleting}>
                            {deleting ? "Removing…" : "Confirm"}
                          </button>
                          <button style={s.cancelBtn} onClick={() => setPendingDeleteCep(null)}>Cancel</button>
                        </span>
                      </div>
                    );
                  }

                  return (
                    <div key={pr.opp.cep}
                      style={{ ...s.pipelineGrid, background: bg, cursor: "pointer" as const, padding: "6px 14px" }}
                      onClick={() => openPipelineDetail(pr)}
                      onMouseEnter={e => { (e.currentTarget as HTMLDivElement).style.background = "#eff6ff"; }}
                      onMouseLeave={e => { (e.currentTarget as HTMLDivElement).style.background = bg; }}>
                      <span style={s.tdMono}>{pr.opp.cep}</span>
                      <span style={{ ...s.td, ...s.ell }} title={pr.opp.opportunityName}>{pr.opp.opportunityName}</span>
                      <span style={{ ...s.td, ...s.ell }}>{pr.opp.customer}</span>
                      <span style={{ ...s.td, ...s.ell }}>{pr.opp.marketPortfolio}</span>
                      <span style={s.td}>{pr.opp.region}</span>
                      <span style={s.td}>{pr.opp.subRegion}</span>
                      <span style={{ ...s.td, ...s.ell }}>{pr.opp.forecastStatus || "—"}</span>
                      <span style={{ ...s.td, ...s.ell }}>{shortType}</span>
                      <span style={{ ...s.td, ...s.ell }}>
                        {pr.dw.selectedOfferId || "—"}
                        {pr.offerDiverged && <span style={s.offerBadge} title="Active offer in CEP differs — open deal to review">?</span>}
                      </span>
                      <span style={s.td}>{formatDate(pr.effectiveOiDate)}</span>
                      <span style={s.td}>
                        {pr.dw.risk && <span title="Risk flagged" style={{ color: "#ca8a04", marginRight: 3 }}>⚠</span>}
                        {pr.anticipationOverdue && (
                          <span title={`Anticipation expired ${formatDate(pr.dw.anticipationExpirationDate)}`}
                            style={{ color: "#854d0e", marginRight: 3 }}>⏰</span>
                        )}
                        {(clCount > 0 || pr.isLeftSync) && (
                          <span style={s.clBadge} title={
                            pr.isLeftSync
                              ? `Left CEP sync — review or remove${clCount > 0 ? ` · ${clCount} open change log event${clCount !== 1 ? "s" : ""}` : ""}`
                              : `${clCount} open change log event${clCount !== 1 ? "s" : ""}`
                          }>
                            {clCount > 0 ? clCount : "•"}
                          </span>
                        )}
                        <span style={s.deleteRowBtn}
                          onClick={e => { e.stopPropagation(); setPendingDeleteCep(pr.dw.oppId); }}
                          onMouseEnter={e => { (e.currentTarget as HTMLSpanElement).style.color = "#dc2626"; }}
                          onMouseLeave={e => { (e.currentTarget as HTMLSpanElement).style.color = "#cbd5e1"; }}
                          title="Remove from pipeline">✕</span>
                      </span>
                    </div>
                  );
                })}
              </div>
            </div>
          )}
        </div>
      )}

      {/* ══════════════════════════════════════════════
          OPPORTUNITIES LIST
      ══════════════════════════════════════════════ */}
      {view === "tabs" && tab === "opp" && (
        <div style={s.content}>
          <div style={s.filterBar}>
            <div style={s.searchBox}>
              <span style={{ fontSize: 15, color: "#94a3b8", flexShrink: 0 }}>⌕</span>
              <input style={s.searchInput} placeholder="Search by CEP, customer, name…"
                value={oppSearch} onChange={e => setOppSearch(e.target.value)} />
              {oppSearch && <button style={s.clearBtn} onClick={() => setOppSearch("")}>✕</button>}
            </div>
            <div style={s.chips}>
              <button style={oppFlaggedOnly ? s.chipActive : s.chip}
                onClick={() => setOppFlaggedOnly(!oppFlaggedOnly)}>
                Flagged only
              </button>
              <MultiSelect label="Forecast status" options={allForecastStatuses} selected={oppFilterForecastStatus} onChange={setOppFilterForecastStatus} />
              <MultiSelect label="Region" options={allRegions} selected={oppFilterRegion} onChange={setOppFilterRegion} />
              <MultiSelect label="Portfolio" options={allPortfolios} selected={oppFilterPortfolio} onChange={setOppFilterPortfolio} />
              <MultiSelect label="Score band" options={allBands} selected={oppFilterBand} onChange={setOppFilterBand} />
            </div>
            <div style={s.resultCount}>
              {filteredOpps.length} of {oppImport.length} opportunities
              {filteredOpps.length > 200 && " · showing first 200"}
            </div>
          </div>

          <div style={s.tableArea}>
            <div style={{ minWidth: 1200 }}>
              <div style={{ ...s.oppGrid, ...s.tableHead }}>
                {["","CEP","Opp. name","Customer","Mkt portfolio","Region","Sub-region","Status","Score Band","MP%","GP%","Exp. OI"].map((h,i) => (
                  <span key={i} style={s.th}>{h}</span>
                ))}
              </div>
              {filteredOpps.slice(0, 200).map((opp, i) => {
                const inPipeline = dwByOppId.get(opp.cep)?.inPipeline ?? false;
                const promoted   = opp.promote ; //|| inPipeline
                const bg         = promoted ? "#eff6ff" : i % 2 === 0 ? "#fff" : "#f8fafc";
                return (
                  <div key={opp.rowIdx}
                    style={{ ...s.oppGrid, background: bg, cursor: "pointer" as const, padding: "5px 14px" }}
                    onClick={() => openOppDetail(opp)}
                    onMouseEnter={e => { (e.currentTarget as HTMLDivElement).style.background = "#eff6ff"; }}
                    onMouseLeave={e => { (e.currentTarget as HTMLDivElement).style.background = bg; }}>
                    <span style={{ cursor: "pointer" as const, display: "flex" as const, alignItems: "center" as const }}
                      onClick={e => { e.stopPropagation(); togglePromote(opp); }}
                      title={opp.promote ? "Remove promotion flag" : "Flag for promotion"}>
                      {promoted
                        ? <span style={s.cbChecked}>✓</span>
                        : <span style={s.cbEmpty} />}
                    </span>
                    <span style={s.tdMono}>{opp.cep}</span>
                    <span style={{ ...s.td, ...s.ell }} title={opp.opportunityName}>{opp.opportunityName}</span>
                    <span style={{ ...s.td, ...s.ell }}>{opp.customer}</span>
                    <span style={{ ...s.td, ...s.ell }}>{opp.marketPortfolio}</span>
                    <span style={s.td}>{opp.region}</span>
                    <span style={s.td}>{opp.subRegion}</span>
                    <span style={{ ...s.td, ...s.ell }}>{opp.stage}</span>
                    {/* <span style={{ ...s.td, fontWeight: 700, color: BAND_COLOR[opp.scoreBand?.charAt(0)] ?? "#94a3b8" }}> */}
                    <span style={{ ...s.td, ...s.ell }}>
                      {opp.scoreBand}
                    </span>
                    <span style={s.td}>{opp.marketProbability}</span>
                    <span style={s.td}>{opp.groupProbability}</span>
                    <span style={s.td}>{formatDate(opp.expectedOiDate)}</span>
                  </div>
                );
              })}
              {filteredOpps.length > 200 && (
                <div style={{ padding:"10px 14px", textAlign:"center" as const, fontSize:14, color:"#94a3b8", background:"#fff" }}>
                  Showing first 200 of {filteredOpps.length} — apply filters to narrow
                </div>
              )}
            </div>
          </div>
        </div>
      )}

      {/* ══════════════════════════════════════════════
          PIPELINE DEAL DETAIL
      ══════════════════════════════════════════════ */}
      {view === "pipeline-detail" && selectedPipeline && (() => {
        const pr              = selectedPipeline;
        const dealComments    = comments.filter(c => c.oppId === pr.opp.cep);
        const dealChangeLogs  = changeLogs.filter(c => c.oppId === pr.opp.cep);
        const openClCount     = dealChangeLogs.filter(c => !c.acknowledgedBy).length;
        // Equipment keys off the planner's SelectedOfferID (the frozen-baseline
        // offer this deal is being planned against), falling back to the live
        // CEP active offer if SelectedOfferID hasn't been set yet.
        const eqOfferId       = pr.dw.selectedOfferId || pr.opp.offerBusinessId;
        const eqAllLines      = equipmentByOppId.get(pr.dw.oppId) ?? [];
        // Split by offer: the selected offer is what the planner actively
        // manages here; when offerDiverged, CEP's live active offer has its
        // own freshly-resolved lines too — shown read-only alongside, so the
        // planner can compare before deciding whether to accept it.
        const eqLines         = eqAllLines.filter(l => l.offerBusinessId === eqOfferId);
        const eqExcluded      = excludedAdjustByOffer.get(eqOfferId) ?? [];
        const eqNewOfferLines = pr.offerDiverged ? eqAllLines.filter(l => l.offerBusinessId === pr.opp.offerBusinessId) : [];
        return (
          <div style={s.detailPage}>
            <div style={s.detailId}>
              <div style={s.detailCep}>{pr.opp.cep} — {pr.opp.opportunityName}</div>
              <div style={s.detailSub}>{pr.opp.customer} · {pr.dw.dealType || "—"}</div>
              {pr.anticipationOverdue && (
                <div style={s.overdueBanner}>
                  ⏰ Anticipation expired {formatDate(pr.dw.anticipationExpirationDate)} — review deal status or update the expiration date.
                </div>
              )}
            </div>

            <div style={s.detailBody}>
              {/* CEP snapshot — live data or DW snapshot for LeftSync deals */}
            <CollapsibleSection
              title="CEP snapshot"
              subtitle={pr.isLeftSync ? "snapshot · deal left CEP sync" : "as of last sync"}>
              {pr.isLeftSync && (
                <div style={s.leftSyncNote}>
                  This deal is no longer in the CEP sync feed — it may have been Won,
                  Abandoned, or Closed. The data below is the last known snapshot.{" "}
                  <span style={{ fontWeight: 600 }}>Review the deal and remove it if it is no longer needed.</span>
                </div>
              )}
              <CrmGrid pairs={oppCepPairs(pr.opp)} />
            </CollapsibleSection>

              {/* Planning */}
              <CollapsibleSection title="Planning">
                <div style={s.twoColForm}>
                  <div style={s.field}>
                    <label style={s.fieldLbl}>Deal status</label>
                    <select style={s.sel} value={pdDealType} onChange={e => setPdDealType(e.target.value)}>
                      <option value="">— select —</option>
                      {DEAL_STATUSES.map(dt => <option key={dt}>{dt}</option>)}
                    </select>
                  </div>
                  <div style={s.field}>
                    <label style={s.fieldLbl}>Selected offer</label>
                    <input style={{ ...s.inp, borderColor: pr.offerDiverged ? "#fde68a" : "#cbd5e1", background: pr.offerDiverged ? "#fffbeb" : "#fff" }}
                      value={pdSelectedOffer} onChange={e => setPdSelectedOffer(e.target.value)} />
                  </div>
                </div>
                {pr.offerDiverged && (
                  <div style={s.offerDivergeBanner}>
                    ⚠ Active in CEP: {pr.opp.offerBusinessId} —{" "}
                    <span style={s.acceptLink} onClick={acceptActiveOffer}>accept</span>
                  </div>
                )}
                <div style={s.twoColForm}>
                  <div style={s.field}>
                    <label style={s.fieldLbl}>CEP OI date</label>
                    <div style={s.readonlyVal}>{formatDate(pr.dw.crmOiDate || pr.opp.expectedOiDate)}</div>
                  </div>
                  <div style={s.field}>
                    <label style={s.fieldLbl}>Adjusted OI <span style={s.fieldHint}>(optional)</span></label>
                    <input type="date" style={s.inp} value={pdAdjustedOi} onChange={e => setPdAdjustedOi(e.target.value)} />
                  </div>
                </div>
                {ANTICIPATION_TYPES.includes(pdDealType) && (
                  <div style={s.twoColForm}>
                    <div style={s.field}>
                      <label style={s.fieldLbl}>Anticipation start <span style={s.fieldHint}>(recommended)</span></label>
                      <input type="date" style={s.inp} value={pdAnticipationStart}
                        onChange={e => setPdAnticipationStart(e.target.value)} />
                    </div>
                    <div style={s.field}>
                      <label style={s.fieldLbl}>Anticipation expiration <span style={s.fieldHint}>(recommended)</span></label>
                      <input type="date" style={s.inp} value={pdAnticipationExpiration}
                        onChange={e => setPdAnticipationExpiration(e.target.value)} />
                    </div>
                  </div>
                )}
                {ANTICIPATION_TYPES.includes(pdDealType)
                  && !!pdAnticipationStart && !!pdAnticipationExpiration
                  && pdAnticipationExpiration <= pdAnticipationStart && (
                  <div style={s.fieldWarning}>⚠ Expiration date should be after the start date.</div>
                )}
                {OI_HEADER_TYPES.includes(pdDealType) && (
                  <div style={s.field}>
                    <label style={s.fieldLbl}>OI Header <span style={s.fieldHint}>(recommended · e.g. H1000000)</span></label>
                    <input style={s.inp} value={pdOiHeader} placeholder="H1000000"
                      onChange={e => setPdOiHeader(e.target.value)} />
                    {!!pdOiHeader && !OI_HEADER_PATTERN.test(pdOiHeader) && (
                      <div style={s.fieldWarning}>⚠ Expected format: a letter, then "1", then 6 digits (e.g. H1000000).</div>
                    )}
                  </div>
                )}
                <div style={{ ...s.field, marginTop: 8 }}>
                  <label style={s.checkLabel}>
                    <input type="checkbox" checked={pdRisk} onChange={e => setPdRisk(e.target.checked)}
                      style={{ width: 16, height: 16, cursor: "pointer" as const }} />
                    Risk flag
                  </label>
                </div>
                <div style={s.saveRow}>
                  <span style={s.lastSaved}>
                    {pr.dw.lastModifiedBy
                      ? `Last saved: ${pr.dw.lastModifiedBy} · ${formatDate(pr.dw.lastModifiedDate)}`
                      : "Not yet saved"}
                  </span>
                  <button style={s.saveBtn} onClick={savePlanningFields} disabled={pdSaving}>
                    {pdSaving ? "Saving…" : "Save"}
                  </button>
                </div>
              </CollapsibleSection>

              {/* Equipment — all edit actions (exclude / override qty / add / remove) live here, not on the Opportunity side */}
              {!!eqOfferId && (
                <CollapsibleSection title="Equipment lines" subtitle={`${eqLines.length} line${eqLines.length === 1 ? "" : "s"}`}>
                  {eqActionNote && <div style={s.eqActionNote}>{eqActionNote}</div>}
                  {pr.offerDiverged && (
                    <div style={s.eqOfferLabel}>Selected offer — {eqOfferId}</div>
                  )}
                  <EquipmentTable
                    lines={eqLines}
                    excluded={eqExcluded}
                    pendingManual={pendingManualByOffer.get(eqOfferId) ?? []}
                    pendingRestore={pendingRestoreByOffer.get(eqOfferId) ?? []}
                    importLineByKey={importLineByKey}
                    l5DescByCode={l5DescByCode}
                    archetypeMasterByCode={archetypeMasterByCode}
                    adjustByAdjustmentId={adjustByAdjustmentId}
                    pendingRemovalKeys={pendingRemovalKeys}
                    editable={true}
                    saving={eqSaving}
                    onExclude={(line) => excludeEquipmentLine(eqOfferId, line)}
                    onRestore={restoreExcludedLine}
                    onRemoveManual={removeManualEquipmentLine}
                    onRemovePending={removePendingLine}
                    onKeepLeftCep={keepLeftCepLine}
                    onRemoveLeftCep={removeLeftCepLine}
                  />

                  {pr.offerDiverged && (
                    <div style={s.eqNewOfferBlock}>
                      <div style={s.eqOfferLabel}>
                        Active in CEP (preview) — {pr.opp.offerBusinessId}{" "}
                        <span style={s.acceptLink} onClick={acceptActiveOffer}>accept this offer</span>
                      </div>
                      <div style={s.fieldHint}>
                        Read-only until accepted. Accepting switches "Selected offer" above and this becomes the
                        editable list on the next sync.
                      </div>
                      <EquipmentTable
                        lines={eqNewOfferLines}
                        excluded={[]}
                        pendingManual={[]}
                        pendingRestore={[]}
                        importLineByKey={importLineByKey}
                        l5DescByCode={l5DescByCode}
                        archetypeMasterByCode={archetypeMasterByCode}
                        adjustByAdjustmentId={adjustByAdjustmentId}
                        pendingRemovalKeys={EMPTY_KEY_SET}
                        editable={false}
                        saving={false}
                        onExclude={() => {}}
                        onRestore={() => {}}
                        onRemoveManual={() => {}}
                        onRemovePending={() => {}}
                        onKeepLeftCep={() => {}}
                        onRemoveLeftCep={() => {}}
                      />
                    </div>
                  )}

                  {eqAddingLine ? (
                    <div style={s.eqAddForm}>
                      <div style={s.twoColForm}>
                        <div style={s.field}>
                          <label style={s.fieldLbl}>Family</label>
                          <select style={s.sel} value={eqNewFamily}
                            onChange={e => { setEqNewFamily(e.target.value); setEqNewProductArchetype(""); }}>
                            <option value="">— select —</option>
                            {familyOptions.map(f => <option key={f} value={f}>{f}</option>)}
                          </select>
                        </div>
                        <div style={s.field}>
                          <label style={s.fieldLbl}>Product archetype</label>
                          <select style={s.sel} value={eqNewProductArchetype} disabled={!eqNewFamily}
                            onChange={e => setEqNewProductArchetype(e.target.value)}>
                            <option value="">{eqNewFamily ? "— select —" : "Pick a family first"}</option>
                            {(productArchetypesByFamily.get(eqNewFamily) ?? []).map(pa => <option key={pa} value={pa}>{pa}</option>)}
                          </select>
                        </div>
                      </div>
                      <div style={s.fieldHint}>Plant is resolved automatically once SyncPipeline runs — you don't pick it.</div>
                      <div style={s.eqAddFormActions}>
                        <button style={s.addBtn} disabled={eqSaving || !eqNewProductArchetype}
                          onClick={() => addManualEquipmentLine(eqOfferId)}>
                          {eqSaving ? "Adding…" : "Add line"}
                        </button>
                        <button style={s.cancelBtn} onClick={() => { setEqAddingLine(false); setEqNewFamily(""); setEqNewProductArchetype(""); }}>Cancel</button>
                      </div>
                    </div>
                  ) : (
                    <button style={s.addBtn} onClick={() => setEqAddingLine(true)}>+ Add equipment line</button>
                  )}
                </CollapsibleSection>
              )}

              {/* Comments */}
              <CollapsibleSection title={`Comments (${dealComments.length})`}>
                {dealComments.length === 0 && <div style={s.emptySmall}>No comments yet.</div>}
                {dealComments.map(c => (
                  <div key={c.rowIdx} style={s.commentCard}>
                    <div style={s.commentMeta}>{c.author} · {c.timestamp?.slice(0,10) || "—"}</div>
                    <div style={s.commentText}>{c.commentText}</div>
                    <button style={s.deleteLink} onClick={() => deleteComment(c)}>Delete</button>
                  </div>
                ))}
                <textarea style={s.textarea} placeholder="Add a comment…" rows={3}
                  value={pdNewComment} onChange={e => setPdNewComment(e.target.value)} />
                <button style={s.addBtn} onClick={addComment}
                  disabled={pdAddingComment || !pdNewComment.trim()}>
                  {pdAddingComment ? "Adding…" : "Add comment"}
                </button>
              </CollapsibleSection>

              {/* Change log */}
              <CollapsibleSection title="Change log"
                badge={openClCount > 0 ? <span style={s.clOpenBadge}>{openClCount} open</span> : undefined}>
                {dealChangeLogs.length === 0 && <div style={s.emptySmall}>No events.</div>}
                {dealChangeLogs.map(cl => (
                  <div key={cl.rowIdx} style={cl.acknowledgedBy ? s.clCardAck : s.clCard}>
                    <div style={s.clCardHeader}>
                      <span style={s.clEventBadge}>{cl.eventType}</span>
                      <span style={s.clDate}>{formatDate(cl.firstDetectedDate)}</span>
                    </div>
                    <div style={s.clDetail}>{cl.eventDetail}</div>
                    {!cl.acknowledgedBy
                      ? <button style={s.ackBtn} onClick={() => acknowledgeChangeLog(cl)}>Acknowledge</button>
                      : <div style={s.ackNote}>Acknowledged by {cl.acknowledgedBy} · {formatDate(cl.acknowledgedDate)}</div>}
                  </div>
                ))}
              </CollapsibleSection>
            </div>
          </div>
        );
      })()}

      {/* ══════════════════════════════════════════════
          OPPORTUNITY DETAIL
      ══════════════════════════════════════════════ */}
      {view === "opp-detail" && selectedOpp && (() => {
        const opp        = selectedOpp;
        const dw         = dwByOppId.get(opp.cep);
        const inPipeline = dw?.inPipeline ?? false;
        const isPromoted = opp.promote;
        // View-only here — exclude / override qty / add / remove all live on
        // the Pipeline detail side (see PIPELINE DEAL DETAIL below). Also only
        // shown at all when there's an offer to have equipment against.
        const eqLines    = equipmentByOppId.get(opp.cep) ?? [];
        const hasOffer   = !!opp.offerBusinessId;

        return (
          <div style={s.detailPage}>
            <div style={s.detailId}>
              <div style={s.detailCep}>{opp.cep} — {opp.opportunityName}</div>
              <div style={s.detailSub}>{opp.customer} · {opp.stage} · Band {opp.scoreBand || "—"}</div>
            </div>

            <div style={s.detailBody}>
              {/* Promote card */}
              <div style={s.promoteCard}>

              <div style={{ flex: 1 }}>
                <div style={s.promoteTitle}>Pipeline promotion</div>
                <div style={s.promoteSub}>
                  {inPipeline
                    ? "Already in the pipeline (DemandWork entry exists)."
                    : "Flag to include in the next SyncPipeline run."}
                </div>
              </div>
              {!inPipeline && (
                <button style={isPromoted ? s.promoteBtnOn : s.promoteBtnOff}
                  onClick={() => togglePromote(opp)} disabled={odPromoting}>
                  {odPromoting ? "…" : isPromoted ? "Flagged ✓" : "Flag for promotion"}
                </button>
              )}
              {inPipeline && <span style={s.inPipelineBadge}>In pipeline</span>}
            </div>
            {isPromoted && !inPipeline && (
              <div style={s.promoteNote}>
                ✓ Flagged — run SyncPipeline in Excel to create the pipeline entry.
              </div>
            )}

            {/* CEP snapshot */}
            <CollapsibleSection title="CEP snapshot" subtitle="read-only · as of last sync">
              <CrmGrid pairs={oppCepPairs(opp)} />
            </CollapsibleSection>

            {/* Equipment — view-only; edits happen in Pipeline detail */}
            {hasOffer && (
              <CollapsibleSection title="Equipment lines" subtitle={`${eqLines.length} line${eqLines.length === 1 ? "" : "s"}`}>
                <EquipmentTable
                  lines={eqLines}
                  excluded={[]}
                  pendingManual={[]}
                  pendingRestore={[]}
                  importLineByKey={importLineByKey}
                  l5DescByCode={l5DescByCode}
                  archetypeMasterByCode={archetypeMasterByCode}
                  adjustByAdjustmentId={adjustByAdjustmentId}
                  pendingRemovalKeys={EMPTY_KEY_SET}
                  editable={false}
                  saving={false}
                  onExclude={() => {}}
                  onRestore={() => {}}
                  onRemoveManual={() => {}}
                  onRemovePending={() => {}}
                  onKeepLeftCep={() => {}}
                  onRemoveLeftCep={() => {}}
                />
              </CollapsibleSection>
            )}
          </div>
          </div>
        );
      })()}

    </div>
  );
}

/* ─── styles ──────────────────────────────────────────────────────────────── */
const s = {
  shell:        { fontFamily: "'Segoe UI',system-ui,sans-serif", fontSize: 16, color: "#1e293b", background: "#f8fafc",height: "100%", minHeight: "0", display: "flex" as const, flexDirection: "column" as const, overflow: "hidden" as const },
  header:       { background: "#0f2942", padding: "10px 14px", display: "flex" as const, alignItems: "center" as const, gap: 8, flexShrink: 0 },
  backBtn:      { background: "none", border: "1px solid #334d6e", borderRadius: 6, padding: "4px 10px", color: "#94a3b8", cursor: "pointer" as const, fontSize: 14, fontFamily: "inherit", whiteSpace: "nowrap" as const, flexShrink: 0 },
  headerTitle:  { color: "#fff", fontSize: 22, fontWeight: 700, letterSpacing: "-0.01em", flex: 1 },
  version:      { fontSize: 13, fontWeight: 400, color: "#94a3b8", marginLeft: 6 },
  headerRight:  { display: "flex" as const, alignItems: "center" as const, gap: 8, flexShrink: 0 },
  syncTime:     { fontSize: 13, color: "#64748b", whiteSpace: "nowrap" as const },
  reloadBtn:    { background: "transparent", border: "1px solid #334d6e", borderRadius: 6, padding: "4px 8px", color: "#94a3b8", cursor: "pointer" as const, fontSize: 15, lineHeight: 1 },
  tabBar:       { display: "flex" as const, background: "#fff", borderBottom: "2px solid #e2e8f0", padding: "0 8px", flexShrink: 0, overflowX: "auto" as const },
  tab:          { fontSize: 14, padding: "9px 12px", border: "none", background: "none", color: "#64748b", cursor: "pointer" as const, borderBottom: "2px solid transparent", marginBottom: -2, fontWeight: 400, whiteSpace: "nowrap" as const, fontFamily: "inherit" },
  tabActive:    { fontSize: 14, padding: "9px 12px", border: "none", background: "none", color: "#0f2942", cursor: "pointer" as const, borderBottom: "2px solid #0284c7", marginBottom: -2, fontWeight: 700, whiteSpace: "nowrap" as const, fontFamily: "inherit" },
  clBadge:      { fontSize: 12, background: "#dc2626", color: "#fff", borderRadius: 10, padding: "1px 5px", marginLeft: 5, fontWeight: 600 },
  statusBar:    { fontSize: 14, padding: "5px 14px", background: "#eff6ff", color: "#1d4ed8", borderBottom: "1px solid #bfdbfe", flexShrink: 0 },
  errorBar:     { fontSize: 14, padding: "5px 14px", background: "#fef2f2", color: "#dc2626", borderBottom: "1px solid #fecaca", flexShrink: 0 },
  actionBar:    { fontSize: 14, padding: "5px 14px", background: "#f0fdf4", color: "#15803d", borderBottom: "1px solid #bbf7d0", flexShrink: 0 },
  content:      { flex: 1, minHeight: 0, overflow: "auto" as const, display: "flex" as const, flexDirection: "column" as const },
  empty:        { padding: "40px 14px", textAlign: "center" as const, color: "#94a3b8", fontSize: 15, lineHeight: 1.6, whiteSpace: "pre-line" as const },
  emptySmall:   { fontSize: 14, color: "#94a3b8", padding: "8px 0" },
  // Filter bar (shared between Pipeline and Opportunities)
  filterBar:    { padding: "10px 14px", background: "#fff", borderBottom: "1px solid #f1f5f9" },
  searchBox:    { display: "flex" as const, alignItems: "center" as const, gap: 8, background: "#f8fafc", border: "1px solid #e2e8f0", borderRadius: 7, padding: "6px 10px", marginBottom: 8 },
  searchInput:  { flex: 1, border: "none", background: "transparent", fontSize: 14, outline: "none", color: "#1e293b", fontFamily: "inherit", minWidth: 0 },
  clearBtn:     { background: "none", border: "none", color: "#94a3b8", cursor: "pointer" as const, fontSize: 14, padding: 0, lineHeight: 1 },
  chips:        { display: "flex" as const, gap: 6, flexWrap: "wrap" as const },
  chip:         { fontSize: 13, padding: "3px 10px", borderRadius: 10, border: "1px solid #e2e8f0", background: "#fff", color: "#64748b", cursor: "pointer" as const, fontFamily: "inherit" },
  chipActive:   { fontSize: 13, padding: "3px 10px", borderRadius: 10, border: "1px solid #0284c7", background: "#eff6ff", color: "#0284c7", cursor: "pointer" as const, fontFamily: "inherit", fontWeight: 600 },
  filterSel:    { fontSize: 13, padding: "3px 6px", borderRadius: 7, border: "1px solid #e2e8f0", background: "#fff", color: "#64748b", fontFamily: "inherit", cursor: "pointer" as const },
  resultCount:  { fontSize: 13, color: "#94a3b8", marginTop: 6 },
  // Summary bar
  summaryBar:   { padding: "8px 14px", background: "#fff", borderBottom: "1px solid #f1f5f9", display: "flex" as const, alignItems: "center" as const, gap: 8, flexWrap: "wrap" as const, fontSize: 14 },
  summaryLabel: { color: "#64748b" },
  summaryValue: { fontWeight: 700, color: "#1e293b" },
  dot:          { color: "#cbd5e1" },
  clWarning:    { color: "#854d0e", fontWeight: 600 },
  // Tables
  //tableWrap:    { overflowX: "auto" as const, overflowY: "auto" as const, maxHeight: "600px" },
  tableWrap:    { overflowX: "auto" as const },
  tableArea: { flex: 1, minHeight: 0, overflowX: "auto" as const, overflowY: "auto" as const },
  tableHead:    { background: "#0f2942", padding: "6px 14px", position: "sticky" as const, top: 0, zIndex: 2 },
  //pipelineGrid: { display: "grid" as const, gridTemplateColumns: "80px 95px 90px 90px 58px 68px 88px 64px 28px", alignItems: "center" as const, borderBottom: "1px solid #f1f5f9" },
  pipelineGrid: { display: "grid" as const, gridTemplateColumns: "0.8fr 2.5fr 2fr 1.5fr 1fr 1fr 1.5fr 1.5fr 1.2fr 1fr 0.5fr", alignItems: "center" as const, borderBottom: "1px solid #f1f5f9" },
  //oppGrid:      { display: "grid" as const, gridTemplateColumns: "26px 74px 98px 85px 80px 56px 68px 76px 65px 50px 50px 64px", alignItems: "center" as const, borderBottom: "1px solid #f1f5f9" },
  oppGrid: { display: "grid" as const, gridTemplateColumns: "0.4fr 0.8fr 2.5fr 2fr 1.5fr 1fr 1fr 1.3fr 0.8fr 0.6fr 0.6fr 0.9fr", alignItems: "center" as const, borderBottom: "1px solid #f1f5f9" },
  th:           { fontSize: 13, color: "#fff", fontWeight: 600, whiteSpace: "nowrap" as const, overflow: "hidden" as const },
  td:           { fontSize: 13, color: "#475569", whiteSpace: "nowrap" as const },
  tdMono:       { fontSize: 13, fontFamily: "monospace", fontWeight: 600, color: "#0284c7", whiteSpace: "nowrap" as const },
  ell:          { overflow: "hidden" as const, textOverflow: "ellipsis" as const },
  offerBadge:   { fontSize: 13, background: "#fffbeb", color: "#854d0e", padding: "1px 5px", borderRadius: 8, marginLeft: 3, fontWeight: 600, border: "1px solid #fde68a" },
  cbChecked:    { display: "inline-flex" as const, width: 15, height: 15, borderRadius: 3, background: "#0284c7", color: "#fff", fontSize: 11, alignItems: "center" as const, justifyContent: "center" as const },
  cbEmpty:      { display: "inline-block" as const, width: 15, height: 15, borderRadius: 3, border: "1px solid #cbd5e1", background: "#fff" },
  // Detail identity
  //detailId:     { padding: "10px 14px", background: "#fff", borderBottom: "1px solid #f1f5f9", position: "sticky" as const, top: 0, zIndex: 9999 },
  detailPage: {
    flex: 1,
    minHeight: 0,
    display: "flex" as const,
    flexDirection: "column" as const,
    overflow: "hidden" as const,
  },

  detailId: {
    padding: "10px 14px",
    background: "#fff",
    borderBottom: "1px solid #f1f5f9",
    flexShrink: 0,
    zIndex: 10,
    boxShadow: "0 2px 4px rgba(15, 41, 66, 0.06)",
  },

  detailBody: {
    flex: 1,
    minHeight: 0,
    overflowY: "auto" as const,
    paddingBottom: 12,
  },
  detailCep:    { fontSize: 15, fontWeight: 700, color: "#0f2942" },
  detailSub:    { fontSize: 13, color: "#64748b", marginTop: 2 },
  // Collapsible section
  section:      { background: "#fff", borderBottom: "1px solid #f1f5f9", marginTop: 8 },
  sectionBtn:   { width: "100%", padding: "10px 14px", background: "none", border: "none", display: "flex" as const, alignItems: "center" as const, justifyContent: "space-between" as const, cursor: "pointer" as const, textAlign: "left" as const, fontFamily: "inherit" },
  sectionTitle: { fontSize: 14, fontWeight: 700 },
  sectionSub:   { fontSize: 13, color: "#94a3b8" },
  sectionBody:  { padding: "0 14px 14px", borderTop: "1px solid #f1f5f9" },
  // CEP snapshot — 2-column layout: label | value | label | value
  crmGrid:      { display: "grid" as const, gridTemplateColumns: "100px 1fr 100px 1fr", gap: "4px 12px", marginTop: 8 },
  crmLabel:     { fontSize: 13, color: "#64748b", fontWeight: 600, padding: "2px 0" },
  crmValue:     { fontSize: 13, color: "#1e293b", padding: "2px 0" },
  // Form fields
  twoColForm:   { display: "grid" as const, gridTemplateColumns: "1fr 1fr", gap: 10, marginTop: 10 },
  field:        { marginTop: 10 },
  fieldLbl:     { display: "block" as const, fontSize: 14, color: "#64748b", marginBottom: 4, fontWeight: 600 },
  fieldHint:    { fontSize: 12, fontWeight: 400, color: "#94a3b8" },
  sel:          { fontSize: 15, padding: "6px 8px", borderRadius: 7, border: "1px solid #cbd5e1", background: "#fff", width: "100%", fontFamily: "inherit" },
  inp:          { fontSize: 15, padding: "6px 8px", borderRadius: 7, border: "1px solid #cbd5e1", background: "#fff", width: "100%", fontFamily: "inherit", boxSizing: "border-box" as const },
  readonlyVal:  { fontSize: 15, padding: "6px 8px", borderRadius: 7, border: "1px solid #f1f5f9", background: "#f8fafc", color: "#94a3b8" },
  checkLabel:   { display: "flex" as const, alignItems: "center" as const, gap: 8, cursor: "pointer" as const, fontSize: 15, color: "#1e293b" },
  fieldWarning: { fontSize: 12, fontWeight: 600, color: "#854d0e", marginTop: 4 },
  saveRow:      { display: "flex" as const, alignItems: "center" as const, justifyContent: "space-between" as const, marginTop: 12, paddingTop: 10, borderTop: "1px solid #f1f5f9" },
  lastSaved:    { fontSize: 13, color: "#94a3b8" },
  saveBtn:      { fontSize: 14, padding: "7px 16px", borderRadius: 8, border: "none", background: "#0f2942", color: "#fff", cursor: "pointer" as const, fontWeight: 600, fontFamily: "inherit" },
  offerDivergeBanner: { display: "flex" as const, alignItems: "center" as const, gap: 5, margin: "6px 0 0", padding: "6px 8px", background: "#fffbeb", borderRadius: 5, border: "1px solid #fde68a", fontSize: 13, color: "#854d0e" },
  overdueBanner: { marginTop: 6, padding: "6px 8px", background: "#fffbeb", borderRadius: 5, border: "1px solid #fde68a", fontSize: 13, color: "#854d0e", fontWeight: 600 },
  acceptLink:   { fontWeight: 600, cursor: "pointer" as const, textDecoration: "underline" as const },
  // Comments
  commentCard:  { padding: "8px 10px", background: "#f8fafc", borderRadius: 7, border: "1px solid #f1f5f9", marginBottom: 8 },
  commentMeta:  { fontSize: 13, color: "#94a3b8", marginBottom: 3 },
  commentText:  { fontSize: 14, color: "#1e293b", lineHeight: 1.5 },
  deleteLink:   { background: "none", border: "none", color: "#94a3b8", fontSize: 13, cursor: "pointer" as const, padding: 0, marginTop: 4, fontFamily: "inherit" },
  textarea:     { fontSize: 14, padding: "8px 10px", borderRadius: 7, border: "1px solid #e2e8f0", background: "#fff", width: "100%", marginTop: 10, resize: "vertical" as const, fontFamily: "inherit", boxSizing: "border-box" as const },
  addBtn:       { fontSize: 14, padding: "6px 14px", borderRadius: 7, border: "none", background: "#0f2942", color: "#fff", cursor: "pointer" as const, fontWeight: 600, marginTop: 6, fontFamily: "inherit" },
  // Change log
  clCard:       { padding: 10, background: "#fffbeb", border: "1px solid #fde68a", borderRadius: 7, marginBottom: 8 },
  clCardAck:    { padding: 10, background: "#f8fafc", border: "1px solid #f1f5f9", borderRadius: 7, marginBottom: 8, opacity: 0.7 },
  clCardHeader: { display: "flex" as const, alignItems: "center" as const, justifyContent: "space-between" as const, marginBottom: 6 },
  clEventBadge: { fontSize: 13, fontWeight: 600, background: "#854d0e", color: "#fff", padding: "2px 8px", borderRadius: 10 },
  clDate:       { fontSize: 13, color: "#94a3b8" },
  clDetail:     { fontSize: 14, color: "#854d0e", lineHeight: 1.5, marginBottom: 8 },
  ackBtn:       { fontSize: 13, padding: "4px 12px", borderRadius: 6, border: "1px solid #fde68a", background: "#fff", color: "#854d0e", cursor: "pointer" as const, fontWeight: 600, fontFamily: "inherit" },
  ackNote:      { fontSize: 13, color: "#94a3b8" },
  clOpenBadge:  { fontSize: 13, background: "#fef2f2", color: "#dc2626", padding: "1px 7px", borderRadius: 10, fontWeight: 600, border: "1px solid #fecaca" },
  // Opportunity detail
  promoteCard:     { margin: "10px 14px 0", padding: 12, background: "#fff", border: "1px solid #e2e8f0", borderRadius: 8, display: "flex" as const, alignItems: "center" as const, gap: 10 },
  promoteTitle:    { fontSize: 14, fontWeight: 600, color: "#0f2942", marginBottom: 2 },
  promoteSub:      { fontSize: 13, color: "#64748b" },
  promoteBtnOff:   { fontSize: 14, padding: "7px 14px", borderRadius: 8, border: "none", background: "#0f2942", color: "#fff", cursor: "pointer" as const, fontWeight: 600, fontFamily: "inherit", whiteSpace: "nowrap" as const, flexShrink: 0 },
  promoteBtnOn:    { fontSize: 14, padding: "7px 14px", borderRadius: 8, border: "1px solid #bbf7d0", background: "#f0fdf4", color: "#15803d", cursor: "pointer" as const, fontWeight: 600, fontFamily: "inherit", whiteSpace: "nowrap" as const, flexShrink: 0 },
  inPipelineBadge: { fontSize: 13, background: "#eff6ff", color: "#0284c7", padding: "3px 10px", borderRadius: 10, fontWeight: 600, border: "1px solid #bfdbfe", whiteSpace: "nowrap" as const },
  promoteNote:     { margin: "6px 14px 0", padding: "6px 8px", background: "#f0fdf4", borderRadius: 5, border: "1px solid #bbf7d0", fontSize: 13, color: "#15803d" },
  equipPlaceholder:{ padding: "16px 0", textAlign: "center" as const, display: "flex" as const, flexDirection: "column" as const, alignItems: "center" as const, gap: 8 },
  equipText:       { fontSize: 14, color: "#94a3b8", lineHeight: 1.5, maxWidth: 280 },
  comingSoonBadge: { fontSize: 13, background: "#f1f5f9", color: "#94a3b8", padding: "1px 8px", borderRadius: 10, border: "1px solid #e2e8f0" },
  // Equipment table (Opportunity detail = view-only, Pipeline detail = editable)
  eqActionNote:    { fontSize: 13, color: "#15803d", background: "#f0fdf4", border: "1px solid #bbf7d0", borderRadius: 5, padding: "6px 8px", marginBottom: 8, lineHeight: 1.4 },
  eqTableWrap:     { overflowX: "auto" as const, marginTop: 4 },
  eqTableHead:     { background: "#0f2942", padding: "6px 14px", borderRadius: "6px 6px 0 0" },
  eqGridView:      { display: "grid" as const, gridTemplateColumns: "2.2fr 1.3fr 0.8fr 1fr 0.9fr 1fr 1.2fr", alignItems: "center" as const, gap: 8, borderBottom: "1px solid #f1f5f9" },
  eqGridEdit:      { display: "grid" as const, gridTemplateColumns: "2fr 1.2fr 0.8fr 0.9fr 0.9fr 1fr 1.1fr 1.4fr", alignItems: "center" as const, gap: 8, borderBottom: "1px solid #f1f5f9" },
  eqCard:          { padding: 10, background: "#f8fafc", borderRadius: 7, border: "1px solid #f1f5f9", marginBottom: 8 },
  eqCardHeader:    { display: "flex" as const, alignItems: "center" as const, gap: 6, marginBottom: 4 },
  eqOriginBadge:   { fontSize: 12, fontWeight: 600, color: "#64748b", background: "#eef2f6", border: "1px solid #e2e8f0", borderRadius: 8, padding: "1px 7px", whiteSpace: "nowrap" as const, flexShrink: 0 },
  eqLabel:         { fontSize: 14, fontWeight: 600, color: "#1e293b" },
  eqMetaRow:       { display: "flex" as const, alignItems: "center" as const, gap: 6, fontSize: 13, color: "#64748b", marginTop: 2 },
  eqArchetypeCode: { fontFamily: "monospace", fontSize: 13, color: "#0284c7", fontWeight: 600 },
  eqStatusRow:     { display: "flex" as const, alignItems: "center" as const, gap: 6, marginTop: 4 },
  eqStatusOk:      { fontSize: 12, fontWeight: 600, color: "#15803d", background: "#f0fdf4", border: "1px solid #bbf7d0", borderRadius: 8, padding: "1px 7px" },
  eqStatusWarn:    { fontSize: 12, fontWeight: 600, color: "#854d0e", background: "#fffbeb", border: "1px solid #fde68a", borderRadius: 8, padding: "1px 7px" },
  eqStatusPending: { fontSize: 12, fontWeight: 600, color: "#1d4ed8", background: "#eff6ff", border: "1px solid #bfdbfe", borderRadius: 8, padding: "1px 7px" },
  eqStatusRemoving:{ fontSize: 12, fontWeight: 600, color: "#dc2626", background: "#fef2f2", border: "1px solid #fecaca", borderRadius: 8, padding: "1px 7px", whiteSpace: "nowrap" as const },
  eqAdvisoryBadge: { fontSize: 12, fontWeight: 600, color: "#9a3412", background: "#fff7ed", border: "1px solid #fed7aa", borderRadius: 8, padding: "1px 7px" },
  eqAdvisoryLeftCep: { fontSize: 12, fontWeight: 600, color: "#854d0e", background: "#fffbeb", border: "1px solid #fde68a", borderRadius: 8, padding: "1px 7px", whiteSpace: "nowrap" as const },
  eqDetailText:    { fontSize: 13, color: "#854d0e", marginTop: 4, lineHeight: 1.4 },
  eqActionRow:     { display: "flex" as const, alignItems: "center" as const, gap: 12, marginTop: 6 },
  eqExcludedRow:   { display: "flex" as const, alignItems: "center" as const, justifyContent: "space-between" as const, padding: "6px 0", borderBottom: "1px solid #f1f5f9", fontSize: 14 },
  eqAddForm:       { marginTop: 10, padding: 10, background: "#f8fafc", borderRadius: 7, border: "1px solid #f1f5f9" },
  eqOfferLabel:    { fontSize: 13, fontWeight: 600, color: "#475569", margin: "10px 0 4px" },
  eqNewOfferBlock: { marginTop: 4, paddingTop: 10, borderTop: "1px dashed #e2e8f0" },
  eqAddFormActions:{ display: "flex" as const, alignItems: "center" as const, gap: 8, marginTop: 8 },
  leftSyncNote:  { fontSize: 13, color: "#854d0e", background: "#fffbeb", border: "1px solid #fde68a", borderRadius: 5, padding: "8px 10px", marginTop: 8, marginBottom: 4, lineHeight: 1.5 },
  // MultiSelect
  multiSelWrap:    { position: "relative" as const },
  multiSelBtn:     { fontSize: 13, padding: "3px 10px", borderRadius: 10, border: "1px solid #e2e8f0", background: "#fff", color: "#64748b", cursor: "pointer" as const, fontFamily: "inherit", whiteSpace: "nowrap" as const },
  multiSelBtnOn:   { fontSize: 13, padding: "3px 10px", borderRadius: 10, border: "1px solid #0284c7", background: "#eff6ff", color: "#0284c7", cursor: "pointer" as const, fontFamily: "inherit", fontWeight: 600, whiteSpace: "nowrap" as const },
  multiSelPanel:   { position: "absolute" as const, top: "calc(100% + 4px)", left: 0, background: "#fff", border: "1px solid #e2e8f0", borderRadius: 8, boxShadow: "0 4px 12px rgba(0,0,0,0.08)", zIndex: 100, minWidth: 180, maxHeight: 240, overflowY: "auto" as const },
  multiSelRow:     { display: "flex" as const, alignItems: "center" as const, gap: 8, padding: "6px 12px", cursor: "pointer" as const, fontSize: 14, whiteSpace: "nowrap" as const },
  multiSelToggle:  { padding: "6px 12px", fontSize: 13, color: "#0284c7", cursor: "pointer" as const, borderBottom: "1px solid #f1f5f9" },
  // Pipeline row delete
  deleteRowBtn:     { cursor: "pointer" as const, color: "#cbd5e1", fontSize: 14, padding: "0 4px", lineHeight: 1, flexShrink: 0 },
  deleteConfirmRow: { gridColumn: "1 / -1", display: "flex" as const, alignItems: "center" as const, gap: 10, fontSize: 14, color: "#dc2626" },
  confirmDeleteBtn: { fontSize: 13, padding: "4px 10px", borderRadius: 6, border: "none", background: "#dc2626", color: "#fff", cursor: "pointer" as const, fontFamily: "inherit", fontWeight: 600 },
  cancelBtn:        { fontSize: 13, padding: "4px 10px", borderRadius: 6, border: "1px solid #e2e8f0", background: "#fff", color: "#64748b", cursor: "pointer" as const, fontFamily: "inherit" },
};