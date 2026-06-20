# SOP Planning — README

A capacity-planning tool for S&OP, built on a co-authored Excel workbook (SharePoint), three Office Scripts (the engine), and an Office Add-in (the visual interface). No IT approval needed — everything runs on tools already in your Microsoft 365 license, plus one free, publicly-hosted file location for the add-in's interface files (no business data is stored there — see §5).

---

## 1. What this is

The workbook holds your demand, routings, capacity, and backlog as plain Excel tables. Three Office Scripts read those tables, compute a finite-capacity production plan (or an unconstrained "ideal date" plan), and write the results back as plain tables. The Office Add-in is a side panel inside Excel that turns those results into a load chart, a capacity heatmap, a schedule view, and editable forms — so you don't have to read raw tables to understand or adjust the plan.

```
Demand files (region input)
        │
        ▼
   DemandImport  ──[Sync Demand script]──▶  Demand  (the master pipeline list)
        │
        ▼
  Routings · CapacityBase · CapacityOverride · Backlog · Adjustments · Calendar
        │
        ▼
   [Rebalance script — constrained or unconstrained]
        │
        ▼
   Results / Schedule  (or ResultsUnconstrained / ScheduleUnconstrained)
        │
        ▼
   Office Add-in (this README's launcher) ── visualises + edits
        │
        ▼
   Power BI (separate, reads Results for reporting)
```

---

## 2. The Excel tables

All tables live in the same SharePoint-hosted workbook. Column names matter — the scripts read by header name, not position.

| Table | Purpose |
|---|---|
| **DemandImport** | Raw import from the regional demand files (one row per opportunity + equipment line) |
| **Demand** | The synced master list — adds `RoutingKey`, `Plant`, `Priority` once a planner assigns them. This is what the engine actually schedules. |
| **Routings** | One row per `Routing Key`, giving the standard load and duration/offset for every stage (ESL, CTO, Assembly, Testing, FAT) plus the milestone day-offsets (KOM, BOM, Frozen, FCA, etc.) |
| **CapacityBase** | Default weekly capacity per Plant + Stage |
| **CapacityOverride** | Sparse exceptions — only the plant/stage/weeks where capacity differs from the default (a ramp-up, a maintenance dip, etc.) |
| **Backlog** | Firm, already-booked load per Plant + Stage + Week — reserved before any pipeline opportunity is scheduled |
| **Adjustments** | Sparse, per-line special cases — override a stage's load or pin its start week for one specific `OppID` + `LineID` |
| **Calendar** | Maps the engine's internal week index (1…156) to real ISO weeks/dates, so the tool can run multi-year without breaking at year boundaries |
| **Results** / **ResultsUnconstrained** | Output: load per Plant + Stage + Week + Status, split into Backlog vs Planned, with utilisation % and overload |
| **Schedule** / **ScheduleUnconstrained** | Output: the planned date of every milestone (KOM, Assembly start/end, Testing, FAT, FCA…) per opportunity + equipment line, plus lead time |

**Constrained vs Unconstrained:** the constrained engine respects capacity — if a week is full, the load slides forward (fill-and-carry). The unconstrained engine places everything at its ideal date regardless of capacity, which is useful for seeing what the *demand* actually wants before capacity reshapes it.

---

## 3. The three Office Scripts

Found under Excel's **Automate** tab. Each is a standalone script you run on demand (or bind to a button).

1. **Sync Demand** — reads `DemandImport`, compares it against the current `Demand` table, and flags every line as New / Changed / Removed / Unchanged. New lines arrive with blank `RoutingKey`/`Plant` — a planner has to assign those (see §6, Demand tab).
2. **Rebalance (constrained)** — the production engine. Reserves `Backlog` first, then schedules `Demand` lines in priority order (manual `Priority` → Status → OI date) against remaining capacity, writing `Results` + `Schedule`.
3. **Rebalance (unconstrained)** — same scheduling logic but ignores capacity entirely, writing `ResultsUnconstrained` + `ScheduleUnconstrained`.

**Run order matters:** Sync Demand → Rebalance (either mode). Running Rebalance without syncing first just re-plans whatever was already in `Demand`.

> **Known limitation:** the Office Add-in cannot trigger these scripts directly — there's no Office.js API for it (a current Microsoft platform gap). The Run tab in the add-in navigates you to the right sheet; you still trigger the script yourself from the Automate tab.

---

## 4. How the add-in is hosted (read this once)

Earlier versions of this tool required every user to run a local development server (`npm start`) on their own machine just to use the add-in. **This is no longer the case.**

The add-in's interface files (the charts, forms, and tabs you see in the task pane) are now hosted on **GitHub Pages**, a free static file host, at:

```
https://ricardo-gama.github.io/sop-planning-tool/
```

This is just a file server for the add-in's *code* — think of it like a website hosting some JavaScript and HTML. **No S&OP data is stored there.** All demand, routings, capacity, and results stay exactly where they always have: in the SharePoint workbook. The add-in reads and writes that data live, through Excel's own connection, regardless of where its own interface files are hosted.

What this means in practice:

- You do **not** need Node.js installed to *use* the add-in.
- You do **not** need to run any local server, batch file, or terminal command.
- The add-in works as long as you have internet access and Excel is signed in normally — no different from any other website your browser loads.
- Only one person (the tool maintainer) needs to rebuild and redeploy the add-in's code when changes are made — see §10 for that process.

For everyone else, installing the add-in is a one-time, two-step process, covered in §5.

---

## 5. Installing the Add-in (one-time, each person)

You only need to do this once per computer. After that, the add-in is just available in Excel like any other built-in feature.

### Step A — Get added to the shared manifest folder

The add-in is distributed via a small SharePoint folder containing a single configuration file (`manifest.xml`). If you don't already have access to it, ask the tool maintainer to share:

```
S&OP - Global Equipment Planning → Shared Documents → AddinManifest
```

Once shared with you, **sync that folder** so it appears as a normal folder on your computer:

1. Open the `AddinManifest` folder link in your browser
2. Click **Sync** (top toolbar)
3. This opens the OneDrive app and adds the folder to your computer — wait for it to finish syncing
4. In File Explorer, locate the synced folder (usually under a "Sidel" or company entry in the left sidebar) and open it until you see `manifest.xml` sitting inside
5. With that folder open, click the **address bar** at the top of File Explorer to reveal the full text path, and copy it — you'll need it in Step B

It will look something like:
```
C:\Users\<yourname>\Sidel\S&OP - Global Equipment Planning - AddinManifest
```

### Step B — Trust the folder in Excel

1. Open Excel (any workbook)
2. **File** → **Options** → **Trust Center** → **Trust Center Settings...**
3. **Trusted Add-in Catalogs** (left sidebar)
4. In the **Catalog Url** box, paste the path from Step A, but reformatted as a network path:
   ```
   \\localhost\C$\Users\<yourname>\Sidel\S&OP - Global Equipment Planning - AddinManifest
   ```
   (Replace `<yourname>` with your actual Windows username, and the rest of the path with whatever you copied in Step A — the drive letter `C:` becomes `C$`, and a `\\localhost\` prefix is added.)
5. Click **Add catalog**
6. Tick the **Show in Menu** checkbox next to the newly added entry — easy to miss, but required
7. Click **OK**, then **OK** again
8. **Fully close Excel** and reopen it (Trust Center changes need a restart)

### Step C — Add the add-in

1. Open the shared S&OP workbook
2. Go to **Insert** → **My Add-ins** (or **Add-ins**, depending on your Excel version)
3. Look for a **SHARED FOLDER** tab at the top of the dialog
4. The add-in should appear there — select it and click **Add**

The task pane should open and load the same Load / Heat / Schedule / Demand / Adjust tabs described in §6. From now on, it's available the same way any other add-in is — no daily startup steps, no terminal, no shortcut to double-click.

---

## 6. Using the Add-in

Five tabs, plus a Constrained/Unconstrained toggle in the header that switches every tab between the two result sets.

- **Load** — stacked area chart of load over time. Filter by plant (multi-select), stage, status (multi-select), and a from/to week range. Backlog always renders at the bottom of the stack, followed by status in priority order. A **Save as PNG** button exports the current chart view. A capacity override form sits below the filters for quick "what if I add capacity here" entries — it writes to `CapacityOverride` immediately, but you need to re-run the Rebalance script to see the effect reflected in the chart.
- **Heat** — a week-by-week utilisation grid for one plant + stage at a time, colour-coded (green/amber/orange/red), with a breakdown table below showing the load split by status.
- **Schedule** — every opportunity + equipment line with its OI date, KOM, FAT, FCA, and lead time, filterable by plant.
- **Demand** — defaults to showing only **unassigned** lines (no `RoutingKey` or `Plant` yet, shown in red). Click **Edit** on a line to set Routing Key (dropdown, sourced from `Routings`), Plant (dropdown, sourced from `CapacityBase` — so it's always in sync with whatever plants you've actually configured), and Priority. Below that, all five stages are listed with their standard load (from the selected routing) and two optional fields each: a load override and a start-week pin. Saving writes everything in one go — `RoutingKey`/`Plant`/`Priority` to `Demand`, stage fields to `Adjustments`.
- **Adjust** — the same per-line adjustment form (for when you already know the `OppID`/`LineID` and don't need the full Demand list), plus the capacity override form, each with a table of existing entries below for quick editing.

The refresh button in the tab bar refreshes all data from the workbook — use it after running a script.

---

## 7. Typical weekly cycle

1. Regional planners update the two demand files (unchanged from before).
2. Run **Sync Demand** in Excel.
3. Open the add-in's **Demand** tab, assign Routing Key + Plant (+ Priority if needed) to any newly unassigned lines.
4. Run **Rebalance (constrained)**.
5. Review the **Load** and **Heat** tabs for overloads; add capacity overrides or line adjustments as needed.
6. Re-run Rebalance to see the effect.
7. Power BI picks up `Results` on its normal refresh.

---

## 8. Troubleshooting

| Symptom | Cause | Fix |
|---|---|---|
| Add-in doesn't appear under Insert → My Add-ins → Shared Folder | Trust Center catalog not added correctly, or Excel wasn't restarted after adding it | Re-check §5 Step B; confirm "Show in Menu" is ticked; fully close and reopen Excel |
| "Sorry, the location you entered isn't correct" when adding the catalog | Trust Center rejects plain `C:\` drive-letter paths | Use the `\\localhost\C$\...` format shown in §5 Step B, not the raw `C:\` path |
| "This add-in is no longer available" | Stale registration from an old localhost-based install (pre-migration) | Remove the old entry: Excel → Insert → Add-ins → My Add-ins → ⋯ → Remove. Then re-add from the Shared Folder tab per §5 Step C |
| Task pane opens but shows a blank/broken page | No internet access, or GitHub Pages is temporarily unreachable | Confirm you have internet access; try opening `https://ricardo-gama.github.io/sop-planning-tool/taskpane.html` directly in a browser to check the host is up |
| TypeScript `Cannot find name 'Excel'` / `'Office'` (maintainer only, during development) | Missing type reference | Add `/// <reference types="office-js" />` at the top of `App.tsx` |
| TypeScript `Set<string> can only be iterated...` (maintainer only) | Set spread (`[...mySet]`) needs a newer TS target | Use `Array.from(mySet)` instead |
| Outlook "Debug Event-based handler" popup | Office runtime shared with Outlook, unrelated to this add-in | Click Cancel — harmless |

---

## 9. Known limitations (current state)

- The add-in can't trigger Office Scripts directly (Microsoft platform gap) — run scripts manually from the Automate tab.
- Capacity overrides entered in the add-in don't recompute the chart live — they write to the table; you re-run Rebalance to see the effect.
- The network-share catalog method (§5) is a Windows-only mechanism and is, per Microsoft's own documentation, intended primarily for testing rather than formal production deployment — though it is fully functional for a small, stable team like ours. If the add-in's *ribbon* (not the task pane content) ever changes — e.g. a new button is added — each person will need to remove and re-add the add-in once.
- Manifest distribution currently goes through a synced SharePoint folder rather than a centrally IT-managed catalog, because centralized deployment requires admin rights not available to this team.

---

## 10. For the maintainer: updating and redeploying the add-in

This section is only relevant if you're making code changes to the add-in itself.

1. Make your changes in `src/taskpane/` as usual.
2. Build for production:
   ```bash
   npm run build
   ```
   This compiles the app and automatically rewrites `manifest.xml` inside the `dist/` folder to point at the live GitHub Pages URL (handled by `webpack.config.js`'s `urlDev` → `urlProd` swap).
3. Deploy the new build:
   ```bash
   npm run deploy
   ```
   This pushes the contents of `dist/` to GitHub Pages. The live URL updates automatically, usually within a minute.
4. If the build's `manifest.xml` changed in a way that affects the **ribbon** (new buttons, renamed commands), copy the updated `dist/manifest.xml` over the one in the shared `AddinManifest` SharePoint folder, and let the team know they'll need to remove and re-add the add-in once (see §8).
5. If only the task pane content changed (charts, forms, tabs) — no action needed beyond steps 2–3. Everyone's already-installed add-in will pick up the new version automatically next time they open it, since it loads fresh from GitHub Pages each time.

---

## 11. File index

| File | Purpose |
|---|---|
| `sop_rebalance.ts` (×2) | The constrained and unconstrained Office Scripts |
| `manifest.xml` (project root) | Source manifest, points at `localhost` — used only for local development |
| `dist/manifest.xml` | Production manifest, points at GitHub Pages — this is the one distributed to the team |
| `src/taskpane/index.tsx` | Add-in entry point — initializes Office, renders `App` |
| `src/taskpane/App.tsx` | The add-in UI — all 5 tabs |
| `webpack.config.js` | Build config — dev server settings, production URL swap, GitHub Pages `publicPath` |
| `package.json` | Includes the `build` and `deploy` scripts used in §10 |