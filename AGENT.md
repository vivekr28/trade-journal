# AGENT.md — Trading Journal App

Context file for any AI agent (or human) picking up work on this project. Read this before making changes.

## What this is

A personal trading journal / capital tracker used to log trades, track partial exits, and analyze performance (P&L curve, drawdown, win rate, R-multiples, setup edge, risk discipline). Two files, no build tooling:

- `trading-journal.html` — the entire frontend. Single-file Vue 3 (Options API) + Vuetify 3 app, loaded via CDN script tags (no npm, no bundler). Chart.js 4.4.4 for all charts. Opened directly in a browser.
- `trading-journal.gs` — Google Apps Script backend, deployed as a web app bound to a Google Sheet. Acts as the "database" (sheets: Trades, Trade_Legs, Config, Transactions/Capital log) and API (`doGet`/`doPost` dispatch to functions like `getAllTrades`, `addNewTrade`, `addExitLeg`, `getRibbonData`, `updateTradeSL`, etc.). The frontend talks to it via the Apps Script web app URL stored in `localStorage` (`gas_url`).

There is no repo-level test suite or CI. Verification is manual/scripted per-change (see "How changes are verified" below).

## Architecture notes

- Vue app root is a single `data()`/`computed`/`methods` object mounted on the page — everything lives in one `<script>` block inside the HTML file. No components, no router, no Vuex/Pinia.
- Key state: `allTrades` (all trades from backend), `openTrades`, `exitLegsCache` (keyed by trade id, populated async after trades load — see "known pitfalls" below), `activeTab` (`trades` / `analytics` / `account`), `tradesViewMode` (`cards` / `table`).
- Trade object fields of note: `symbol`, `setup` (free-text/select, e.g. `SVRO3`, `SVRO15`, `SNB`, `VCP60`), `type` (`L`/`S`), `contract` (`CNC`/`MIS`), `status` (`Open`/`Closed`), `buyDate`/`buyTime`/`buyPrice`, `initialSL`/`currentSL`, `exitDate`, `avgExitPrice`, `pnl`, `rocePct`, `invested`, `capital` (capital at time of trade), plus market-context fields (`marketEnv`, `rvol`, `adRatio`, `causeMove`, `consolMove`, `maUndercut`, `lfRise`) used in the Setup Edge & Market Context analytics.
- Partial exits are modeled as "legs" (`Trade_Legs` sheet). `getRData(t)` / `calcRData(t)` derive realized qty/R/PnL from `exitLegsCache[t.id]` (array of `{qty, price, date, time, legR, legPnl}`).
- Analytics tab computeds (`analyticsCore`, `pnlCurveData`, `capitalMarkers`, `drawdownData`, `monthlyPnlData`, `monthlyStreakData`, `edgeGroupings`, `contextGroupings`, `riskTrendData`, `streakTimelineData`, etc.) feed Chart.js renders in `renderAnalyticsCharts()` → `renderPnlCurveChart`, `renderMonthlyPnlChart`, `renderDrawdownChart`, `renderRiskTrendChart`, `renderStreakChart`.
- Trades tab has two views: card list (`.trade-card-open` / `.trade-card-closed`) and a dense `Trades Table` (`tradeTableRows` computed → `<table class="trades-table">`), toggled via `tradesViewMode`. Table columns: Symbol, Setup, Entry Date/Time, Avg Buy Price, **Cur. Invested** (current remaining invested = `rd.remQty * rd.netBuyPrice`, `null`/`—` for closed trades), **Cur. Allocation %** (current remaining invested / capital * 100, `null`/`—` for closed trades), Exit Date/Time, Avg Sell/Exit Price, Overall P&L, Duration, ROCE %, R Value.

## Known pitfalls (hit during earlier sessions — don't repeat)

1. **Race condition on load**: `fetchAllTrades()` assigns `this.allTrades` synchronously but `exitLegsCache` populates asynchronously afterward (inside a `Promise.all`). Anything that renders off `exitLegsCache` (e.g. analytics charts) must wait for that to resolve — triggered via `this.$nextTick(() => this.renderAnalyticsCharts())` placed *after* the exit-legs fetch, not on `allTrades` assignment.
2. **Flexbox `white-space:nowrap` + fixed `min-width`**: adding `min-width` / `flex-shrink:0` to a label without also giving its sibling value `min-width:0` causes the value to overflow and visually overlap the next cell (nowrap text won't shrink below its own content width). Always pair a shrink-locked label with a `min-width:0` + `overflow:hidden;text-overflow:ellipsis` value.
3. **CSS Grid alignment across visually-stacked-but-separate grid containers doesn't work**: if a "section" is built from multiple independent `display:grid` divs stacked vertically (one per row), each computes its own column widths from its own content — rows will NOT align vertically even with identical `grid-template-columns`. Real alignment needs either one shared grid container for the whole section, or fixed-pixel (not `fr`/`auto`) column tracks. The `.exit-leg-row` grid (`18px minmax(0,105px) minmax(0,125px) 48px auto`) is the working pattern to copy for anything similar.

## Style conventions (established through iteration, don't regress)

- **[17 Sep 2026] Re-themed to match wealthlab.in — supersedes the old "flat white text" rule below.** Design tokens (shadcn/ui-style dark theme):
  - Page background `#0F172A`, card/surface background `#1E293B`, borders `#334155` (subtle dividers use `rgba(51,65,85,0.6)`).
  - Primary/value text `#F1F5F9` (near-white), **muted/label text `#94A3B8`** — labels and secondary/meta text now intentionally use this dim gray, values stay bright. This *reverses* the earlier flat-white-everywhere decision below; the user explicitly re-approved dim-label/bright-value differentiation to match wealthlab.in's look when asked directly about the conflict.
  - Accent blue `#3B82F6` (primary, already matched pre-restyle), light accent `#60A5FA`, green (gains/success) `#4ADE80`, red (losses/destructive) `#F87171` (unchanged), yellow/warning `#FBBF24` (unchanged).
  - Font: **Inter** everywhere (replaced DM Sans for body text and JetBrains Mono for numeric/mono values). Numeric values that used to get `font-family:'JetBrains Mono'` now use `font-variant-numeric:tabular-nums` on Inter instead (see `.mono`, `.ribbon-val`, `.tc-val`, `.tc-cell-val`, `.r-big`, `.exit-leg-r`, `.exit-leg-pnl`, `.aj-dash-cell .val`). Google Fonts link is now just `Inter:wght@400;500;600;700;800`.
  - Vuetify theme colors (`createVuetify` config near the bottom of the script) were updated to the same tokens: `background:#0F172A`, `surface:#1E293B`, `surface-variant:#334155`, `on-background`/`on-surface:#F1F5F9`, `success:#4ADE80`, others unchanged.
  - Border-radius/spacing/shadows were left as-is — they already matched wealthlab's conventions (10px card radius, 16px dialogs, soft hover shadows) before the restyle.
  - ~~All in-card text (labels, values, titles) is flat bright white `#FFFFFF`~~ — **superseded, see above.** (Old rationale kept for history: the user had previously rejected differentiated dim/bright label-vs-value coloring in favor of one flat brightness. That preference was explicitly and deliberately overridden by the wealthlab.in restyle request.)
- Dark theme, background `#141925` for cards, `#0B0F17`-ish page background. — **outdated, see wealthlab retheme above**; the actual current values are page `#0F172A` / card `#1E293B`.
- Card borders: brightened from the original dim `1px solid #252D3F`. Trade cards use `2px solid #475569` (open) / `2px solid #64748B` (closed) plus a `4px` left accent. Analytics section cards (Core Performance, Setup Edge & Market Context, Risk & Discipline) now match at `2px solid #475569`, keeping their own `16px` border-radius (vs. trade cards' `10px`).
- No orange/green background tinting on trade card sections — removed per explicit request as "not professional."
- Profit/loss coloring uses `colorClass(value)` (green positive / red negative) — keep this for R-multiples, P&L, ROCE%; don't apply it to plain informational text.
- Trade card sections use only vertical divider lines between adjacent sections (`.trade-section + .trade-section { border-left:1px solid #334155; }`) — no per-section boxes/backgrounds.

## Recent UI changes (session log)

- **Trades table — Cur. Invested column**: Added between Avg Buy Price and Allocation %. Shows `rd.remQty * rd.netBuyPrice` (current remaining deployed capital) for open/partial trades; shows `—` for closed trades. `tradeTableRows` return object carries `invested: isClosed ? null : (rd.remQty||0) * (rd.netBuyPrice||0)`. colspan on empty-state row updated to 12.
- **Current Position card (cards view) — Position %**: Added `Position %` cell after the existing `Invested` cell in the first `tc-grid` of the Current Position section. Formula: `((rd.remQty * rd.netBuyPrice) / (t.capital || capital) * 100).toFixed(2)%`. Mirrors the pattern already present in the Initial Position card.
- **Trades table — Cur. Allocation % column**: Renamed from "Allocation %" to "Cur. Allocation %" and updated formula to use current remaining invested (`rd.remQty * rd.netBuyPrice`) instead of original `t.invested`. Shows `—` for closed trades (`allocationPct` is `null` when `isClosed`). `curInvested` is computed alongside `allocationPct` in `tradeTableRows`. TD cell renders `row.allocationPct !== null ? row.allocationPct.toFixed(2)+'%' : '—''.'
- **Setup dropdown — VCP60 added**: Added `VCP60` to the new-trade dialog's Setup `v-select` items list (line 799). Items now: `['SVRO3','SVRO15','SNB','VCP60']`. No analytics changes needed — "By Setup" grouping picks up setup names dynamically.
- **Org Pos / Cur Pos card title restructure**: Renamed "Initial Position" → "Org Pos", "Current Position" → "Cur Pos". Title row is now inside `tc-split-main` as a `tc-grid` row: col 1 = title with Pos% inline (no decimals, `Math.round`), col 2 = Qty (using `tc-cell-label` styling to align with data labels below), col 3 = Inv. Data rows use all 3 columns: Org Pos shows Buy Price / Net Buy / Initial SL in row 1, Abs Risk / On Capital / SL% in row 2. Cur Pos shows Exited Qty / Open R / Open P&L in row 1, P&L% in row 2. Open Qty removed from Cur Pos data rows (already shown in title).
- **1R Unit / Current SL aside fixed width**: `.tc-split-aside` changed from `min-width:70px` to fixed `width:130px; min-width:130px; max-width:130px` so the vertical separator aligns consistently across all cards regardless of value width. Sized to fit ₹10,000.00.
- **Current SL edit buttons**: Save/Cancel buttons replaced with green check-circle and red close-circle icons (`mdi-check-circle` / `mdi-close-circle`).
- **Realized card title**: Renamed "Realized & Combined" → "Realized". Title row now shows: `REALIZED ₹capital +R +₹P&L ROCE%` — capital at time of trade, realized R, realized P&L, and ROCE% on capital, all inline and color-coded.
- **Trades table — Capital column**: Added `Capital` column after Cur. Invested, showing `t.capital || this.capital` for each trade. `tradeTableRows` return object carries `tradeCapital`. colspan updated to 13.

## Recent analytics changes (session log)

- **Monthly P&L includes partial exits**: `monthlyPnlData()` now includes realised P&L from partial exit legs on still-open trades, bucketed by each leg's exit date month. Previously it only counted fully closed trades, causing a mismatch with the ribbon's Overall P&L (which includes partial-exit P&L via `getRibbonData()` in the backend). The fix iterates `allTrades.filter(status==='Open')`, calls `getRData(t)` to get `legDetails`, and adds each leg's `legPnl` to the month keyed by `leg.date.slice(0,7)`. Each partial exit leg also increments the month's `count` and `wins` (if `legPnl > 0`).

- **Win/Loss Streaks chart (was "Losing Streak Timeline")**: Replaced the per-trade losing-streak-counter bar chart with a monthly stacked bar chart showing wins (green, bottom) and losses (red, top) per month. New computed `monthlyStreakData()` aggregates win/loss counts per month from `closedTradesSorted`. Chart uses Chart.js stacked bar with `stack:'streak'`. The old `streakTimelineData` computed is retained (still used by `longestLosingStreak` and `thresholdCrossCount` stats).

## How changes are verified

This file is edited live on the user's machine via a device-bridge shell (Python heredoc scripts doing exact-string `content.count(old) == 1` replacements — never blind regex replace-all on ambiguous strings). After every edit to `trading-journal.html`:

1. Extract the `<script>...</script>` block and run `node --check` on it to catch JS syntax errors.
2. Check brace balance (`{` vs `}`) across the whole file.
3. Check HTML tag balance for whatever tags were touched (`<div>`/`</div>`, `<v-card>`/`</v-card>`, `<th>`/`</th>`, `<td>`/`</td>`, `<tr>`/`</tr>`, `<table>`, `<thead>`, `<tbody>`, etc.) — note `<v-card>` is often self-closing/short-form so an open/close mismatch there isn't necessarily a bug, cross-check by reading context.
4. For risky layout/CSS changes (alignment, grid restructuring), build and screenshot a standalone reduced test case with Playwright *before* touching the live file, at both a comfortable and a deliberately narrow width, to confirm no overlap/breakage. This was adopted after a CSS change once broke live rendering (text overlap) — don't skip this step for anything touching flex/grid sizing.

No automated test suite exists otherwise; there's no CI. Changes are validated manually against the live file structure as above, and the user visually confirms in-browser.

## Backend (`trading-journal.gs`)

Google Apps Script v4, ~1193 lines. Deployed as a web app bound to a Google Sheet. Changes here require **redeploying the Apps Script web app** for the frontend to see them — always flag this to the user, it's a manual step.

### Sheet schema

**Trades** (37 cols, A–AK):

| Col | Name | Notes |
|-----|------|-------|
| A | Trade_ID | `T-yyyyMMdd-HHmmss` |
| B | Type | `L` / `S` |
| C | Contract | `CNC` / `MIS` |
| D | Symbol | uppercased |
| E | Lots | |
| F | Lot_Size | |
| G | Buy_Date | yyyy-MM-dd |
| H | Buy_Time | HH:mm |
| I | Buy_Price | |
| J | Initial_SL | |
| K | Setup | |
| L | RVol | |
| M | PreMkt_AD_Ratio | |
| N | Cause_%_Move | |
| O | Consol_%_Move | |
| P | MA_Undercut | |
| Q | Invested_Amount | formula `=E*F*I` |
| R | Position_Size_% | formula using AK capital |
| S | SL_% | formula |
| T | Risk_of_Capital_% | formula using AK capital |
| U | Timestamp | |
| V | Remaining_Qty | decremented on each exit leg |
| W | Status | `Open` / `Closed` |
| X | Exit_Date | |
| Y | Avg_Exit_Price | |
| Z | P&L | |
| AA | P&L_% | decimal (e.g. -0.0524 = -5.24%) |
| AB | Result | |
| AC | LF_Rise | |
| AD | Entry_Chart_URL | Google Drive URL |
| AE | Buy_Charges | from `calcCharges()` |
| AF | Net_Buy_Price | from `calcCharges()` |
| AG | Current_SL | starts as Initial_SL; updated by `updateTradeSL` |
| AH | Entry_Notes | |
| AI | Market_Env | |
| AJ | ROCE_% | decimal |
| AK | Capital | capital snapshot at trade-creation time (from Summary.B2) |

**Trade_Legs** (14 cols): Leg_ID, Trade_ID, Date, Time, Action, Price, Quantity, Amount, Running_Remaining, Timestamp, Exit_Chart_URL, Exit_Charges, Net_Exit_Price, Exit_Notes. Leg_ID format: `{Trade_ID}-X{n}` (0-indexed exit count).

**Capital_Transactions** (7 cols): Txn_ID, Date, Time, Type, Amount, Running_Balance, Timestamp.

**Config** (Key/Value/Description): brokerage and STT rates for MIS and CNC (buy/sell, % and cap), plus `Streak_Threshold`. Loaded by `calcCharges()` at runtime.

**Summary**: formula-only sheet. `B2` = current capital (referenced when snapshotting capital on `addNewTrade`).

### API — `doPost` actions (JSON body with `action` field)

| Action | Function | Key params |
|--------|----------|------------|
| `addNewTrade` | `addNewTrade(data)` | symbol, buyPrice, initialSL, lots, lotSize, buyDate/Time, setup, contract, type, rvol, adRatio, causeMove, consolMove, maUndercut, lfRise, entryNotes, entryChartUrl |
| `addExitLeg` | `addExitLeg(data)` | tradeId, exitPrice, exitQty, exitDate, exitTime, exitNotes, exitChartUrl, contract |
| `getExitLegs` | `getExitLegs(tradeId)` | tradeId |
| `getAllTrades` | `getAllTrades()` | — |
| `getOpenTrades` | `getOpenTrades()` | — |
| `getRibbonData` | `getRibbonData()` | — |
| `getCapital` | `getCapital()` | — |
| `updateCapital` | `updateCapital(amount, status, date)` | amount, status, date |
| `getTransactions` | `getTransactions()` | — |
| `getInvested` | `getInvested()` | — |
| `updateTradeSL` | `updateTradeSL(tradeId, newSL)` | tradeId, newSL — updates col AG |
| `updateTradeFields` | `updateTradeFields(data)` | tradeId + any of: entryPrice, lots, lotSize, rvol, adRatio, causeMove, consolMove, maUndercut, lfRise, entryChartUrl, entryNotes, marketEnv; exitLegs array with legIndex, exitPrice, exitQty, exitNotes, chartUrl |
| `updateConfig` | `updateConfig(key, value)` | key, value |
| `uploadChart` | `uploadChart(imageData, fileName, mimeType)` | base64 imageData, fileName, mimeType |

`doGet` supports: `getRibbonData`, `getCapital` (query param `action`). Default returns `{status:'ok', version:'v4'}`.

### Key function notes

- `calcCharges(rawPrice, qty, contract, side)` — reads live Config sheet for brokerage/STT rates; returns `{brokerage, stt, totalCharges, netPrice}`.
- `addNewTrade` — snapshots current capital from `Summary.B2` into col AK at creation time so allocation % is always relative to capital-at-entry, not current capital.
- `addExitLeg` — decrements `Remaining_Qty` (col V) on the Trades row; marks Status `Closed` and writes exit summary columns (X–AB, AJ) when remaining hits 0. Validates exit qty doesn't exceed remaining. Supports contract override (e.g. CNC→MIS same-day exit).
- `getRibbonData` — computes aggregate stats (open P&L, invested, win rate, streak, etc.) and calls `calcStreaks()`.
- `setupSheet()` — idempotent; safe to re-run any time. Adds missing columns/sheets without destroying data.
- `uploadChart` — stores chart images in a dedicated Google Drive folder (`getOrCreateChartFolder()`); returns a shareable URL stored in col AD (entry) or Trade_Legs col K (exit).

### Migration functions (one-time, already run — do not re-run)

- `migratePnlPctToDecimal()` — converted P&L_% from multiplied-by-100 format to true decimals.
- `migrateBackfillRoce()` — backfilled ROCE_% for closed trades.
- `migrateBackfillCapital()` — backfilled Capital col (AK) with ₹1,00,000 for older trades.

## Session: 15 Sep 2026 — Streak sort bug fix & FOUC fix

### Bug: Winning streak showing 1 despite last trade being a loss

**Root cause**: `calcStreaks()` in `trading-journal.gs` sorted closed trades by `Exit_Date` (column X) only — a date with no time component. When two trades exited on the same date (e.g. DATAPATTNS at 09:39 and EMMVEE at 15:01 on 15 Sep 2026), the sort order was non-deterministic. The winning trade could land first in the sorted array, making the function report `consecWins: 1` instead of `losingStreak: 1`.

**Fix (backend — `trading-journal.gs`)**: `calcStreaks()` now reads exit date+time from `Trade_Legs` sheet to build a `lastExitMap` (keyed by Trade_ID → last exit leg's datetime as timestamp). The sort uses this full datetime, falling back to `Exit_Date` from Trades sheet for trades without legs. Additionally, the `Time` column in `Trade_Legs` is stored by Google Sheets as a **Date object** (e.g. `Sat Dec 30 1899 15:03:00 GMT+0530`), not an "HH:mm" string — the code handles both formats via `instanceof Date` check before extracting hours/minutes.

**Fix (frontend — `trading-journal.html`)**: `closedTradesSorted()` computed property now uses a `lastExitTs()` helper that reads exit date+time from `exitLegsCache[t.id]` (each leg has `.date` and `.time` fields). Falls back to `exitDate`/`buyDate` for trades without cached legs. This fixes the Win/Loss Streaks chart in the analytics tab which depends on `closedTradesSorted`.

**Important note on Trade_Legs Time column**: Google Sheets stores time values as Date objects, not strings. Any future code reading the `Time` column (index 3) from `Trade_Legs` must handle `instanceof Date` (use `.getHours()` / `.getMinutes()`) as well as string format ("HH:mm") as a fallback.

### Fix: FOUC (Flash of Unstyled Content) — raw mustache tags visible on load

**Problem**: Before Vue.js initialized, the browser briefly rendered raw `{{losingStreak}}`, `{{consecWins}}` etc. as plain text.

**Fix**: Added `v-cloak` directive to `<div id="app">` and CSS rule `[v-cloak] { display: none !important; }`. Vue automatically removes `v-cloak` once the app mounts, keeping the page hidden until then.

### Fix: "Not Connected" flash on every page load

**Problem**: `connected` data property starts as `false`. On mount, `refreshAll()` is called which eventually sets `connected = true` after the first API response. Between Vue mounting (removing `v-cloak`) and the API response, the page renders the "Not Connected" state briefly.

**Fix**: In `mounted()`, set `this.connected = true` immediately before `await this.refreshAll()` when a saved URL exists. If the API call fails, `gas()` sets it back to `false`.

### Files changed

- `trading-journal.gs` — `calcStreaks()` rewritten with Trade_Legs time lookup
- `trading-journal.html` — `closedTradesSorted()` updated with exit time sort, `v-cloak` added, optimistic `connected` flag in `mounted()`

### Deployment note

Both `.gs` and `.html` changes require redeployment: paste `.gs` into Apps Script editor → save → Deploy → Manage deployments → edit → New version → Deploy.

## Session: 16 Sep 2026 — Edit popup: editable entry/exit prices & quantity

### Feature: Editable Trade Parameters in Edit Dialog

**What**: The edit popup now allows editing entry price, lots, lot size, and exit leg prices/quantities. When any of these change, all dependent values recalculate automatically — solving the problem where manual sheet edits left formulas and derived fields stale.

### HTML changes (`trading-journal.html`)

- **Trade Parameters section** added at top of edit dialog: three fields — Entry Price (₹), Lots, Lot Size — with a live preview showing Total Qty and Invested as the user types.
- **Exit Legs section** updated: each exit leg now has editable Exit Price and Exit Qty fields (previously these were read-only display text).
- **`openEditDialog(t)`**: now populates `et.entryPrice`, `et.lots`, `et.lotSize` from the trade object.
- **`saveEdit()`**: payload now includes `entryPrice`, `lots`, `lotSize`; each exit leg update includes `exitPrice` and `exitQty`.

### Backend changes (`trading-journal.gs` — `updateTradeFields`)

**Entry price/lots/lotSize editing** (~80 lines added):
- When any of E (Lots), F (Lot_Size), I (Buy_Price) change:
  - Re-applies formulas for Q (Invested `=E*F*I`), R (Position_Size_%), S (SL_%), T (Risk_of_Capital_%)
  - Recalculates AE (Buy_Charges) and AF (Net_Buy_Price) via `calcCharges()`
  - Adjusts V (Remaining_Qty) to preserve exited quantity: `newRemaining = newTotalQty - oldExitedQty`
  - Updates the Buy leg in Trade_Legs (price, quantity, amount, running remaining)
  - For **closed trades**: recalculates Y (Avg_Exit_Price), Z (P&L), AA (P&L_%), AB (Result), AJ (ROCE_%) from exit legs

**Exit leg price/qty editing** (~100 lines added):
- When an exit leg's price or qty changes in Trade_Legs:
  - Updates F (Price), G (Quantity), H (Amount) on the leg row
  - Recalculates L (Exit_Charges) and M (Net_Exit_Price) via `calcCharges()`
  - Recalculates I (Running_Remaining) for all legs of that trade
  - Updates V (Remaining_Qty) on the Trades row
  - For **closed trades**: recalculates Y, Z, AA, AB, AJ (same as above)

**Exit leg notes/chart saving** (previously broken — was in HTML payload but GS ignored it):
- `updateTradeFields` now processes `data.exitLegs` array: writes `exitNotes` to col N and `chartUrl` to col K of each matching sell leg in Trade_Legs.

### Data safety

- Values are always read from the sheet before comparison; recalculation only triggers when values actually differ (`priceChanged` / `exitLegsChanged` flags).
- If only notes/charts are edited, no price recalculation runs.
- Manual P&L overrides in the sheet will be replaced if entry/exit prices are edited on a closed trade (recalculated from exit legs — this is the correct value).

### Files changed

- `trading-journal.html` — edit dialog UI, `openEditDialog()`, `saveEdit()`
- `trading-journal.gs` — `updateTradeFields()` expanded from ~30 lines to ~180 lines

## Session: 17 Sep 2026 — Dialog polish to match wealthlab.in (New/Edit/Exit/View Trade popups)

**What**: The color/token retheme earlier in this session covered cards/ribbon/table but not the four trade dialogs' structural chrome. Reviewed wealthlab.in's marketing screenshots (icon+label section headers, colored pill badges, rounded circular icon badges, gradient accent washes, pill-shaped CTA buttons) — no login access to their actual dashboard modals, so this is a grounded-but-not-pixel-exact match.

**Changes** (all 4 dialogs: New Trade, Edit Trade, Exit Trade, View Trade):
- Title bars: added a circular `.dialog-icon-badge` (36px, rounded-full, tinted bg matching accent) + a subtle diagonal gradient wash behind the whole title bar (`.dialog-title-accent.accent-{blue,red,green,slate}`), replacing the flat `v-card-title` + plain `v-divider`. Colors: New Trade/Edit Trade = blue (`mdi-plus`/`mdi-pencil-outline`), Exit Trade = red (`mdi-logout-variant`), View Trade = dynamic green/red arrow-up/down based on `vt.type` (Long/Short).
- Every `.dialog-section-title` now has a small 14px `mdi` icon prefix (e.g. Trade Info → `mdi-information-outline`, Market Context → `mdi-chart-line`, Deep Dive → `mdi-radar`, Price & Risk → `mdi-shield-alert-outline`, Exit Legs → `mdi-arrow-down-box-outline`, etc.) — CSS updated to `display:flex;align-items:center;gap:6px`.
- Primary footer buttons made `rounded="pill"` with a leading/trailing icon: Submit Trade (arrow-right, trailing), Save Changes (content-save, leading), Exit Position (logout-variant, leading). View dialog's Edit Trade / Exit Trade buttons also pill+icon for consistency.
- Removed the redundant plain `<v-divider>` that used to sit directly under each dialog's `v-card-title` — the new gradient header already carries its own `border-bottom`.
- Exit dialog's title color was hardcoded `#EF4444` (off-token) — corrected to the established red token `#F87171`.

**Verified**: `node --check` on extracted `<script>` block passed; div/v-card/v-card-title/v-card-text/v-card-actions/v-dialog/v-icon/v-btn/v-chip/v-row/v-col tag counts balanced. Visually verified via Playwright-less built-in browser preview (static-preview server + forcing `connected=true` and dialog flags via the Vue proxy for debugging only, not a UI-implementation shortcut) — screenshotted all 4 dialogs including View dialog in both Long/Open (green) and Short/Closed (red) states. No console errors.

**Files changed**: `trading-journal.html` only (CSS block ~104-125, and the 4 dialog templates ~790-1300).
