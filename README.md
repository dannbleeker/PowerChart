# PowerChart

An open-source, think-cell-style chart add-in for PowerPoint, built on **Office.js**.

PowerChart gives you the charts consultants reach for think-cell to make — waterfalls,
Mekko/Marimekko, stacked and clustered columns, 100% charts, lines and areas — with
think-cell's signature annotations (**CAGR arrows, difference arrows, value lines,
column totals, smart segment labels**), inserted into the slide as **native, fully
editable PowerPoint shapes**, never pictures or opaque OLE objects.

![Demo gallery](docs/gallery.png)

## Feature overview

**[User manual](docs/MANUAL.md)** — how to use the pane, the datasheet
conventions, and every option.

| think-cell feature | PowerChart |
|---|---|
| Stacked / clustered / 100% column charts | ✅ |
| Stacked waterfall (multi-series deltas) | ✅ |
| Clustered-stacked charts (blank datasheet rows split stacks) | ✅ |
| Rotation for waterfall & Mekko (not just columns) | ✅ |
| Global label de-collision pass (outside labels never overlap) | ✅ |
| Datasheet formulas (`=B2-B3`, `SUM`/`AVG`/`MIN`/`MAX`, ranges) | ✅ |
| Saved chart templates | ✅ (localStorage) |
| Corporate style file (import/export JSON defaults) | ✅ |
| Same Scale on the current selection | ✅ |
| Scatter partition lines, trend line, group legend | ✅ (`X line`/`Y line`/`Trend` rows) |
| Gantt weeks + weekend shading, section headers, indent, remarks | ✅ (`>` prefix, `Activity \| Owner \| Remark`) |
| Gantt holidays + bracket annotations | ✅ (`Holiday`, `Bracket <label>` rows) |
| Category sorting by total | ✅ |
| Combo secondary (right) value axis | ✅ |
| Label content menu on line & scatter labels | ✅ |
| Connector lines on stacked columns/bars | ✅ (`decorations.connectors`) |
| Single data-point highlight color | ✅ (per-cell `series.colors`) |
| Speech-bubble callouts on a value | ✅ (`decorations.callouts`) |
| Shaded background bands (axis regions) | ✅ (`decorations.bands`, incl. scatter) |
| Exploding pie slice | ✅ (`pie.explode`) |
| Footnote/source line + "100% = N" note | ✅ (`footnote`, `decorations.hundredPercentNote`) |
| Rule-based tables (no side borders, group gaps, totals row) | ✅ (default table style; `{style:"grid"}` keeps the old look) |
| In-cell table effects (harvey balls, trend arrows, semantic colors) | ✅ (`[hb:0.75]`, `[up]`/`[down]`/`[flat]`, `[good]`/`[bad]`) |
| Trend line fit statistics (R², p-value) | ✅ (automatic with the `Trend` row) |
| Pattern/hatch fills on series | ✅ (`series.pattern`, SVG/preview; solid in PPT) |
| Boxplot (five-number rows or raw samples with Tukey whiskers) | ✅ (`boxplot` kind, vertical & horizontal, shared axis across boxes) |
| Radar/spider chart (translucent polygons) | ✅ (`radar` kind; filled in SVG+pptx, outline in live add-in) |
| Heatmap (sequential/diverging color matrix) | ✅ (`heatmap` kind) |
| Map charts as tile-grid cartograms (US / EU / Europe / World) | ✅ (`tilemap` kind, auto-detected layouts) |
| Tufte datamark axes (ticks only, range-frame option) | ✅ (`valueAxis: "datamarks"`, `tickMode: "data"`) |
| Error bars (symmetric or asymmetric) | ✅ (`Error` / `Error+` / `Error-` datasheet rows) |
| Palette from the presentation theme (Accent1-6) | ✅ (pane → "Use deck theme", PowerPointApi 1.10) |
| Cascade / decomposition chart (stage-by-stage breakdown) | ✅ (`cascade` kind) |
| Bullet charts (target ticks + range bands) | ✅ (`Target` row + `decorations.bands`) |
| Combo with clustered or 100% columns | ✅ (`combo.columns`) |
| Bubble size legend (automatic) | ✅ |
| Forecast styling on lines (dashed + hollow markers) | ✅ (`decorations.forecastFrom`) |
| Scatter quadrant matrix (zones + labels in one step) | ✅ (`decorations.quadrants`) |
| Funnel / pyramid chart with conversion rates | ✅ (`funnel` kind) |
| Lollipop, dot-plot, and dumbbell-range styles | ✅ (`decorations.barStyle` on clustered) |
| Gantt progress fill + plan-vs-actual baselines | ✅ (`% Complete`, `Baseline start`/`end` rows) |
| Grouped boxplots (two categorical dimensions) | ✅ (`Min \| 2024`-style row suffixes) |
| Waterfall budget-vs-actual (hatched gap to target) | ✅ (`Target` row on waterfalls) |
| Confidence bands + plan-vs-actual ribbons on lines | ✅ (`Band low`/`Band high` rows, `decorations.fillBetween`) |
| Heatmap marginal totals | ✅ (`heatmap.totals`) |
| Slope chart (before/after, labels both ends) | ✅ (`decorations.slope` on line) |
| Waffle chart (10×10 unit grid, one dominant share) | ✅ (`waffle` kind) |
| KPI / number tile (big number + delta arrow) | ✅ (Elements section, `↓ good` flip) |
| Bar-of-pie breakout (top-N + "Other" detailed) | ✅ (`pie.breakout`) |
| Small multiples (grid of panels, one shared scale) | ✅ (`multiples: {columns?}`) |
| Stepped line & area (staircase, jump before/after/center) | ✅ (`decorations.stepped`) |
| Column gap width & clustered bar overlap (Excel-style) | ✅ (`gapWidth` 0–500, `overlap` −100…100) |
| Butterfly value ticks & gridlines on both flanks | ✅ (`valueAxis`/`gridlines` on `butterfly`) |
| Area charts with negative values (dip below the baseline) | ✅ (`area`, positives up / negatives down) |
| Scatter/bubble trajectory trail (Gapminder-style path) | ✅ (`decorations.trajectory`) |
| Boxplot jittered raw-data dots over the box | ✅ (`boxplot.jitter`, raw-sample mode) |
| Scatter/bubble continuous color scale (a third variable) | ✅ (`Color` row → sequential ramp + gradient legend) |
| Smoothed line charts (Catmull-Rom curves) | ✅ (`decorations.smooth`) |
| Waterfall grouping spacers (section gaps in the bridge) | ✅ (`waterfall.spacerIndices`) |
| Gantt auto-summary bars on section rows | ✅ (`decorations.summaryBars`) |
| Notched boxplots (median confidence interval) | ✅ (`boxplot.notch`, raw-sample mode) |
| Radar min–max "peer range + us" band | ✅ (`decorations.radarBand`) |
| Automatic "Other" bucket for long-tail series | ✅ (`otherBucket: {max}`) |
| Calendar heatmap (weekday × week contributions grid) | ✅ (`heatmap.calendar`) |
| Butterfly with stacked flanks (>1 series per side) | ✅ (`butterfly: {split}`) |
| Radar per-spoke scales (mixed KPI units) | ✅ (`radar: {perSpoke}`) |
| Missing-data bridge on line charts | ✅ (`decorations.bridgeGaps`) |
| Transparent floating column segments | ✅ (series `color: "transparent"`) |
| Combo with a waterfall or Mekko base under the lines | ✅ (`combo.columns: "waterfall" \| "mekko"`) |
| Combo with independent per-line axes (mixed-unit KPIs) | ✅ (`combo.lineAxes: "independent"`) |
| Hexagonal tile maps | ✅ (`tilemap.shape: "hex"`) |
| Per-region mini-glyphs (bars) on tile maps | ✅ (`tilemap.glyph: "bars"`) |
| 100% charts with negative segments (below the axis) | ✅ (`stacked100`) |
| Selection awareness (select a chart → edit banner) | ✅ |
| Insert into selected placeholder bounds; chart size controls | ✅ |
| Datasheet keyboard navigation + insert/delete at cursor | ✅ |
| Auto-update mode (edits push to the slide live) | ✅ |
| Ribbon menu with per-chart-type entries | ✅ (`?kind=` deep links) |
| Dark-mode pane, busy/status feedback, German UI localization | ✅ |
| **Claude Agent Skill** (charts as native shapes from any Claude surface) | ✅ download from the [skill-latest release](../../releases/tag/skill-latest) (rebuilt by CI on every merge), or `npm run skill` locally |
| Claude integration research & plan | 📋 [docs/CLAUDE-INTEGRATION.md](docs/CLAUDE-INTEGRATION.md) |
| Combo chart (stacked columns + line series) | ✅ (`type: "line"` on a series) |
| Pie & doughnut | ✅ (SVG exact; PowerPoint via triangle fans, needs 1.9 rotation) |
| Bar charts as rotated column charts, butterfly charts | ✅ (`horizontal` toggle, `butterfly` kind) |
| Waterfall with computed totals (`e` cells) and connectors | ✅ |
| Mekko (Marimekko) with %-axis and column totals | ✅ |
| Mekko with units (`X extent` row) | ✅ |
| Line & area charts | ✅ |
| Smart segment labels (auto-hidden when they don't fit, auto contrast ink) | ✅ |
| Column totals | ✅ |
| CAGR arrow (`+x.x% p.a.`) | ✅ |
| Total & level difference arrows (%/absolute delta) | ✅ |
| Value lines (multiple; fixed values or mean Ø) | ✅ |
| Series labels placed at the last column (de-overlapped) | ✅ |
| Excel-style datasheet with TSV paste, transpose, `100%=` row | ✅ (task pane grid) |
| Scatter & bubble with collision-avoiding point labels | ✅ (`X`/`Y`/`Size`/`Group` rows) |
| Gantt / timeline with milestones — numeric or **calendar dates** | ✅ (`Start`/`End`/`Milestone` rows; ISO/`dd.mm.yyyy`/"Jan 2026" dates) |
| Value-axis breaks (compress an out-of-scale range) | ✅ |
| Same Scale across all charts in the deck | ✅ (re-renders every chart to a shared scale) |
| Segment order menu (sheet / reversed / ascending / descending) | ✅ |
| Manual value-axis scale (pin min and/or max) | ✅ |
| Number format control (decimals, suffix, locale) | ✅ |
| Label content menu (value / % / series / category combos) | ✅ |
| Palette presets + per-series color pickers | ✅ |
| Axis title, log scale, date-spaced line x-axis | ✅ |
| Difference arrows anchored to value lines, per-series CAGR | ✅ |
| JSON automation (export/import/batch insert + `npm run render` CLI) | ✅ (open take on `.ppttc`) |
| Datasheet undo/redo (Ctrl+Z / Ctrl+Y) | ✅ |
| Agenda / chapter slides (one per chapter, current highlighted) | ✅ (appended to the deck) |
| Harvey balls, checkboxes, process flow, table element | ✅ (Elements section) |
| Gantt: responsible column, dependency arrows, today line, quarter scale | ✅ (`Activity \| Owner`, `After`/`Today` rows) |
| Excel data bridge (selection → chart JSON → paste to refresh) | ✅ (`manifest-excel.xml` companion) |
| Manual label nudging (config-driven) | ✅ (`labelOffsets`) |
| Pie leader lines for outside labels | ✅ |
| Visual-regression snapshots + fuzz tests in CI | ✅ |
| Visual chart gallery (Elements-style thumbnails) | ✅ |
| Output as native, editable PowerPoint shapes | ✅ (grouped) |
| Re-edit inserted charts (config persisted in shape tags) | ✅ ("Edit selected chart") |
| *Live* Excel data links, in-canvas drag handles | 🚧 out of Office.js reach (see docs/RESEARCH.md; the Excel bridge above is the sandbox-safe substitute) |

## How it works

```
datasheet / config ──▶ layout engine (pure TS) ──▶ scene graph ──▶ SVG renderer (preview, tests)
                                                              └──▶ Office.js renderer (native shapes)
```

- **`src/core`** — the layout engine. Pure TypeScript, no Office.js dependency:
  chart layouts, value scales with nice ticks, label fitting/contrast, decorations.
  Deterministic and fully unit-testable.
- **`src/render/svg.ts`** — renders a scene to SVG for the live preview, the demo
  gallery, and visual tests.
- **`src/render/powerpoint.ts`** — renders the same scene as native PowerPoint
  shapes via the Office.js Shape API (`addGeometricShape`, `addLine`, `addTextBox`),
  then groups them. Every bar, label and arrow stays editable in PowerPoint.
- **`src/taskpane`** — the add-in UI: chart gallery, editable datasheet (paste
  straight from Excel), decoration toggles, live preview, insert button.

See [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md) for design details and
[docs/RESEARCH.md](docs/RESEARCH.md) for the deep research on think-cell that
informed this clone.

## Getting started

```bash
npm install
npm run dev        # http://localhost:3000 — demo gallery (no PowerPoint needed)
npm test           # unit tests for the layout engine
npm run build      # production bundle
```

### Try it without PowerPoint

Open `http://localhost:3000/` for the gallery, or
`http://localhost:3000/src/taskpane/taskpane.html` for the full task-pane UI —
outside PowerPoint the **Download SVG** button replaces slide insertion.

### Sideload into PowerPoint

1. Serve the add-in over HTTPS (`npx office-addin-dev-certs install`, then point
   `server.https` in `vite.config.ts` at the generated certs).
2. Sideload `manifest.xml`:
   - **PowerPoint on the web**: Insert → Add-ins → Upload My Add-in.
   - **Windows/Mac desktop**: see the
     [Office add-in sideloading docs](https://learn.microsoft.com/office/dev/add-ins/testing/test-debug-office-add-ins).
3. Open the **PowerChart** button on the Home tab, pick a chart, paste data,
   and hit **Insert into slide**.

Requires PowerPoint with **PowerPointApi 1.4+** (Microsoft 365 desktop or web).
Grouping uses 1.8+ and arrowhead rotation 1.9+ when available; both degrade
gracefully on older hosts.

## Datasheet conventions

- Row 1 = category names, column A = series names — the same mental model as
  think-cell's internal datasheet. **⇄ Transpose** swaps the two.
- Paste a range straight from Excel (TSV) into any cell.
- **Waterfall**: one series of deltas; type `e` (or `=`) in a cell to draw a
  computed running total at that category, exactly like think-cell's `e` cells.
- **`100%=` row**: name a row `100%=` to set per-category denominators for
  100% charts — columns whose series sum to less stay short of full height.
- **`X extent` row**: name a row `X` or `X extent` to give a Mekko explicit
  column widths (think-cell's "Mekko with units").
- **Re-editing**: select an inserted PowerChart on the slide and press
  **Edit selected chart** — the pane reloads its data and "Update chart"
  replaces it in place.

## Disclaimer

PowerChart is an independent open-source project. It is not affiliated with,
endorsed by, or derived from think-cell Software GmbH. "think-cell" is a
trademark of its owner; it is referenced here only to describe compatibility
of concepts.
