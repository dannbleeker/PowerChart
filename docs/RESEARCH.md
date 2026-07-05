# Deep research: think-cell, and how to clone it

This report condenses a fan-out research run (5 search angles, 22 sources fetched,
59 claims extracted, top 25 adversarially verified 3-0 each — 0 refuted) into the
findings that shaped PowerChart. Confidence is high for everything below unless
flagged; caveats at the end.

## 1. What think-cell actually is

- **A native C++ COM add-in, ~1M lines of C++**, for installed desktop
  PowerPoint/Excel only — there is no web/Office.js think-cell. Its own
  engineering page states a dedicated reverse-engineering team (IDA/Hex-Rays)
  built a function-hooking engine that **patches Office executables in memory at
  every startup**, locating assembly signatures to survive Office updates. That
  is why think-cell breaks on new Office builds until compatibility updates ship
  — and why no sandboxed add-in can replicate its integration depth.
  *(Sources: think-cell.com/en/career/tech; corroborated by their KB breakage
  pattern. LOC figure is self-reported.)*
- **The founding thesis is layout, not charting.** think-cell began (2002,
  Fraunhofer FIRST spin-off) as an algorithm for automatic slide layout; the
  company names its hard problems as layout quality, algorithms fast enough for
  interactive editing, and a UI to match. Its core IP is automatic label
  placement (US Patent 7,292,244; Müller & Schödl 2005 on scatter-chart
  labeling). **A clone's differentiating work is the layout/labeling engine,
  not drawing rectangles.**

## 2. Feature model (verified specifics)

- **One unified stacked chart type.** think-cell does not distinguish simple vs
  stacked column/bar charts — a simple chart is a stacked chart with one series.
  Bar charts are *rotated column charts*: users drag a rotation handle to
  convert column↔bar (works for Mekko and waterfall too), and butterfly charts
  are two back-to-back bar charts, one rotated 180°, with the same scale applied.
  → One stacked engine + an orientation transform covers column/bar/butterfly.
- **Mekko math.** Mekko with %-axis = a 100% stacked column chart whose column
  widths are proportional to column totals, so **segment area ∝ absolute
  value**. A second variant, "Mekko with units," takes explicit column widths
  from an `X extent` datasheet row scaled to the chart width.
- **Datasheet semantics.** The internal Excel-like datasheet opens automatically
  on insertion and reopens on double-click; supports **transpose rows/columns**;
  values are absolute by default with each column summing to 100% for percentage
  math; a special **`100%=` row** lets users enter percentages and have absolute
  values derived.
- **Labels and totals are automatic**, placed collision-free at creation with no
  user action; totals can be toggled via "Add Total"; segment stacking order is
  user-controllable (sheet order / reversed / ascending / descending).
- **Axes are direct-manipulation**: auto-scaled by default, rescaled by dragging
  handles at axis ends; **axis breaks** are inserted at the mouse position and
  resized by dragging the compressed range's bounding lines; **Same Scale**
  links scales across multiple charts, re-adjusting all of them whenever any
  chart's data changes.
- **Decorations are live data-bound objects, not static annotations.** Two kinds
  of difference arrows — *level* (between segments/data points) and *total*
  (between column totals). CAGR arrows compute from column totals, taking the
  date range from the datasheet cells behind the category labels. Value lines
  are horizontal lines at a value (several per chart allowed), and difference
  arrows may anchor to segments, data points, **or a value line**. Everything
  recomputes on every data edit.

## 3. Ecosystem: there is no open-source clone

- **No open-source reimplementation of think-cell's rendering exists.** The
  closest project, **ThinkcellBuilder** (Python, GitHub), only automates data
  entry: it emits `.ppttc` files — think-cell's official JSON automation format
  (IANA-registered `application/vnd.think-cell.ppttc+json`, published JSON
  schema) — and still requires a licensed think-cell install to render.
- Lesson 1: the rendering pipeline is closed; a real clone must build its own.
- Lesson 2: the **JSON-in → deck-out automation surface is a proven, valued
  interface** worth replicating (PowerChart's `ChartConfig` is JSON-shaped for
  exactly this reason).
- Commercial alternatives (UpSlide, Empower, Aploris, Vizzlo, Zebra BI) exist,
  but no specific claims about their internals survived verification — treat
  anything written about them as unvetted.

## 4. Office.js feasibility (what PowerChart is built on)

- **Shape construction is fully supported**: `ShapeCollection.addGeometricShape`
  (GeometricShapeType + position/size in points from the slide's top-left),
  `addLine` (ConnectorType; left/top is the start point, width/height the delta
  to the endpoint — negatives allowed), `addTextBox`, plus fill/line formatting
  and text via `textFrame.textRange`. Grouping: `addGroup` / `ShapeGroup.ungroup`.
- **Requirement sets**: shape creation/format/delete arrived only in
  **PowerPointApi 1.4** (Windows 2207+, Mac 16.62+, web; not iPad); **grouping
  needs 1.8**; **1.3 added persistent string tags** on presentations, slides and
  shapes — the mechanism for storing a serialized chart model on the generated
  group so charts stay re-editable. `insertSlidesFromBase64` (1.2) is a limited
  OOXML fallback for older hosts, but without in-place editing.
- **Key limitation**: Office.js lines are positioned geometry only — there is no
  API to glue endpoints to shape connection sites (unlike VBA's `AddConnector`),
  so "universal connectors" require the add-in to maintain its own anchor logic.
  No freeform/path shapes either (hence triangle-shape arrowheads).
- **Verified architecture conclusion**: render charts as grouped native shapes,
  persist the data model in tags, recompute layout and re-emit shapes on every
  edit. This is exactly PowerChart's pipeline.

## 5. How the findings map to PowerChart

| Finding | Status in PowerChart |
|---|---|
| Layout/labeling engine is the core IP | Pure-TS engine in `src/core`, unit-tested |
| Native grouped shapes + tags persistence | `src/render/powerpoint.ts` (tags: `POWERCHART_CONFIG` on the group) |
| Unified stacked model | `layoutColumns` handles stacked/clustered/100% from one code path |
| Mekko: width ∝ total, area ∝ value | `layoutMekko` implements the %-axis variant |
| Waterfall with computed totals | `e` cells → running-total bars + dashed connectors |
| Live decorations from anchors | CAGR/difference/value-line computed from `LayoutAnchors`, re-derived on every render |
| Datasheet: auto data grid + paste | Task-pane grid with Excel TSV paste |
| JSON automation surface | `ChartConfig` is plain JSON in/out |

**Parity ledger (final).** Everything feasible from the research is built —
rotation (column↔bar for columns/waterfall/Mekko/boxplot), butterfly, Mekko
with units, datasheet transpose + `100%=`, level/total/value-line-anchored
difference arrows, multiple value lines, pinned scales, axis breaks, Same
Scale, segment order, Gantt, agenda, plus chart types think-cell lacks
(boxplot raw-sample mode, radar, heatmap, tile-grid maps). What remains
un-cloneable from a sandboxed add-in is the trio think-cell memory-patches
Office for: live Excel data links, in-canvas drag manipulation, and the
slide-layout engine. The README feature table is the authoritative list.

## Caveats

- think-cell.com blocks automated fetches (403), so manual quotes were verified
  via search-index snippets; some cited paths may have drifted.
- The C++/reverse-engineering claims come from think-cell's own recruiting pages
  (self-reported, behaviorally corroborated).
- Gantt/agenda/process-flow, smart-labeling internals, Excel data links, and
  commercial-competitor details produced no *verified* claims — absence of
  verification, not absence of the features.
- How think-cell physically represents charts inside `.pptx` (grouped shapes vs
  OLE vs custom XML parts) remains an open question; likewise the exact
  algorithms behind interactive-speed label placement, and how Excel data links
  could work from a sandboxed PowerPoint add-in.

## Primary sources

- think-cell manual: column/line/area, Mekko, axes, chart decorations, data
  entry pages (think-cell.com/en/resources/manual/…)
- think-cell engineering: think-cell.com/en/career/tech
- Microsoft Learn: PowerPoint add-in shapes guide, `ShapeCollection` /
  `ShapeGroup` API reference, PowerPoint API requirement sets, tags guide
- github.com/Philistino/ThinkcellBuilder; static.think-cell.com/ppttc/ppttc-schema.json
- Slide Science tutorials (UX corroboration); Peltier Tech (Marimekko geometry)
