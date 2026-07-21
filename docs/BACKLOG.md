# Backlog

Curated candidate work, from a research sweep (July 2026) comparing every
existing chart kind against think-cell, Excel, Highcharts, Datawrapper, and
Mekko Graphics, plus a chart-type survey across the Zelazny / FT Visual
Vocabulary taxonomies and competitor add-ins (Zebra BI, Vizzlo, UpSlide).

**This is the only backlog document.** Items graduate from here into PRs and are
deleted when they ship — what has shipped is recorded by the README feature
table and by git, not here. Rejected ideas stay in §3 so they aren't
re-proposed.

Feasibility is judged against the live-add-in constraint: rects, lines, text,
ellipses, and any of PowerPoint's 177 preset geometries (all of them
PowerPointApi 1.4), plus polygon *outlines* — no freeform curves, and no
images. The SVG and skill-pptx renderers additionally have filled polygons and
patterns.

## 1. Open

- **SVG/PNG image output as an alternative render mode** — offer, beside the
  native-shapes renderer, a mode that inserts the whole chart as a *single image
  object*. The scene→SVG renderer already exists (`src/render/svg.ts`) and the
  canvas rasterisation primitive now ships too (the pane's PNG download does
  `drawImage`→`toDataURL` on the preview SVG), so the pixels are free on both
  paths; the remaining work is the *insert* side. The scene graph is pure
  geometry/text with no `foreignObject` or remote assets, so the canvas stays
  untainted. The motive is performance/reliability, not looks — one object
  sidesteps the PowerPoint-**web** wall where the live canvas stops answering past
  ~20 shapes per `context.sync()` (office-js #4272, #5022, #6498), which is
  exactly what caps the densest kinds today (`DEMO_SHAPE_BUDGET = 90`,
  `SHAPES_PER_SYNC = 10`). Insert paths, all shipped: `setSelectedDataAsync(png,
  {coercionType: Image})` (ImageCoercion 1.1, widest host reach) or
  `addGeometricShape` + `ShapeFill.setImage(png)` (PowerPointApi 1.8) — the latter
  returns a tracked `Shape`, so the `POWERCHART_CONFIG` tag round-trips onto the
  picture exactly as it does now, keeping the chart re-editable. Native-vector SVG
  via `setSelectedDataAsync(svg, {coercionType: XmlSvg})` (ImageCoercion 1.2) is
  the best quality/size answer — PowerPoint can even "Convert to Shape" it — but
  it's flaky (office-js #4967 open: vertical shift; #2881: complex SVG; #412: no
  iOS), so gate it behind a flag with a PNG fallback. The cost is real: an image
  is not editable in PowerPoint, ignores theme colours, and blurs on rescale
  unless rasterised high-DPI — so pair it with an **"explode to native shapes"**
  command that reads the tag, deletes the picture, and draws the real shapes on
  demand (paying the web cost only when the user actually edits). This is distinct
  from the rejected per-point **image / icon node** in §3: that was one bitmap
  *per data point* inside the scene graph, unreachable at the 1.4 pin; this is one
  image for the *whole chart* as an output format, with the `1.8`/ImageCoercion
  gates degrading to the native-shapes path on older hosts (like grouping
  already does). Feasibility: medium — needs Phase-2 real-host validation first,
  since the `setSelectedDataAsync` base64 size cap (office-js #225, fixed, no
  published threshold) and the shape-tag value-size limit are both undocumented.

Otherwise nothing new: the second sweep's six candidates all shipped (grand total
label, IBCS scenario notation + the stroke-only hollow-column primitive, IBCS
variance tier, polynomial scatter trendlines, PNG export, copy-config-as-URL) —
see the README feature table and git for what landed. Whatever else surfaces
starts from a fresh research pass.

**Considered and dropped in that sweep** (so they aren't re-proposed): a
screen-reader data-table alternative to `describeChart` — real WCAG best practice
beyond ~6 points, but the primary output is native PPT shapes where alt text is
linear only, so the gain is confined to the downloaded SVG; **CSV file import** —
already covered by the datasheet's TSV clipboard paste, which is what
Excel/Sheets put on the clipboard; **variance/integrated bars as a new kind**,
tornado, icicle, fan, Venn — recipes of shipped kinds or off-genre/curve-bound
(see §3's standing rejections of recipe-of-existing-kind proposals).

## 2. Defects from the July 2026 adversarial bug hunt

A 12-lens hunt (each finding independently re-verified by an adversarial
verifier that defaulted to REFUTED) over the areas the #174-#185 review did not
cover. Every item below was **reproduced by execution** against the real source —
the file:line and the observed wrong output are recorded so each can be picked up
cold. The percent-as-date / ISO-date-time / blank-aggregate-arg defects this hunt
found are already fixed (#187), as are the pie/treemap/violin theme defects (#186).

### High (10)

- **Axis tick labels lose decimals on a narrow-range axis: resolveFormat picks precision from tick magnitude, not tick step → duplicate and wrong labels** — `src/core/format.ts:72`
  Observed: Value-axis text nodes = ["99","100","100","101","101","102"] for the tick set [99,99.5,100,100.5,101,101.5] — "100" and "101" each appear twice at different heights, and the 101.5
- **Baseline (plan-vs-actual) rows are excluded from the timeline extent, so the ghost bar is drawn off the plot and off the canvas** — `src/core/layout/gantt.ts:140`
  Observed: The `all` array at gantt.ts:140-147 folds in starts/ends/milestones/today/holidays/brackets but NOT baseStarts/baseEnds, so lo/hi (and hence t0/t1) ignore the baseline. sceneToSvg
- **In-place update orphans the whole chart when it is not grouped — shape count grows every edit** — `src/render/powerpoint.ts:270`
  Observed: groupAndTagAll (line 918 / fallback line 947) anchors the config tag on `created[0]` — ONE of the chart's 13 shapes — because there is no group to hold it. updateChartsInSlides the
- **CAGR control silently drops decorations.cagr.series — the arrow then prints a different growth rate** — `src/taskpane/app.ts:728`
  Observed: decorations.cagr goes from {"from":0,"to":2,"series":0} to {"from":0,"to":2"} (PROBE-K). The rendered cagr-label text changes from "+100.0% p.a." (series 0: 10→40 over 2 periods) t
- **Hex tilemap renders invisible (white-on-white) in the live Office.js add-in — the whole value encoding is lost** — `src/core/layout/tilemap.ts:128`
  Observed: The scene nodes are correct — tile-CA {kind:"polygon", fill:"#e2e7f6", stroke:"#ffffff"}, tile-TX {fill:"#2a78d6", stroke:"#ffffff"} — but Office.js emits 6 edge shapes per tile an
- **Calendar heatmap legend runs off the bottom and right of the frame for any date range under ~13 weeks** — `src/core/layout/heatmap.ts:563`
  Observed: legend-step-0 is {x:22, y:283, w:29.94, h:29.94} → bottom = 312.94 on a 300-tall chart, and sceneToSvg emits viewBox "0 0 480 300" with the lowest rect bottom at 310.4, so the lege
- **Value-axis tick labels collapse to duplicates because tick precision is derived from tick MAGNITUDE, not tick STEP** — `src/core/layout/frame.ts:345`
  Observed: The five value-axis text nodes read ["7.4","7.5","7.5","7.5","7.5"] — 5 distinct gridlines (7.44/7.45/7.46/7.47/7.48), only 2 distinct labels, and the top tick 7.48 is printed as "
- **Shipped render-svg.mjs crashes on a kind:"agenda" config and throws away every SVG in the batch** — `scripts/render-batch.mjs:34`
  Observed: Unhandled `TypeError: Cannot read properties of undefined (reading 'hundredPercent')` from buildChart, process aborts on config #1, out/ contains ZERO svg files — the waterfall and
- **A "transparent" series colour (the documented floating-segment idiom) renders as an opaque grey block in the pptx deck and hands an invalid colour to Office.js** — `src/core/layout/mekko.ts:70`
  Observed: SVG emits `<rect ... fill="transparent" ... data-name="seg-0-0"/>` (invisible = correct). Rendering the same config through skill/scripts/render-pptx.mjs produced slide1.xml with t
- **Hex tile map renders as ~360 white-on-white shapes in the live add-in — the entire value colouring is lost** — `src/core/layout/tilemap.ts:124`
  Observed: Each tile is a PolygonNode with the data-carrying `fill` and `stroke: style.background` ("#ffffff"). Office.js has no freeform fill, so powerpoint.ts:1198 draws the outline with `n

### Medium (31)

- **⇄ Transpose does not remap A1 references, so formula cells silently compute different numbers after the swap** — `src/taskpane/datasheet.ts:231`
  Observed: Before: series North=[10,20], South=[30,40], Total=[40,60]. Transposed cells: [["","North","South","Total"],["Q1","10","30","=SUM(B2:B3)"],["Q2","20","40","=SUM(C2:C3)"]] — the for
- **A comma-decimal number pasted from a European Excel is silently mangled ("1.234,5" → 1.2345, a 1000× error)** — `src/taskpane/datasheet.ts:191`
  Observed: value = 1.2345. Also "1,5" → 15 and "1 234" (narrow/space group separator) → null. The parser does Number(raw.replace(/,/g,"")), which turns "1.234,5" into "1.2345".
- **formatPercent ignores numberFormat.locale (and has no thousands grouping) — every percent label is en-US inside a localized chart** — `src/core/format.ts:76`
  Observed: One chart, two number systems. Funnel stage values (formatNumber, locale-aware) = ["12.000","4.300","1.250","337"], but the conversion labels (formatPercent) = ["▾ 35.8%","▾ 29.1%"
- **niceTicks .toPrecision(12) collapses every tick to one value past 12 significant digits — degenerate axis, top tick below the data max** — `src/core/format.ts:113`
  Observed: Value-axis labels = ["10,000,000,000,000","10,000,000,000,000","10,000,000,000,000","10,000,000,000,000","10,000,000,000,000"] and all five gridline y1 values = 302 (every gridline
- **Gantt never reserves the footnote row, so the footnote text prints on top of the Today label and the last task row** — `src/core/layout/gantt.ts:131`
  Observed: gantt.ts:131-137 computes bottomH = today != null ? fs*1.6 : 6 and never subtracts footnoteH(cfg, style, decor) (every other footnote-aware layout — pie, radar, funnel, cascade, he
- **A zero-length (single-day) or reversed task draws nothing at all — the whole activity vanishes with no marker** — `src/core/layout/gantt.ts:479`
  Observed: The `e > s` guard at gantt.ts:479 drops both. Rendered bar nodes were only ['bar-2']; Cutover and Review produced no bar, no hairline, no section band (`sections: []`) — they occup
- **In-cell effect glyphs ([up]/[hb:]) land at the cell's LEFT edge while the cell text is right-aligned — the glyph is orphaned from its value and reads as belonging to the previous column** — `src/core/elements.ts:381`
  Observed: cellEffectNodes is always called at `x + 5` (the cell's left edge) but every column except column 0 sets align:"right" (line 393). Measured ink spans in the produced scene: the [up
- **Arrowhead glyphs are anchored by their TIP on the text's vertical centre, so an [up] arrow hangs entirely below the line and a [down] arrow entirely above it (7.6–8.1pt apart)** — `src/core/elements.ts:309`
  Observed: The arrowhead node's (x,y) is its TIP in all three renderers (svg.ts:175 draws `M 0 0 L -1.8s ...` translated to (x,y); geometry.ts:41 offsets the Office box so the tip lands on it
- **Process-flow labels never shrink: from ~10 steps adjacent labels overlap each other and from 13 steps the first label's ink starts at a negative x (outside the scene)** — `src/core/elements.ts:96`
  Observed: fontSize is fixed at Math.min(11, height*0.3) with no fit-to-width pass (contrast buildKpiTile:137, which shrinks its value to fit). Measured: n=10 → label box 40.2pt but "Discover
- **numberFormat.decimals outside the pane's 4 options is rewritten to 0 the moment the user edits the suffix box** — `src/taskpane/app.ts:874`
  Observed: exportConfig().numberFormat goes from {"decimals":3,"forceSign":true} to {"decimals":0,"suffix":"%","forceSign":true} (PROBE-E). currentConfig does Number("") === 0, so every label
- **?tab= deep-link value is interpolated raw into a CSS selector — a quote throws and aborts the rest of pane boot** — `src/taskpane/app.ts:1815`
  Observed: Uncaught `SyntaxError: Invalid selector .tabs .tab[data-tab="chart"]"]` at module top level, which aborts everything after line 1815 (PROBE-A1): #build-stamp stays "" (never stampe
- **Renaming a series in the datasheet silently drops its colour, combo type, pattern, per-point colours and IBCS scenario** — `src/taskpane/app.ts:236`
  Observed: exportConfig().data.series goes from [{"name":"Rev",...},{"name":"Margin %",...,"color":"#ff0000","type":"line","pattern":"dots","scenario":"FC"}] to [{"name":"Rev",...},{"name":"M
- **Tilemap mini-glyph bars clamp negative values to zero height — a decline is indistinguishable from no change** — `src/core/layout/tilemap.ts:147`
  Observed: glyph-TX-0 h=0.00 (value -20) and glyph-NY-1 h=0.00 (value -5), while glyph-CA-0 h=5.30 (value 10) and glyph-NY-0 h=15.91 (value 30). The two negative bars render as zero-height re
- **Waffle silently truncates cells when the 100%= denominator is smaller than the parts, giving unequal areas for equal values** — `src/core/layout/waffle.ts:68`
  Observed: Cells by fill: Region A = 62 cells, Region B = 38 cells — two identical values (50 and 50, both 62.5% of the 80 denominator) drawn as wildly different areas — while the legend labe
- **logScale axis floor is always dataMax/1000 (the `dataMin > 0` branch is unreachable), adding a phantom sub-1 decade that prints as "0"** — `src/core/layout/frame.ts:58`
  Observed: Axis labels are ["0","1","10","100","1,000"] — a log axis with a tick labelled "0", which is impossible. The tick is really 0.1: `valueScale(frame,0,300,…,log)` returns ticks [0.1,
- **Radar's private legend never wraps and no other kind's legend behaves that way — chips and labels are emitted outside the canvas** — `src/core/layout/radar.ts:236`
  Observed: legend-chip-8 is a rect at x=483.8..490.8 and legend-8 ("Austria") is a text at x=493.8..537.6 on a 480pt-wide canvas — both entirely outside the chart shape (with 10 series, 4 nod
- **A manual scale whose min is at or above the auto max collapses the axis span to zero, and the `|| 1` fallback renders columns ~83 canvas-heights off the shape** — `src/core/layout/frame.ts:79`
  Observed: After tick filtering, `max` collapses onto `min` (ticks filtered to [100], so min=100 and max=ticks[last]=100), `max - min || 1` becomes a 1pt-per-unit axis, and the columns are em
- **Named CSS colours collapse to one mid grey in the headless pptx — two differently-coloured series become indistinguishable** — `skill/scripts/render-pptx.mjs:139`
  Observed: pptx slide1 srgbClr set = ['0B0B0B','808080','A5A49E','FFFFFF'] — BOTH series render as the same 808080 grey, so the legend colours and the bars no longer distinguish Alpha from Be
- **style.fontFamily is documented in the skill reference but is a complete no-op in every renderer** — `skill/reference.md:157`
  Observed: The generated pptx contains only `typeface="Segoe UI"` (render-pptx.mjs:213 hardcodes the fallback because no layout ever sets TextNode.fontFamily). A byte-comparison probe on the
- **Headless pptx renderer hardcodes a white slide background, so a dark-styled chart exports white-on-white** — `skill/scripts/render-pptx.mjs:370`
  Observed: slide1.xml carries `<p:bg><p:bgPr><a:solidFill><a:srgbClr val="FFFFFF"/>` while the title/labels/axis ink is `srgbClr val="FFFFFF"` — the chart's whole text layer is white on a whi
- **Office.js drops the chart's alt text entirely on any host without PowerPointApi 1.8 (the documented ungrouped degradation)** — `src/render/powerpoint.ts:923`
  Observed: 13 native shapes are created and ZERO carry altTextDescription/altTextTitle (probe output: "shapes created = 13 groups = 0", "shapes carrying altTextDescription = 0", "shapes carry
- **pptx exporter emits no alt text at all — undocumented break of the SVG/Office.js/pptx parity contract** — `skill/scripts/render-pptx.mjs:174`
  Observed: `descr=` occurrences: 0, `title=` occurrences: 0; every shape is `<p:cNvPr id="n" name="Text 0">` with no alt text. The other two renderers DO carry it: svg.ts:69-70 emits <title>/
- **Gantt bar label ink is hardcoded #ffffff instead of contrastInk(fill), so it vanishes on a light palette colour** — `src/core/layout/gantt.ts:522`
  Observed: bar-0 fill = "#ffe066", bar-label-0 color = "#ffffff" → contrast 1.30:1 (WCAG needs 4.5:1). The identical fill in the column family picks the opposite ink: a stacked chart with the
- **Gantt percent-complete overlay is a hardcoded #1b4e8a, ignoring the palette — it is not "a darker inner bar" and reads as a foreign series** — `src/core/layout/gantt.ts:506`
  Observed: progress-0 fill = "#1b4e8a" in all three runs (default palette, red palette, dark background) while bar-0 fill tracks the palette (#2a78d6 / #e34948). With the red palette a BLUE b
- **sequentialScale / divergingScale / NO_DATA hardcode a white canvas, so heatmap and tilemap are the only surfaces that ignore dark theme** — `src/core/color.ts:148`
  Observed: Heatmap cell fills are byte-identical on both themes: ["#f1f4fb","#b3c3e9","#2a78d6","#e4e9f7","#9fb4e5","#6590db"]. On the dark canvas the lowest cell #f1f4fb is a near-white bloc
- **The pane preview, Download SVG and Download PNG force a white canvas, so a dark-theme config previews as unreadable and the preview contradicts the deck** — `src/taskpane/app.ts:1080`
  Observed: sceneToSvg(scene, {background:"#ffffff"}) paints `<rect width="100%" height="100%" fill="#ffffff"/>` under a scene whose title node colour is #f2f1ec → contrast 1.13:1 against the
- **legendRow drops Series.pattern (and the IBCS scenario restyle), so the legend key cannot be matched to the bars it explains** — `src/core/layout/column.ts:962`
  Observed: Legend chips: [{f:"#2a78d6"},{f:"#2a78d6"}] — two identical solid squares, no pattern field. The plotted segments are [{seg-0-0 f:#2a78d6},{seg-1-0 f:#2a78d6 p:diagonal}] — the SVG
- **radar perSpoke: the `Math.max(1, …)` floor defeats per-spoke normalisation for any spoke whose maximum is below 1** — `src/core/layout/radar.ts:49`
  Observed: spokeMax = Math.max(1, ...values) floors the divisor at 1, so the "Conversion rate" spoke normalises against 1 instead of against its own max 0.5. Measured rim radius = 131.0; draw
- **candlestick columnTop falls back to toY(0) on a blank High cell, throwing decorations thousands of points off-canvas** — `src/core/layout/candlestick.ts:85`
  Observed: `columnTop: scale.toY(high[c] ?? 0)` maps the missing cell to value 0, but the OHLC scale is deliberately zero-free (zeroFloor:false, domain ~99–106), so columnTop[1] lands ~3490.
- **violin columnTop uses Math.max(...[0]) for an empty category, putting decorations 600pt below the canvas** — `src/core/layout/violin.ts:128`
  Observed: `scale.toY(Math.max(...(samplesOf(c).length ? samplesOf(c) : [0])))` substitutes value 0, but the violin scale is data-driven (zeroFloor:false, domain 50–70), so columnTop[1] ≈ 950
- **radar's hand-rolled legend never wraps (both the polygon and the radial-bar variant), running 72pt off the canvas** — `src/core/layout/radar.ts:274`
  Observed: drawLegend advances `x += chip + 3 + textWidth + 12` with no maxX check, so legend-5 ("Middle East") spans x 476.4→552.6 on a 480pt canvas — 72.6pt past the right edge (clipped in

### Low (18)

- **Calendar-Gantt rows other than Start/End/Milestone are written back to the datasheet as raw epoch-day integers** — `src/taskpane/datasheet.ts:148`
  Observed: cells = [["","A","B"],["Start","2026-01-05","2026-02-01"],["End","2026-01-23","2026-03-01"],["Today","20486",""],["Baseline start","20454",""],["Holidays","20459",""],["After","","
- **Datasheet round trip re-writes Today / Holiday / Baseline / Bracket dates as raw epoch-day integers** — `src/taskpane/datasheet.ts:148`
  Observed: GANTT_DATE_ROW = /^(start|end|milestone)$/i, so only those three rows are re-serialised as ISO. The generated sheet was: ["Start","2026-01-05","2026-02-02"], ["End","2026-01-30","2
- **Dependency arrow is silently dropped when the predecessor is a milestone row** — `src/core/layout/gantt.ts:595`
  Observed: gantt.ts:595-597 reads only `ends[p]`, which is null for a milestone-only row, and returns. No dep-v/dep-h/dep-head nodes were emitted at all (`C dep nodes []`), so the arrow the a
- **Gutter "Column" rows are not budget-checked, so enough of them drive plot.w negative and the bars collapse to zero width left of the plot** — `src/core/layout/gantt.ts:132`
  Observed: Each column is capped at cfg.width*0.12 but their count is not, and catW takes another 0.32*width, so plot = {x:440.4, y:16, w:-97.86, h:218}. Bars came out as {"x":440.4,"w":0} an
- **A long agenda chapter title is never shrunk or wrapped and runs off the right edge of the slide** — `src/core/agenda.ts:94`
  Observed: fs is derived only from the chapter COUNT (agenda.ts:32), never from title length. The chapter text node is x=133.2 w=740.4 (box ends 873.6) at fontSize 18 bold; a 95-char title's
- **buildTableScene([]) returns a zero-height scene whose bottom rule is drawn at y=−0.5, above the top of the scene** — `src/core/elements.ts:403`
  Observed: Scene height is 0 and sceneToSvg emits `<svg width="480" height="0" viewBox="0 0 480 0">` containing `rule-top` at y=0.5 and `rule-bottom` at y=−0.5 — the bottom rule is ABOVE the
- **updateChartsInSlides resurrects a chart whose shape the user deleted** — `src/render/powerpoint.ts:279`
  Observed: Phase 1 filters out targets whose SLIDE is gone (`live`, line 260) with the explicit rationale "a chart whose slide is gone is not an error, it is nothing to do", but a target whos
- **"+ Row" inserts a fully blank row, which the model reads as a stack separator — the stacked chart instantly splits into two stacks** — `src/taskpane/datasheet.ts:312`
  Observed: exportConfig().data.series immediately becomes [{"name":"S1","values":[1,2],"stack":0},{"name":"S2","values":[3,4],"stack":1}] (PROBE-N) — sheetToData's blank-row rule (datasheet.t
- **detectLayout has no runner-up margin despite its docstring, so codes valid in two grids silently pick the US map** — `src/core/layout/tilemap-layouts.ts:178`
  Observed: 51 tiles are emitted, i.e. the US grid: Germany's value is painted onto Delaware, Malta's onto Montana, and so on — a European dataset rendered as a US cartogram with no warning. s
- **rgba() with a percentage alpha makes the pptx scale the RGB channels by 2.55, turning any colour white** — `skill/scripts/render-pptx.mjs:132`
  Observed: pptx bar fill = FFFFFF with alpha 50% (pure white). `/%/.test(m[1])` tests the WHOLE argument list, so the `50%` alpha flips the RGB scale factor to 2.55: 100*2.55=255, 150*2.55=25
- **reference.md's style schema omits five shipped, output-affecting ChartStyle ink fields** — `skill/reference.md:157`
  Observed: A probe that renders the same waterfall with each field overridden shows text, mutedText, axis and gridline each CHANGE the output (and background changes it on stacked charts — ve
- **Shipped skill's render-svg.mjs prints a usage line naming a file that does not exist in the package** — `scripts/render-batch.mjs:18`
  Observed: The script prints `usage: node scripts/render-batch.mjs <charts.json> [outDir]`, and its header comment (line 7) likewise says `node scripts/render-batch.mjs examples/charts.json o
- **Chart alt text (#137/#129) is silently dropped by the headless pptx renderer** — `skill/scripts/render-pptx.mjs:373`
  Observed: grep over skill/scripts/render-pptx.mjs finds no altText/descr usage, and no `descr=` attribute appears in any slide XML of the 121-slide deck I rendered from showcase.json. The SV
- **An untitled chart renders role="img" with a <desc> but no accessible name** — `src/render/svg.ts:62`
  Observed: Emitted markup: `<svg … role="img"> <desc>stacked column chart. 3 data series: …</desc> <text …>`. No <title>, no aria-label, no aria-labelledby. Under the SVG accessibility mappin
- **butterfly's stacked-flank legend is a third non-wrapping copy and overflows the canvas by 25pt** — `src/core/layout/butterfly.ts:112`
  Observed: The stacked-legend loop advances `lx += chip + 3 + textWidth(s.name, fs) + 12` unbounded: legend-4 ("Latin America") spans x 429.0→505.2 on a 480pt canvas (25.2pt overflow, clipped
- **butterfly plot height is never floored, so a short chart emits negative-height bar rects (SVG drops them, Office.js clamps — canvas ≠ export)** — `src/core/layout/butterfly.ts:39`
  Observed: `plot.h = cfg.height - titleH - headerH - 6 - axisH` goes negative (-7.3 here), so slotH and barH follow: every seg rect gets h=-2.44. Emitted SVG is `<rect x="69.5" y="39.39" widt
- **funnel prints a down-arrow and a >100% conversion for the ascending ordering its own doc comment recommends** — `src/core/layout/funnel.ts:96`
  Observed: The between-stage labels read "▾ 500.0%" and "▾ 0.0%". The ▾ glyph asserts a drop while the number states a 5x increase — the label contradicts itself on a documented, supported or
- **waterfall: any delta whose incoming running total happens to be 0 loses its "+" sign** — `src/core/layout/waterfall.ts:194`
  Observed: `const floating = !b.isTotal && b.segs[0]?.from !== 0;` treats the third column as a base column because its stack starts at 0, so forceSign is suppressed and the labels come out "

## 3. Rejected or already covered (do not re-propose)

- **An image / icon node** — not reachable in the live add-in, so nothing can
  be built on it. PowerPoint's `ShapeCollection` exposes exactly
  `addGeometricShape`, `addGroup`, `addLine`, `addTable`, `addTextBox`;
  `addImage` exists only in the Excel namespace. The one route to pixels is
  `ShapeFill.setImage`, which is **PowerPointApi 1.8** against a manifest
  pinned at **1.4** — it would vanish silently on any older host. It also
  bloats twice over: `ChartConfig` round-trips through a shape tag verbatim
  (`tagData: JSON.stringify(cfg)`), and PptxgenJS does not dedupe media, so an
  identical icon is re-embedded once per point. The only thing presets cannot
  do is a **logo** (a competitor-positioning map). If that is ever wanted, the
  shape is: user-supplied data URI (no asset library, so no licensing call),
  gated on `supports("1.8")` like grouping already is, and drawn *additively
  over* a marker so an old host degrades to a plain point rather than a
  missing one.
- **Scatter/bubble point icons** — the primitive above is unavailable, and the
  real itch was not bitmaps: shape as a categorical channel is what Excel and
  Highcharts mean by point markers, and PowerPoint's presets give it filled and
  native at 1.4. Shipped as `scatter.markers`.
- **Heatmap per-cell icon overlays** — a heatmap's rows ARE its data series, so
  a "Glyph" row yields one row of glyphs, not a per-cell matrix; carrying a
  genuine second dimension would mean inventing a datasheet convention for
  every heatmap, for a want nobody has stated. The demonstrable gap was
  narrower and is fixed: a diverging scale states direction in hue alone, and
  its strongest + and − sit at 1.12:1 in greyscale — the same tone — with no
  label to fall back on under `sizeEncode`. Shipped as `heatmap.symbols`.
- **An X / cross marker, and a five-point star** — a marker shape has to
  reproduce its OOXML preset *exactly*, because the SVG renderer draws the
  points while the PowerPoint renderers name the preset, and `markerScale`
  measures the area off those same points. A shape that only approximates its
  preset therefore breaks the bubble's "area ∝ size" claim in the deck while
  keeping it in the preview — the worst kind of divergence, since the preview
  looks right. `mathMultiply` fails on angled arms (and is redundant with
  `plus` at 3-4pt). `star5` fails on its `hf`/`vf` stretch: the preset widens
  itself 1.05146x and heightens itself 1.10557x so the star fills its box
  (a 5-point star spans 1.902R by 1.809R, and those factors are exactly
  2/1.902 and 2/1.809), which makes PowerPoint's star **16.2% larger in area**
  than an inscribed SVG one. Reproducing that faithfully is possible on paper
  but not verifiable here without a PowerPoint rasteriser, so the set stops at
  the shapes whose geometry is exact: circle, square, diamond, triangle, plus.

- **Waterfall connector re-routing** (drag to skip columns) — the drag is out of
  Office.js reach, and the rendering feature underneath is ill-posed: the
  connector is not an object with authored endpoints, it is a derived assertion
  that "this level carries into the next bar". Re-routing it is truthful only
  where the skipped columns leave the level unchanged — and there it already
  works, via `spacerIndices` and `totalIndices`. Anywhere else the line would
  end in mid-air pointing at a bar it does not touch. What was actually wanted
  shipped instead as `waterfall.detailGroups`.
- **Scatter-on-combo** — a continuous-x scatter over categorical columns needs a
  second x scale that means nothing beside category slots. The coherent reading
  — unconnected marks at category positions — shipped as `type: "marker"`
  (overlay) and already existed as `decorations.barStyle: "dot"` (clustered).
- **Free 2D bubble repulsion** — both axes carry data, so moving a marker in 2D
  corrupts two readings at once with nothing to bound them. The honest version
  shipped as `scatter.spread`: one named axis, hard-capped, cap printed in the
  footnote.
- **Gantt resource capacity vs load** — a per-resource per-week histogram against
  a capacity line. That is a stacked column chart with a `Target` row, both
  shipped; it needs a value axis, which a timeline does not have. Recipe in
  docs/MANUAL.md. (Lane grouping itself shipped as `gantt.lanes`.)
- **Dial / needle gauge** — bullet charts replace it deliberately (Few); low
  deck demand. Note the *semi-circle scorecard* gauge did ship (`doughnut` +
  `pie.semi`); what stays rejected is the dial-with-needle and its threshold
  bands.
- **Sankey / chord / arc** — need curved ribbons; infographic genre.
- **Ridgeline** — stacked density curves; academic register. (The single-column
  `violin` kind shipped at owner request; ridgeline stacking not pursued.)
- **Stream graph** — feasible but editorial aesthetic; no deck demand.
- **Pictogram with icon libraries** — needs the image node above, which is not
  reachable in the live add-in at all (see the first entry), plus an asset
  library. Waffle is the deliberate substitute: it covers the part-to-whole
  genre with square cells, it does not render icons.
- **Histogram as a kind** — the look is `clustered` + `gapWidth: 0`, both
  shipped. If this is ever revisited, auto-binning raw samples into categories
  is the only real gap; the bar geometry is not. (`histogramBins` in
  src/core/format.ts bins over a fixed domain, but nothing derives categories.)
- **Choropleth maps, 3D, drill-down interactivity** — out of scope by design
  (see CLAUDE.md). Tilemap proportional-area cartograms and tilemap drill-down
  fall here too: hard/infeasible.
- **Population pyramid, plain dot chart** — already covered by `butterfly`
  (+ `butterfly.split`) and `decorations.barStyle: "dot"`.
- **Radar vertex markers** — already there: radar emits `marker-*` ellipse
  nodes, which the Office.js renderer draws, so they appear in the live add-in
  too.
- ~~Candlestick / OHLC~~ — shipped as the `candlestick` kind at owner request,
  despite the thin consulting-demand signal.
