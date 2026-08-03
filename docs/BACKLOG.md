# Backlog

Curated candidate work, from a research sweep (July 2026) comparing every
existing chart kind against think-cell, Excel, Highcharts, Datawrapper, and
Mekko Graphics, plus a chart-type survey across the Zelazny / FT Visual
Vocabulary taxonomies and competitor add-ins (Zebra BI, Vizzlo, UpSlide).

**This is the only backlog document.** Items graduate from here into PRs and are
deleted when they ship — what has shipped is recorded by the README feature
table and by git, not here. Rejected ideas stay in §2 so they aren't
re-proposed.

Feasibility is judged against the live-add-in constraint: rects, lines, text,
ellipses, and any of PowerPoint's 177 preset geometries (all of them
PowerPointApi 1.4), plus polygon *outlines* — no freeform curves, and no
images. The SVG and skill-pptx renderers additionally have filled polygons and
patterns.

## 1. Open

### Deck-level style, so a shared deck keeps its branding

**Researched:** 2026-08-01 (`docs/RESEARCH.md` §4b).

The imported style file and saved templates live in `localStorage`
(`powerchart-style`, `powerchart-templates`), so they follow the **browser**,
not the deck. Send a branded deck to a colleague and their charts do not match
yours — and every chart they add drifts further.

**Shape of the fix:** a presentation-scoped custom XML part
(`Presentation.customXmlParts`), which is the one thing it is genuinely the
right tool for. Chart config stays in shape tags — it has to travel with the
shape, through copy/paste into another deck and through PowerPoint's own
Duplicate Slide, which a presentation-scoped part cannot do.

**Known trap:** Office.js enumerates only parts related from
`ppt/presentation.xml`. A part written at the package root
(`/customXml/itemN.xml`) is invisible to it. If the generated deck writes one,
it must be related from the presentation, not dropped at the root.

**Priority:** low. Nothing is broken today; this is a sharing gap, and it needs
a decision about precedence (deck style vs the user's own imported style) before
any code.

### Drive PowerPoint on the web from Playwright, so the manual run is a command

**Researched:** 2026-08-02.

The standing test run is down to four items because the in-host battery
absorbed the rest. The remaining ones need a human because nothing can click
the pane — but something can. Playwright reaches an add-in task pane, and there
is a public working reference implementation to copy from
(`kzarzycki/powerpoint-mcp`, its `e2e/` directory).

**What makes it tractable here:** `manifest.xml` already points at
`https://localhost:3000`, which is the constraint Microsoft's own (undocumented)
sideload seam enforces — `office-addin-dev-settings` appends
`wdaddindevserverport` / `wdaddinmanifestfile` / `wdaddintest` to an
Office-on-the-web document URL, gated behind a `WEB_SIDELOAD_TEST` env var that
exists for Microsoft's internal CI. `playwright-core` is already a
devDependency.

**Blockers, all defeated in that reference repo:** headless user agents are
sniffed and sideloading silently skipped (spoof UA + `sec-ch-ua`); Chrome's
Private Network Access blocks the WAC origin from reaching loopback (intercept
with `context.route` and answer `Access-Control-Allow-Private-Network`); the
`wdaddintest` flag does *not* in fact suppress the developer-mode dialogs, so
they must be clicked. Reach the pane with `page.frames()` — it is a nested
cross-origin iframe, so `frameLocator()` chaining cannot get there, and Selenium
cannot reach it at all (office-js#5264).

**The real wall is login**, not the add-in machinery. That reference repo primes
a browser profile by hand in a headed window and runs headless afterwards; its
own CI does not run the suite. Headless would need Entra certificate-based auth
(Playwright `clientCertificates`, ≥1.46 — Microsoft's Power Platform samples do
exactly this), which needs a test user with CBA that the owner probably cannot
self-serve on a corporate tenant.

**Priority:** medium — but only the local half. Run locally it turns the
remaining manual items into `npm run test:e2e` and pays for itself without
solving login. CI is a separate, owner-gated question and may never be worth it.
Covers PowerPoint **on the web** only; desktop stays human either way.

**What is left for it to absorb, now that the battery has taken the rest.** The
selection round trip is a scenario again (*edit the chart YOU click*), but it
waits for a **human** click and records the verdict — so what still needs a
person is the click and the drag, not the checking.

Do not have such a driver call `setSelectedShapes` to supply that click. On the
web it wedges the host's selection subsystem (`docs/RESEARCH.md` §4b), and the
wedge is in PowerPoint rather than in how it is called. A Playwright click on
the **canvas** is a real click and might dodge it entirely — the one thing worth
trying first if this is ever built.

### A golden-image gate on the generated deck

**Researched:** 2026-08-02 (measured against `examples/showcase.pptx`).

The schema half of this shipped — `scripts/validate-ooxml.mjs` gates CI, and
`verify-deck.mjs` gained the duplicate-`cNvPr`-id check that neither tool had.
What is left is the visual half.

LibreOffice headless renders the deck to PDF in 7 s and to 122 PNGs in 17.6 s,
and — the part that makes a gate possible at all — the output is **byte-identical
across fresh profiles**, so it can be compared by hash rather than by a fuzzy
pixel diff. Its usual PPTX weak spots do not apply here: this generator emits no
gradients, no pattern fills, no `normAutofit` and no effects, which is every
category where LibreOffice is known to diverge.

**What it needs first:** a pinned container. The deck asks for Segoe UI and
Calibri, neither present in CI, so every render is a substituted render —
self-consistent, but different from a laptop that has the real fonts, and
different again after a LibreOffice minor bump. Without pinning, the baseline
churns and the gate gets switched off. Consider asserting structure (page count,
no all-white page) rather than hashes: version-tolerant, no committed baselines,
and it still catches "the chart rendered to nothing".

**Frame it as** "did our own output change", never as "does this match
PowerPoint" — no FOSS renderer is close enough to PowerPoint for the second
claim, and a gate that overclaims gets ignored.

**Do not** use a LibreOffice round-trip as a PowerPoint proxy: converting the
showcase deck back to `.pptx` silently deleted all 122 `ppt/tags/*.xml` parts,
leaving 121 charts non-re-editable. Anything learned that way is a fact about
LibreOffice.

**Priority:** low. The cheap, deterministic half is already in CI; this one buys
less and costs a container to maintain.

## 2. Rejected or already covered (do not re-propose)

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
- **Alt text in the headless pptx renderer** — the hunt flagged the missing
  `descr=` as a break of the SVG/Office.js/pptx parity contract, but pptxgenjs
  exposes alt text only on pictures and native charts, and every PowerChart
  shape is an `addShape`/`addText` autoshape. There is no seam to write it
  through short of hand-patching the generated OOXML. Recorded as a documented
  limit in `skill/reference.md` (#197) alongside the fixed deck font.
