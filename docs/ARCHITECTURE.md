# Architecture

## Design goals

1. **Native shapes, not pictures.** think-cell's key trick is that its charts are
   real PowerPoint objects users can tweak. Office.js can't build think-cell's
   deep COM integration, but it *can* create geometric shapes, lines and text
   boxes with exact positions — so we lay charts out ourselves and emit shapes.
2. **A pure layout core.** Everything hard (stacking, scales, label fitting,
   decorations) lives in `src/core` with zero Office.js imports. It runs in
   Node for unit tests, in the browser for the preview, and inside PowerPoint
   for insertion — one code path for all three.
3. **Renderer-agnostic scene graph.** Layouts emit `rect | line | text |
   arrowhead` nodes in points (PowerPoint's native unit, 1pt = 1/72"). Renderers
   are dumb: SVG maps 1pt→1px; Office.js maps nodes to shapes 1:1.

## Data flow

```
ChartConfig (kind, data, decorations, style)
        │  buildChart()                        src/core/chart.ts
        ▼
layout (per kind)                              src/core/layout/*.ts
  · computeFrame()  — reserve margins for title/totals/axis/series labels
  · categorySlots() — think-cell's 2:1 column:gap rhythm
  · valueScale()    — nice-tick linear scale spanning zero
  · emit segment rects + labels + chrome
  · return LayoutAnchors (column tops, centers, totals, baseline)
        │
        ▼
decorationNodes()                              src/core/decor.ts
  · CAGR arrow, difference arrow, value line — computed from anchors,
    identical across chart kinds
        │
        ▼
Scene { nodes: SceneNode[] }                   src/core/scene.ts
        │                          │
        ▼                          ▼
sceneToSvg()               insertSceneIntoSlide()
src/render/svg.ts          src/render/powerpoint.ts
```

## think-cell behaviors replicated

- **Segment labels** are centered in their segment, hidden when the segment is
  shorter than ~1.25× the font size or narrower than the label, and switch
  between black/white ink by background luminance (`contrastInk`).
- **Chart-wide label precision**: decimals are resolved once per chart from the
  data's magnitude (`resolveFormat`), so "+14" never sits next to "+9.0".
- **Waterfall semantics**: a single series of deltas builds a running total;
  `e` cells draw a bar from the baseline to the running total; consecutive bar
  levels are joined with dashed connectors; floating deltas get signed labels,
  base/total bars don't.
- **Series labels** sit to the right of the last column at segment midpoints and
  are greedily pushed apart (then clamped) so they never overlap — think-cell's
  placement, simplified.
- **Decorations from anchors**: CAGR/difference arrows and value lines only need
  `columnTop/columnValue/categoryX`, so every current and future chart type gets
  them for free.

## Office.js constraints that shaped the design

| Constraint | Consequence |
|---|---|
| No freeform/path shapes | Arrowheads are triangles; rotation is applied when the host supports `Shape.rotation`, otherwise they stay axis-aligned |
| No line arrowhead properties in `ShapeLineFormat` | Same — arrowheads are separate shapes |
| `addGroup` only on newer hosts | Grouping is best-effort in a try/catch |
| No native chart-object API in PowerPoint Office.js | Exactly why we draw shapes — also what makes output editable |
| Shape units are points | The whole engine works in points end-to-end |
| Area fills can't be arbitrary polygons | Stacked areas are approximated with thin vertical slabs |

## Testing strategy

- `npm test` — vitest over the pure core: scale math, stacking invariants,
  waterfall running totals, Mekko width proportionality, CAGR/difference labels,
  and a smoke test that every sample builds inside its frame.
- Demo gallery (`index.html`) — renders every chart kind through the same engine
  for visual review; screenshot it in CI with Playwright if desired.
- The Office.js renderer is deliberately thin (a switch statement) so nearly all
  logic is covered by the Node tests.

## Roadmap

- Scatter/bubble with smart label placement
- Gantt/timeline (think-cell's second pillar)
- Agenda (chapter slides)
- Excel data links (Office.js Excel add-in companion + shared workbook binding)
- Theme-color palettes read from the presentation
- Butterfly/tornado, combo charts, error bars
