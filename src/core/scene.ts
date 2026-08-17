/**
 * Renderer-agnostic scene graph. Layouts emit these nodes; the SVG renderer
 * (preview/tests) and the Office.js renderer (native PowerPoint shapes)
 * consume them. Coordinates are in points, origin top-left.
 *
 * ── Renderer parity contract ────────────────────────────────────────────────
 * Three renderers consume this graph: SVG (src/render/svg.ts), Office.js
 * (src/render/powerpoint.ts, the live add-in), and PptxgenJS
 * (skill/scripts/render-pptx.mjs, the headless skill). SVG is the reference —
 * it can draw anything — so where the two PowerPoint renderers differ it is
 * because Office.js and OOXML presets cannot express what SVG can. The
 * divergences are intentional; each is noted on the field or kind it affects so
 * a later change does not "fix" an approximation into a regression:
 *
 *  - Pattern fills (rect.pattern): SVG only; solid elsewhere.
 *  - Polygon fills (polygon.fill/fillOpacity): SVG + pptx custGeom; Office.js
 *    has no freeform fill and degrades to the stroked outline.
 *  - Wedge geometry: SVG + pptx draw the exact arc; Office.js approximates with
 *    a triangle/rectangle fan (no adjustable pie geometry).
 *  - Dash arrays (line.dash): SVG honours the exact array; the PowerPoint
 *    renderers expose enums, so they map to the nearest native style via
 *    `dashKind` (dotted → roundDot/sysDot, else dash) rather than the exact rhythm.
 *    WHETHER a line is dashed at all is not approximate, and all three sinks must
 *    answer it the same way — they ask `dashKind`, which returns `none` for an
 *    array carrying no positive finite length. Do not re-guard on `dash` being
 *    truthy at a call site: `[]` is truthy, and that divergence drew a solid line
 *    in the preview and a dotted one in both decks.
 *  - TEXT alpha (text.color carrying one, e.g. `#0b0b0b80` or an `rgba()`):
 *    SVG and pptx honour it — pptxgenjs takes the same 0-100 transparency on a
 *    text run as on a shape — and Office.js cannot: `font.color` is a hex
 *    string with nowhere to put an alpha, so the live add-in draws such a label
 *    OPAQUE. Fills and strokes honour their alpha in all three, so this is the
 *    one paint channel that does not, and it is a host limit rather than a
 *    choice. Muted ink is how a chart de-emphasises a label, so a chart using
 *    it reads as flatter in the add-in than in the preview or the skill's deck.
 *  - Chevron point depth and arrowhead proportions: SVG draws its own geometry;
 *    the PowerPoint renderers name a native preset whose default proportions
 *    differ slightly (see the notes on those kinds). Reproducing the preset
 *    geometry exactly is not verifiable without a PowerPoint rasteriser, so the
 *    preview approximates a shape the deck draws natively — deliberately, the
 *    same call made for the rejected star5 marker.
 * ────────────────────────────────────────────────────────────────────────────
 */

export interface RectNode {
  kind: "rect";
  x: number;
  y: number;
  w: number;
  h: number;
  fill: string;
  stroke?: string;
  strokeWidth?: number;
  /** Hatch/dot pattern overlaid on the fill (SVG renderer; solid elsewhere). */
  pattern?: "diagonal" | "crosshatch" | "dots" | "horizontal";
  name?: string;
}

export interface PolygonNode {
  kind: "polygon";
  points: { x: number; y: number }[];
  /** Fill color; rendered translucent via fillOpacity in SVG. PowerPoint
   * renderers degrade to the stroked outline only (no freeform fills). */
  fill?: string;
  fillOpacity?: number;
  stroke?: string;
  strokeWidth?: number;
  name?: string;
}

export interface LineNode {
  kind: "line";
  x1: number;
  y1: number;
  x2: number;
  y2: number;
  stroke: string;
  strokeWidth?: number;
  /**
   * Dash pattern in points, e.g. [2, 2]. SVG renders the exact array; the
   * PowerPoint renderers have only an enum of named styles, so they collapse it
   * to the nearest one via `dashKind` (a dotted [1.5,1.5] stays dotted;
   * everything else is a dash). The rhythm is approximate in the deck by design.
   *
   * An array with no positive finite length (`[]`, `[0, 0]`, `[-5, -5]`, `[NaN]`)
   * means NOT DASHED in every sink — ask `dashKind`, never `if (n.dash)`.
   */
  dash?: number[];
  name?: string;
}

export interface TextNode {
  kind: "text";
  /** Bounding box; alignment applies within it. */
  x: number;
  y: number;
  w: number;
  h: number;
  text: string;
  fontSize: number;
  color: string;
  bold?: boolean;
  align: "left" | "center" | "right";
  valign: "top" | "middle" | "bottom";
  fontFamily?: string;
  name?: string;
}

export interface EllipseNode {
  kind: "ellipse";
  cx: number;
  cy: number;
  rx: number;
  ry: number;
  fill: string;
  stroke?: string;
  strokeWidth?: number;
  name?: string;
}

/**
 * Pie/doughnut wedge. Angles in degrees, 0 = 12 o'clock, clockwise.
 * SVG renders an exact path; PowerPoint approximates with a triangle fan
 * (Office.js exposes no adjustable pie geometry).
 */
export interface WedgeNode {
  kind: "wedge";
  cx: number;
  cy: number;
  r: number;
  /** Inner radius for doughnuts; 0 for pies. */
  innerR: number;
  startAngle: number;
  endAngle: number;
  fill: string;
  stroke?: string;
  strokeWidth?: number;
  name?: string;
}

/**
 * Process-flow chevron / pentagon-arrow (PowerPoint's chevron & homePlate).
 * SVG draws the arrow with its notch at a fixed fraction of the height; the
 * PowerPoint renderers name the native chevron/homePlate preset, whose own
 * default point depth differs slightly — so the arrow's point is a touch
 * deeper/shallower in the deck than the preview. Intentional (see the parity
 * contract at the top): the preset can't be matched pixel-for-pixel here.
 */
export interface ChevronNode {
  kind: "chevron";
  x: number;
  y: number;
  w: number;
  h: number;
  fill: string;
  /** First step in a flow has a flat left edge (homePlate), the rest are chevrons. */
  flatLeft?: boolean;
  name?: string;
}

/**
 * Filled marker symbol centred on (cx, cy), inscribed in a `2*size` square.
 *
 * Shape is an encoding channel that survives what color does not: greyscale
 * printing and red-green color blindness both flatten a palette, and a deck
 * gets printed. A PolygonNode would render the same outline in SVG but
 * degrades to an unfilled outline in PowerPoint (no freeform fills there),
 * so a symbol is its own kind: each shape maps to a native preset geometry
 * and stays filled in all three renderers. See `symbolPoints` / `SYMBOL_PRESET`.
 */
export interface SymbolNode {
  kind: "symbol";
  shape: SymbolShape;
  cx: number;
  cy: number;
  /** Half the box side, so it reads like an ellipse's radius. */
  size: number;
  fill: string;
  stroke?: string;
  strokeWidth?: number;
  name?: string;
}

/**
 * Filled triangle with tip at (x, y), pointing along `angle` (degrees, 0 = east,
 * clockwise). SVG draws a narrow isosceles triangle; the PowerPoint renderers
 * name the native `triangle` preset in a `2*size` square (see `arrowheadBox`),
 * which is a touch broader. The tip anchor and angle match across all three;
 * only the triangle's proportions differ, intentionally (see the parity contract).
 */
export interface ArrowheadNode {
  kind: "arrowhead";
  x: number;
  y: number;
  angle: number;
  size: number;
  fill: string;
  name?: string;
}

export type SceneNode =
  RectNode | LineNode | TextNode | EllipseNode | WedgeNode | ChevronNode | ArrowheadNode | PolygonNode | SymbolNode;

// Circle/wedge math lives in ./geometry (shared with the renderers); re-exported
// here so scene consumers (layouts) keep importing `polar` from the scene module.
export { polar } from "./geometry";
import { wedgeFanSteps, type SymbolShape } from "./geometry";
import { toRgb } from "./color";
export type { SymbolShape };

/**
 * Drop any node whose geometry is not a finite number.
 *
 * The scene is the contract between the chart engine and three renderers, and
 * "every coordinate is a real number" was a property it happened to have
 * rather than one it promised. The SVG renderer defends itself — every numeric
 * goes through `num()` — but the two PowerPoint renderers do not, and they are
 * the ones that write a file. A `NaN` there lands in `addGeometricShape({left:
 * NaN, …})` and in OOXML as an EMU value, so the produced .pptx is one
 * PowerPoint may simply refuse to open. Nine chart kinds could do it, from
 * values a datasheet cell can hold (`1e308`, `5e-324`).
 *
 * DROPPED rather than zeroed. A node whose position could not be computed has
 * no right position to fall back to, and zeroing puts a stray bar in the
 * corner of the chart — wrong in a way that looks deliberate. Leaving it out
 * loses that one node and keeps the rest of the chart, which is what a reader
 * can actually interpret.
 *
 * For a valid config this drops nothing, so it is a floor and not a filter.
 */
export function finiteNodes(nodes: SceneNode[]): SceneNode[] {
  return nodes.filter((n) => allNumbersFinite(n) && !degeneratePolygon(n));
}

/**
 * Text clipped to what the chart can actually hold — the backstop under every
 * layout's own fitting.
 *
 * A label is drawn at whatever size its layout chose, in a box that layout sized
 * from the frame, and neither PowerPoint renderer wraps or clips a text box. So
 * any label wider than the room in front of it draws straight off the chart:
 * invisible in a picture-mode render, and lying across whatever sits beside the
 * chart on a slide.
 *
 * At a thumbnail frame that was not a corner case but the normal outcome — 18 of
 * the 25 kinds put ink outside their own frame at 120x90, by as much as 124pt on
 * a 120pt-wide chart. Most of it came from a handful of SHARED nodes (the title,
 * the footnote, the series labels) and each is now fitted where it is built,
 * which is better than clipping because shrinking keeps the whole word. This
 * catches what those did not, once, instead of in twenty-five layouts — the
 * per-site fixes stop being a list somebody has to finish.
 *
 * Only the horizontal axis: a label too TALL for its frame cannot be rescued by
 * shortening it, and the layouts that had that problem now reserve for it.
 *
 * Clipped from the anchor the node was placed by, so alignment is preserved: a
 * left-aligned label keeps its left edge, a right-aligned one its right, a
 * centred one its centre. A label already inside the frame is returned
 * untouched and byte-identical, which is why no snapshot moves.
 */
export function clipTextToFrame<T extends SceneNode>(nodes: T[], width: number): T[] {
  for (const n of nodes) {
    if (n.kind !== "text") continue;
    const t = n as unknown as TextNode;
    const ink = textWidth(t.text, t.fontSize, t.bold);
    const x = t.align === "right" ? t.x + t.w - ink : t.align === "center" ? t.x + (t.w - ink) / 2 : t.x;
    if (x >= -0.5 && x + ink <= width + 0.5) continue;
    // The room in front of the anchor this node was positioned by. A centred
    // label may only grow to twice its distance from the nearer edge before one
    // side leaves the frame.
    const centre = t.x + t.w / 2;
    const room =
      t.align === "right" ? t.x + t.w : t.align === "center" ? 2 * Math.min(centre, width - centre) : width - t.x;
    // The same ellipsis walk as `clipToWidth` in `elements.ts`, inlined rather
    // than imported: that module imports `textWidth` and `finiteNodes` from
    // here, so reaching back for it would put this file in an import cycle with
    // it for four lines of loop.
    if (room <= 0) {
      t.text = "";
      continue;
    }
    let cut = t.text;
    while (cut.length > 0 && textWidth(`${cut}…`, t.fontSize, t.bold) > room) cut = cut.slice(0, -1);
    t.text = cut ? `${cut}…` : "";
  }
  return nodes;
}

/**
 * A polygon with nothing to draw — and the hole this gate had.
 *
 * `allNumbersFinite` asks whether every number in a node is finite, and an
 * EMPTY point list satisfies that trivially: there are no numbers to fail. So a
 * polygon carrying `points: []` sailed through a filter whose whole job is to
 * keep un-openable files from being written, and then broke exactly the
 * renderer the filter exists to protect. `pptx-paint.mjs` takes the polygon's
 * bounding box with `Math.min(...xs)`, which is `Infinity` for no points, and
 * writes `x="Infinity"` into the OOXML — not a number, not an Int64, and
 * Microsoft's own validator rejects the deck.
 *
 * Found by rendering 3033 hostile configs through the skill's headless
 * renderer: `{kind: "radar", data: {}}` lays out its grid rings before it knows
 * it has no axes, and emits two of them empty.
 *
 * Two points, not one, because a polygon is a closed path: one point has no
 * edges and cannot be a shape either. The rest of the chart is kept, which is
 * the same trade the filter above already makes.
 */
function degeneratePolygon(n: SceneNode): boolean {
  return n.kind === "polygon" && (n.points?.length ?? 0) < 2;
}

/** Every number anywhere in the node — including a polygon's point list. */
function allNumbersFinite(value: unknown, depth = 0): boolean {
  // Scene nodes are shallow; the bound is a cycle guard, not a shape claim.
  if (depth > 6) return true;
  if (typeof value === "number") return Number.isFinite(value);
  if (Array.isArray(value)) return value.every((v) => allNumbersFinite(v, depth + 1));
  if (value && typeof value === "object") {
    for (const v of Object.values(value)) if (!allNumbersFinite(v, depth + 1)) return false;
  }
  return true;
}

export interface Scene {
  width: number;
  height: number;
  nodes: SceneNode[];
  /**
   * Accessible name (the chart title) and a one-line text alternative
   * summarising the data. Emitted by the SVG renderer as `<title>`/`<desc>`
   * under `role="img"`, so a screen reader announces the chart instead of
   * silence. Set by buildChart; optional so hand-built scenes stay valid.
   */
  title?: string;
  desc?: string;
}

/**
 * How many NATIVE shapes a scene becomes on the Office.js host — NOT the node
 * count. A wedge fans out into `wedgeFanSteps` shapes (+2 stroke edges); a polygon
 * draws one line per edge. So a 10-node pie is ~50 shapes and a 10-node violin
 * ~250. This is the number the web shape budget cares about (counting nodes waved
 * both past it and the host choked) and the number the demo's contents table shows.
 */
export function estimateOfficeShapes(scene: Scene): number {
  let total = 0;
  for (const n of scene.nodes) {
    if (n.kind === "wedge") {
      const span = n.endAngle - n.startAngle;
      total += wedgeFanSteps(n.r, span).steps + (n.stroke && span < 359.9 ? 2 : 0);
    } else if (n.kind === "polygon") {
      total += n.points.length; // one line per edge, closed
    } else {
      total += 1;
    }
  }
  return total;
}

/**
 * Approximate rendered text width in points (average glyph ≈ 0.54 em for UI sans).
 *
 * Coerces, for the same reason `xmlText` and `paintText` do: the type says
 * `string` and the value came out of a file someone pasted. Sixty-odd call
 * sites across every layout ask this question about a title, a category, a
 * series name or a table cell, and a non-string used to answer two different
 * ways, both silent:
 *
 * - `null`/`undefined` THREW `Cannot read properties of null (reading
 *   'length')`, taking the whole chart down. `buildKpiTile({})` — a tile with
 *   no value yet — and `buildProcessFlow([null])` did exactly that, and both
 *   are exported from `src/index.ts` as the skill's public API.
 * - a NUMBER returned `NaN`, because `(2024).length` is `undefined`. That is
 *   the worse one. Every fit-to-width test here is a comparison, and each of
 *   them is FALSE against NaN — so shrink-to-fit silently stopped shrinking,
 *   and a width built as `Math.max(w, textWidth(...))` became NaN and had its
 *   whole node dropped by `finiteNodes`. `valueAxisTitle: 99` — a units label
 *   of `99`, or any year — vanished from the chart in all 25 kinds, with no
 *   error anywhere: the safety net turned a crash into a disappearance.
 *
 * `String()` rather than a rejection, matching `xmlText`: the renderers will
 * draw `2024`, so measuring it as four characters is the honest answer, and
 * measuring a missing string as zero is what every caller already means by it.
 */
export function textWidth(text: string, fontSize: number, bold = false): number {
  return String(text ?? "").length * fontSize * (bold ? 0.58 : 0.54);
}

/**
 * Pick black or white ink for a given fill so segment labels stay readable.
 * Parses via the shared `toRgb` — this used to carry its own hex-only copy, which
 * read every rgb()/hsl() fill as pure black and so chose WHITE ink for a
 * near-white segment.
 */
export function contrastInk(fill: string): string {
  const [r, g, b] = toRgb(fill).map((c) => {
    const s = c / 255;
    return s <= 0.03928 ? s / 12.92 : Math.pow((s + 0.055) / 1.055, 2.4);
  });
  const lum = 0.2126 * r + 0.7152 * g + 0.0722 * b;
  return lum > 0.35 ? "#0b0b0b" : "#ffffff";
}

/**
 * Move a node horizontally by `dx`.
 *
 * Every horizontal coordinate a `SceneNode` can carry, which is four scalars
 * (`x`, `x1`, `x2`, `cx`) and one array — `points`, the polygon's own geometry.
 * The array is the whole reason this lives here rather than beside its caller.
 *
 * It began as a duck-typed loop over the four scalar names in the demo
 * gallery, a shape that can never fail to COMPILE when a node kind gains a
 * coordinate, and `points` was already missing from it: a polygon stayed where
 * it was while everything around it moved. That was latent — the gallery only
 * composes Harvey balls and checkboxes, and `src/core/elements.ts` emits no
 * polygon — but the helper's contract is "shift this node" and it quietly did
 * not, for one kind, with nothing to say so.
 *
 * Here it sits next to the node contract it has to keep up with, and it is
 * reachable by a test, which the gallery module is not: that file touches the
 * DOM at import time.
 */
export function shiftNodeX<T extends SceneNode>(n: T, dx: number): T {
  const node = n as unknown as Record<string, number>;
  for (const k of ["x", "x1", "x2", "cx"]) if (typeof node[k] === "number") node[k] += dx;
  const pts = (n as unknown as { points?: { x: number; y: number }[] }).points;
  if (Array.isArray(pts)) for (const p of pts) p.x += dx;
  return n;
}
