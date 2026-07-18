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
export type { SymbolShape };

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

/** Approximate rendered text width in points (average glyph ≈ 0.54 em for UI sans). */
export function textWidth(text: string, fontSize: number, bold = false): number {
  return text.length * fontSize * (bold ? 0.58 : 0.54);
}

/** Pick black or white ink for a given fill so segment labels stay readable. */
export function contrastInk(fill: string): string {
  const hex = fill.replace("#", "");
  const n = parseInt(hex.length === 3 ? hex.replace(/./g, "$&$&") : hex, 16);
  const [r, g, b] = [(n >> 16) & 255, (n >> 8) & 255, n & 255].map((c) => {
    const s = c / 255;
    return s <= 0.03928 ? s / 12.92 : Math.pow((s + 0.055) / 1.055, 2.4);
  });
  const lum = 0.2126 * r + 0.7152 * g + 0.0722 * b;
  return lum > 0.35 ? "#0b0b0b" : "#ffffff";
}
