/**
 * Renderer-agnostic scene graph. Layouts emit these nodes; the SVG renderer
 * (preview/tests) and the Office.js renderer (native PowerPoint shapes)
 * consume them. Coordinates are in points, origin top-left.
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
  /** Dash pattern in points, e.g. [2, 2]. */
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

/** Process-flow chevron / pentagon-arrow (PowerPoint's chevron & homePlate). */
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

/** Filled triangle with tip at (x, y), pointing along `angle` (degrees, 0 = east, clockwise). */
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
