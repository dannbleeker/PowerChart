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
  | RectNode
  | LineNode
  | TextNode
  | EllipseNode
  | WedgeNode
  | ChevronNode
  | ArrowheadNode;

/** Point on a circle for wedge geometry (0° = 12 o'clock, clockwise). */
export function polar(cx: number, cy: number, r: number, angleDeg: number): { x: number; y: number } {
  const a = ((angleDeg - 90) * Math.PI) / 180;
  return { x: cx + r * Math.cos(a), y: cy + r * Math.sin(a) };
}

export interface Scene {
  width: number;
  height: number;
  nodes: SceneNode[];
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
