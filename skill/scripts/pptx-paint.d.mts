// Types for the pure paint/node helpers in pptx-paint.mjs. Hand-written (the
// module is plain JS shipped in the skill zip) so the in-process test and any TS
// caller get real signatures instead of an implicit `any`. Not copied into the
// skill package — build-skill.mjs ships only the .mjs.

/** Points → inches. */
export const IN: number;

/** Normalise any allow-listed paint to exactly six hex digits (no leading #). */
export function hex(color: string | null | undefined): string;

/** Opacity 0..1 carried by a paint; 1 when opaque, 0 for `transparent`. */
export function alphaOf(color: string | null | undefined): number;

/** A pptxgenjs solid fill: `{color}`, `{color, transparency}`, or `{type:"none"}`. */
export function fillOf(
  color: string | null | undefined,
  fillOpacity?: number,
): { color: string; transparency?: number } | { type: "none" };

/** True only for a present, non-transparent paint. */
export function visible(paint: string | null | undefined): boolean;

/** HSL (h in degrees, s/l in 0–100) → six hex digits. */
export function hslToHex(h: number, s: number, l: number): string;

/** CSS Color 4 named colours as name → rrggbb. */
export const CSS_NAMES: Record<string, string>;

/** The four engine helpers a node mapping binds. */
export interface PptxEngine {
  dashKind(dash: number[]): "dot" | "dash";
  annularSectorPoints(
    cx: number,
    cy: number,
    innerR: number,
    r: number,
    startAngle: number,
    endAngle: number,
  ): { x: number; y: number }[];
  SYMBOL_PRESET: Record<string, string>;
  arrowheadBox(
    x: number,
    y: number,
    size: number,
    angle: number,
  ): { left: number; top: number; size: number; rotation: number };
}

/** A slide sink capturing the PptxgenJS calls a node mapping makes. */
export interface PptxSlide {
  addShape(type: string, opts: Record<string, unknown>): void;
  addText(text: string, opts: Record<string, unknown>): void;
}

/**
 * Bind the engine helpers and return `addNode(slide, n, dx, dy)` — a pure mapper
 * from one scene node to PptxgenJS calls at a slide offset (inches). The node is
 * intentionally loose (`any`): it is a discriminated scene node whose full union
 * lives in the TS core, not this JS module.
 */
export function makeAddNode(
  engine: PptxEngine,
): (slide: PptxSlide, n: unknown, dx: number, dy: number) => void;
