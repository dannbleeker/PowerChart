/**
 * Greedy point-label placement (the flavor of problem think-cell's scatter
 * labeling algorithm solves): try candidate positions around each anchor,
 * keep the first that collides with nothing already placed; hide otherwise.
 * Points are processed as given (callers can pre-sort by importance).
 */

import { BoxHash, gridCellFor } from "./grid";

export interface Box {
  x: number;
  y: number;
  w: number;
  h: number;
}

export interface LabelRequest {
  /** Anchor point the label belongs to. */
  cx: number;
  cy: number;
  /** Keep-out radius around the anchor (marker size). */
  r: number;
  w: number;
  h: number;
}

export interface PlacedLabel {
  index: number;
  box: Box;
  /** Which candidate slot won (for debugging/tests). */
  slot: number;
}

const overlaps = (a: Box, b: Box) =>
  a.x < b.x + b.w && a.x + a.w > b.x && a.y < b.y + b.h && a.y + a.h > b.y;

/** Candidate offsets: E, W, N, S, NE, NW, SE, SW of the anchor. */
function candidates(req: LabelRequest, pad: number): Box[] {
  const { cx, cy, r, w, h } = req;
  const d = r + pad;
  return [
    { x: cx + d, y: cy - h / 2, w, h },
    { x: cx - d - w, y: cy - h / 2, w, h },
    { x: cx - w / 2, y: cy - d - h, w, h },
    { x: cx - w / 2, y: cy + d, w, h },
    { x: cx + d * 0.8, y: cy - d * 0.8 - h, w, h },
    { x: cx - d * 0.8 - w, y: cy - d * 0.8 - h, w, h },
    { x: cx + d * 0.8, y: cy + d * 0.8, w, h },
    { x: cx - d * 0.8 - w, y: cy + d * 0.8, w, h },
  ];
}

/**
 * Place labels greedily. `bounds` clips candidates to the plot; `obstacles`
 * are boxes that must stay clear (e.g. all markers). Returns only the labels
 * that found a spot — the rest are hidden, as think-cell does when a chart
 * gets too dense.
 */
export function placeLabels(
  requests: LabelRequest[],
  bounds: Box,
  obstacles: Box[] = [],
  pad = 2,
): PlacedLabel[] {
  const placed: PlacedLabel[] = [];
  // Spatial hash over the taken boxes so a dense scatter's candidate test stays
  // near-linear instead of scanning every placed label/marker (the old
  // `taken.some(...)` was 8·n² across all points). The exact `overlaps` test
  // still decides, so placement is byte-identical to the full scan.
  const cell = gridCellFor([...obstacles, ...requests.map((r) => ({ w: r.w, h: r.h }))]);
  const taken = new BoxHash<Box>(cell);
  for (const o of obstacles) taken.insert(o, o);
  requests.forEach((req, index) => {
    const options = candidates(req, pad);
    for (let slot = 0; slot < options.length; slot++) {
      const box = options[slot];
      if (
        box.x < bounds.x ||
        box.y < bounds.y ||
        box.x + box.w > bounds.x + bounds.w ||
        box.y + box.h > bounds.y + bounds.h
      ) {
        continue;
      }
      if (taken.some(box, (t) => overlaps(box, t))) continue;
      placed.push({ index, box, slot });
      taken.insert(box, box);
      return;
    }
  });
  return placed;
}
