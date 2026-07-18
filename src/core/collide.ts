import type { SceneNode, TextNode } from "./scene";
import { textWidth } from "./scene";
import { BoxHash, gridCellFor } from "./grid";

/**
 * Global label de-collision pass (a lightweight take on think-cell's
 * guaranteed non-overlapping labels): after layout + decorations, outside
 * labels (totals, CAGR/difference/value-line labels, series labels) are
 * nudged upward until they clear every other label's tight bounding box.
 * Inside-segment and axis labels stay fixed — they own their space.
 */

interface Box {
  x: number;
  y: number;
  w: number;
  h: number;
}

/** Movable label name prefixes, in nudge-priority order (later = moves). */
const MOVABLE = [/^total-/, /^value-line-label/, /^series-label-/, /^combo-series-label-/, /^diff-label$/, /^cagr-label$/];

const movableRank = (name: string | undefined): number => {
  if (!name) return -1;
  return MOVABLE.findIndex((re) => re.test(name));
};

/** Actual painted extent of a text node, given its alignment. */
function tightBox(n: TextNode): Box {
  const w = Math.min(n.w, textWidth(n.text, n.fontSize, n.bold));
  const h = Math.min(n.h, n.fontSize * 1.25);
  const x = n.align === "left" ? n.x : n.align === "right" ? n.x + n.w - w : n.x + (n.w - w) / 2;
  const y = n.valign === "top" ? n.y : n.valign === "bottom" ? n.y + n.h - h : n.y + (n.h - h) / 2;
  return { x, y, w, h };
}

const overlaps = (a: Box, b: Box) =>
  a.x < b.x + b.w && a.x + a.w > b.x && a.y < b.y + b.h && a.y + a.h > b.y;

export function resolveLabelCollisions(nodes: SceneNode[]): void {
  const texts = nodes.filter((n): n is TextNode => n.kind === "text" && !!n.text);
  const fixed: Box[] = [];
  const movable: { node: TextNode; rank: number }[] = [];
  for (const t of texts) {
    const rank = movableRank(t.name);
    if (rank >= 0) movable.push({ node: t, rank });
    else fixed.push(tightBox(t));
  }
  // Lower rank settles first and becomes an obstacle for the rest.
  movable.sort((a, b) => a.rank - b.rank);

  // Spatial hash over the settled boxes so each movable label tests only its
  // neighbourhood, not the whole settled set on every one of its ≤10 nudges. The
  // exact `overlaps` test still decides, so the nudged positions are identical.
  const cell = gridCellFor([...fixed, ...movable.map((m) => tightBox(m.node))]);
  const settled = new BoxHash<Box>(cell);
  for (const b of fixed) settled.insert(b, b);
  for (const { node } of movable) {
    let box = tightBox(node);
    let tries = 0;
    while (settled.some(box, (s) => overlaps(box, s)) && tries < 10) {
      node.y -= node.fontSize * 0.55; // nudge upward
      box = tightBox(node);
      tries++;
    }
    settled.insert(box, box);
  }
}
