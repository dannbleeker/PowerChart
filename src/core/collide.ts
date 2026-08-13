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
const MOVABLE = [
  /^total-/,
  /^value-line-label/,
  /^series-label-/,
  /^combo-series-label-/,
  /^diff-label$/,
  /^cagr-label$/,
  // The combo line's point labels, LAST so they are the ones that move: the
  // column's value label is anchored to the bar it sits on, where the line's
  // label is a floater above a point. They were absent, so on an ordinary
  // 480x300 combo at the default font a line label and a column label were drawn
  // through each other — this pass exists for exactly that and was simply never
  // told about them.
  /^combo-label-/,
];

/**
 * Labels that may be flipped BELOW their mark when there is no room above.
 *
 * The nudge only goes up, for a good reason — of two labels naming different
 * things, pushing one down crosses the one beneath it and they end up naming
 * each other's marks. That reason does not apply to a label anchored to a
 * single POINT: below its point is as true as above it, which is where
 * think-cell puts one when the room above is spent.
 *
 * Without the flip these labels had nowhere to go. On a 480x300 combo at a 26pt
 * font a line's point label sat on the column total beneath it, was nudged 143
 * points up over ten tries, and came to rest inside the TITLE — every nudge
 * legal, the budget exhausted, and the collision merely relocated onto a worse
 * partner. Restricted to this list rather than opened to every movable label,
 * because a series label really can be reordered by moving down.
 *
 * The CAGR caption is here for the same reason and is worth its own note,
 * because the OTHER way of getting it off the title was tried and reverted:
 * flooring it at the title's bottom turned five `title x cagr-label` overlaps
 * into eight against the column totals. That was an unconditional move, and
 * this is not — the flip only takes a position already clear of everything
 * settled, and the totals rank ahead of it and are settled first. So it cannot
 * buy its way off the title by landing on a total; it stays put instead. Worth
 * 13 of the sweep's overlapping pairs against the earlier attempt's −3.
 */
const FLIPPABLE = [/^combo-label-/, /^cagr-label$/];

const canFlip = (name: string | undefined) => !!name && FLIPPABLE.some((re) => re.test(name));

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

const overlaps = (a: Box, b: Box) => a.x < b.x + b.w && a.x + a.w > b.x && a.y < b.y + b.h && a.y + a.h > b.y;

export function resolveLabelCollisions(nodes: SceneNode[], canvasH?: number): void {
  const texts = nodes.filter((n): n is TextNode => n.kind === "text" && !!n.text);
  const fixed: Box[] = [];
  const movable: { node: TextNode; rank: number }[] = [];
  for (const t of texts) {
    const rank = movableRank(t.name);
    if (rank >= 0) movable.push({ node: t, rank });
    else fixed.push(tightBox(t));
  }
  // Lower rank settles first and becomes an obstacle for the rest; within a
  // rank, the LOWEST label settles first.
  //
  // The order within a rank used to be the order the layout emitted, and the
  // nudge only ever goes up — so of two overlapping labels the one that
  // happened to come first stayed put and the other was pushed past it. On a
  // chart dense enough that its series labels cannot all fit at their minimum
  // gap, that reordered them: every label still on the canvas, each naming the
  // wrong line. Settling bottom-up makes the nudge move a label AWAY from the
  // one below it, which can never cross.
  movable.sort((a, b) => a.rank - b.rank || b.node.y - a.node.y);

  // Spatial hash over the settled boxes so each movable label tests only its
  // neighbourhood, not the whole settled set on every one of its ≤10 nudges. The
  // exact `overlaps` test still decides, so the nudged positions are identical.
  const cell = gridCellFor([...fixed, ...movable.map((m) => tightBox(m.node))]);
  const settled = new BoxHash<Box>(cell);
  for (const b of fixed) settled.insert(b, b);
  for (const { node } of movable) {
    const startY = node.y;
    let box = tightBox(node);
    let tries = 0;
    let clear = !settled.some(box, (s) => overlaps(box, s));
    while (!clear && tries < 10) {
      node.y -= node.fontSize * 0.55; // nudge upward
      box = tightBox(node);
      tries++;
      if (box.y < 0) {
        // Nudging would push the label off the top of the canvas (y=0 for every
        // chart). An overlapping label still reads; an off-canvas one is lost — so
        // give up and restore. Without this, a column total sharing the totals row
        // with the fixed grand-total label was nudged clean off the top.
        node.y = startY;
        box = tightBox(node);
        break;
      }
      clear = !settled.some(box, (s) => overlaps(box, s));
    }
    // Up was blocked or spent, so try BELOW the mark — from the label's own
    // starting point, never from wherever the failed climb left it. Only for a
    // label anchored to a single point (see FLIPPABLE), and only while the
    // canvas height is known: without it there is no bottom edge to refuse at,
    // and this would push labels off the foot of the chart the way the guard
    // above stops it pushing them off the head.
    if (!clear && canvasH != null && canFlip(node.name)) {
      node.y = startY;
      for (let down = 0; down < 10; down++) {
        node.y += node.fontSize * 0.55;
        const cand = tightBox(node);
        if (cand.y + cand.h > canvasH) break;
        if (!settled.some(cand, (s) => overlaps(cand, s))) {
          box = cand;
          clear = true;
          break;
        }
      }
      // Neither direction cleared: leave it where the layout put it. A label
      // moved and still colliding is strictly worse than one that never moved —
      // it collides somewhere its author did not choose.
      if (!clear) node.y = startY;
      box = tightBox(node);
    }
    settled.insert(box, box);
  }
}

/**
 * A combo chart's two families of value labels — the column totals and the
 * line's point labels — can both want the same band on a short frame, and on a
 * short frame neither can move.
 *
 * `resolveLabelCollisions` above resolves every one of these if the canvas is
 * unbounded: the point label is FLIPPABLE, so it goes below its point, landing
 * at y 57-75 on a 60pt-tall chart. That destination is off the bottom, so the
 * flip is correctly refused and the label stays where the layout put it —
 * drawn through the total. Shrinking does not help either, because the two are
 * centred at the same y whenever the column total and the line value coincide,
 * which is a property of the DATA.
 *
 * So one of them has to go, and `decor.tightLabelPriority` says which. This runs
 * AFTER de-collision on purpose: only labels that are still colliding once every
 * legal move has been tried are dropped, so a roomy chart keeps both and so does
 * a short chart whose data happens to separate them.
 *
 * Returns the nodes to drop rather than mutating, so the caller owns the scene
 * array and this stays a pure decision — the same split as `positionalSweepPlan`
 * and `chooseGroupMembers`.
 */
export function unplaceableComboLabels(nodes: SceneNode[], priority: "columns" | "line" = "columns"): Set<SceneNode> {
  // `!!n.text`, never `n.text.trim()`: `text` is `string` in the types and a
  // number in a config someone pasted, so calling a string method on it throws.
  // `resolveLabelCollisions` above tests truthiness for the same reason and
  // `textWidth` coerces with `String(text ?? "")`. The hostile-input sweep
  // caught this within one run of adding it.
  const texts = nodes.filter((n): n is TextNode => n.kind === "text" && !!n.text);
  const totals = texts.filter((t) => /^total-/.test(t.name ?? ""));
  const points = texts.filter((t) => /^combo-label-/.test(t.name ?? ""));
  const drop = new Set<SceneNode>();
  if (!totals.length || !points.length) return drop;
  // The loser is dropped, the winner is left exactly as de-collision settled it.
  const [losers, keepers] = priority === "line" ? [totals, points] : [points, totals];
  for (const l of losers) {
    const lb = tightBox(l);
    if (keepers.some((k) => overlaps(lb, tightBox(k)))) drop.add(l);
  }
  return drop;
}
