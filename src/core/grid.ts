/**
 * Uniform spatial hash for axis-aligned box overlap queries.
 *
 * Exact for ANY cell size: a box is registered in every grid cell it overlaps,
 * and a query returns every item sharing a cell with the query box — so any true
 * overlap (whose non-empty intersection lies in some shared cell) is guaranteed
 * to be returned. The grid only PRUNES the candidate set; the caller still runs
 * the precise overlap test, so the decision is identical to a full scan, just
 * without the O(n²) `taken.some(...)` sweep the label/collision passes used.
 *
 * Cell size affects only speed (bigger cells → more candidates per query, more
 * cells spanned per big box), never correctness. Pick roughly the typical box
 * size so a query touches O(1) cells.
 */
export interface GridBox {
  x: number;
  y: number;
  w: number;
  h: number;
}

/**
 * Cells one box may occupy before it is treated as degenerate.
 *
 * The index sizes its cell to the largest label, so a real label covers a
 * handful of cells and a chart's worth of them covers a few hundred. Four
 * thousand is far past anything a layout produces and far short of anything
 * that costs time.
 */
const MAX_CELLS_PER_BOX = 4096;

export class BoxHash<T> {
  private readonly inv: number;
  private readonly buckets = new Map<string, T[]>();

  constructor(cell: number) {
    this.inv = 1 / Math.max(1e-6, cell);
  }

  private forEachCell(b: GridBox, fn: (key: string) => void): void {
    const x0 = Math.floor(b.x * this.inv);
    const x1 = Math.floor((b.x + b.w) * this.inv);
    const y0 = Math.floor(b.y * this.inv);
    const y1 = Math.floor((b.y + b.h) * this.inv);
    // A box whose corners are not finite has no cells to visit, and a box
    // spanning more cells than any label could is not a label. Both used to be
    // enumerated one cell at a time: a label placed at 1e300 — which a value of
    // 1e308 in the datasheet produces — made this loop about 1e298 times,
    // allocating a Map key on each pass. Not slow: unbounded, in both time and
    // memory, from a number someone can type into a cell.
    if (!Number.isFinite(x0) || !Number.isFinite(x1) || !Number.isFinite(y0) || !Number.isFinite(y1)) return;
    if ((x1 - x0 + 1) * (y1 - y0 + 1) > MAX_CELLS_PER_BOX) {
      // Degenerate geometry still gets ONE cell, so it neither disappears from
      // the index nor drags the rest of the pass down with it. The cost is a
      // label that may overlap something — which is what it was already doing.
      fn(`${x0},${y0}`);
      return;
    }
    for (let gx = x0; gx <= x1; gx++) {
      for (let gy = y0; gy <= y1; gy++) fn(`${gx},${gy}`);
    }
  }

  insert(b: GridBox, item: T): void {
    this.forEachCell(b, (key) => {
      const arr = this.buckets.get(key);
      if (arr) arr.push(item);
      else this.buckets.set(key, [item]);
    });
  }

  /**
   * Run `test` against every item sharing a cell with `b`, short-circuiting on
   * the first truthy result — the pruned equivalent of `taken.some(test)`. An
   * item spanning several of the query's cells may be visited more than once;
   * the test is a pure predicate, so that only costs a repeat check, never a
   * wrong answer.
   */
  some(b: GridBox, test: (item: T) => boolean): boolean {
    let hit = false;
    this.forEachCell(b, (key) => {
      if (hit) return;
      const arr = this.buckets.get(key);
      if (arr) {
        for (const item of arr) {
          if (test(item)) {
            hit = true;
            return;
          }
        }
      }
    });
    return hit;
  }
}

/** A cell size ~ the largest box, so any overlapping pair lands within a step. */
export function gridCellFor(boxes: { w: number; h: number }[], min = 8): number {
  let m = min;
  for (const b of boxes) {
    if (b.w > m) m = b.w;
    if (b.h > m) m = b.h;
  }
  return m;
}
