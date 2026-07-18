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
