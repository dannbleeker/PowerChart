import { describe, expect, it } from "vitest";
import { BoxHash, gridCellFor, type GridBox } from "../src/core/grid";

const overlaps = (a: GridBox, b: GridBox) => a.x < b.x + b.w && a.x + a.w > b.x && a.y < b.y + b.h && a.y + a.h > b.y;

function rng(seed: number) {
  return () => {
    seed |= 0;
    seed = (seed + 0x6d2b79f5) | 0;
    let t = Math.imul(seed ^ (seed >>> 15), 1 | seed);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

describe("BoxHash", () => {
  it("some() matches a brute-force scan for every query (exactness invariant)", () => {
    const rand = rng(20260718);
    for (let trial = 0; trial < 40; trial++) {
      const boxes: GridBox[] = Array.from({ length: 60 }, () => ({
        x: rand() * 400,
        y: rand() * 300,
        w: 2 + rand() * 40,
        h: 2 + rand() * 20,
      }));
      // Vary the cell size across trials — correctness must not depend on it.
      const cell = trial % 3 === 0 ? 4 : trial % 3 === 1 ? gridCellFor(boxes) : 200;
      const hash = new BoxHash<GridBox>(cell);
      for (const b of boxes) hash.insert(b, b);
      for (let q = 0; q < 30; q++) {
        const query: GridBox = { x: rand() * 400, y: rand() * 300, w: 2 + rand() * 40, h: 2 + rand() * 20 };
        const brute = boxes.some((b) => overlaps(query, b));
        const viaHash = hash.some(query, (b) => overlaps(query, b));
        expect(viaHash, `cell=${cell}`).toBe(brute);
      }
    }
  });

  it("finds a large box overlapping a small query far from its origin cell", () => {
    const hash = new BoxHash<GridBox>(8); // cell far smaller than the box
    const big: GridBox = { x: 0, y: 0, w: 500, h: 5 };
    hash.insert(big, big);
    expect(hash.some({ x: 480, y: 1, w: 4, h: 3 }, () => true)).toBe(true);
    expect(hash.some({ x: 480, y: 100, w: 4, h: 3 }, () => true)).toBe(false);
  });

  it("gridCellFor returns at least the floor and grows with the biggest box", () => {
    expect(gridCellFor([])).toBe(8);
    expect(gridCellFor([{ w: 3, h: 4 }])).toBe(8);
    expect(gridCellFor([{ w: 50, h: 4 }])).toBe(50);
    expect(gridCellFor([{ w: 5, h: 42 }])).toBe(42);
  });
});
