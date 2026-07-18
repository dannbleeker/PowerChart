import { describe, expect, it } from "vitest";
import { buildChart } from "../src/core/chart";
import type { ChartConfig } from "../src/core/types";
import type { RectNode } from "../src/core/scene";

/**
 * squarify now carries the row's running max/min/sum instead of re-slicing and
 * re-scanning each step (O(L²) → O(L) per strip). It's byte-identical — same
 * max/min/sum feed the same aspect-ratio test — so the snapshot suite pins the
 * exact layout; this adds a structural invariant on a larger treemap the samples
 * don't cover.
 */
describe("treemap squarify stays correct and deterministic", () => {
  const cfg: ChartConfig = {
    kind: "treemap",
    width: 640,
    height: 400,
    data: {
      categories: Array.from({ length: 60 }, (_, i) => `T${i}`),
      series: [{ name: "S", values: Array.from({ length: 60 }, (_, i) => ((i * 17) % 50) + 1) }],
    },
  };

  it("emits one non-degenerate tile per item", () => {
    const tiles = buildChart(cfg).nodes.filter(
      (n): n is RectNode => n.kind === "rect" && !!n.name?.startsWith("tile-"),
    );
    expect(tiles).toHaveLength(60);
    for (const t of tiles) {
      expect(t.w).toBeGreaterThan(0);
      expect(t.h).toBeGreaterThan(0);
    }
  });

  it("is deterministic across rebuilds", () => {
    expect(JSON.stringify(buildChart(cfg).nodes)).toBe(JSON.stringify(buildChart(cfg).nodes));
  });
});
