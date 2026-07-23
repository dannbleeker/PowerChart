import { describe, expect, it } from "vitest";
import { buildChart } from "../src/core/chart";
import { paletteColor } from "../src/core/style";
import type { ChartConfig } from "../src/core/types";

/** Palette resolution and per-category / per-group colouring. */

describe("reordering categories carries the per-category colors", () => {
  /** The fill of each column, left to right. */
  const fills = (scene: { nodes: any[] }) =>
    scene.nodes
      .filter((n) => n.kind === "rect" && /^(seg|col|bar)-/.test(n.name ?? ""))
      .sort((a, b) => a.x - b.x)
      .map((n) => n.fill);

  const data = {
    categories: ["A", "B", "C"],
    // The highlight is declared on A, the SMALLEST value — so any sort moves it.
    series: [{ name: "S1", values: [10, 50, 30], colors: ["#ff0000", null, null] }],
  };

  it("categorySort moves a highlight with its data point", () => {
    const scene = buildChart({
      kind: "clustered",
      width: 480,
      height: 300,
      categorySort: "descending",
      data,
    } as ChartConfig);
    const order = scene.nodes
      .filter((n: any) => n.kind === "text" && n.name?.startsWith("category-"))
      .sort((a: any, b: any) => a.x - b.x)
      .map((n: any) => n.text);
    expect(order).toEqual(["B", "C", "A"]);
    // Red belongs to A, now rightmost. It used to stay at position 0 and paint B.
    const red = fills(scene)
      .map((f, i) => [i, f])
      .filter(([, f]) => f === "#ff0000");
    expect(red).toEqual([[order.indexOf("A"), "#ff0000"]]);
  });

  it("pareto moves a highlight with its data point", () => {
    const scene = buildChart({ kind: "clustered", width: 480, height: 300, pareto: true, data } as ChartConfig);
    const order = scene.nodes
      .filter((n: any) => n.kind === "text" && n.name?.startsWith("category-"))
      .sort((a: any, b: any) => a.x - b.x)
      .map((n: any) => n.text);
    expect(order).toEqual(["B", "C", "A"]);
    const red = fills(scene)
      .map((f, i) => [i, f])
      .filter(([, f]) => f === "#ff0000");
    expect(red).toEqual([[order.indexOf("A"), "#ff0000"]]);
  });
});

describe("paletteColor wraps by the palette's actual length", () => {
  it("indexes a short custom palette by its length, never % 8", () => {
    const pal = ["#111111", "#222222", "#333333"];
    expect(paletteColor(pal, 4)).toBe("#222222"); // 4 % 3 = 1, not undefined
    expect(paletteColor(pal, 3)).toBe("#111111");
  });
  it("guards an empty palette", () => {
    expect(paletteColor([], 0)).toBe("#888888");
  });
});
