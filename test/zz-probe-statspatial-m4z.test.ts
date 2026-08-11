import { describe, it } from "vitest";
import { buildChart } from "../src/core/chart";
import type { ChartConfig } from "../src/core/types";

const base = { width: 600, height: 400 };
const names = (s: { nodes: unknown[] }) => (s.nodes as { name?: string }[]).map((n) => n.name);

describe("probe", () => {
  it("tilemap map key reaching Object.prototype", () => {
    for (const key of ["constructor", "toString", "__proto__", "hasOwnProperty", "valueOf"]) {
      const cfg = {
        ...base,
        kind: "tilemap",
        map: key,
        data: { categories: ["CA", "TX", "NY"], series: [{ name: "v", values: [1, 2, 3] }] },
      } as unknown as ChartConfig;
      let out: string;
      try {
        const s = buildChart(cfg);
        out = JSON.stringify(names(s));
      } catch (e) {
        out = "THREW: " + (e as Error).message;
      }
      console.log("MAPKEY", key, out);
    }
  });

  it("tilemap unknown map string (control)", () => {
    const cfg = {
      ...base,
      kind: "tilemap",
      map: "usa",
      data: { categories: ["CA", "TX", "NY"], series: [{ name: "v", values: [1, 2, 3] }] },
    } as unknown as ChartConfig;
    const s = buildChart(cfg);
    console.log("MAPKEY usa", JSON.stringify(names(s)));
  });

  it("treemap single item", () => {
    const cfg = {
      ...base,
      kind: "treemap",
      data: { categories: ["only"], series: [{ name: "v", values: [42] }] },
    } as unknown as ChartConfig;
    const s = buildChart(cfg);
    console.log("TREEMAP1", JSON.stringify(s.nodes.filter((n) => n.kind === "rect")));
  });

  it("treemap two items skewed", () => {
    const cfg = {
      ...base,
      kind: "treemap",
      data: { categories: ["a", "b", "c"], series: [{ name: "v", values: [1000000, 1, 1] }] },
    } as unknown as ChartConfig;
    const s = buildChart(cfg);
    console.log("TREEMAPSKEW", JSON.stringify(s.nodes.filter((n) => n.kind === "rect")));
  });

  it("mekko zero column", () => {
    const cfg = {
      ...base,
      kind: "mekko",
      data: {
        categories: ["A", "B", "C"],
        series: [
          { name: "s1", values: [10, 0, 5] },
          { name: "s2", values: [20, 0, 5] },
        ],
      },
    } as unknown as ChartConfig;
    const s = buildChart(cfg);
    console.log(
      "MEKKO0",
      JSON.stringify(s.nodes.filter((n) => /^(seg|total|category)-/.test((n as { name?: string }).name ?? ""))),
    );
  });

  it("mekko units with a zero extent", () => {
    const cfg = {
      ...base,
      kind: "mekko",
      data: {
        categories: ["A", "B"],
        xExtent: [10, 0],
        series: [
          { name: "s1", values: [10, 8] },
          { name: "s2", values: [20, 4] },
        ],
      },
    } as unknown as ChartConfig;
    const s = buildChart(cfg);
    console.log(
      "MEKKOUNITS",
      JSON.stringify(s.nodes.filter((n) => /^(seg|total)-/.test((n as { name?: string }).name ?? ""))),
    );
  });
});
