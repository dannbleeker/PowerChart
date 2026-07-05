// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from "vitest";
import { buildChart, DEFAULT_SIZE } from "../src/core/chart";
import { loadThemePalette } from "../src/render/powerpoint";
import type { ChartConfig } from "../src/core/types";
import type { LineNode, RectNode } from "../src/core/scene";

/** Error bars from Error rows + theme palette from the host deck. */

const col = (extra: { name: string; values: (number | null)[] }[]): ChartConfig => ({
  kind: "stacked",
  ...DEFAULT_SIZE,
  data: {
    categories: ["A", "B", "C"],
    series: [{ name: "S", values: [20, 30, 25] }, ...extra],
  },
});

describe("error bars", () => {
  it("Error row draws symmetric whiskers at the column total and never renders as a segment", () => {
    const s = buildChart(col([{ name: "Error", values: [4, 6, null] }]));
    // Only the data series renders segments.
    expect(s.nodes.filter((n) => n.name?.startsWith("seg-1"))).toHaveLength(0);
    const bars = s.nodes.filter((n): n is LineNode => n.kind === "line" && !!n.name?.startsWith("error-") && !n.name.includes("cap"));
    expect(bars).toHaveLength(2); // C has no delta
    const bar = bars[0];
    const seg = s.nodes.find((n) => n.name === "seg-0-0") as RectNode;
    // Whisker is centered on the column and symmetric around its top.
    expect(bar.x1).toBeCloseTo(seg.x + seg.w / 2);
    expect(seg.y - bar.y1).toBeCloseTo(bar.y2 - seg.y, 1);
    expect(s.nodes.some((n) => n.name === "error-cap-hi-0")).toBe(true);
    expect(s.nodes.some((n) => n.name === "error-cap-lo-0")).toBe(true);
  });

  it("Error+/Error− rows give asymmetric whiskers with only their own caps", () => {
    const s = buildChart(col([
      { name: "Error+", values: [8, null, null] },
      { name: "Error-", values: [2, null, null] },
    ]));
    const bar = s.nodes.find((n): n is LineNode => n.kind === "line" && n.name === "error-0")!;
    const seg = s.nodes.find((n) => n.name === "seg-0-0") as RectNode;
    expect(seg.y - bar.y1).toBeGreaterThan((bar.y2 - seg.y) * 2); // +8 vs −2
    const onlyPlus = buildChart(col([{ name: "Error+", values: [8, null, null] }]));
    expect(onlyPlus.nodes.some((n) => n.name === "error-cap-hi-0")).toBe(true);
    expect(onlyPlus.nodes.some((n) => n.name === "error-cap-lo-0")).toBe(false);
  });

  it("widens the auto scale so whiskers stay inside the plot", () => {
    const plain = buildChart(col([]));
    const withErr = buildChart(col([{ name: "Error", values: [null, 40, null] }]));
    const topOf = (sc: typeof plain, name: string) => (sc.nodes.find((n) => n.name === name) as RectNode).y;
    // Column tops sit lower (larger y) when the scale grew to fit +40.
    expect(topOf(withErr, "seg-0-1")).toBeGreaterThan(topOf(plain, "seg-0-1"));
    const bar = withErr.nodes.find((n): n is LineNode => n.kind === "line" && n.name === "error-1")!;
    expect(bar.y1).toBeGreaterThan(0); // inside the chart
  });

  it("works on line charts (first series) and is ignored where unsupported", () => {
    const line = buildChart({ ...col([{ name: "Error", values: [3, 3, 3] }]), kind: "line" });
    expect(line.nodes.filter((n) => n.name === "error-0")).toHaveLength(1);
    const pie = buildChart({ ...col([{ name: "Error", values: [3, 3, 3] }]), kind: "pie" });
    // Unsupported kind: the row stays ordinary data, no whiskers.
    expect(pie.nodes.some((n) => n.name?.startsWith("error-"))).toBe(false);
  });
});

describe("loadThemePalette", () => {
  afterEach(() => vi.unstubAllGlobals());

  const host = (getThemeColor: (c: string) => { value: string }) => {
    const context = {
      presentation: {
        getSelectedSlides: () => ({ getItemAt: () => ({ themeColorScheme: { getThemeColor } }) }),
      },
      sync: async () => {},
    };
    vi.stubGlobal("PowerPoint", { run: async <T,>(cb: (c: typeof context) => Promise<T>) => cb(context) });
  };

  it("reads Accent1–6 and normalizes to #rrggbb", async () => {
    host((c) => ({ value: { Accent1: "2A78D6", Accent2: "#1BAF7A" }[c] ?? "EDA100" }));
    const palette = await loadThemePalette();
    expect(palette).toHaveLength(6);
    expect(palette![0]).toBe("#2a78d6");
    expect(palette![1]).toBe("#1baf7a");
  });

  it("returns null on hosts without the API or without a selection", async () => {
    vi.stubGlobal("PowerPoint", {
      run: async () => {
        throw new Error("ThemeColorScheme requires PowerPointApi 1.10");
      },
    });
    expect(await loadThemePalette()).toBeNull();
  });
});
