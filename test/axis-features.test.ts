import { describe, expect, it } from "vitest";
import { valueScale } from "../src/core/layout/frame";
import { formatDay, monthStarts, parseDateToken } from "../src/core/format";
import { valueExtent } from "../src/core/chart";
import { layoutColumns } from "../src/core/layout/column";
import { layoutGantt } from "../src/core/layout/gantt";
import { DEFAULT_DECOR, DEFAULT_STYLE } from "../src/core/style";
import type { ChartConfig } from "../src/core/types";
import type { TextNode } from "../src/core/scene";

const frame = { x: 0, y: 0, w: 100, h: 100 };

function cfg(partial: Partial<ChartConfig>): ChartConfig {
  return { kind: "stacked", width: 480, height: 300, data: { categories: [], series: [] }, ...partial };
}

describe("axis break", () => {
  const scale = valueScale(frame, 0, 100, undefined, { from: 10, to: 90 });

  it("compresses the broken range into a small band", () => {
    const band = scale.toY(10) - scale.toY(90);
    expect(band / frame.h).toBeCloseTo(0.06, 5);
  });

  it("keeps segments outside the break proportional", () => {
    // below (0..10) and above (90..100) are equal spans → equal pixel spans.
    const below = scale.toY(0) - scale.toY(10);
    const above = scale.toY(90) - scale.toY(100);
    expect(below).toBeCloseTo(above, 5);
    expect(scale.toY(0)).toBeCloseTo(frame.y + frame.h, 5);
    expect(scale.toY(100)).toBeCloseTo(frame.y, 5);
  });

  it("drops ticks inside the break and exposes the band", () => {
    expect(scale.ticks.every((t) => t <= 10 || t >= 90)).toBe(true);
    expect(scale.breakBand).toBeTruthy();
  });

  it("emits break markers in column charts", () => {
    const c = cfg({
      kind: "clustered",
      data: { categories: ["A", "B"], series: [{ name: "S", values: [10, 600] }] },
      axisBreak: { from: 50, to: 550 },
    });
    const { nodes } = layoutColumns(c, DEFAULT_STYLE, DEFAULT_DECOR);
    expect(nodes.find((n) => n.name === "axis-break")).toBeTruthy();
    expect(nodes.find((n) => n.name === "axis-break-lo")).toBeTruthy();
  });
});

describe("calendar dates", () => {
  it("parses common date formats to epoch days", () => {
    const iso = parseDateToken("2026-01-15")!;
    expect(new Date(iso * 86400000).toISOString().slice(0, 10)).toBe("2026-01-15");
    expect(parseDateToken("15.01.2026")).toBe(iso);
    expect(parseDateToken("2026-01")).toBe(iso - 14);
    expect(parseDateToken("42")).toBeNull();
    expect(parseDateToken("")).toBeNull();
  });

  it("enumerates month starts and formats labels", () => {
    const jan = parseDateToken("2026-01-01")!;
    const apr = parseDateToken("2026-04-01")!;
    const months = monthStarts(jan, apr);
    expect(months).toHaveLength(4);
    expect(formatDay(jan, true)).toBe("Jan 26");
    expect(formatDay(parseDateToken("2026-02-05")!)).toBe("5 Feb");
  });

  it("renders a month timeline in calendar Gantt", () => {
    const jan = parseDateToken("2026-01-05")!;
    const may = parseDateToken("2026-05-20")!;
    const c = cfg({
      kind: "gantt",
      data: {
        categories: ["A"],
        series: [
          { name: "Start", values: [jan] },
          { name: "End", values: [may] },
        ],
        dates: true,
      },
    });
    const { nodes } = layoutGantt(c, DEFAULT_STYLE, DEFAULT_DECOR);
    const header = nodes.filter((n): n is TextNode => n.kind === "text" && n.name === "timeline");
    expect(header.length).toBeGreaterThanOrEqual(4);
    expect(header.some((h) => /Jan|Feb|Mar/.test(h.text))).toBe(true);
  });
});

describe("valueExtent (Same Scale)", () => {
  it("uses stacked totals for stacked charts", () => {
    const e = valueExtent(
      cfg({
        data: {
          categories: ["A", "B"],
          series: [
            { name: "S1", values: [10, 20] },
            { name: "S2", values: [5, 15] },
          ],
        },
      }),
    );
    expect(e).toEqual({ min: 0, max: 35 });
  });

  it("tracks the running total for waterfalls", () => {
    const e = valueExtent(
      cfg({
        kind: "waterfall",
        data: { categories: ["a", "b", "c"], series: [{ name: "D", values: [100, -30, 50] }] },
      }),
    );
    expect(e).toEqual({ min: 0, max: 120 });
  });

  it("returns null for charts without a value axis", () => {
    expect(
      valueExtent(cfg({ kind: "mekko", data: { categories: ["A"], series: [{ name: "S", values: [1] }] } })),
    ).toBeNull();
  });
});

describe("valueScale zero floor", () => {
  const frame2 = { x: 0, y: 0, w: 100, h: 100 };

  it("includes zero by default (column charts baseline at 0)", () => {
    const s = valueScale(frame2, 40, 95);
    expect(s.min).toBeLessThanOrEqual(0);
    expect(s.max).toBeGreaterThanOrEqual(95);
  });

  it("keeps the domain data-driven when zeroFloor is false (distributions)", () => {
    // A boxplot/violin/candlestick of 40–95 must not be squashed against 0.
    const s = valueScale(frame2, 40, 95, undefined, undefined, undefined, false);
    expect(s.min).toBeGreaterThan(0); // niceTicks rounds down from 40, still well above 0
    expect(s.max).toBeGreaterThanOrEqual(95);
    // The data occupies a real fraction of the plot, not a thin sliver at the top.
    const frac = (s.toY(40) - s.toY(95)) / frame2.h;
    expect(frac).toBeGreaterThan(0.6);
  });

  it("still honours a manual cfg.scale override with zeroFloor off", () => {
    const s = valueScale(frame2, 40, 95, { min: 0, max: 100 }, undefined, undefined, false);
    expect(s.min).toBe(0);
    expect(s.max).toBe(100);
  });
});

describe("degenerate manual scale", () => {
  const frame2 = { x: 0, y: 0, w: 100, h: 100 };

  it("keeps a positive span when the scale cannot hold the data", () => {
    // scale.min at/above the auto max filtered the ticks down to one value, so
    // `max - min || 1` mapped one data unit to one point.
    for (const override of [{ min: 100 }, { min: 100, max: 100 }, { min: 0, max: 0 }]) {
      const s = valueScale(frame2, 0, 20, override);
      expect(s.max, JSON.stringify(override)).toBeGreaterThan(s.min);
      // ...and the ink lands on the plot, not tens of canvas heights away.
      for (const v of [0, 10, 20]) {
        expect(s.toY(v), JSON.stringify(override)).toBeGreaterThanOrEqual(frame2.y - frame2.h);
        expect(s.toY(v), JSON.stringify(override)).toBeLessThanOrEqual(frame2.y + 2 * frame2.h);
      }
    }
  });

  it("draws the columns inside the chart for scale:{min:100} on data 10/20", () => {
    const c = cfg({
      kind: "clustered",
      data: { categories: ["A", "B"], series: [{ name: "S", values: [10, 20] }] },
      scale: { min: 100 },
    });
    const { nodes } = layoutColumns(c, DEFAULT_STYLE, DEFAULT_DECOR);
    const segs = nodes.filter((n) => n.name?.startsWith("seg-")) as { y: number; h: number }[];
    expect(segs.length).toBe(2);
    for (const s of segs) {
      expect(s.y).toBeGreaterThanOrEqual(0);
      expect(s.y + s.h).toBeLessThanOrEqual(c.height);
    }
  });
});

describe("value-axis tick labels", () => {
  const labelsOf = (partial: Partial<ChartConfig>): string[] =>
    layoutColumns(cfg({ kind: "clustered", ...partial }), DEFAULT_STYLE, { ...DEFAULT_DECOR, valueAxis: true })
      .nodes.filter((n): n is TextNode => n.kind === "text" && n.name === "value-axis")
      .map((n) => n.text);

  it("are distinct on a narrow axis (precision from the step, not the magnitude)", () => {
    // 7.444–7.471 used to print ["7.4","7.5","7.5","7.5","7.5"]: five gridlines,
    // two labels, and a top tick named as a value outside the scale.
    const labels = labelsOf({
      data: { categories: ["Mon", "Tue"], series: [{ name: "DKK/EUR", values: [7.4442, 7.4708] }] },
      scale: { min: 7.44, max: 7.48 },
    });
    expect(labels).toEqual(["7.44", "7.45", "7.46", "7.47", "7.48"]);
  });

  it("never prints a log-axis decade below 1 as '0'", () => {
    const labels = labelsOf({
      data: { categories: ["A", "B", "C"], series: [{ name: "S", values: [3, 30, 300] }] },
      logScale: true,
    });
    expect(labels).not.toContain("0");
    expect(new Set(labels).size).toBe(labels.length);
  });

  it("still reads plain on an ordinary axis", () => {
    expect(labelsOf({ data: { categories: ["A", "B"], series: [{ name: "S", values: [10, 2500] }] } })).toEqual([
      "0",
      "1,000",
      "2,000",
      "3,000",
    ]);
  });
});
