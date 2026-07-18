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
