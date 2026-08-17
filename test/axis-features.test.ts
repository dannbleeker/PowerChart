import { describe, expect, it } from "vitest";
import { buildChart } from "../src/core/chart";
import { sceneToSvg } from "../src/render/svg";
import { valueScale } from "../src/core/layout/frame";
import { formatDay, formatDayRange, monthStarts, parseDateToken } from "../src/core/format";
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

/**
 * A day number JS cannot make a Date of.
 *
 * `formatDay` had no finite guard, so `MONTHS[NaN]` was `undefined` and
 * `getUTCDate()` was `NaN`, and the result — the literal text
 * `"NaN undefined"` — went straight into the scene. `finiteNodes` does not
 * catch it: that net filters non-finite NUMBERS, and this is a string.
 *
 * The route is ordinary. A Gantt whose dates arrived as epoch SECONDS (or
 * milliseconds) is what an export or an agent produces, and `data.dates` is a
 * plain passthrough flag rather than something the engine derives — so the
 * chart drew 18 of its 21 text nodes as "NaN undefined", into the SVG preview,
 * the live add-in, and the .pptx the skill hands back with exit 0.
 */
describe("a date label that cannot be computed", () => {
  it("is blank rather than the words NaN undefined", () => {
    for (const bad of [1e12, -1e12, NaN, Infinity, -Infinity, 1e8 + 1]) {
      expect(formatDay(bad), `formatDay(${bad})`).toBe("");
      expect(formatDay(bad, true), `formatDay(${bad}, withYear)`).toBe("");
    }
    // Unchanged for a day number that is a real date.
    expect(formatDay(20000)).toMatch(/[A-Z][a-z]{2}/);
  });

  it("keeps that text out of a chart built from seconds-since-epoch", () => {
    // The end-to-end case, in the units the mistake actually arrives in.
    const cfg = {
      kind: "gantt",
      title: "Q1 plan",
      data: {
        dates: true,
        categories: ["Design", "Build"],
        series: [
          { name: "Start", values: [1767225600, 1767225660] },
          { name: "End", values: [1767225660, 1767225700] },
        ],
      },
    } as unknown as ChartConfig;
    const svg = sceneToSvg(buildChart(cfg));
    expect(svg, "a chart shipped the words NaN undefined as its labels").not.toContain("NaN undefined");
    expect(svg).not.toContain("undefined");
  });
});

/**
 * A day that does not EXIST in its month is not a date.
 *
 * `Date.UTC` and `Date.parse` normalise rather than refuse, so `Feb 29 2023` —
 * a plausible mistake about a leap year — came back as 1 March 2023, and
 * `Apr 31` as 1 May. On a Gantt row that is a task silently starting a month
 * late, with nothing in the chart or the log to say so.
 *
 * The parser was already inconsistent about it: `Jan 32 2024` was refused,
 * because 32 is not a day number at all, while `Apr 31` was accepted and
 * quietly moved.
 */
describe("a date that does not exist", () => {
  it("is refused rather than rolled into the next month", () => {
    expect(parseDateToken("Feb 29 2023"), "a non-leap 29 February became 1 March").toBeNull();
    expect(parseDateToken("Apr 31 2024"), "31 April became 1 May").toBeNull();
    expect(parseDateToken("2024-02-30"), "an ISO 30 February became 1 March").toBeNull();
    expect(parseDateToken("30.02.2024"), "a dotted 30 February became 1 March").toBeNull();
    expect(parseDateToken("Sep 31"), "31 September became 1 October").toBeNull();
  });

  it("still parses every real date, including the leap day", () => {
    // The other half, or the rule above would be satisfied by refusing
    // everything — and this parser's whole job is to say yes to a date cell.
    expect(parseDateToken("Feb 29 2024"), "the leap day was refused").not.toBeNull();
    expect(parseDateToken("29.02.2024"), "the dotted leap day was refused").not.toBeNull();
    expect(parseDateToken("2024-02-29")).not.toBeNull();
    expect(parseDateToken("Apr 30 2024")).not.toBeNull();
    // Forms that name no unambiguous day, or none at all, are untouched.
    expect(parseDateToken("Jan-24"), "the mmm-yy form stopped parsing").not.toBeNull();
    expect(parseDateToken("2024-01-15T10:00:00Z"), "an ISO date-time stopped parsing").not.toBeNull();
    expect(parseDateToken("15 Jan 2024")).toBe(parseDateToken("Jan 15 2024"));
  });
});

/**
 * A span's ends are two specific DAYS, and `formatDay` is not the way to say so.
 *
 * `formatDay` collapses the first of a month to the bare month name, which is
 * right for the tick strip — a month gridline reads `Mar`, and a day number on
 * every one of them would be noise. `spanLabel` in the Gantt reused it for a
 * bar's date RANGE, where the day number is the whole point: a task running
 * 1 Jan to 31 Mar was labelled `Jan–31 Mar`, one running 1 Jan to 1 Apr — the
 * ordinary shape of a quarterly plan — was labelled `Jan–Apr`, and a roadmap
 * whose phases ran whole years read `Jan–Jan`, carrying no dates at all.
 */
describe("a calendar span names both of its days", () => {
  const d = (s: string) => parseDateToken(s)!;

  it("keeps the day number on a month start", () => {
    expect(formatDayRange(d("2026-01-01"), d("2026-03-31"))).toBe("1 Jan–31 Mar");
    expect(formatDayRange(d("2026-03-01"), d("2026-03-08"))).toBe("1 Mar–8 Mar");
    expect(formatDayRange(d("2026-01-01"), d("2026-04-01"))).toBe("1 Jan–1 Apr");
  });

  it("adds the year when the span crosses one", () => {
    expect(formatDayRange(d("2000-01-01"), d("2030-01-01"))).toBe("1 Jan 00–1 Jan 30");
    expect(formatDayRange(d("2026-12-20"), d("2027-01-10"))).toBe("20 Dec 26–10 Jan 27");
    // …and leaves a within-year span alone, so an ordinary plan is unchanged.
    expect(formatDayRange(d("2026-02-05"), d("2026-02-19"))).toBe("5 Feb–19 Feb");
  });

  it("says nothing it cannot compute", () => {
    // Same answer `formatDay` gives a day number JS cannot make a Date of: a
    // blank end rather than "NaN undefined" in the .pptx.
    expect(formatDayRange(NaN, d("2026-01-01"))).toBe("–1 Jan");
    expect(formatDayRange(d("2026-01-01"), Infinity)).toBe("1 Jan–");
  });
});
