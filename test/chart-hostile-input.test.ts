import { describe, it, expect } from "vitest";
import { buildChart } from "../src/core/chart";
import { niceTicks } from "../src/core/format";
import { BoxHash } from "../src/core/grid";
import { sampleConfig, CHART_KINDS } from "../src/core/samples";
import type { ChartConfig } from "../src/core/types";
import { sceneToSvg } from "../src/render/svg";
import { buildDeckBase64 } from "../src/render/pptx-deck";
import { toRgb, alphaOf } from "../src/core/color";

/**
 * Values a datasheet cell can hold, against every chart kind.
 *
 * Both bugs this file pins are the same shape: a loop whose bound came from
 * the data, with nothing between the number someone typed and the number of
 * iterations. Neither was slow — both were unbounded, in time and in memory,
 * and the tab dies before the chart does.
 */

const HOSTILE: [string, (c: ChartConfig) => ChartConfig][] = [
  ["huge values", (c) => vals(c, () => 1e308)],
  ["tiny values", (c) => vals(c, () => 5e-324)],
  ["NaN value", (c) => vals(c, () => NaN)],
  ["Infinity value", (c) => vals(c, () => Infinity)],
  ["-Infinity value", (c) => vals(c, () => -Infinity)],
  ["all zero", (c) => vals(c, () => 0)],
  ["negative values", (c) => vals(c, (v) => -Math.abs(Number(v) || 1))],
  ["no series", (c) => ({ ...c, data: { ...c.data, series: [] } })],
  ["empty values", (c) => ({ ...c, data: { ...c.data, series: c.data.series.map((s) => ({ ...s, values: [] })) } })],
  ["one category", (c) => ({ ...c, data: { ...c.data, categories: c.data.categories.slice(0, 1) } })],
  ["no categories", (c) => ({ ...c, data: { ...c.data, categories: [] } })],
  ["zero size", (c) => ({ ...c, width: 0, height: 0 })],
  ["negative size", (c) => ({ ...c, width: -100, height: -50 })],
  ["nan size", (c) => ({ ...c, width: NaN, height: NaN })],
];

function vals(c: ChartConfig, f: (v: unknown) => number): ChartConfig {
  return { ...c, data: { ...c.data, series: c.data.series.map((s) => ({ ...s, values: s.values.map(f) })) } };
}

describe("a hostile number in a cell cannot hang the chart engine", () => {
  it("builds every kind against every hostile input, and returns", () => {
    // The whole grid, timed as one. Before the two fixes below this did not
    // fail — it never finished, and vitest killed the run at its timeout.
    const t0 = Date.now();
    for (const { kind } of CHART_KINDS) {
      for (const [name, mutate] of HOSTILE) {
        expect(() => buildChart(mutate(sampleConfig(kind))), `${kind} / ${name}`).not.toThrow();
      }
    }
    expect(Date.now() - t0).toBeLessThan(30_000);
  }, 60_000);

  it("does not fill memory generating ticks for an axis near MAX_VALUE", () => {
    // `ceil(1.7e308 / 5e307)` is 4, and 4 x 5e307 is Infinity — so the loop
    // bound was Infinity while both its inputs were finite. The entry guard
    // checks the inputs; these are outputs.
    for (const max of [1e300, 1e307, 1.7e308, Number.MAX_VALUE]) {
      const ticks = niceTicks(0, max);
      expect(ticks.length, `max=${max}`).toBeGreaterThan(0);
      expect(ticks.length, `max=${max}`).toBeLessThan(1100);
      for (const t of ticks) expect(Number.isFinite(t), `max=${max} tick ${t}`).toBe(true);
    }
  });

  it("does not enumerate a grid cell per unit of a runaway box", () => {
    // The label index sizes its cell to the largest label, so a real label
    // covers a handful of cells. A label placed at 1e300 — which a value of
    // 1e308 produces — spans about 1e298 of them, and each one allocated a Map
    // key. One box, unbounded memory.
    const hash = new BoxHash<number>(8);
    const t0 = Date.now();
    hash.insert({ x: 0, y: 0, w: 1e300, h: 1e300 }, 1);
    hash.insert({ x: NaN, y: NaN, w: NaN, h: NaN }, 2);
    hash.insert({ x: 0, y: 0, w: Infinity, h: Infinity }, 3);
    expect(Date.now() - t0).toBeLessThan(1000);
    // The finite-but-huge box is still findable — degenerate, not discarded.
    expect(hash.some({ x: 0, y: 0, w: 8, h: 8 }, (v) => v === 1)).toBe(true);
  });

  it("still indexes an ordinary label across the cells it really covers", () => {
    // The negative control: the cap must not have turned the index into a
    // single bucket, which would make every label collide with every other.
    const hash = new BoxHash<string>(8);
    hash.insert({ x: 0, y: 0, w: 20, h: 20 }, "near");
    hash.insert({ x: 400, y: 400, w: 20, h: 20 }, "far");
    expect(hash.some({ x: 0, y: 0, w: 4, h: 4 }, (v) => v === "near")).toBe(true);
    expect(hash.some({ x: 0, y: 0, w: 4, h: 4 }, (v) => v === "far")).toBe(false);
  });
});

describe("no non-finite geometry leaves the chart engine", () => {
  /** Every number anywhere in a node — including a polygon's point list. */
  function badNumbers(v: unknown, path: string, out: string[], depth = 0): void {
    if (depth > 6) return;
    if (typeof v === "number") {
      if (!Number.isFinite(v)) out.push(`${path} = ${v}`);
      return;
    }
    if (Array.isArray(v)) return v.forEach((e, i) => badNumbers(e, `${path}[${i}]`, out, depth + 1));
    if (v && typeof v === "object")
      for (const [k, x] of Object.entries(v)) badNumbers(x, `${path}.${k}`, out, depth + 1);
  }

  it("holds for every kind against every hostile input", () => {
    // Nine kinds could emit `rect.y = -Infinity`, `wedge.endAngle = NaN`,
    // `polygon.points[0].x = Infinity` and so on. The SVG renderer defends
    // itself — every numeric goes through `num()` — but the two PowerPoint
    // renderers do not, and they are the ones that write a file. A NaN there
    // lands in `addGeometricShape({left: NaN, …})` and in OOXML as an EMU
    // value, so the produced .pptx is one PowerPoint may refuse to open.
    const bad: string[] = [];
    for (const { kind } of CHART_KINDS) {
      for (const [name, mutate] of HOSTILE) {
        const scene = buildChart(mutate(sampleConfig(kind)));
        scene.nodes.forEach((n, i) => badNumbers(n, `${kind}/${name}#${i}(${n.kind})`, bad));
      }
    }
    expect(bad.slice(0, 20)).toEqual([]);
  }, 60_000);

  it("drops nothing from a chart whose numbers are all real", () => {
    // The floor is not a filter. If this ever shrinks a valid chart, the guard
    // has started eating geometry rather than guarding it — and the snapshots
    // would move, which is the louder half of the same alarm.
    for (const { kind } of CHART_KINDS) {
      const scene = buildChart(sampleConfig(kind));
      expect(scene.nodes.length, `${kind} lost nodes`).toBeGreaterThan(0);
      const bad: string[] = [];
      scene.nodes.forEach((n, i) => badNumbers(n, `${kind}#${i}`, bad));
      expect(bad).toEqual([]);
    }
  });

  it("repairs a poisoned font size instead of laying out around it", () => {
    // `style.fontSize` is arithmetic, not decoration: label boxes, decoration
    // lifts and half the axis geometry derive from it. A string there — and a
    // config arrives from JSON import, from a template, from a config tag
    // authored in another deck — turned every one of those into NaN. Dropping
    // the poisoned nodes would leave an empty chart; repairing the value on
    // the way in leaves the chart someone asked for.
    const hostile = { ...sampleConfig("clustered"), style: { fontSize: '10"><script>' } } as unknown as ChartConfig;
    const poisoned = buildChart(hostile);
    const clean = buildChart(sampleConfig("clustered"));
    expect(poisoned.nodes.length).toBe(clean.nodes.length);
  });
});

/**
 * The STYLE half of the same question.
 *
 * The suite above pins what a hostile *cell* does. Nothing pinned what a
 * hostile *style* does, and the answer turned out to be "throws": a palette of
 * numbers reached `toRgb`, whose `(color ?? "").trim()` guarded null and
 * undefined and nothing else, and the renderer died with
 * `TypeError: … .trim is not a function`.
 *
 * Every route into a style is user JSON, and none of them is exotic. The pane's
 * style import stores whatever parses and applies it to every chart from then
 * on; a whole config can be pasted into the JSON box; the skill takes one from
 * an agent. The type says `string[]`, and TypeScript checks the code rather
 * than the file someone pastes.
 */
const HOSTILE_STYLES: [string, unknown][] = [
  ["palette of numbers", { palette: [1, 2, 3] }],
  ["palette of nulls", { palette: [null, undefined] }],
  ["palette of objects", { palette: [{}, []] }],
  ["palette is a string", { palette: "red" }],
  ["palette is an object", { palette: { 0: "#fff" } }],
  ["palette is empty", { palette: [] }],
  ["negative is a number", { negative: 1 }],
  ["neutral is an array", { neutral: ["#fff"] }],
  ["fontFamily is a number", { fontFamily: 12 }],
  ["fontSize is Infinity", { fontSize: Infinity }],
];

describe("a hostile style cannot take the renderer down", () => {
  it("builds and draws every kind against every hostile style", () => {
    for (const { kind } of CHART_KINDS) {
      for (const [name, style] of HOSTILE_STYLES) {
        const cfg = { ...sampleConfig(kind), style } as unknown as ChartConfig;
        expect(() => sceneToSvg(buildChart(cfg)), `${kind} / ${name}`).not.toThrow();
      }
    }
  }, 60_000);

  it("reads a non-string paint as an unrecognised one, not as a crash", () => {
    // The specific fix, at the specific function, so a later refactor of the
    // sweep above cannot quietly lose it. Mid grey is what an unrecognised
    // paint already resolved to (named CSS colours do the same), so a bad one
    // degrades exactly like an unknown one instead of taking the chart with it.
    for (const junk of [1, null, undefined, {}, [], true, NaN]) {
      expect(() => toRgb(junk as unknown as string), `toRgb(${String(junk)})`).not.toThrow();
      expect(toRgb(junk as unknown as string)).toEqual([128, 128, 128]);
      expect(() => alphaOf(junk as unknown as string), `alphaOf(${String(junk)})`).not.toThrow();
      expect(alphaOf(junk as unknown as string)).toBe(1);
    }
  });
});

/**
 * Every top-level config key, holding the wrong type.
 *
 * The two suites above ask what a bad CELL and a bad STYLE do. This asks the
 * broader question — what does a config with the wrong type ANYWHERE do — and
 * it found three crashes the narrower ones could not:
 *
 * - `title: 2024` and `valueAxisTitle: [...]` threw `s.replace is not a
 *   function` out of `xmlText`. A number for a title is what someone writes for
 *   a year, not an exotic mistake.
 * - `numberFormat: null` threw on `.decimals`. A default parameter only fires
 *   for `undefined`, and `null` is what a serialiser writes for an absent
 *   field.
 *
 * Both renderers, because they are separate code that has disagreed before: the
 * SVG one the preview uses, and the pptx one the skill and the file path use.
 * Ten kinds rather than all of them keeps this inside a few seconds while still
 * covering every family — the failures found were never kind-specific.
 */
const TOP_LEVEL_KEYS =
  "horizontal scale segmentOrder categorySort secondaryAxis axisBreak valueAxisTitle labelOffsets logScale gapWidth overlap footnote pie pareto multiples boxplot map heatmap otherBucket tilemap butterfly scatter gantt radar combo width height title decorations waterfall numberFormat labels render".split(
    " ",
  );

const WRONG_TYPES: [string, unknown][] = [
  ["string", "x"],
  ["number", 7],
  ["negative", -7],
  ["NaN", NaN],
  ["Infinity", Infinity],
  ["true", true],
  ["null", null],
  ["array", [1, "a", null]],
  ["object", { a: 1 }],
  ["nested junk", { max: "x", min: [], value: {} }],
];

describe("a config key of the wrong type cannot take a renderer down", () => {
  for (const key of TOP_LEVEL_KEYS) {
    it(`${key}`, async () => {
      const bad: string[] = [];
      for (const { kind } of CHART_KINDS.slice(0, 10)) {
        for (const [label, value] of WRONG_TYPES) {
          const cfg = { ...sampleConfig(kind), [key]: value } as unknown as ChartConfig;
          try {
            const scene = buildChart(cfg);
            sceneToSvg(scene);
            await buildDeckBase64([{ scene, title: "t", configJson: "{}", slot: 0, run: "r" }], {
              width: 720,
              height: 405,
            });
          } catch (e) {
            bad.push(`${kind}/${key}=${label}: ${e instanceof Error ? e.message : String(e)}`);
          }
        }
      }
      // Sliced so a broad regression reports the first few rather than a wall.
      expect(bad.slice(0, 6)).toEqual([]);
    }, 30_000);
  }
});

/**
 * The DATA's shape, rather than the numbers inside it.
 *
 * `normalizeData` already coerced four things — categories not an array, series
 * not an array, values not an array, non-finite cells. It stopped one level
 * short of the ones that actually turn up:
 *
 * - **`categories: [2023, 2024]`.** Not a hostile input at all — years,
 *   quarters, ids — and it threw `raw.trim is not a function` and
 *   `c.split is not a function` out of the layout.
 * - **`series[].name: 5`.** Same, from the same authors: `s.name.trim`.
 * - **A null entry in `series`.** It survived as `{values: [...]}` with no
 *   name, and every consumer reaching for `s.name` died on it. JSON arrays
 *   hold nulls.
 * - **`data: null`.** Two of the four reads used `data?.`; the other two did
 *   not, so `data.hundredPercent` threw. Half a guard.
 *
 * Fixed in `normalizeData` rather than at each consumer, because that function
 * exists to make the layout's types honest and these were the cases it was
 * missing — patching the consumers would have been the same fix written eleven
 * times and forgotten on the twelfth.
 */
const HOSTILE_SHAPES: [string, (c: ChartConfig) => unknown][] = [
  ["categories are numbers", (c) => ({ ...c, data: { ...c.data, categories: [1, 2, 3] } })],
  ["categories are null", (c) => ({ ...c, data: { ...c.data, categories: [null, null] } })],
  ["categories are objects", (c) => ({ ...c, data: { ...c.data, categories: [{}, []] } })],
  ["categories is a string", (c) => ({ ...c, data: { ...c.data, categories: "abc" } })],
  ["categories is null", (c) => ({ ...c, data: { ...c.data, categories: null } })],
  ["series is null", (c) => ({ ...c, data: { ...c.data, series: null } })],
  ["series is an object", (c) => ({ ...c, data: { ...c.data, series: { a: 1 } } })],
  ["series entries are null", (c) => ({ ...c, data: { ...c.data, series: [null, null] } })],
  ["series name is a number", (c) => withSeries(c, (s) => ({ ...s, name: 5 }))],
  ["series values is a string", (c) => withSeries(c, (s) => ({ ...s, values: "abc" }))],
  ["series values is null", (c) => withSeries(c, (s) => ({ ...s, values: null }))],
  ["series values are strings", (c) => withSeries(c, (s) => ({ ...s, values: s.values.map(() => "x") }))],
  ["series type is unknown", (c) => withSeries(c, (s) => ({ ...s, type: "zzz" }))],
  ["series colors is a string", (c) => withSeries(c, (s) => ({ ...s, colors: "red" }))],
  ["series pattern is a number", (c) => withSeries(c, (s) => ({ ...s, pattern: 3 }))],
  ["series scenario is an object", (c) => withSeries(c, (s) => ({ ...s, scenario: {} }))],
  ["data is null", (c) => ({ ...c, data: null })],
  ["data is an array", (c) => ({ ...c, data: [] })],
];

function withSeries(c: ChartConfig, f: (s: ChartConfig["data"]["series"][number]) => unknown) {
  return { ...c, data: { ...c.data, series: c.data.series.map(f) } };
}

describe("a hostile data SHAPE cannot take the renderer down", () => {
  for (const [name, mutate] of HOSTILE_SHAPES) {
    it(name, () => {
      const bad: string[] = [];
      for (const { kind } of CHART_KINDS) {
        try {
          sceneToSvg(buildChart(mutate(sampleConfig(kind)) as ChartConfig));
        } catch (e) {
          bad.push(`${kind}: ${e instanceof Error ? e.message : String(e)}`);
        }
      }
      expect(bad.slice(0, 4)).toEqual([]);
    });
  }
});
