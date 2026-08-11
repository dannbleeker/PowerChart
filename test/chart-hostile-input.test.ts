import { describe, it, expect } from "vitest";
import { readFileSync } from "fs";
import { buildChart } from "../src/core/chart";
import { niceTicks } from "../src/core/format";
import { BoxHash } from "../src/core/grid";
import { sampleConfig, CHART_KINDS } from "../src/core/samples";
import type { ChartConfig } from "../src/core/types";
import { sceneToSvg } from "../src/render/svg";
import { buildDeckBase64 } from "../src/render/pptx-deck";
import { toRgb, alphaOf } from "../src/core/color";
import { finiteNodes, textWidth } from "../src/core/scene";
import type { SceneNode } from "../src/core/scene";

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
 *   not, so `data.hundredPercent` threw — half a guard. That one is now refused
 *   BY NAME rather than by accident (see below), because two tests depended on
 *   a config with no data failing, and they were right to: the batch renderer
 *   has to be able to tell its caller which chart was empty.
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
];

function withSeries(c: ChartConfig, f: (s: ChartConfig["data"]["series"][number]) => unknown) {
  return { ...c, data: { ...c.data, series: c.data.series.map(f) } };
}

describe("a config with no data at all is refused, in words", () => {
  it("names the problem instead of throwing a TypeError from the layout", () => {
    // The refusal is the point, and so is the message. `{kind: "pie"}` with no
    // data used to die on `Cannot read properties of null (reading
    // 'hundredPercent')` deep inside normalisation, which is the same outcome
    // reached by accident — and an accident is one refactor away from becoming
    // a silent empty chart, which is what briefly happened.
    for (const data of [null, undefined, [], "x", 7]) {
      const cfg = { ...sampleConfig("pie"), data } as unknown as ChartConfig;
      expect(() => buildChart(cfg), `data=${JSON.stringify(data)}`).toThrow(/no .?data.? object/);
    }
    // An EMPTY data object is not the same thing and stays legal — "no series"
    // and "no categories" are states with their own tests above.
    expect(() => buildChart({ ...sampleConfig("pie"), data: {} } as unknown as ChartConfig)).not.toThrow();
  });
});

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

/**
 * Decoration keys, holding the wrong type.
 *
 * The top-level sweep replaced `decorations` wholesale, which never reached
 * inside it. One key was unguarded: `labelContent` is a LIST, and a config
 * writing a bare `"value"` instead of `["value"]` — an easy thing to hand-write
 * or to generate — threw `parts.map is not a function`.
 *
 * It needed fixing twice. `segmentLabel` is the shared consumer, but scatter
 * builds its own label from the same key and so had its own copy of the bug —
 * which is exactly why this sweep runs every kind rather than a representative
 * few.
 */
/**
 * The keys, read from the type rather than remembered.
 *
 * `DECOR_KEYS` used to be a hand-written string of the decoration names
 * somebody thought of, and it happened to contain only the SCALAR ones. The
 * three list-valued decorations — `valueLines`, `callouts`, `bands`, each read
 * as `decor.<key>.forEach` with each entry a record of named fields — were
 * simply never in it, and all three crashed the renderer on a value of the
 * wrong shape: nine kinds, thirteen kinds, fifteen kinds, plus a null entry
 * dying one field later on `Cannot read properties of null`. Twelve crash modes
 * behind a sweep that reported clean.
 *
 * A sweep is only as good as the list it was given, and a hand-written list
 * goes stale in silence — the same shape of gap as the three colour sinks,
 * where each hole was found by a sweep aimed at that renderer alone. So the
 * list comes from `types.ts` now: add a decoration to the interface and it is
 * swept, with nothing to remember.
 *
 * The parse is asserted too. A regex that quietly matched nothing would turn
 * this whole file into a sweep of zero keys that passes in 4ms.
 */
function keysOf(iface: string): string[] {
  const src = readFileSync(new URL(`../src/core/types.ts`, import.meta.url), "utf8");
  const body = new RegExp(`export interface ${iface} \\{([\\s\\S]*?)\\n\\}`).exec(src)?.[1] ?? "";
  return [...body.matchAll(/^ {2}([a-zA-Z][a-zA-Z0-9]*)\??:/gm)].map((m) => m[1]);
}

const DECOR_KEYS = keysOf("Decorations");
const SERIES_KEYS = keysOf("Series");
const DATA_KEYS = keysOf("ChartData");

describe("the hostile sweep covers what the types declare", () => {
  it("reads the key lists out of types.ts", () => {
    // Counts, not contents: the point is that the parse worked and the
    // interfaces are being read, not that they have a particular shape.
    expect(DECOR_KEYS.length, "no decoration keys parsed — the sweep below would test nothing").toBeGreaterThan(30);
    expect(SERIES_KEYS.length, "no series keys parsed").toBeGreaterThan(6);
    expect(DATA_KEYS.length, "no data keys parsed").toBeGreaterThan(4);
    // Spot-checks for the three that were missing, so a parser that finds only
    // the first field of an interface cannot pass the counts above.
    for (const k of ["valueLines", "callouts", "bands", "fillBetween"]) expect(DECOR_KEYS).toContain(k);
  });
});

describe("a decoration of the wrong type cannot take the renderer down", () => {
  for (const key of DECOR_KEYS) {
    it(key, () => {
      const bad: string[] = [];
      for (const { kind } of CHART_KINDS) {
        for (const [label, value] of WRONG_TYPES) {
          const base = sampleConfig(kind);
          const cfg = { ...base, decorations: { ...base.decorations, [key]: value } } as unknown as ChartConfig;
          try {
            sceneToSvg(buildChart(cfg));
          } catch (e) {
            bad.push(`${kind}/${key}=${label}: ${e instanceof Error ? e.message : String(e)}`);
          }
        }
      }
      expect(bad.slice(0, 4)).toEqual([]);
    });
  }
});

/**
 * The same question of a SERIES field and a DATA field, which nothing asked.
 *
 * `Series.name` is `string` in the type and five places in `chart.ts` act on it
 * unguarded — `CARRIED_ROW.test(s.name.trim())` and its siblings. A series
 * written as `{ values: [1, 2, 3] }`, which is the obvious way to write a
 * single-series chart, therefore crashed seventeen of the twenty-five kinds.
 * That is not a hostile input at all; it is the shape a person writes by hand.
 *
 * And `data.xExtent: NaN` survived `normalizeData` untouched — the pad guard
 * was `arr ? … : arr`, which hands a FALSY non-array straight back — so the
 * mekko layout died on `data.xExtent?.some is not a function`, where the
 * optional chain does not help because NaN is not nullish.
 */
describe("a series field of the wrong type cannot take the renderer down", () => {
  for (const key of SERIES_KEYS) {
    it(key, () => {
      const bad: string[] = [];
      for (const { kind } of CHART_KINDS) {
        for (const [label, value] of WRONG_TYPES) {
          const base = sampleConfig(kind);
          const series = base.data.series.map((s) => ({ ...s, [key]: value }));
          const cfg = { ...base, data: { ...base.data, series } } as unknown as ChartConfig;
          try {
            sceneToSvg(buildChart(cfg));
          } catch (e) {
            bad.push(`${kind}/${key}=${label}: ${e instanceof Error ? e.message : String(e)}`);
          }
        }
      }
      expect(bad.slice(0, 4)).toEqual([]);
    });
  }

  it("a series with no name at all renders every kind", () => {
    // Stated on its own, because it is the ordinary case rather than an
    // adversarial one and it deserves to fail by name.
    const bad: string[] = [];
    for (const { kind } of CHART_KINDS) {
      const base = sampleConfig(kind);
      const series = base.data.series.map((s) => {
        const o = { ...s } as Partial<typeof s>;
        delete o.name;
        return o;
      });
      const cfg = { ...base, data: { ...base.data, series } } as unknown as ChartConfig;
      try {
        sceneToSvg(buildChart(cfg));
      } catch (e) {
        bad.push(`${kind}: ${e instanceof Error ? e.message : String(e)}`);
      }
    }
    expect(bad.slice(0, 4)).toEqual([]);
  });
});

describe("a data field of the wrong type cannot take the renderer down", () => {
  for (const key of DATA_KEYS) {
    it(key, () => {
      const bad: string[] = [];
      for (const { kind } of CHART_KINDS) {
        for (const [label, value] of WRONG_TYPES) {
          const base = sampleConfig(kind);
          const cfg = { ...base, data: { ...base.data, [key]: value } } as unknown as ChartConfig;
          try {
            sceneToSvg(buildChart(cfg));
          } catch (e) {
            bad.push(`${kind}/${key}=${label}: ${e instanceof Error ? e.message : String(e)}`);
          }
        }
      }
      expect(bad.slice(0, 4)).toEqual([]);
    });
  }
});

/**
 * Not a crash — a DISAPPEARANCE. The half the sweeps above cannot see.
 *
 * Every suite in this file asks the same question: does the wrong type take a
 * renderer DOWN. That question has an answer for `valueAxisTitle: 99` — no, it
 * does not — and the answer was wrong about the chart. `textWidth` returned
 * `NaN` for a number (`(99).length` is `undefined`), the axis-title node was
 * built as `w: Math.max(…, NaN)`, and `finiteNodes` — the safety net whose own
 * comment calls itself "a floor and not a filter" — dropped the node. The units
 * label vanished from all 25 kinds with no error anywhere, and a guard that
 * only watches for throws reported everything fine.
 *
 * So this asserts the content SURVIVES, not merely that nothing exploded. A
 * value the renderers will happily draw must reach them.
 */
describe("a config value the renderer can draw must not silently vanish", () => {
  // Asserted through the RENDERER, not off the node: the scene may carry the
  // number verbatim (all three sinks coerce their own text, which is the rule
  // this repo already settled on), so what matters is that the label reaches
  // the output at all. The bug was that the node never got there.
  const axisTitleText = (cfg: ChartConfig) =>
    buildChart(cfg)
      .nodes.filter((n) => n.name === "value-axis-title")
      .map((n) => String((n as { text?: unknown }).text));

  it("keeps a numeric axis title, exactly as it keeps the string form", () => {
    const base = sampleConfig("stacked");
    expect(axisTitleText({ ...base, valueAxisTitle: 99 } as unknown as ChartConfig)).toEqual(["99"]);
    expect(axisTitleText({ ...base, valueAxisTitle: "99" } as ChartConfig)).toEqual(["99"]);
    // And it survives all the way into the SVG a reader actually looks at.
    expect(sceneToSvg(buildChart({ ...base, valueAxisTitle: 99 } as unknown as ChartConfig))).toContain(
      'data-name="value-axis-title"',
    );
  });

  it("measures numeric text as its digits, so no layout decision reads NaN", () => {
    // The root. Every fit-to-width test in the engine is a comparison, and each
    // one is FALSE against NaN — so shrink-to-fit stops shrinking and a width
    // built with Math.max becomes NaN. Sixty-odd call sites share this.
    expect(textWidth(2024 as unknown as string, 10)).toBe(textWidth("2024", 10));
  });

  it("draws every kind's numeric axis title, not just the one sampled above", () => {
    const missing: string[] = [];
    for (const { kind } of CHART_KINDS) {
      const cfg = { ...sampleConfig(kind), valueAxisTitle: 42 } as unknown as ChartConfig;
      const asString = { ...sampleConfig(kind), valueAxisTitle: "42" } as ChartConfig;
      // Only kinds that draw the label at all — the comparison is against the
      // string form, so a kind with no value axis is simply "both empty".
      if (axisTitleText(asString).length && !axisTitleText(cfg).length) missing.push(kind);
    }
    expect(missing).toEqual([]);
  });
});

/**
 * A node with no numbers passes a filter that checks numbers.
 *
 * `finiteNodes` is the engine's last gate before the scene reaches three
 * renderers, and its own comment says why: the PowerPoint renderers write a
 * file, so a bad coordinate produces a .pptx PowerPoint may refuse to open. It
 * asked whether every number in a node is finite — and an EMPTY point list
 * satisfies that trivially, because there are no numbers to fail.
 *
 * So a `polygon` carrying `points: []` went straight through the gate and broke
 * exactly the renderer the gate exists to protect. `pptx-paint.mjs` takes the
 * bounding box with `Math.min(...xs)`, which is `Infinity` for no points, and
 * writes `x="Infinity"` into the OOXML. Not an Int64; Microsoft's own validator
 * rejects the deck.
 *
 * Found by rendering 3033 hostile configs through the skill's headless
 * renderer and scanning the output for numeric attributes that are not numbers.
 * Four of them, on one slide, from `{kind: "radar", data: {}}`.
 */
describe("a polygon with nothing to draw", () => {
  it("is dropped rather than handed to a renderer that will divide by nothing", () => {
    const poly = (name: string, points: { x: number; y: number }[]): SceneNode => ({
      kind: "polygon",
      points,
      stroke: "#000000",
      name,
    });
    const empty = poly("grid-2", []);
    const one = poly("grid-3", [{ x: 1, y: 2 }]);
    const real = poly("series-0", [
      { x: 0, y: 0 },
      { x: 5, y: 5 },
    ]);
    const kept = finiteNodes([empty, one, real]);
    expect(
      kept.map((n) => n.name),
      "a degenerate polygon reached the renderers",
    ).toEqual(["series-0"]);
  });

  it("never leaves one in a real chart's scene", () => {
    // The producer: a radar lays out its grid rings before it knows it has no
    // axes to hang them on. Asserted across every kind, because any layout that
    // emits a ring, web or outline can do the same from empty data.
    const bad: string[] = [];
    for (const { kind } of CHART_KINDS) {
      for (const data of [{}, { series: [] }, { series: [], categories: [] }]) {
        const scene = buildChart({ ...sampleConfig(kind), data } as unknown as ChartConfig);
        for (const n of scene.nodes)
          if (n.kind === "polygon" && n.points.length < 2) bad.push(`${kind}: ${n.name ?? "(unnamed)"}`);
      }
    }
    expect([...new Set(bad)], "a chart shipped a polygon with no geometry").toEqual([]);
  });
});
