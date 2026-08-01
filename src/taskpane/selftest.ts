/**
 * Five things that have never once run against a real PowerPoint.
 *
 * The demo deck covers inserting; nothing covers what happens AFTERWARDS —
 * inserting on top of an earlier run, a slide duplicated so two claim one slot,
 * redrawing a chart the user is looking at, a deck-wide rescale, turning a
 * degraded picture back into shapes. Every one of those paths is guarded, and
 * every guard has only ever been checked against a fake host. The list has sat
 * in `docs/PUBLISHING.md` as five separate things for a human to remember to
 * try, which in practice means five separate sessions: deploy, click through
 * it, save the deck, upload, read.
 *
 * So: one click, five scenarios, each recorded with what it actually observed.
 * They run in order and they leave their slides in the deck, because the point
 * is a file someone can open, look at, and hand to `npm run triage`.
 *
 * **What this deliberately does NOT test.** Office.js has no way to select a
 * SHAPE, so the pane's selection-driven entry points — "Edit selected chart",
 * "Explode" as the user reaches them — cannot be scripted. What is exercised
 * here is the machinery underneath them, reached by the same targets
 * `listChartsInDeck` hands the pane. A scenario that passes here can still be
 * broken at the selection layer; a scenario that fails here is broken for
 * everyone.
 */
import type { ChartConfig, ChartKind } from "../core/types";
import type { Scene } from "../core/scene";
import { buildChart } from "../core/chart";
import { sampleConfig } from "../core/samples";
import { buildDeckBase64 } from "../render/pptx-deck";
import {
  canInsertPicture,
  canInsertSlidesFromBase64,
  insertSlidesFromPptx,
  listChartsInDeck,
  newRunId,
  reconcileDeck,
  showSlide,
  slideCount,
  slideHoldsOnlyChart,
  updateChartInSlide,
  errorText,
} from "../render/powerpoint";
import { trace } from "../core/trace";

/** One scenario's verdict, as it goes into the run log. */
export interface ScenarioResult {
  name: string;
  /** False means the scenario ran and the host got it wrong. */
  ok: boolean;
  /**
   * True when the scenario could not run at all — the host lacks the API, or
   * a previous scenario left nothing to work with. Kept apart from a failure
   * on purpose: "we did not check" and "we checked and it is broken" are
   * different answers, and reporting the first as the second is how a
   * requirement-set gap gets diagnosed as a bug.
   */
  skipped?: boolean;
  /** What was actually observed — the sentence a diagnosis starts from. */
  detail: string;
  ms: number;
}

const cfg = (title: string, kind: ChartKind = "clustered"): ChartConfig => ({
  ...sampleConfig(kind),
  title,
});

/** A small deck of tagged charts, built once and inserted however a scenario needs. */
async function buildProbe(
  run: string,
  titles: string[],
): Promise<{ base64: string; scenes: Scene[]; shapesPerSlide: number[] }> {
  const configs = titles.map((t) => cfg(t));
  const scenes = configs.map(buildChart);
  const built = await buildDeckBase64(
    configs.map((c, i) => ({
      scene: scenes[i],
      title: titles[i],
      configJson: JSON.stringify(c),
      slot: i,
      run,
    })),
  );
  return { base64: built.base64, scenes, shapesPerSlide: built.shapesPerSlide };
}

/** Charts this self-test put in the deck, by the title it gave them. */
async function probeCharts(prefix: string) {
  const all = await listChartsInDeck();
  return all
    .map((c) => ({ ...c, cfg: JSON.parse(c.configJson) as ChartConfig }))
    .filter((c) => typeof c.cfg.title === "string" && c.cfg.title.startsWith(prefix));
}

type Scenario = (prefix: string) => Promise<{ ok: boolean; detail: string; skipped?: boolean }>;

/**
 * Inserting a demo deck on top of an earlier one.
 *
 * Before the run token existed this was destructive: every item matched two
 * slides, and the repair pass deleted a healthy run as the other one's
 * duplicate. The token is what makes it safe, and nothing has ever confirmed
 * that against a real host.
 */
const insertTwice: Scenario = async (prefix) => {
  if (!canInsertSlidesFromBase64()) return { ok: false, skipped: true, detail: "host has no insertSlidesFromBase64" };
  const titles = [`${prefix} twice A`, `${prefix} twice B`];
  const before = await slideCount();
  for (const _ of [0, 1]) {
    const { base64 } = await buildProbe(newRunId(), titles);
    await insertSlidesFromPptx(base64, titles.length);
  }
  const after = await slideCount();
  const found = await probeCharts(`${prefix} twice`);
  const ok = after - before === 4 && found.length === 4;
  return {
    ok,
    detail: `deck grew by ${after - before} (want 4); ${found.length} of 4 charts re-editable`,
  };
};

/**
 * Two slides claiming one (run, slot) — what duplicating a demo slide produces.
 *
 * The same bytes inserted twice carry the same token, which is exactly what
 * PowerPoint's own Duplicate Slide does to a slot tag. The repair pass is
 * supposed to drop one and keep a working one; getting that backwards deletes
 * the user's chart.
 */
const duplicateSlot: Scenario = async (prefix) => {
  if (!canInsertSlidesFromBase64()) return { ok: false, skipped: true, detail: "host has no insertSlidesFromBase64" };
  const titles = [`${prefix} dup A`, `${prefix} dup B`];
  const run = newRunId();
  const { base64, shapesPerSlide } = await buildProbe(run, titles);
  const before = await slideCount();
  await insertSlidesFromPptx(base64, titles.length);
  await insertSlidesFromPptx(base64, titles.length);
  const afterInsert = await slideCount();
  const outcome = await reconcileDeck(
    // The count the GENERATOR reported, not an estimate of what Office.js
    // would have drawn: the two disagree by design (a pie is one custGeom
    // wedge in a file and a sixteen-triangle fan through Office.js), and
    // estimating here measured five perfect charts as wreckage on a real run.
    titles.map((t, i) => ({ slot: i, title: t, shapes: shapesPerSlide[i], chart: true, wroteTag: true })),
    { before, after: afterInsert },
    () => undefined,
    { run },
  );
  const settled = await slideCount();
  const kept = await probeCharts(`${prefix} dup`);
  // Four slides in, two out, and both survivors still charts. A pass that
  // deleted both copies of one item would leave two slides and read as
  // success on the count alone — hence the second half of this.
  const ok = settled - before === 2 && kept.length === 2;
  return {
    ok,
    detail: `4 slides inserted, ${settled - before} kept, ${kept.length} of 2 still re-editable; ${outcome.plan.summary.duplicates} queued as duplicates`,
  };
};

/**
 * Redrawing a chart the user is looking at.
 *
 * The add-in's worst case: every shape replaced, on the one slide guaranteed
 * to be on screen. A real run died on its first batch here. The scenario puts
 * the view on the slide FIRST — redrawing off-screen would be testing the easy
 * path, which is the one that already works.
 */
const editOnVisibleSlide: Scenario = async (prefix) => {
  const [chart] = await probeCharts(prefix);
  if (!chart) return { ok: false, skipped: true, detail: "no probe chart in the deck to edit" };
  const shown = await showSlide(chart.target.slideId);
  const next = { ...chart.cfg, title: `${chart.cfg.title} (edited)` };
  const target = await updateChartInSlide(buildChart(next), chart.target, { tagData: JSON.stringify(next) });
  if (!target) return { ok: false, detail: "the chart was gone from the slide after the update" };
  const again = await probeCharts(prefix);
  const round = again.find((c) => c.cfg.title === next.title);
  return {
    ok: !!round,
    detail: round
      ? `redrawn ${shown ? "on the visible slide" : "(host would not move the view)"} and the config round-tripped`
      : "the edited chart is no longer re-editable — its config did not survive the redraw",
  };
};

/**
 * A deck-wide rescale, with a chart on screen.
 *
 * Every chart's shapes are deleted and redrawn in one request context. A chart
 * whose redraw stalls is left blank, so what matters is not only that it
 * finishes but that nothing is emptied on the way.
 */
const sameScaleAcrossDeck: Scenario = async (prefix) => {
  const charts = await probeCharts(prefix);
  if (charts.length < 2) return { ok: false, skipped: true, detail: "fewer than two probe charts to scale together" };
  await showSlide(charts[0].target.slideId);
  // Every finite value across every probe chart. `Math.max()` of nothing is
  // -Infinity, and JSON.stringify turns that into `null` — so an empty or
  // all-blank data set wrote `{"scale":{"max":null}}` into the tag and then
  // compared it against -Infinity, which never matches. The scenario reported
  // a failure that was its own arithmetic.
  const values = charts
    .flatMap((c) => c.cfg.data?.series ?? [])
    .flatMap((s) => s.values)
    .map(Number)
    .filter((v) => Number.isFinite(v));
  if (!values.length) return { ok: false, skipped: true, detail: "probe charts carry no numbers to scale" };
  const max = Math.max(...values);
  for (const c of charts) {
    const next: ChartConfig = { ...c.cfg, scale: { max } };
    await updateChartInSlide(buildChart(next), c.target, { tagData: JSON.stringify(next) });
  }
  const after = await probeCharts(prefix);
  const scaled = after.filter((c) => c.cfg.scale?.max === max).length;
  return {
    ok: scaled === charts.length && after.length === charts.length,
    detail: `${scaled} of ${charts.length} charts carry the shared scale; ${after.length} still re-editable`,
  };
};

/**
 * Turning a degraded picture back into native shapes.
 *
 * The picture path is what the web host falls back to, and Explode is the only
 * way out of it. The claim it rests on — that a picture keeps its config and
 * can become shapes again — has never been checked anywhere but a fake.
 */
const explodePicture: Scenario = async (prefix) => {
  if (!canInsertPicture()) return { ok: false, skipped: true, detail: "host has no picture fill (PowerPointApi 1.8)" };
  if (!rasterizer) return { ok: false, skipped: true, detail: "no rasteriser — cannot make a picture to explode" };
  const [chart] = await probeCharts(prefix);
  if (!chart) return { ok: false, skipped: true, detail: "no probe chart in the deck to explode" };
  // Collapse it to a picture first, the way a struggling host would, then ask
  // for it back as shapes. Both halves matter: a config lost on the way IN is
  // indistinguishable from one lost on the way OUT if only one is exercised.
  //
  // The picture comes from a real rasterisation. `render: "image"` on the
  // config does NOT make one — the renderer takes the picture path only when
  // handed `pictureBase64` — so an earlier version of this passed `undefined`
  // and quietly performed two ordinary shape updates while reporting that the
  // picture round-trip worked. A scenario that cannot fail is worse than one
  // that is missing: it reads as evidence.
  const asPicture: ChartConfig = { ...chart.cfg, render: "image" };
  const png = await rasterizer(buildChart(asPicture));
  if (!png) return { ok: false, skipped: true, detail: "the browser would not rasterise the chart" };
  const pictured = await updateChartInSlide(buildChart(asPicture), chart.target, {
    tagData: JSON.stringify(asPicture),
    pictureBase64: png,
  });
  if (!pictured) return { ok: false, detail: "the chart vanished while being collapsed to a picture" };
  // One shape is what a picture IS. More than one means the renderer drew
  // shapes instead and the rest of this scenario would prove nothing.
  const asOne = await slideHoldsOnlyChart(pictured.slideId);
  if (!asOne) trace("selftest", "picture may not be a single shape", { slide: pictured.slideId });
  const asShapes: ChartConfig = { ...chart.cfg, render: "shapes" };
  const exploded = await updateChartInSlide(buildChart(asShapes), pictured, { tagData: JSON.stringify(asShapes) });
  if (!exploded) return { ok: false, detail: "the picture vanished while being exploded back to shapes" };
  const back = (await probeCharts(prefix)).find((c) => c.target.shapeId === exploded.shapeId);
  return {
    ok: !!back,
    detail: back
      ? "collapsed to a picture and exploded back, config intact"
      : "exploded, but the config did not survive",
  };
};

const SCENARIOS: { name: string; run: Scenario }[] = [
  { name: "insert on top of an earlier run", run: insertTwice },
  { name: "two slides claiming one slot", run: duplicateSlot },
  { name: "edit a chart on the visible slide", run: editOnVisibleSlide },
  { name: "same scale across the deck", run: sameScaleAcrossDeck },
  { name: "explode a degraded picture", run: explodePicture },
];

/**
 * Run every scenario, in order, and report all of them.
 *
 * A scenario that throws is a FAILED scenario, not a failed run: the four
 * after it are exactly the ones nobody has data for, and stopping at the first
 * error would spend a real-host session to learn one thing. Each verdict is
 * traced as it lands, so a run that takes the tab down still leaves a record
 * of how far it got.
 */
/**
 * How the battery gets a picture, when it needs one.
 *
 * Rasterising needs a canvas, which lives in the pane — the same split
 * `insertDemoDeck` already uses for its degrade-to-picture path. Absent, the
 * picture scenario reports SKIPPED rather than pretending.
 */
let rasterizer: ((scene: Scene) => Promise<string | undefined>) | undefined;

export function setSelfTestRasterizer(fn: (scene: Scene) => Promise<string | undefined>): void {
  rasterizer = fn;
}

export async function runSelfTest(prefix = `selftest ${newRunId()}`): Promise<ScenarioResult[]> {
  const out: ScenarioResult[] = [];
  for (const { name, run } of SCENARIOS) {
    const t0 = Date.now();
    let result: ScenarioResult;
    try {
      const r = await run(prefix);
      result = { name, ...r, ms: Date.now() - t0 };
    } catch (err) {
      result = { name, ok: false, detail: `threw: ${errorText(err)}`, ms: Date.now() - t0 };
    }
    trace("selftest", result.skipped ? "scenario skipped" : result.ok ? "scenario passed" : "scenario FAILED", {
      name,
      detail: result.detail,
      ms: result.ms,
    });
    out.push(result);
  }
  return out;
}

/** The one-line summary the pane shows. */
export function describeSelfTest(results: ScenarioResult[]): string {
  const ran = results.filter((r) => !r.skipped);
  const failed = ran.filter((r) => !r.ok);
  const skipped = results.length - ran.length;
  const parts = [`${ran.length - failed.length} of ${ran.length} scenarios passed`];
  if (failed.length) parts.push(`failed: ${failed.map((f) => f.name).join(", ")}`);
  if (skipped) parts.push(`${skipped} skipped (host cannot run them)`);
  return `Self-test — ${parts.join(" · ")}.`;
}
