/**
 * Nine things that have never once run against a real PowerPoint, on one click —
 * plus one experiment that only makes sense on its own.
 *
 * The demo deck covers inserting onto slides it added BLANK. Nothing covers
 * what happens elsewhere — inserting on top of an earlier run, a slide
 * duplicated so two claim one slot, redrawing a chart the user is looking at,
 * editing the chart the user SELECTED, drawing onto a slide that already has
 * content, a deck-wide rescale, stopping part-way, turning a degraded picture
 * back into shapes, and whether any of it is visible. Every one of those paths
 * is guarded, and every guard has only ever been checked against a fake host.
 * The list sat in `docs/PUBLISHING.md` as separate things for a human to
 * remember to try, which in practice meant a session each: deploy, click
 * through it, save the deck, upload, read.
 *
 * The blank-slide bias is not incidental. `insertSceneIntoSlide` — the
 * everyday "put a chart on the slide I am looking at" — draws onto whatever
 * the user already has there, and no test anywhere touched that until
 * `insertOntoUsedSlide` below. That is exactly where the worst bug of the
 * session lived, and why a demo run could not have found it.
 *
 * Scenarios run in order and leave their slides in the deck, because the point
 * is a file someone can open, look at, and hand to `npm run triage`. The one
 * exception is `chartIsVisible`, which takes its control slide away again.
 *
 * **On selection.** This file used to say Office.js had no way to select a
 * shape, and that the pane's selection-driven entry points therefore could not
 * be scripted. It was wrong, in the most expensive way a comment can be: it
 * justified a hole with a fact, so nobody re-checked it for months.
 * `Slide.setSelectedShapes(shapeIds)` has been GA since **PowerPointApi 1.5** —
 * the same set this add-in already requires for `getSelectedShapes` and
 * `setSelectedSlides`. The mistaken belief was that it would live on
 * `Presentation` beside `setSelectedSlides`; it is one class down, on `Slide`.
 * `editViaSelection` is what that sentence had been preventing.
 *
 * **Order is load-bearing, because of two live web-host bugs.** On PowerPoint
 * on the web `setSelectedShapes([])` does not clear the selection
 * (office-js#3083), and a picture cannot be inserted while another shape is
 * selected (office-js#3698). So a scenario that leaves a chart selected breaks
 * the PICTURE scenario rather than its own — a failure that reads, in a run
 * log, as a bug in code that is fine. Every selecting scenario clears up after
 * itself through `clearShapeSelection`, which re-selects the slide rather than
 * trusting the empty-array clear.
 */
import type { ChartConfig, ChartKind } from "../core/types";
import type { Scene } from "../core/scene";
import { buildChart } from "../core/chart";
import { sampleConfig } from "../core/samples";
import { buildDeckBase64 } from "../render/pptx-deck";
import {
  addScratchSlide,
  canInsertPicture,
  canSelectShapes,
  clearShapeSelection,
  deleteSlideById,
  isStopped,
  isStopRequested,
  isTimeout,
  readbackTimeoutMs,
  loadChartFromSelection,
  requestStop,
  resetStop,
  selectShape,
  selectionLadder,
  slideImageBase64,
  canInsertSlidesFromBase64,
  insertSceneIntoSlide,
  insertSlidesFromPptx,
  listChartsInDeck,
  newRunId,
  reconcileDeck,
  showSlide,
  slideCount,
  slideSize,
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

/**
 * How long the selection scenario waits on a host that has stopped answering.
 *
 * Short on purpose. The default budget exists to break an infinite wait; this
 * one exists because we already KNOW what the wait ends in on the web, and
 * three minutes of a person's evening to re-learn it is not a test.
 *
 * Capped by the default rather than replacing it: ten seconds is shorter than
 * the ninety this normally sits under, and far longer than the milliseconds a
 * test shortens that to.
 */
const selectionBudgetMs = (): number => Math.min(10_000, readbackTimeoutMs());

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
    // At the destination's size, like every other deck build — the self-test
    // exists to exercise the real path, and a 16:9 file inserted into a 4:3
    // deck is exactly the mismatch it should be catching rather than creating.
    await slideSize(),
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

/**
 * Drawing charts onto a slide that already has content on it.
 *
 * The everyday action — "insert a chart on the slide I am looking at" — and
 * until now the only path with no real-host coverage at all. Every other
 * scenario here, and the whole demo deck, works on slides the run added BLANK.
 *
 * That gap is not theoretical. The worst bug of the session lived in exactly
 * it: `insertSceneIntoSlide` could not ask for a fresh proxy re-fetch, because
 * the re-fetch identified a chart's shapes as "the last N on the slide" — true
 * of a blank slide a run just added, false of the slide a user is working on.
 * So a chart over ten shapes lost its group AND its config tag on the web,
 * silently, on the single most-used action in the add-in. A demo run would not
 * have caught it: the demo path opted into the re-fetch and the everyday path
 * could not.
 *
 * Two charts, onto a slide that starts with one. The second lands on a slide
 * holding two already, which is where a positional heuristic goes wrong.
 */
const insertOntoUsedSlide: Scenario = async (prefix) => {
  const [host] = await probeCharts(prefix);
  if (!host) return { ok: false, skipped: true, detail: "no probe chart to insert alongside" };
  const shown = await showSlide(host.target.slideId);
  const before = await slideCount();
  const added = [`${prefix} onto A`, `${prefix} onto B`];
  for (const [n, title] of added.entries()) {
    const c = cfg(title);
    // Cascaded, the way a repeated insert does, so the second does not land
    // exactly on top of the first.
    await insertSceneIntoSlide(buildChart(c), {
      tagData: JSON.stringify(c),
      left: 40 + n * 24,
      top: 40 + n * 24,
    });
  }
  const after = await slideCount();
  const mine = await probeCharts(`${prefix} onto`);
  // The chart that was already there must still be a chart. A run that swept
  // the slide's existing shapes into its own group would take this with it —
  // and delete it on the next edit, because the parts tag would claim it.
  const survivor = (await probeCharts(prefix)).find((c) => c.cfg.title === host.cfg.title);
  const problems = [
    after !== before && `the deck grew by ${after - before} — the charts did not go ON the slide`,
    mine.length !== added.length && `${mine.length} of ${added.length} new charts are re-editable`,
    !survivor && "the chart already on the slide is no longer re-editable",
  ].filter(Boolean);
  return {
    ok: problems.length === 0,
    detail: problems.length
      ? problems.join("; ")
      : `${added.length} charts drawn ${shown ? "onto the visible slide" : "(host would not move the view)"}, ` +
        `all re-editable, the one already there untouched`,
  };
};

/**
 * The pane's actual entry point: a chart the USER has selected.
 *
 * Everything above reaches the machinery through targets `listChartsInDeck`
 * hands over. That is not how anyone uses the add-in. A user clicks a chart and
 * presses "Edit it", and the pane answers from `loadChartFromSelection` — a
 * different read, on a different collection (`getSelectedShapes`), which no
 * scenario here has ever exercised because of a comment that said it could not
 * be done. It can, at PowerPointApi 1.5, and this is the whole round trip:
 * select the shape, read it back as the pane does, edit through THAT target,
 * and confirm what comes back is the same chart.
 *
 * Deselects at the end by re-selecting the slide, never by
 * `setSelectedShapes([])` — see `clearShapeSelection`. A chart left selected
 * breaks the picture scenario below rather than this one.
 */
const editViaSelection: Scenario = async (prefix) => {
  if (!canSelectShapes()) return { ok: false, skipped: true, detail: "host cannot select shapes (PowerPointApi 1.5)" };
  const [chart] = await probeCharts(prefix);
  if (!chart) return { ok: false, skipped: true, detail: "no probe chart in the deck to select" };
  if (!(await selectShape(chart.target.slideId, chart.target.shapeId, selectionBudgetMs()))) {
    return { ok: false, skipped: true, detail: "the host would not select the chart" };
  }
  try {
    // Exactly what the pane does when someone presses "Edit it".
    //
    // On a short budget, because on PowerPoint on the WEB this call does not
    // come back at all once a shape has been selected programmatically —
    // measured, twice: 90 seconds to the bound, then the very next
    // `setSelectedSlides` also silent. At the full budget this scenario costs
    // three minutes to learn a thing we already know, so it is given ten
    // seconds and a name for what it found.
    const picked = await loadChartFromSelection(selectionBudgetMs());
    if (!picked) return { ok: false, detail: "the selected chart did not read back as a PowerChart" };
    const was = JSON.parse(picked.configJson) as ChartConfig;
    if (was.title !== chart.cfg.title) {
      return { ok: false, detail: `the selection read back a different chart: "${was.title}"` };
    }
    // And edit through the target the SELECTION produced, not the one the deck
    // scan produced. The two are built by different code from different reads,
    // and only one of them is what a user's edit actually travels on.
    const next = { ...was, title: `${was.title} (via selection)` };
    const target = await updateChartInSlide(buildChart(next), picked.target, { tagData: JSON.stringify(next) });
    if (!target) return { ok: false, detail: "the chart was gone from the slide after editing the selection" };
    const round = (await probeCharts(prefix)).find((c) => c.cfg.title === next.title);
    return {
      ok: !!round,
      detail: round
        ? "selected, read back, edited through the selection's own target, still re-editable"
        : "edited through the selection, and the result is no longer re-editable",
    };
  } catch (err) {
    // Silence is a fact about the HOST, not a defect of ours, and the two must
    // not read alike in a report. `setSelectedShapes` is GA at PowerPointApi
    // 1.5 and takes the call, but on the web it leaves the selection subsystem
    // unable to answer anything afterwards — `getSelectedShapes` and
    // `setSelectedSlides` both went silent for the full budget in the same
    // run. Two selection bugs are already open against the web host
    // (office-js#3083, #3698); this is the same family and worse.
    //
    // Reported SKIPPED rather than FAILED, with the reason: nothing in the
    // add-in is broken by it, because nothing but this battery ever selects a
    // shape programmatically. A red line here would send the next diagnosis
    // after our own code.
    if (isTimeout(err)) {
      return {
        ok: false,
        skipped: true,
        detail:
          "the host stopped answering selection calls after a programmatic select — known web-host limitation, " +
          "same family as office-js#3083 / #3698; the pane's own Edit-it path is unaffected",
      };
    }
    throw err;
  } finally {
    // Also on the short budget: once the subsystem is wedged this call is
    // silent too, and waiting the full budget for a deselect we already know
    // will not answer is ninety seconds spent on nothing.
    await clearShapeSelection(chart.target.slideId, selectionBudgetMs());
  }
};

/**
 * Stopping a run part-way, and what the deck looks like afterwards.
 *
 * Stop is cooperative — Office.js has no abort, so a sync already handed to
 * PowerPoint runs to completion whatever anyone wants, and the only honest
 * promise is "nothing further after this batch". That makes the interesting
 * question not "did it stop" but "what did it leave": a stop is indistinguishable
 * from a stall to every layer above it, and both leave a partial chart on the
 * slide. This asserts the two things a user is owed — the call ends, and it
 * ends by SAYING it stopped rather than reporting a chart it did not finish.
 *
 * The stop is requested before the render rather than mid-flight, which is the
 * one part a script cannot time: the pane's Stop button is pressed by a human
 * some seconds in. What is exercised is the same cooperative path, taken at its
 * first batch boundary instead of its fifth.
 */
const stopPartWay: Scenario = async (prefix) => {
  // If the USER has already asked to stop, this scenario must not run at all —
  // and above all must not clear their flag. It is the only code in the add-in
  // that calls `resetStop()`, and doing that on the way out of a scenario the
  // user had already cancelled would resume a battery they had stopped. The
  // guard in `runSelfTest` means this branch should be unreachable; it is here
  // because "should be unreachable" is how the flag got clobbered in the first
  // place.
  if (isStopRequested()) return { ok: false, skipped: true, detail: "not reached — the run was stopped" };
  const before = await slideCount();
  const c = cfg(`${prefix} stopped`);
  requestStop();
  let outcome: string;
  try {
    const target = await insertSceneIntoSlide(buildChart(c), { tagData: JSON.stringify(c) });
    outcome = target ? "the insert ran to completion and reported a chart" : "the insert finished without a chart";
  } catch (err) {
    outcome = isStopped(err) ? "stopped" : `threw something other than a stop: ${errorText(err)}`;
  } finally {
    resetStop();
  }
  const after = await slideCount();
  const problems = [
    outcome !== "stopped" && outcome,
    after !== before && `the deck grew by ${after - before} — a stopped insert added a slide`,
    // A stop must not leave a chart the pane would offer to edit: half a chart
    // that claims to be whole is worse than a visible mess.
    (await probeCharts(`${prefix} stopped`)).length > 0 && "a stopped insert left a re-editable chart behind",
  ].filter(Boolean);
  return {
    ok: problems.length === 0,
    detail: problems.length
      ? problems.join("; ")
      : "stopped at a batch boundary, nothing added, nothing left claiming to be a chart",
  };
};

/**
 * Whether anything is actually VISIBLE.
 *
 * Every other assertion in this file counts shapes and reads tags. All of them
 * pass for a chart drawn entirely in white, or at zero size, or off the edge of
 * the slide — a chart that is structurally perfect and invisible. Nothing in
 * this project has ever checked otherwise except a human looking at a deck.
 *
 * `Slide.getImageAsBase64` (PowerPointApi 1.8) is the host's own rasteriser, so
 * this compares a slide holding a chart against the same slide before it had
 * one. Deliberately not a pixel comparison against the SVG renderer: two
 * different text shapers never agree, and a check that cries wolf gets ignored.
 * The question here is the crude one nobody was asking — did drawing the chart
 * change what the slide looks like at all.
 */
const chartIsVisible: Scenario = async (prefix) => {
  // Its own slide, taken away afterwards. Before-and-after on ONE slide is the
  // only comparison that isolates the chart: two different slides differ for a
  // dozen reasons a rasteriser can see and this scenario should not care about.
  const slideId = await addScratchSlide();
  if (!slideId) return { ok: false, skipped: true, detail: "the host would not add a slide to draw on" };

  const measure = async (): Promise<{ ok: boolean; detail: string; skipped?: boolean }> => {
    const blank = await slideImageBase64(slideId, 640);
    if (!blank) return { ok: false, skipped: true, detail: "host will not rasterise a slide (PowerPointApi 1.8)" };
    const c = cfg(`${prefix} visible`);
    const drawn = await insertSceneIntoSlide(buildChart(c), { slideId, tagData: JSON.stringify(c) });
    if (!drawn) return { ok: false, detail: "nothing was drawn, so there is nothing to look at" };
    const withChart = await slideImageBase64(slideId, 640);
    if (!withChart) return { ok: false, detail: "the host rasterised the empty slide but not the one with a chart" };
    return {
      ok: withChart !== blank,
      detail:
        withChart !== blank
          ? `drawing the chart changed what the slide looks like (${blank.length} → ${withChart.length} bytes)`
          : "the slide renders identically with and without the chart — nothing is visible",
    };
  };

  // The verdict is computed first, then the slide is given back, and only then
  // are the two combined. Not a `finally` that appends to the detail — a
  // `return` inside `try` evaluates its expression BEFORE the `finally` runs,
  // so the warning would be stitched onto a string that had already left. That
  // is what the first version of this did, and the test written to prove the
  // warning appears is the only reason it did not ship that way.
  let verdict: { ok: boolean; detail: string; skipped?: boolean };
  let removed: boolean;
  try {
    verdict = await measure();
  } finally {
    // Unlike every other scenario, this one cleans up: its slide is a control
    // surface, not a result anyone would want to open. `deleteSlideById` now
    // verifies from a fresh read rather than assuming, so false here means the
    // slide is genuinely still in the deck — carrying a config tag, and so a
    // chart the pane would offer to edit.
    removed = await deleteSlideById(slideId);
    if (!removed) trace("selftest", "could not remove the visibility scenario's scratch slide", { slideId });
  }
  if (removed) return verdict;
  return {
    ...verdict,
    detail: `${verdict.detail} — WARNING: the scratch slide it drew on could not be removed, and is still in the deck`,
  };
};

/**
 * Order: the scenarios with a track record first, the new ones last.
 *
 * A battery is only worth its whole run if it survives to the end, and this one
 * has twice now failed to — a wedge at 1819s and a host killed outright at
 * 108s. When that happens every scenario AFTER the failure reports nothing, so
 * where an unproven scenario sits in this list decides how much of the run its
 * failure costs.
 *
 * The three added most recently — selection, stop, visibility — therefore run
 * at the end, and the heaviest and least proven of them (`chartIsVisible`,
 * which asks the host to rasterise a slide twice) runs dead last. It was 6th,
 * which meant a crash there cost the verdicts of three scenarios that had
 * worked on a real host that morning. New code must not be able to take the
 * evidence for old code with it.
 *
 * The first two remain first because everything else needs the probe charts
 * they insert. Beyond that the constraint is `explodePicture`: it inserts a
 * PICTURE, and on the web a picture cannot be inserted while a shape is
 * selected (office-js#3698) — so every scenario that selects must clear up
 * after itself, which `clearShapeSelection` does, wherever it runs.
 */
/**
 * Which selection call wedges the host — asked once, properly.
 *
 * Two accounts, different culprits. This project measured
 * `setSelectedShapes([id])` being taken and every selection call after it going
 * silent; office-js#3698 says it is `setSelectedShapes([])` whose promise never
 * resolves. Both are plausible, neither is verified, and the gate above was
 * written against my reading of one 159-second run.
 *
 * `selectionLadder` climbs from the least invasive call to the most and stops
 * at the first silence, because after a wedge every call is silent and a report
 * of four timeouts names nothing. What comes back is one sentence: the last
 * call the host answered, and the first one it did not.
 *
 * Reported `ok` whenever the ladder RAN, whatever it found. It is an
 * experiment, not an assertion — a host that answers every rung is a genuine
 * result (the wedge is gone, or is not on this host), and marking that as a
 * failure would be the same mistake as marking the wedge itself as one.
 */
const whichSelectionCallWedges: Scenario = async (prefix) => {
  if (!canSelectShapes()) return { ok: false, skipped: true, detail: "host cannot select shapes (PowerPointApi 1.5)" };
  const [chart] = await probeCharts(prefix);
  if (!chart) return { ok: false, skipped: true, detail: "no probe chart in the deck to select" };
  const rungs = await selectionLadder(chart.target.slideId, chart.target.shapeId, selectionBudgetMs());
  for (const r of rungs) trace("selftest", "ladder rung", { step: r.step, outcome: r.outcome, ms: r.ms });
  const stopped = rungs.find((r) => r.outcome === "silent");
  const answered = rungs.filter((r) => r.outcome === "ok").length;
  if (!stopped) {
    const refused = rungs.filter((r) => r.outcome === "refused").map((r) => r.step);
    return {
      ok: true,
      detail:
        `the host answered all ${rungs.length} rung(s) — nothing wedged` +
        (refused.length ? `; refused (but kept talking): ${refused.join(", ")}` : ""),
    };
  }
  // The rung before the silence is the last thing the host was willing to do,
  // and naming both is the whole point: "silent at X, after Y" is an answer,
  // "something in the selection API hangs" is what we already had.
  const before = rungs[rungs.indexOf(stopped) - 1];
  return {
    ok: true,
    detail:
      `SILENT at "${stopped.step}" after ${stopped.ms}ms` +
      (before ? `, the rung before it ("${before.step}") answered in ${before.ms}ms` : ", the very first rung") +
      ` — ${answered} rung(s) answered before it`,
  };
};

const SCENARIOS: {
  name: string;
  run: Scenario;
  /**
   * Left out of a full run, offered by the picker.
   *
   * For a scenario whose result is only meaningful on a host nothing else has
   * touched yet. Running it alongside the rest does not merely cost time — it
   * produces a WRONG answer, which is worse than no answer.
   */
  pickedOnly?: boolean;
}[] = [
  { name: "insert on top of an earlier run", run: insertTwice },
  { name: "two slides claiming one slot", run: duplicateSlot },
  { name: "edit a chart on the visible slide", run: editOnVisibleSlide },
  { name: "insert onto a slide that already has content", run: insertOntoUsedSlide },
  { name: "same scale across the deck", run: sameScaleAcrossDeck },
  { name: "explode a degraded picture", run: explodePicture },
  { name: "edit the chart the user selected", run: editViaSelection },
  { name: "stop a run part-way", run: stopPartWay },
  { name: "the chart is actually visible", run: chartIsVisible },
  // PICKED ONLY, and the reason is the experiment's own subject matter.
  //
  // The first draft put this last, reasoning that a scenario which provokes a
  // wedge should not damage the ones after it. That was backwards, and a test
  // caught it before a real host did: `editViaSelection` provokes the SAME
  // wedge six scenarios earlier, so by the time the ladder ran the host was
  // already wedged and it reported silence on its own first rung — an answer
  // about nothing.
  //
  // Both orderings are wrong, because the constraint is not "last" or "first":
  // it is that the ladder has to be the ONLY thing that has touched the
  // selection subsystem. That is not a position in a list, it is a run of its
  // own. Picked from the Scenario menu it gets exactly that — the two probe
  // inserts, then a clean host — and the routine battery neither slows down nor
  // gets its results contaminated.
  { name: "which selection call wedges the host", run: whichSelectionCallWedges, pickedOnly: true },
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

/**
 * Every scenario's name, in run order — for a picker that cannot drift.
 *
 * Derived from `SCENARIOS` rather than repeated beside it, because a list of
 * strings maintained by hand next to the list it describes is a rename away
 * from offering a scenario that does not exist.
 */
/**
 * Everything the Scenario picker offers — including what a full run leaves out.
 *
 * The picker's whole purpose is reaching one scenario without the ones in front
 * of it, so a `pickedOnly` scenario has to be here. It is the ONLY way to reach
 * one.
 */
export const SCENARIO_NAMES: readonly string[] = SCENARIOS.map((s) => s.name);

/** What a full run actually runs. A strict subset — see `pickedOnly`. */
export const ROUTINE_SCENARIO_NAMES: readonly string[] = SCENARIOS.filter((s) => !s.pickedOnly).map((s) => s.name);

/**
 * Run the battery, or ONE scenario of it.
 *
 * `only` exists because of what a full round costs. The first run that produced
 * a usable trace took 496 seconds to reach scenario seven and wedged there, and
 * the scenario before it left four charts on the deck as loose piles of shapes.
 * Iterating on the seventh scenario by running the six in front of it is eight
 * minutes and real damage per attempt, which is not iteration.
 *
 * The first two are inserted anyway whatever is picked: every other scenario
 * needs the probe charts they create, and one that finds none reports SKIPPED —
 * an honest answer, and a useless round. So "only" means "only this, plus what
 * it needs", and the report says which is which.
 */
export async function runSelfTest(prefix = `selftest ${newRunId()}`, only?: string): Promise<ScenarioResult[]> {
  const wanted = only ? SCENARIOS.filter((s, i) => i < 2 || s.name === only) : SCENARIOS.filter((s) => !s.pickedOnly);
  const out: ScenarioResult[] = [];
  for (const { name, run } of wanted) {
    // The battery had no stop check of its own — none at all. So even where a
    // scenario ended promptly, Stop could not end the RUN: the pane switched
    // its button to "Stopping…" and the next scenario started anyway. A
    // battery is the longest thing in the Testing panel and the likeliest to
    // be abandoned half way, which makes this the one loop that most needed it.
    //
    // Scenarios not reached are reported as skipped, with the reason, rather
    // than dropped: a report missing its last four lines looks like a battery
    // that crashed, and that is a different diagnosis from one that was
    // stopped.
    if (isStopRequested()) {
      out.push({ name, ok: false, skipped: true, detail: "not reached — the run was stopped", ms: 0 });
      continue;
    }
    // Announced BEFORE it runs, not only after it finishes.
    //
    // Every verdict this battery emits is a past-tense record, so a run that
    // dies mid-scenario leaves the PREVIOUS scenario's line as its last word —
    // off by one, and pointing at the one thing that demonstrably worked. Twice
    // now a real-host failure has been diagnosed from a screenshot rather than
    // a log, because the log only exists once the run ends and these runs did
    // not. This line is what makes such a screenshot name the right scenario.
    trace("selftest", "scenario starting", { name });
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
