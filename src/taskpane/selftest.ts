/**
 * Eleven things that have never once run against a real PowerPoint, on one
 * click — plus two that only make sense on their own.
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
  isStopped,
  isStopRequested,
  isTimeout,
  stepOf,
  readbackTimeoutMs,
  loadChartFromSelection,
  requestStop,
  resetStop,
  selectShape,
  selectionLadder,
  awaitSelectedChart,
  canWatchSelection,
  slideImageBase64,
  canInsertSlidesFromBase64,
  insertSceneIntoSlide,
  insertSlidesFromPptx,
  listChartsInDeck,
  scanIsComplete,
  scanGap,
  newRunId,
  reconcileDeck,
  showSlide,
  slideCount,
  slideSize,
  slideShapeList,
  timeShapeRounds,
  updateChartInSlide,
  errorText,
  // A live binding, read fresh on every access — the whole point is to diff it
  // around a scenario.
  deadlinesFired,
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
  /**
   * The skip was caused by the DECK SCAN going blind, not by a missing API.
   *
   * A third state, because the summary used to call every skip a host
   * capability gap and five of the nine scenarios skip on an empty scan
   * instead. A run in which the host refused every deck read therefore said, in
   * green, "2 of 2 scenarios passed · 6 skipped (host cannot run them)" — a
   * total loss of deck visibility reported as a feature gap, which is the sort
   * of line someone files and moves on from.
   */
  blind?: boolean;
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

/**
 * Charts this self-test put in the deck, by the title it gave them — AND
 * whether the scan that found them could see the whole deck.
 *
 * Five scenarios draw a conclusion from this coming back empty ("no probe chart
 * to edit", "fewer than two to scale together", and so on) and one draws a
 * conclusion from it coming back empty that it reports as a PASS. All six were
 * reading a blind scan as a fact about the deck. A real host produced
 * `unread=8 slides=8` between two scans that worked, so total blindness here is
 * transient and lands on whichever scenario happens to ask during it.
 */
async function probeCharts(prefix: string) {
  const scan = await listChartsInDeck();
  return {
    found: scan.charts
      .map((c) => ({ ...c, cfg: JSON.parse(c.configJson) as ChartConfig }))
      .filter((c) => typeof c.cfg.title === "string" && c.cfg.title.startsWith(prefix)),
    /** The scan could not see the whole deck — nothing it did NOT find is news. */
    blind: !scanIsComplete(scan),
    gap: scanGap(scan),
  };
}

/**
 * A skip caused by the deck scan going blind, phrased so it cannot be read as a
 * capability gap.
 *
 * `describeSelfTest` used to attribute every skip to "host cannot run them",
 * and five of the nine skip on an empty scan rather than on a missing API. A
 * run in which the host refused every deck read therefore reported, in green,
 * "2 of 2 scenarios passed · 6 skipped (host cannot run them)" — a total loss of
 * deck visibility filed as a feature gap.
 */
const blindSkip = (gap: string) => ({
  ok: false,
  skipped: true,
  blind: true,
  detail: `the deck scan could not see the whole deck (${gap}), so nothing it did not find means anything`,
});

type Scenario = (prefix: string) => Promise<{ ok: boolean; detail: string; skipped?: boolean; blind?: boolean }>;

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
  const { found } = await probeCharts(`${prefix} twice`);
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
  const { found: kept } = await probeCharts(`${prefix} dup`);
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
  const { found, blind, gap } = await probeCharts(prefix);
  const [chart] = found;
  if (!chart)
    return blind ? blindSkip(gap) : { ok: false, skipped: true, detail: "no probe chart in the deck to edit" };
  const shown = await showSlide(chart.target.slideId);
  const next = { ...chart.cfg, title: `${chart.cfg.title} (edited)` };
  const target = await updateChartInSlide(buildChart(next), chart.target, { tagData: JSON.stringify(next) });
  if (!target) return { ok: false, detail: "the chart was gone from the slide after the update" };
  const { found: again } = await probeCharts(prefix);
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
  const { found: charts, blind, gap } = await probeCharts(prefix);
  if (charts.length < 2)
    return blind
      ? blindSkip(gap)
      : { ok: false, skipped: true, detail: "fewer than two probe charts to scale together" };
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
  // ABOVE the tallest bar in the deck, not level with it.
  //
  // `Math.max(...values)` is the ceiling these charts are ALREADY drawn at:
  // every probe chart is the same `sampleConfig("clustered")`, whose axis tops
  // out at its own largest value. Writing that back as `scale.max` therefore
  // produced a config that rendered to the byte-identical scene, and a deck
  // saved out of a real PowerPoint proved it — a rescaled chart and an
  // untouched copy of it had the same bar geometry to the EMU. The scenario is
  // named for a deck-wide rescale and was measuring a deck-wide RE-TAG: it
  // could only ever catch a config that failed to round-trip, never a host
  // that took the config and redrew the old picture.
  //
  // Headroom fixes that at the source. Every bar has to shrink to fit a
  // ceiling nothing in the data reaches.
  const HEADROOM = 1.5;
  const max = Math.round(Math.max(...values) * HEADROOM);
  // And prove the ceiling bites, rather than trusting the arithmetic above to
  // stay true. The no-op survived for as long as it did because nothing ever
  // checked that the new config DREW anything different; a future sample whose
  // data reaches this ceiling would put the scenario straight back to
  // measuring nothing, silently, and the next person would find out from a
  // saved deck the way this one was found.
  const wasInk = barInk(buildChart(charts[0].cfg));
  const nowInk = barInk(buildChart({ ...charts[0].cfg, scale: { max } }));
  if (wasInk > 0 && nowInk === wasInk)
    return {
      ok: false,
      detail: `scale.max=${max} redraws the chart identically — the rescale would prove nothing about the host`,
    };
  // What the UPDATE said, kept rather than dropped on the floor.
  //
  // `updateChartInSlide` already reports a chart it redrew and could not tag
  // (`lost: "no-config"`) and one whose new shape it never got an id for
  // (`"unknown-shape"`). This scenario threw the return value away and then
  // re-counted the deck, so its verdict — "4 of 8 charts carry the shared
  // scale" — named a number and no cause, and the deck it produced took a
  // session to read afterwards. The renderer knew which four and why while it
  // was happening.
  const lost: Record<string, number> = {};
  for (const c of charts) {
    const next: ChartConfig = { ...c.cfg, scale: { max } };
    const back = await updateChartInSlide(buildChart(next), c.target, { tagData: JSON.stringify(next) });
    const why = back === null ? "chart-gone" : back.lost;
    if (why) lost[why] = (lost[why] ?? 0) + 1;
  }
  const { found: after } = await probeCharts(prefix);
  const scaled = after.filter((c) => c.cfg.scale?.max === max).length;
  const shrunk = wasInk > 0 ? `, bars redraw at ${Math.round((nowInk / wasInk) * 100)}% of their height` : "";
  const why = Object.entries(lost)
    .map(([k, n]) => `${n}×${k}`)
    .join(", ");
  return {
    ok: scaled === charts.length && after.length === charts.length,
    detail:
      `${scaled} of ${charts.length} charts carry the shared scale (max=${max}${shrunk}); ` +
      `${after.length} still re-editable${why ? `; the update reported ${why}` : ""}`,
  };
};

/**
 * Total height of a scene's bars — the one number a rescale has to move.
 *
 * Summed rather than taken off the tallest bar, because a chart that honours a
 * new ceiling for one series and ignores it for another is still drawn wrong,
 * and a single-bar reading would call that fixed.
 */
function barInk(scene: Scene): number {
  let total = 0;
  for (const n of scene.nodes) if (n.kind === "rect" && n.name?.startsWith("seg-")) total += n.h;
  return total;
}

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
  const { found, blind, gap } = await probeCharts(prefix);
  const [chart] = found;
  if (!chart)
    return blind ? blindSkip(gap) : { ok: false, skipped: true, detail: "no probe chart in the deck to explode" };
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
  // Read the slide BEFORE the collapse, so what the collapse produced can be
  // told apart from what was already there. See below for what that cost.
  const before = await slideShapeList(chart.target.slideId);
  const pictured = await updateChartInSlide(buildChart(asPicture), chart.target, {
    tagData: JSON.stringify(asPicture),
    pictureBase64: png,
  });
  if (!pictured) return { ok: false, detail: "the chart vanished while being collapsed to a picture" };
  // One shape is what a picture IS. More than one means the renderer drew
  // shapes instead and the rest of this scenario would prove nothing — the same
  // trap as the `undefined` png above, one step later, so it has to be a
  // verdict and not a note.
  //
  // Ask what the collapse ADDED, by shape id — never how many shapes the slide
  // holds. Two earlier versions asked the slide and both were wrong:
  //
  //  - `slideHoldsOnlyChart` is the slide-SWAP gate, which answers no for a
  //    slide it could not read, so a fine picture was logged as suspect
  //    whenever this host refused a collection (which is routinely);
  //  - counting the slide's shapes replaced that with a false FAILURE on its
  //    first real round. The slide held three, and all three were charts: the
  //    battery deliberately piles them onto one slide two scenarios earlier,
  //    and `a selected shape survives an insert` reports the same three. A
  //    neighbouring chart is not a broken picture.
  //
  // The id delta has neither problem. It is one shape when a picture lands
  // beside anything at all, and it is N when the renderer falls through to
  // native shapes.
  const after = await slideShapeList(pictured.slideId);
  const delta =
    before && after
      ? {
          added: after.filter((s) => !new Set(before.map((b) => b.id)).has(s.id)),
          was: before.length,
          now: after.length,
        }
      : undefined;
  if (delta && delta.added.length !== 1)
    return {
      ok: false,
      // What was observed, not what caused it: a fall-through to native shapes
      // and a picture landing next to undeleted leftovers both show up here,
      // and the names are what a reader needs to tell them apart.
      detail:
        `the collapse added ${delta.added.length} shapes (${delta.added.map((s) => s.name).join(", ") || "none"}) — ` +
        `a picture is one; the slide went from ${delta.was} to ${delta.now}`,
    };
  // …and when the host will not say, the scenario still runs: the config
  // round-trip below is worth checking either way. It just may not have been a
  // picture making the trip, and the verdict says so rather than claiming one.
  const confirmed = delta !== undefined;
  if (!confirmed) trace("selftest", "the host would not confirm the picture is one shape", { slide: pictured.slideId });
  const asShapes: ChartConfig = { ...chart.cfg, render: "shapes" };
  const exploded = await updateChartInSlide(buildChart(asShapes), pictured, { tagData: JSON.stringify(asShapes) });
  if (!exploded) return { ok: false, detail: "the picture vanished while being exploded back to shapes" };
  const back = (await probeCharts(prefix)).found.find((c) => c.target.shapeId === exploded.shapeId);
  return {
    ok: !!back,
    detail: back
      ? confirmed
        ? "collapsed to a picture and exploded back, config intact"
        : "collapsed and exploded back, config intact — but the host would not confirm it was a picture"
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
  const { found: hosts, blind, gap } = await probeCharts(prefix);
  const [host] = hosts;
  if (!host) return blind ? blindSkip(gap) : { ok: false, skipped: true, detail: "no probe chart to insert alongside" };
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
  const { found: mine, blind: mineBlind, gap: mineGap } = await probeCharts(`${prefix} onto`);
  // The chart that was already there must still be a chart. A run that swept
  // the slide's existing shapes into its own group would take this with it —
  // and delete it on the next edit, because the parts tag would claim it.
  //
  // Both readings take the scan's OWN account of whether it saw the deck.
  // Without that, "I did not find your chart" and "I could not look" are the
  // same sentence — and this scenario said the first while meaning the second
  // on a real host, on a run where the previous scenario alone had taken 102
  // seconds and the host was answering deck reads short. It is the trap
  // `stop a run part-way` already has a rung for: "found nothing, therefore
  // nothing was left behind" is an assertion a blind scan satisfies for free.
  //
  // A blind scan is reported as SKIPPED, not as a pass and not as a failure:
  // nobody has evidence either way, and a verdict that cannot be attributed
  // costs a whole real-host round to chase.
  const rescan = await probeCharts(prefix);
  if (mineBlind || rescan.blind) return blindSkip(rescan.blind ? rescan.gap : mineGap);
  const survivor = rescan.found.find((c) => c.cfg.title === host.cfg.title);
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
  // The ladder runs immediately before this and has just asked, properly, which
  // selection call this host stops answering. If it found one, spending a
  // budget here to be told nothing is pure cost: the answer is already in the
  // report, one line up, in more detail than this scenario could produce.
  //
  // Reported as SKIPPED rather than failed, because that is what it is. A
  // wedged host is a fact about the host; calling it a failure of the pane's
  // most-used read is the misattribution the ladder exists to prevent.
  if (selectionWedged) {
    return { ok: false, skipped: true, detail: `not attempted — the ladder just found this host ${selectionWedged}` };
  }
  const { found, blind, gap } = await probeCharts(prefix);
  const [chart] = found;
  if (!chart)
    return blind ? blindSkip(gap) : { ok: false, skipped: true, detail: "no probe chart in the deck to select" };
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
    const round = (await probeCharts(prefix)).found.find((c) => c.cfg.title === next.title);
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
    // ...but only when it was a SELECTION call that went silent.
    //
    // This used to accept any timeout, and `isTimeout` is just as true of a
    // draw that stalled. A round on 2026-08-08 skipped with the sentence below
    // while its own trace read `gave up waiting what=drawing shapes 1-10 of 24`
    // and `at=redrawing the chart's shapes` — the selection subsystem was fine,
    // the host stopped answering during a redraw, and the report sent the reader
    // to two selection bugs that had nothing to do with it.
    //
    // A diagnosis is worth having only if it can be wrong.
    const at = stepOf(err) ?? (err instanceof Error ? err.message : undefined);
    if (isTimeout(err) && wedgedSelection(err)) {
      return {
        ok: false,
        skipped: true,
        detail:
          "the host stopped answering selection calls after a programmatic select — known web-host limitation, " +
          "same family as office-js#3083 / #3698; the pane's own Edit-it path is unaffected",
      };
    }
    // A timeout somewhere else is a real FAILURE and says where. Skipping it
    // would hide a stall behind a known-limitation label, which is how a host
    // problem stops being counted.
    if (isTimeout(err)) {
      return { ok: false, detail: `the host stopped answering${at ? ` at: ${at}` : ""} — not a selection failure` };
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
  // The third assertion is "found nothing, therefore nothing was left behind" —
  // which a blind scan satisfies for free. On a real host this scenario reported
  // PASS off a scan that had just read 1 of 8 slides (`unread=7 slides=8`, the
  // line immediately before its own verdict), while two other scenarios reported
  // SKIPPED under exactly the same blindness. An assertion that cannot fail is
  // not a guard, and this one is guarding a promise — that Stop is
  // non-destructive — which nothing else checks.
  const leftovers = await probeCharts(`${prefix} stopped`);
  if (leftovers.blind && leftovers.found.length === 0) return blindSkip(leftovers.gap);
  const problems = [
    outcome !== "stopped" && outcome,
    after !== before && `the deck grew by ${after - before} — a stopped insert added a slide`,
    // A stop must not leave a chart the pane would offer to edit: half a chart
    // that claims to be whole is worse than a visible mess.
    leftovers.found.length > 0 && "a stopped insert left a re-editable chart behind",
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
/**
 * Whether a timed-out step was a SELECTION call, and so the known web-host wedge.
 *
 * Pure and exported because it is the whole of a diagnosis, and a diagnosis
 * that cannot be tested is a label. The steps that count are the ones this
 * scenario makes against the selection subsystem; a draw or a redraw that
 * stalls is a different fact with a different cause, and calling it a selection
 * limitation buries a host stall under a citation.
 *
 * Unknown (`undefined`) is deliberately NOT a selection wedge. An error with no
 * step attached could have come from anywhere, and the honest report for that is
 * a failure that says so rather than a confident wrong answer.
 */
export function wedgedSelection(err: unknown): boolean {
  const step = stepOf(err);
  // The MESSAGE too, and it is not belt-and-braces: a bounded wait rejects with
  // `PowerPoint did not respond while <what>`, and on the path this scenario
  // actually takes nothing attaches a step to that error at all. Reading only
  // the step made every wedge look like "somewhere unknown", which fails the
  // scenario that this branch exists to skip.
  const msg = err instanceof Error ? err.message : typeof err === "string" ? err : "";
  return /select/i.test(step ?? "") || /select/i.test(msg);
}

/**
 * What two rasters and a naming outcome amount to.
 *
 * Pure and exported for the reason `targetWithNoTagResult` is: the state that
 * exposed the bug needs several simultaneous host failures to reach through the
 * fake, and every attempt to arm them overshot into a different failure. The
 * rule is the thing that was wrong; it should be checkable without a
 * PowerPoint.
 *
 * `named` is whether the insert handed back an edit target. It is NOT whether
 * anything was drawn, and reading it that way is the bug this replaces:
 * `insertSceneIntoSlide` returns null when it has no id to hand back, which on
 * PowerPoint web happens after a perfectly good draw. The 2026-08-08 round
 * committed all three batches (`upTo=24 total=24`) and then failed to name the
 * group — `shapes.getItem("51")`, annotated by the host as originally an
 * `addGroup`, refused with InvalidParam — and the scenario reported "nothing
 * was drawn, so there is
 * nothing to look at" over twenty-four shapes that were on the slide. Calling a
 * drawn chart undrawn is the one mistake a visibility check must not make.
 *
 * So the image decides, and `named` only colours the answer:
 *
 * - changed, named — the plain pass.
 * - changed, unnamed — still a pass. The chart is visible; that it carries no
 *   config is a real defect, and a different scenario's business.
 * - unchanged, named — the chart is on the slide and cannot be seen. This is
 *   the defect the scenario was written for.
 * - unchanged, unnamed — "nothing drew" and "it drew invisibly" both fit and
 *   the image cannot separate them. Say so rather than pick one.
 */
export function visibilityVerdict(before: string, after: string, named: boolean): { ok: boolean; detail: string } {
  if (after !== before)
    return {
      ok: true,
      detail:
        `drawing the chart changed what the slide looks like (${before.length} → ${after.length} bytes)` +
        (named ? "" : " — though the host would not name the chart afterwards, so it carries no config"),
    };
  return {
    ok: false,
    detail: named
      ? "the slide renders identically with and without the chart — nothing is visible"
      : "the slide renders identically and the host would not name a chart on it — " +
        "cannot tell a chart that never drew from one that drew invisibly",
  };
}

const chartIsVisible: Scenario = async (prefix) => {
  // Say what is about to be tried, BEFORE trying it.
  //
  // `scenario starting` names the scenario and nothing finer, and this scenario
  // makes five host calls — any of which could be the one that does not come
  // back. A real host proved the difference matters: the run log's last line
  // was this scenario announcing itself, and nothing after it was ever written,
  // so the evidence narrowed the cause to "somewhere in here" and stopped. Five
  // extra entries in a two-thousand-entry ring is a cheap price for a log that
  // names the call instead of the scenario.
  const attempt = async <T>(what: string, fn: () => Promise<T>): Promise<T> => {
    trace("selftest", "visibility step", { what });
    return fn();
  };

  // A slide this run put in the deck EARLIER — never one added moments ago.
  //
  // It used to take its own scratch slide, and that is the call which has now
  // killed PowerPoint on the web five rounds running. The fifth was the
  // experiment that says so: picked alone, the run reached this scenario at
  // 61.5s with only the two inserts it depends on in front of it, added a
  // scratch slide at 61.5s, logged `rasterising the empty slide` at 61.8s, and
  // the tab died. The four rounds before it died ten minutes and nine
  // scenarios in, which is why "the scenario kills the host" and "ten minutes
  // of drawing kills the host, and this is what happened to be running" both
  // fit. They do not both fit any more.
  //
  // A fresh slide is the worst surface this host offers — its id does not
  // round-trip through `getItem`, its shapes will not list, and
  // `getImageAsBase64` has now failed on one in five distinct ways: a
  // GeneralException, a call taken that produced nothing, a sync never
  // answered, a ninety-second stall, and this. Every one of those is a fresh
  // slide; a pre-existing slide has never been asked.
  //
  // So ask the question that has not been asked, on the surface that has not
  // failed. Before-and-after on ONE slide is still what isolates the chart —
  // that part was never the problem — and the slide simply need not be empty
  // for the comparison to mean anything. It also drops the scratch add and the
  // delete, and the delete is separately implicated: one earlier round died on
  // it rather than on the rasterise.
  const { found, blind, gap } = await probeCharts(prefix);
  const host = found[0];
  if (!host)
    return blind ? blindSkip(gap) : { ok: false, skipped: true, detail: "no slide from this run to draw a chart on" };
  const slideId = host.target.slideId;

  // Named apart from the old step on purpose. If a tab still dies here, the
  // crash log says `rasterising a slide that already existed` and that is the
  // remaining reading — this host cannot rasterise AT ALL — rather than another
  // repeat of the one now settled.
  const before = await attempt("rasterising a slide that already existed", () => slideImageBase64(slideId, 640));
  if (!before) return { ok: false, skipped: true, detail: "host will not rasterise a slide (PowerPointApi 1.8)" };
  const c = cfg(`${prefix} visible`);
  const drawn = await attempt("drawing the chart", () =>
    insertSceneIntoSlide(buildChart(c), { slideId, tagData: JSON.stringify(c) }),
  );
  // A null target is NOT "nothing was drawn", so the rasterise happens either
  // way — see `visibilityVerdict`, which is where that used to be decided
  // wrongly. The measurement is the IMAGE, and the slide id came from the
  // caller rather than from the draw.
  const after = await attempt("rasterising the slide with the chart", () => slideImageBase64(slideId, 640));
  if (!after) return { ok: false, detail: "the host rasterised the slide before the chart but not after it" };
  // No cleanup. Every other scenario leaves its slides in the deck, and this
  // one only cleaned up because a scratch slide is a control surface nobody
  // would want to open. A chart on a slide the run already owns is an ordinary
  // result, so the delete — which cost one round on its own — simply goes.
  return visibilityVerdict(before, after, !!drawn);
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
 * The two added most recently — selection and stop — therefore run at the end.
 * The heaviest and least proven, `chartIsVisible`, ran dead last for the same
 * reason (it was 6th, and a crash there cost the verdicts of three scenarios
 * that had worked on a real host that morning). Last was not far enough: it
 * killed the tab in four consecutive rounds without ever returning a verdict,
 * and a battery that never returns never writes its report. It is `pickedOnly`
 * now — see its entry. New code must not be able to take the evidence for old
 * code with it, and "last" only buys that up to the point where the crash costs
 * the report itself.
 *
 * The first two remain first because everything else needs the probe charts
 * they insert. Beyond that the constraint is `explodePicture`: it inserts a
 * PICTURE, and on the web a picture cannot be inserted while a shape is
 * selected (office-js#3698) — so every scenario that selects must clear up
 * after itself, which `clearShapeSelection` does, wherever it runs.
 */
/**
 * Consecutive host-trouble scenarios before the battery stops asking.
 *
 * Three, not one: a single failure is usually the finding this whole file
 * exists to produce. Three in a row is not a finding about a scenario, it is a
 * finding about the host — and every real-host run so far has continued for
 * minutes past that point and then died, taking its own verdicts with it.
 */
const SICK_LIMIT = 3;

/**
 * Whether a finished scenario says the HOST is in trouble, as opposed to
 * having found something.
 *
 * Three facts, and `timedOut` is the one that matters: a bounded wait that hit
 * its deadline, a throw, or a deck scan that went blind. A plain FAILED verdict
 * is none of them — that is the battery working. Nor is a capability skip,
 * which says nothing about how the host is feeling.
 *
 * `timedOut` is passed IN, measured by diffing `deadlinesFired` around the
 * scenario, because the previous version read it out of the verdict's prose:
 * `/did not respond|gave up/` over `result.detail`. A real run walked straight
 * through that. Two scenarios timed out and used those words; a third timed out
 * after 49.8 seconds and reported *"the host stopped answering selection calls
 * after a programmatic select"* — true, specific, and matching neither phrase.
 * The counter reset one short of the limit, the battery spent two more
 * scenarios on a host that had been dead for three, and the tab died with the
 * remaining verdicts inside it.
 *
 * Adding that third phrase to the pattern would have been the same bug with a
 * longer list. A scenario chooses its own words; it does not choose whether a
 * deadline fired.
 *
 * A SKIP counts when a deadline fired inside it — that third scenario reported
 * skipped, and it was the sickest of the three: it had waited out the whole
 * budget to say so.
 */
export function hostSeemsSick(result: ScenarioResult, timedOut: boolean): boolean {
  return timedOut || result.blind === true || result.detail.startsWith("threw:");
}

/** How long the battery waits for a person to click a chart. */
let CLICK_WAIT_MS = 30_000;

/** Test-only: a suite cannot spend thirty seconds waiting for nobody. */
export function _setClickWaitForTest(ms: number): void {
  CLICK_WAIT_MS = ms;
}

/**
 * The one path a real user travels that nothing can script.
 *
 * `setSelectedShapes` is Office.js selecting a shape. A human clicking one is
 * the same call in theory, and on PowerPoint on the web demonstrably not the
 * same in practice — the programmatic version is taken and then the selection
 * subsystem stops answering, which is why `editViaSelection` reports *skipped*
 * there. That leaves the pane's most-used read — the one behind "Edit it" —
 * with no coverage at all on the host where all the bugs are.
 *
 * So ask a person. One click, and then the whole chain runs and is measured:
 * read the selection back, confirm it is the chart that was clicked, edit
 * through the target THAT read produced, and confirm the result is still
 * re-editable.
 *
 * That is not a workaround for being unable to script it. It is the more
 * faithful test, and this project has the measurement to say so: the scripted
 * version does not behave like the real one on the host that matters.
 *
 * Listens via `DocumentSelectionChanged` — a Common API event that does not go
 * through the wedging subsystem, and the same one the pane's own selection
 * banner has always used. Nothing here calls `setSelectedShapes`.
 *
 * Picked only, for the obvious reason: it blocks on a human.
 */
const editViaRealClick: Scenario = async () => {
  if (!canWatchSelection()) {
    return { ok: false, skipped: true, detail: "host does not raise DocumentSelectionChanged" };
  }
  prompt?.("Click a PowerChart on a slide now — the self-test is waiting for it.");
  trace("selftest", "WAITING FOR YOU: click a chart on the slide", { seconds: CLICK_WAIT_MS / 1000 });
  const {
    chart: picked,
    sawClick,
    readFailed,
  } = await awaitSelectedChart(CLICK_WAIT_MS, selectionBudgetMs(), (left) => {
    prompt?.(`Click a PowerChart on a slide — ${left}s left.`);
    trace("selftest", "still waiting for a click", { secondsLeft: left });
  });
  if (!picked) {
    // Two different findings, and only one of them is about the person. A
    // click the host would not describe is a HOST result and belongs in the
    // report as one; calling it "nobody clicked" blames the reader for a
    // failure they can see they did not cause, which is how a report stops
    // being trusted.
    if (readFailed) {
      return {
        ok: false,
        skipped: true,
        detail: "you clicked, and the host would not say what was selected — the selection read never came back",
      };
    }
    return {
      ok: false,
      skipped: true,
      detail: sawClick
        ? "you clicked, but what you clicked is not a PowerChart — nothing was checked"
        : `nobody clicked a PowerChart within ${CLICK_WAIT_MS / 1000}s — nothing was checked`,
    };
  }
  prompt?.("Got it — editing the chart you clicked.");
  const was = JSON.parse(picked.configJson) as ChartConfig;
  const title = typeof was.title === "string" ? was.title : "(untitled)";
  // The read is the half this scenario exists for, and it has just succeeded —
  // so it is stated first and unconditionally. Everything after it is the
  // ordinary edit path, which other scenarios already cover; what none of them
  // can cover is that the config came back from a HUMAN's click.
  const read = `read "${title}" back from a real click`;
  // Edit through the target the SELECTION produced. That target is built by
  // different code, from a different read, than the one the deck scan makes —
  // and only this one is what a user's edit actually travels on.
  const next = { ...was, title: `${title} (via a real click)` };
  const target = await updateChartInSlide(buildChart(next), picked.target, { tagData: JSON.stringify(next) });
  if (!target) return { ok: false, detail: `${read} — and the host would not edit through that target` };
  const round = await loadChartFromSelection(selectionBudgetMs()).catch(() => null);
  return {
    ok: true,
    detail:
      `${read}, edited through the selection's own target, and it is ` +
      (round?.configJson ? "still re-editable" : "no longer readable from the selection (the view may have moved)"),
  };
};

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
/**
 * What the ladder found, in a form the scenario after it can act on.
 *
 * A sentence in a report is for a person. This is for `editViaSelection`, which
 * runs next and would otherwise spend a budget re-discovering it. Reset per run,
 * because a stale answer from the last round is worse than none.
 */
let selectionWedged: string | null = null;

const whichSelectionCallWedges: Scenario = async (prefix) => {
  if (!canSelectShapes()) return { ok: false, skipped: true, detail: "host cannot select shapes (PowerPointApi 1.5)" };
  const { found, blind, gap } = await probeCharts(prefix);
  const [chart] = found;
  if (!chart)
    return blind ? blindSkip(gap) : { ok: false, skipped: true, detail: "no probe chart in the deck to select" };
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
  selectionWedged = `goes silent at "${stopped.step}"`;
  return {
    ok: true,
    detail:
      `SILENT at "${stopped.step}" after ${stopped.ms}ms` +
      (before ? `, the rung before it ("${before.step}") answered in ${before.ms}ms` : ", the very first rung") +
      ` — ${answered} rung(s) answered before it`,
  };
};

/**
 * Does drawing a chart destroy a shape the user had selected?
 *
 * office-js#2775, reported against PowerPoint on the web and still open: adding
 * a text box deletes whatever shape was selected, and the reporter notes it
 * "works fine on desktop". Every chart this add-in draws contains text boxes,
 * and the insert path deliberately leaves the user's selection alone because
 * that is how it learns where to put the chart — so if the bug is live on the
 * owner's build, selecting a picture and inserting a chart against it destroys
 * the picture. Silently, on the everyday path.
 *
 * `dropShapeSelection` is the guard, and it went in without waiting for an
 * answer because the cost is one selection call and the failure is a user
 * losing their own content. This is the question that says whether that guard
 * is load-bearing on THIS host or merely cheap insurance — and it is worth
 * knowing, because a host that does this will do it on every other add-in the
 * owner uses too.
 *
 * Asked on a scratch slide with a shape of our own, never on anything of the
 * user's. Placed after the ladder for the reason the ladder's own comment
 * gives: it calls `setSelectedShapes`, so it must not be the first thing in the
 * run that does.
 */
const selectionSurvivesInsert: Scenario = async (prefix) => {
  if (!canSelectShapes()) return { ok: false, skipped: true, detail: "host cannot select shapes (PowerPointApi 1.5)" };
  if (selectionWedged) {
    return { ok: false, skipped: true, detail: `not attempted — the ladder found this host ${selectionWedged}` };
  }
  // A probe chart on a settled slide, never a fresh scratch slide.
  //
  // The first draft added its own slide and selected a shape on it, and the
  // fake refused: `selectShape` resolves the slide, syncs, then reaches through
  // that same handle — and a freshly-added slide's handle is good for exactly
  // one sync. That is this repo's oldest gotcha biting a new caller, and the
  // scenario is better for the correction anyway: a shape on a slide the user
  // has had for a while is the situation #2775 actually describes.
  const { found, blind, gap } = await probeCharts(prefix);
  const [victim] = found;
  if (!victim) return blind ? blindSkip(gap) : { ok: false, skipped: true, detail: "no probe chart to select" };
  const slideId = victim.target.slideId;
  const inventory = async () =>
    (await listChartsInDeck({ withInventory: true })).inventory?.find((s) => s.slideId === slideId);
  const before = await inventory();
  if (!before) return { ok: false, skipped: true, detail: "the host would not say what is on the slide" };
  if (!(await selectShape(slideId, victim.target.shapeId, selectionBudgetMs()))) {
    return { ok: false, skipped: true, detail: "the host would not select the chart" };
  }
  try {
    // And now draw, with that shape still selected. Text boxes and all — which
    // is exactly #2775's repro, and exactly what the pane did on every insert
    // before `dropShapeSelection`.
    const extra = cfg(`${prefix} drawn while selected`);
    await insertSceneIntoSlide(buildChart(extra), { slideId, tagData: JSON.stringify(extra) });
  } finally {
    // Never leave a chart selected: on the web a picture cannot be inserted
    // while one is (office-js#3698), so the scenario after this would fail
    // instead of this one.
    await clearShapeSelection(slideId, selectionBudgetMs());
  }
  const after = await inventory();
  if (!after) return { ok: false, skipped: true, detail: "the host would not say what is on the slide afterwards" };
  const kept = new Set(after.shapes.map((s) => s.id));
  const lost = before.shapes.filter((s) => s.id && !kept.has(s.id));
  return {
    ok: lost.length === 0,
    detail:
      lost.length === 0
        ? `all ${before.shapes.length} shape(s) already on the slide survived an insert made while one was selected`
        : `${lost.length} of ${before.shapes.length} shape(s) VANISHED when a chart was drawn with one selected ` +
          `— office-js#2775 is live on this host, and dropShapeSelection is what stands between it and the user`,
  };
};

/** Fewer rounds than this and the slope is noise, not a curve. */
const MIN_ROUNDS_FOR_A_VERDICT = 4;

/**
 * How much a series' last third cost more than its first third, in ms.
 *
 * Thirds rather than last-minus-first, because one slow round is not a trend
 * and a two-sample comparison cannot tell the two apart. `null` when there is
 * not enough of a curve to say anything, which the caller reports rather than
 * rounding to "no change" — an unmeasured slope and a flat one are different
 * answers, and this file exists because they kept getting confused.
 *
 * **Median of each third, not the mean.** A third is three samples, and one
 * garbage collection is enough to move a mean of three. Run against a fake host
 * with no slowdown at all, the mean version read
 * `[6, 6, 6, 6, 10, 7, 17, 6]ms` — one 17ms hiccup — as "THE CONTEXT degrades,
 * +67%". A confident wrong answer from noise is the worst thing this experiment
 * could produce, because the whole reason it exists is that the cheap answers
 * were untrustworthy.
 */
function growth(ms: number[]): { head: number; delta: number } | null {
  if (ms.length < MIN_ROUNDS_FOR_A_VERDICT) return null;
  const k = Math.ceil(ms.length / 3);
  const median = (xs: number[]) => [...xs].sort((a, b) => a - b)[Math.floor(xs.length / 2)];
  const head = median(ms.slice(0, k));
  return { head, delta: median(ms.slice(-k)) - head };
}

/**
 * How much growth counts as degradation, as a fraction of the reference round.
 *
 * Half again over the measured span, and chosen rather than derived — nothing
 * here has the sample size to justify a real threshold, and pretending
 * otherwise would be worse than saying so. It sits well above the round-to-round
 * noise of a healthy host and far below the tenfold blowups the real artefacts
 * show, which is all it has to separate.
 */
const GREW = 0.5;

/**
 * And how much growth counts in plain milliseconds, whatever the fraction says.
 *
 * The relative test alone is unsafe at the bottom of the scale: against a 6ms
 * round, 3ms of scheduler noise is "+50%". Both bars have to clear, so the
 * verdict needs a trend that is both proportionally large AND big enough to
 * matter — and 25ms is far below the hundreds of milliseconds every real
 * degraded run shows, while being comfortably above what a busy machine adds to
 * a round on its own.
 */
const GREW_MS = 25;

export interface DegradationVerdict {
  /** Growth across the long-context arm, as a fraction of the reference round. */
  oneContext: number;
  /** The same, for the arm that took a fresh context per round. */
  freshContext: number;
  suspect: "context" | "host" | "both" | "none" | "unknown";
  /** The sentence the report leads with. */
  headline: string;
}

/**
 * Read the two curves.
 *
 * Both arms grow the deck by the same amount in the same order, and both age
 * the tab. The ONLY thing that differs is whether the context is re-created
 * between rounds, so:
 *
 * - only the long-context arm grows → the request context is what degrades, and
 *   the fix is a shape of code we already know how to write (shorter contexts).
 * - the fresh-context arm grows too → the host itself is slowing as the deck
 *   grows or the tab ages, and shortening contexts will not help.
 * - both grow, the long one much harder → both, and the first is still worth
 *   fixing.
 * - neither grows → whatever kills a long run is not in this loop, which rules
 *   out three suspects at once and is the most useful answer of the four.
 *
 * **Growth in milliseconds against one shared reference, never each arm's own
 * ratio.** The first version of this compared tail/head within each arm, and it
 * gets the commonest case exactly wrong: under a host that slows linearly the
 * arms are equally affected in ms, but the second arm starts from a higher
 * baseline, so its RATIO is smaller. Worked through with the real numbers the
 * long arm read ×2.64 and the fresh arm ×1.47 for a slowdown neither arm caused
 * — a confident "the context did it" about a host. Both deltas divided by one
 * common reference are order-blind, and that case comes out equal.
 *
 * Residual limit, stated because the raw curves are in the report for exactly
 * this: the arms run one after the other, so a host whose cost grows STRONGLY
 * non-linearly is not fully cancelled — a logarithmic trend still hands the
 * first arm the steeper stretch. A linear trend, which is what a deck growing
 * by a fixed number of shapes per round should produce, cancels exactly.
 */
export function readDegradation(oneContextMs: number[], freshContextMs: number[]): DegradationVerdict {
  const one = growth(oneContextMs);
  const fresh = growth(freshContextMs);
  if (!one || !fresh) {
    return {
      oneContext: NaN,
      freshContext: NaN,
      suspect: "unknown",
      headline: `not enough rounds to read a slope (${oneContextMs.length} in one context, ${freshContextMs.length} in fresh ones)`,
    };
  }
  // The earliest, cheapest measurement in the experiment, and the same divisor
  // for both arms — that sameness is the whole point. Floored at a millisecond
  // because a round the clock reads as free is below the instrument, not free:
  // dividing by it would turn timer jitter into an infinite slowdown.
  const ref = Math.max(one.head, 1);
  const oneGrew = one.delta / ref;
  const freshGrew = fresh.delta / ref;
  const pct = (g: number) => `${g >= 0 ? "+" : ""}${Math.round(g * 100)}%`;
  const both = `one context ${pct(oneGrew)}, fresh contexts ${pct(freshGrew)}, against a ${Math.round(one.head)}ms round`;
  const verdict = (suspect: DegradationVerdict["suspect"], headline: string) => ({
    oneContext: oneGrew,
    freshContext: freshGrew,
    suspect,
    headline,
  });
  // Both bars, never either alone — see GREW_MS.
  const rose = (g: number, deltaMs: number) => g >= GREW && deltaMs >= GREW_MS;
  const oneRose = rose(oneGrew, one.delta);
  const freshRose = rose(freshGrew, fresh.delta);
  if (oneRose && !freshRose) return verdict("context", `THE CONTEXT degrades — ${both}`);
  if (oneRose && freshRose) {
    return oneGrew >= freshGrew * 1.5
      ? verdict("both", `THE HOST is slowing and the context makes it worse — ${both}`)
      : verdict("host", `THE HOST is slowing whatever the context does — ${both}`);
  }
  if (freshRose) {
    return verdict("host", `THE HOST is slowing, and holding a context open did not add to it — ${both}`);
  }
  // Phrased as a statement about the THRESHOLD, not about reality. "Neither
  // curve grew" is simply false in front of two curves that each climbed a
  // tenth, which is a thing this has already printed — and a summary that
  // contradicts the numbers under it is how a reader learns to stop reading
  // either.
  return verdict("none", `no degradation worth the name — ${both}`);
}

/** How big the degradation experiment is. Shrunk by tests, which need the shape and not the minutes. */
let degradeRounds = 8;
let degradePerRound = 12;

/** Test-only: run the experiment at a size a suite can afford. */
export function _setDegradeSizeForTest(rounds: number, perRound: number): void {
  degradeRounds = rounds;
  degradePerRound = perRound;
}

/**
 * Does a long run get slower because of the CONTEXT, the DECK, or the TAB?
 *
 * Every real-host artefact this project owns degrades — 496s to reach scenario
 * seven, a 38-item run whose later inserts cost multiples of its first, a tab
 * PowerPoint eventually killed — and not one of them can say which of the three
 * it was, because all three grow together in an ordinary run. Reasoning about it
 * has already cost this project two sessions; the rule the repo adopted from
 * that is to ask a question that separates the answers instead, which is what
 * this is.
 *
 * Two arms, each drawing the same shapes onto a slide of its own, differing in
 * one thing: whether the request context is re-created between rounds. Separate
 * slides on purpose — sharing one would make the second arm draw onto a slide
 * already holding the first arm's shapes, and a fat slide is a fourth variable.
 *
 * Picked only, and for both of the reasons that word exists here. It costs
 * `2 × rounds` syncs and a few hundred shapes, which is not something to put in
 * front of every routine run; and its subject is a host that has been fatigued,
 * so running it after eight other scenarios would measure them instead of it.
 */
const degradesOverTime: Scenario = async (prefix) => {
  const rounds = degradeRounds;
  const perRound = degradePerRound;
  const slideA = await addScratchSlide();
  const slideB = slideA ? await addScratchSlide() : null;
  if (!slideA || !slideB) {
    return {
      ok: false,
      skipped: true,
      detail: `the host would not give the experiment two slides to draw on (${slideA ? "one" : "none"} of two)`,
    };
  }
  const deckBefore = await slideCount();
  const one = await timeShapeRounds(slideA, {
    rounds,
    perRound,
    oneContext: true,
    label: `${prefix} one-context`,
  });
  const fresh = await timeShapeRounds(slideB, {
    rounds,
    perRound,
    oneContext: false,
    label: `${prefix} fresh-context`,
  });
  const oneMs = one.rounds.map((r) => r.ms);
  const freshMs = fresh.rounds.map((r) => r.ms);
  const verdict = readDegradation(oneMs, freshMs);
  trace("selftest", "degradation curves", {
    oneContext: oneMs,
    freshContext: freshMs,
    suspect: verdict.suspect,
    deckBefore,
  });
  // The raw curves go in the detail, not just the verdict. A ratio is a
  // conclusion and this file has been wrong about conclusions before; the
  // numbers are what someone can re-read when the threshold above turns out to
  // be the wrong one.
  const curves =
    `one context: [${oneMs.join(", ")}]ms` +
    (one.cutShort ? ` (cut short — ${one.cutShort})` : "") +
    `; fresh contexts: [${freshMs.join(", ")}]ms` +
    (fresh.cutShort ? ` (cut short — ${fresh.cutShort})` : "");
  return {
    ok: verdict.suspect === "none",
    // "unknown" means the host did not let the experiment run far enough, which
    // is a skip — the same distinction every other scenario here draws between
    // "we did not check" and "we checked and it is bad".
    ...(verdict.suspect === "unknown" ? { skipped: true } : {}),
    detail:
      `${verdict.headline} — ${rounds}×${perRound} shapes per arm, ` +
      `deck was ${deckBefore} slide(s) at the start. ${curves}`,
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
  // The ladder, and the position is the whole argument.
  //
  // It used to be picked-only, on the grounds that it must be the ONLY thing
  // that has touched the selection subsystem — a run of its own, not a position
  // in a list. That was too strong, and it cost the owner a separate five-minute
  // round every time.
  //
  // The property that actually matters is narrower: **the ladder has to be the
  // first `setSelectedShapes` in the run.** That is the call this project
  // measured wedging and the one office-js#3698 names; `setSelectedSlides`,
  // which three earlier scenarios call through `showSlide`, has never wedged
  // anything and the ladder's own second rung is there to notice if it starts.
  // Sitting immediately before `editViaSelection` gives exactly that, because
  // `editViaSelection` is the only other caller.
  //
  // Position 3 was considered and is worse than either. The ladder can wedge the
  // host, so putting it early means six scenarios run against a wedged host
  // instead of two — strictly less coverage than today, dressed up as a saving.
  { name: "which selection call wedges the host", run: whichSelectionCallWedges },
  // Also after the ladder, and for the ladder's own reason: it calls
  // `setSelectedShapes`, so it must not be the first thing in the run that does.
  { name: "a selected shape survives an insert", run: selectionSurvivesInsert },
  { name: "edit the chart the user selected", run: editViaSelection },
  { name: "stop a run part-way", run: stopPartWay },
  // Picked only because it is 0 for 4 and takes the tab with it every time.
  //
  // Four real-host rounds, four different builds (a5b032d, 618b8d8, cedbc6c,
  // cacf58a), and every one of them ends inside this scenario: at 602s, 631s,
  // 603s and 645s, always within a step or two of `adding a scratch slide`. It
  // has never once returned a verdict. Run 7 got as far as a 90-second rasterise
  // timeout; run 8 was handed a rasterise that "took the call and returned
  // nothing"; runs 6 and 9 simply stopped writing.
  //
  // Running last already protects the other scenarios' verdicts — that is why it
  // was put there — but it does not protect the ROUND. The battery's report is
  // written when the battery returns, and it has not returned yet, so every
  // round so far has reached the owner as a crash file and a screenshot rather
  // than a report. A scenario with no verdicts and a perfect record of killing
  // the tab is not providing coverage; the routine list is the wrong place for
  // it.
  //
  // Picked, it also became the experiment nobody had run. Every crash so far
  // happened around the ten-minute mark with nine scenarios' worth of load
  // behind it, so "this scenario kills the host" and "ten minutes of drawing
  // kills the host, and this is merely what was running" both fit — the same
  // two-readings trap this project keeps paying for. Running it ALONE, on a
  // fresh host, separates them in one short round.
  //
  // IT RAN, 2026-08-08, and it settled: the scenario. Picked alone it was
  // reached at 61.5 seconds with only its two inserts in front of it, took a
  // scratch slide at 61.5s, logged `rasterising the empty slide` at 61.8s, and
  // PowerPoint died. Elapsed time and volume of drawing are both out — the same
  // two inserts run at the head of every routine round and kill nothing. The
  // scenario no longer makes that call; see `chartIsVisible`.
  //
  // AND THE RE-RUN CAME BACK, 2026-08-08, on `e49cca8`: no crash. It rasterised
  // a pre-existing slide, drew its 24 shapes, rasterised again, and returned a
  // verdict — the first this scenario has ever produced in six rounds. So the
  // last reading is closed too: this host rasterises perfectly well, and it was
  // the FRESH slide all along, both directions now measured.
  //
  // The verdict it returned was wrong, and that was ours: the host refused to
  // name the group afterwards, `insertSceneIntoSlide` handed back null, and the
  // scenario called twenty-four committed shapes "nothing was drawn". See
  // `visibilityVerdict`.
  //
  // Still `pickedOnly` for one more round, and the criterion is now that it
  // comes back PASSING rather than merely returning. One survival is one
  // sample — the same discipline `UNSTABLE_ANSWERS` applies to answers that
  // flipped, applied here to a scenario that killed the tab five times and then
  // did not.
  //
  // Note what does NOT depend on this: `npm run visible-charts` rasterises every
  // sample in a real browser on every CI run and fails on a chart that is drawn
  // but invisible. What is lost here is the check against PowerPoint's OWN
  // rasteriser, and only from the routine round.
  { name: "the chart is actually visible", run: chartIsVisible, pickedOnly: true },
  // Picked only for the plainest reason there is: it blocks on a human.
  { name: "edit the chart YOU click", run: editViaRealClick, pickedOnly: true },
  { name: "what makes a long run slow down", run: degradesOverTime, pickedOnly: true },
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

/**
 * How a scenario speaks to the person running it.
 *
 * Only one scenario needs it, and it needs it badly: a battery that blocks
 * waiting for a click, without saying so, is indistinguishable from one that
 * has hung — and this project has spent two rounds failing to tell those
 * apart. The trace carries the same words to the Live steps list, so a pane
 * that never wires this up is quieter but not silent.
 */
let prompt: ((message: string) => void) | undefined;

/** Let the pane put a scenario's request in front of the user. */
export function setSelfTestPrompt(fn: (message: string) => void): void {
  prompt = fn;
}

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
  // Per run, never carried between them: a stale "this host wedges" from the
  // last round would make the next one skip a scenario on evidence it no longer
  // has — which is exactly the sort of quiet, sticky wrong answer this file is
  // built to avoid.
  selectionWedged = null;
  const out: ScenarioResult[] = [];
  /** Consecutive scenarios that told us the host is in trouble. */
  let sick = 0;
  /** Set once the breaker has tripped, so the rest report why. */
  let abandoned: string | null = null;
  for (const { name, run } of wanted) {
    // Stop asking a host that has stopped answering.
    //
    // The battery had no rung for this, and every real-host artefact this
    // project owns ends the same way. The last one: at 261.6s a scenario failed
    // ("the chart was gone"), at 261.8s the deck scan read 0 of 8 slides — and
    // the battery then ran Same Scale for 95 seconds (six charts × 24 shapes,
    // six consecutive 5010 failures, two 90-second waits), then a picture
    // round-trip, then two more scenarios, until PowerPoint killed the tab at
    // ~396s. Roughly 130 seconds and 150 shapes were issued AFTER three
    // distinct signals that the host was already gone, and the tab took the
    // remaining verdicts with it.
    //
    // Three in a row rather than one, because a single failure is often the
    // finding — the battery exists to produce those. Three consecutive is not a
    // finding about a scenario, it is a finding about the host.
    if (!abandoned && sick >= SICK_LIMIT) {
      abandoned = `${sick} scenarios in a row could not get an answer out of the host`;
      trace("selftest", "giving up on the host", { after: out.length, why: abandoned });
    }
    if (abandoned) {
      out.push({ name, ok: false, skipped: true, blind: true, detail: `not reached — ${abandoned}`, ms: 0 });
      continue;
    }
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
    // Deadlines already fired before this scenario ran, so the diff below is
    // this scenario's own. Read here rather than once per run: what matters is
    // whether THIS one could not get an answer, not whether the run ever
    // couldn't.
    const deadlinesBefore = deadlinesFired;
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
    const timedOut = deadlinesFired > deadlinesBefore;
    if (timedOut) trace("selftest", "the host missed a deadline in this scenario", { name, sick: sick + 1 });
    sick = hostSeemsSick(result, timedOut) ? sick + 1 : 0;
    out.push(result);
  }
  return out;
}

/** Whether anything in this run needs looking at — a failure, or a blind scan. */
export function selfTestNeedsAttention(results: ScenarioResult[]): boolean {
  return results.some((r) => (!r.ok && !r.skipped) || r.blind);
}

/** The one-line summary the pane shows. */
export function describeSelfTest(results: ScenarioResult[]): string {
  const ran = results.filter((r) => !r.skipped);
  const failed = ran.filter((r) => !r.ok);
  const blind = results.filter((r) => r.skipped && r.blind).length;
  const skipped = results.length - ran.length - blind;
  const parts = [`${ran.length - failed.length} of ${ran.length} scenarios passed`];
  if (failed.length) parts.push(`failed: ${failed.map((f) => f.name).join(", ")}`);
  if (skipped) parts.push(`${skipped} skipped (host cannot run them)`);
  // Counted apart, and never folded into the line above. These are not a
  // capability gap: the host would not let the add-in see the deck, which is a
  // finding, and the one time it happened it was reported in green.
  if (blind) parts.push(`${blind} could not run — the deck scan went blind`);
  return `Self-test — ${parts.join(" · ")}.`;
}
