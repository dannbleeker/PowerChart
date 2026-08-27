import { describe, it, expect } from "vitest";
import { readFileSync } from "node:fs";
import { estimateOfficeShapes } from "../src/core/scene";
import { buildChart } from "../src/core/chart";
import {
  isSlowInsert,
  worthOwnSlide,
  estimateInsertMs,
  offerSentence,
  describeMs,
  SLOW_INSERT_MS,
} from "../src/core/insert-cost";
import { sampleConfig } from "../src/core/samples";

/**
 * The slow-insert offer: when the pane speaks up, and what it must never do.
 *
 * Adding a chart to a slide that already holds content costs several times what
 * the same chart costs on an empty one — the host draws shape by shape, and the
 * cost per shape climbs with what is already there. The pane said nothing about
 * it, so a loaded slide simply looked hung.
 *
 * (The archive's famous 32x is a DIFFERENT comparison — a generated deck file
 * against a live insert — and a single insert cannot take the deck path. Mixing
 * the two is what made the first draft promise "nearly instant".)
 *
 * The decision points are covered here as pure functions. The DOM half —
 * `offerOwnSlide` resolving on a click — is wired in `app.ts` and exercised by
 * the pane suites; what matters most is that the RULE is right, because a
 * warning that fires on the everyday case is worse than none.
 *
 * EQUIVALENT MUTANT, recorded rather than chased: widening `worthOwnSlide`'s
 * halving ratio anywhere from 1.0 to about 3.6 changes no answer these tests can
 * see. The archive's buckets are 3886 / 5490 / 13995 / 18074ms, so the ratio
 * between an empty slide and a loaded one lands at 1.4 or 3.6 and never between.
 * A test pinning 2.0 exactly would be testing the constant, not the behaviour.
 */
const shapesOf = (kind: Parameters<typeof sampleConfig>[0]) => estimateOfficeShapes(buildChart(sampleConfig(kind)));

describe("when the pane offers a slide of its own", () => {
  it("never offers a new slide when the CURRENT one is empty", () => {
    // THE BUG THIS TEST FOUND. A big chart takes ~15s on a completely empty
    // slide too — a normal insert draws shape by shape wherever it lands — so
    // "is it slow" alone fires here and offers a new slide that would be exactly
    // as slow. An offer that cannot deliver spends the user's attention and
    // their trust in the next warning.
    for (const kind of ["stacked", "pie", "waterfall"] as const) {
      expect(worthOwnSlide(shapesOf(kind), 0), `${kind} offered a new slide from an empty one`).toBe(false);
    }
  });

  it("speaks up for the same chart on a loaded slide", () => {
    // 35 shapes present is the archive's 21-50 bucket, ~14s a batch — the case
    // the user currently waits through with no explanation, and the one where
    // moving really does help.
    const shapes = shapesOf("stacked");
    expect(worthOwnSlide(shapes, 35), "stayed quiet about a slow insert").toBe(true);
    expect(estimateInsertMs(shapes, 35)).toBeGreaterThan(SLOW_INSERT_MS);
    // Moving has to at least halve it, or it is not worth asking.
    expect(estimateInsertMs(shapes, 0)).toBeLessThanOrEqual(estimateInsertMs(shapes, 35) / 2);
  });

  it("stays quiet about a small chart, however loaded the slide", () => {
    // The halving ratio alone would fire here: ten shapes onto a 35-shape slide
    // is ~14.0s against ~3.9s on a blank one, a clear win by ratio. But 14s is
    // not a wait worth interrupting for, and the offer costs a slide. BOTH tests
    // have to pass before the pane says anything.
    expect(isSlowInsert(10, 35), "a ten-shape chart counted as slow").toBe(false);
    expect(worthOwnSlide(10, 35), "nagged about a ten-shape chart").toBe(false);
  });

  it("is still slow on an empty slide — it just cannot be helped", () => {
    // Keeping the two ideas apart on purpose: the insert IS slow, and saying so
    // in progress is fine. Offering to move it is what would be dishonest.
    const shapes = shapesOf("stacked");
    expect(isSlowInsert(shapes, 0)).toBe(true);
    expect(worthOwnSlide(shapes, 0)).toBe(false);
  });

  it("never fires when the occupancy is unknown", () => {
    // `getSlideShapeBounds` answers null when the host will not describe the
    // slide. The pane treats that as UNKNOWN, never as empty or as loaded: a
    // warning that fires because a read FAILED is worse than one that never
    // fires. The guard is `occupied && …` in app.ts, asserted here as source
    // because the alternative is mounting the whole pane for one boolean.
    const src = readFileSync("src/taskpane/app.ts", "utf8");
    expect(src, "the null-occupancy guard is gone").toMatch(/if \(occupied && worthOwnSlide\(/);
  });

  it("quotes both waits rather than promising the new slide is instant", () => {
    // The first draft said a new slide was "nearly instant". It is not: that
    // 0.75s belongs to the DECK path, which hands the host a generated file. A
    // single insert onto a blank slide still draws shape by shape — for a chart
    // big enough to earn the offer, around fifteen seconds. The user would press
    // the button, wait anyway, and discount every later estimate.
    const shapes = shapesOf("stacked");
    const here = estimateInsertMs(shapes, 35);
    const fresh = estimateInsertMs(shapes, 0);
    const said = offerSentence(35, here, fresh);
    expect(said).toContain(describeMs(here));
    expect(said, "the new-slide wait is not quoted").toContain(describeMs(fresh));
    expect(said, "it went back to promising instant").not.toMatch(/instant|immediate/i);
    // The two numbers have to be DIFFERENT, or the sentence quotes one wait
    // twice and reads as though moving buys nothing.
    expect(describeMs(fresh)).not.toEqual(describeMs(here));
  });

  it("counts one shape in the singular", () => {
    // A guard for the smallest kind of tell. The pane speaks carefully
    // everywhere else; "1 shapes" in the one message that interrupts the user
    // is the sentence they would remember.
    expect(offerSentence(1, 30_000, 5_000)).toContain("holds 1 shape,");
    expect(offerSentence(2, 30_000, 5_000)).toContain("holds 2 shapes,");
    expect(offerSentence(0, 30_000, 5_000)).toContain("holds 0 shapes,");
  });

  it("takes the copy from the pane's own numbers", () => {
    // The sentence is only honest if the pane feeds it the real pair: what this
    // insert costs HERE, and what the same chart costs on a blank slide.
    const src = readFileSync("src/taskpane/app.ts", "utf8");
    expect(src).toMatch(/estimateInsertMs\(sceneShapes, occupied\.length\)/);
    expect(src, "the new-slide figure became a constant").toMatch(/estimateInsertMs\(sceneShapes, 0\)/);
  });

  it("falls back to the original slide when the add is lost", () => {
    // This host DROPS `slides.add()`. `addSlideForChart` returns null when even
    // `addSlides`' retries could not land one, and the insert must then go where
    // the user asked — slowly, but it goes. Losing the offer is a nuisance;
    // losing the chart is not acceptable.
    const src = readFileSync("src/taskpane/app.ts", "utf8");
    expect(src).toMatch(/ownSlideId = \(await addSlideForChart\(\)\) \?\? undefined/);
    expect(src, "a failed add does not tell the user").toMatch(/Could not add a slide/);
    // And the slideId is only passed when one really exists.
    expect(src).toMatch(/ownSlideId \? \{ slideId: ownSlideId \} : \{\}/);
  });

  it("recomputes placement for the empty slide", () => {
    // `at` was cascaded around the shapes on the CROWDED slide. Carrying it to a
    // blank one would offset the chart from obstacles that are not there.
    const src = readFileSync("src/taskpane/app.ts", "utf8");
    expect(src).toMatch(/const place = ownSlideId \? \{ left: 60, top: 90 \} : \{ left: at\.left, top: at\.top \}/);
  });

  it("adds the slide through the retrying path, not a bare slides.add", () => {
    // `addSlideForChart` must go THROUGH `addSlides`, which retries and books
    // the shortfall into `lastAddsLost`. A hand-rolled add would silently put a
    // chart nowhere and report success.
    const src = readFileSync("src/render/powerpoint.ts", "utf8");
    const fn = src.slice(src.indexOf("export async function addSlideForChart"));
    const body = fn.slice(0, fn.indexOf("\n}\n"));
    expect(body, "it stopped going through addSlides").toMatch(/await addSlides\(context, 1,/);
    expect(body, "it grew a bare slides.add()").not.toMatch(/slides\.add\(/);
  });
});
