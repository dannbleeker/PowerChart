/**
 * How long an insert onto a given slide is likely to take.
 *
 * SSF Charts is fast exactly when it CREATES a slide and slow exactly when it
 * ADDS to one: about 0.75s against about 24s, roughly 32x, measured over 2,917
 * timed batches in 169 rounds. The cause is the host's, not ours — a deck is
 * handed over as a generated file, where adding to a live slide has to draw
 * shape by shape.
 *
 * TWO DIFFERENT COMPARISONS, AND ONLY ONE OF THEM IS THIS FILE'S. That 0.75s is
 * the DECK path. A single insert cannot take it: even onto a blank slide the
 * pane draws shape by shape, so a 40-shape chart costs ~15s there too. What this
 * file prices is the second axis — the same shape-by-shape draw, onto a slide
 * that is empty against one that is crowded. Confusing the two is what made the
 * first version of the offer promise "nearly instant"; see `offerSentence`.
 *
 * The user cannot see any of that. This is the number that lets the pane say so.
 *
 * THE CURVE IS THE ARCHIVE'S, NOT A MODEL. These four points are the measured
 * medians, per ten-shape batch, bucketed by how many shapes the target slide
 * already held. Interpolating between them is the most honest estimate
 * available — a fitted formula would look more precise and be less true, since
 * the curve visibly saturates and no fit this project has data for explains why.
 *
 *     shapes already on the slide    median batch
 *       0                              3886ms
 *       1-20                           5490ms
 *      21-50                          13995ms
 *      51-100                         18074ms
 *
 * It is an ESTIMATE and the pane must present it as one. Real spread inside a
 * bucket is wide, the archive's own noise floor is 14% IQR between rounds, and
 * a host having a bad minute can double any of this.
 *
 * WHAT THE CURVE LEAVES OUT, MEASURED 2026-09-05 OVER 4,188 TIMED BATCHES.
 *
 * THE EFFECT IS REAL. Measured the one way composition cannot vary — PAIRED,
 * inside a single draw, where the only thing changing between two consecutive
 * batches is that ten more shapes are on the slide. Batch 2 is slower than
 * batch 1 in **1,270 of 1,330** 24-shape draws (95%), median 3,693 -> 5,548, a
 * 1.5x climb for ten shapes. The 103-shape chart runs the series out to ten
 * positions, n=16 at every one, without a single reversal: 1608 2107 2882 3664
 * 4310 5141 6081 6658 7394 8293. Occupancy costs what this file says it costs.
 *
 * THESE FOUR CONSTANTS ARE NOT A CLEAN MEASURE OF IT, and "the curve has held
 * up" was a circular claim that has now been tested and failed. The constants
 * were fitted at `cc59d4d` on the 261 rounds then archived; 103 rounds have
 * been added since. Re-measured on the ROUNDS IT WAS FITTED ON, the curve
 * matches to -4.2 / +2.1 / -1.8 / -0.3%, which is what fitting means. Held out
 * on the 103 rounds it never saw:
 *
 *     bucket    coded    held-out (n=722)     off by
 *       0        3886      7717 (n=236)       +99%
 *       1-20     5490      6813 (n=230)       +24%
 *      21-50    13995     17236 (n=189)       +23%
 *      51-100   18074      7097 (n=66)        -61%
 *
 * Non-monotonic, and the anchor the pane quotes for a blank slide is 2x low.
 *
 * The reason is that the buckets are CHART-MIX buckets. In the fitted set,
 * bucket 0 is 89% 24-shape charts (the cheap kind) and bucket 51-100 is 98%
 * 16-shape charts (the dear kind), so the climb across buckets is partly just
 * the mix changing. In the held-out set bucket 51-100 is 98% 103-shape charts
 * and reads 7,097ms. Same occupancy, different charts, 2.5x apart.
 *
 * So there are TWO real effects here and one axis to carry them. Occupancy is
 * one, and the paired numbers above are what it actually costs. The other is
 * what the chart IS, and this file has no term for it at all:
 *
 * `estimateInsertMs` multiplies `batchMs` by a batch COUNT, which says
 * every ten-shape batch costs the same. At a chart's first batch, held to a
 * slide this run had genuinely drawn nothing on, the medians run:
 *
 *      16 shapes                    10,098ms   n=258
 *      24 shapes                     3,617ms   n=1,202
 *     103 shapes                     1,584ms   n=15
 *
 * 2.8x between the first two, backwards to shape count, and 6.4x across the
 * table. Confounds ruled out: batch position (a 24-shape chart contributes
 * cheap later batches a 16-shape one never reaches, so this is batch 1 only),
 * draw path (16-shape charts are slow on both), and in-place updates, which
 * emit no `batch issued` line at all and so were never in the sample.
 *
 * READ "GENUINELY DREW NOTHING ON" LITERALLY — the first version of this table
 * did not, and was wrong by 13x instead of 2.8x. A batch-1 line is written
 * BEFORE the host has answered `slide.load("id")`, so it keys on the
 * `(visible)` sentinel, and the retag empties that sentinel after every draw.
 * All 775 such lines therefore report `onSlide 0`, and 516 of them — 66.6% —
 * were on a slide this run had already put a median of 16 and up to 88 shapes
 * on. The true prior is recoverable only from the NEXT batch
 * (`batch2.onSlide` minus what batch 1 drew), which is how the table above is
 * built. The 11-shape rotated chart that headed the first version has NO
 * true-zero readings at all — all 45 are on occupied slides — so its 20,662ms
 * is an occupancy number and cannot sit in this table.
 *
 * So: what a shape IS does cost more than how many there are, but by ~3x, not
 * ~13x, and nothing here prices it — this understates a small ornate chart and
 * overstates a big plain one.
 *
 * TREAT THE OUTPUT AS A FLOOR, NOT A FORECAST, and do not tighten anything
 * against it. Re-expressing the 90-shape budget as a time threshold is ON HOLD
 * for exactly this reason; see BACKLOG item 3. What would fix it is a term for
 * the chart, and a refit that is validated on rounds it was not fitted to —
 * which no version of these constants has ever been.
 *
 * HALF THE BATCHES ARE NEVER TIMED, and not at random. `prevBatchMs` is only
 * readable from the following batch, so every draw's LAST batch is unmeasured —
 * 4,583 of 8,771 — and 1,869 single-batch draws contribute nothing whatever.
 * Every one of the 4,188 readings behind these constants drew exactly ten
 * shapes; not one is a partial tail. So no chart small enough to fit in a
 * single batch appears here at all: the sizes drawn are 7, 9, 10, 11, 14, 16,
 * 24 and 103, and all 1,799 draws at 7, 9 or 10 shapes are untimed —
 * `estimateInsertMs` prices exactly those from a sample containing none.
 *
 * AND `present` MEANS TWO DIFFERENT THINGS EITHER SIDE OF THIS FILE. The curve
 * is indexed by the renderer's `onSlide`, which counts only the shapes THIS RUN
 * drew on that slide. The pane calls `estimateInsertMs` with `occupied.length`
 * — the shapes actually on the slide, whoever put them there. So a user's slide
 * holding forty shapes we never touched is priced from readings taken where we
 * had just drawn forty ourselves, into a context still warm from drawing them.
 * Nothing in 360 rounds measures the other case, and the two need not cost the
 * same. Fixing this is a round that inserts onto a genuinely pre-loaded slide,
 * not an edit here.
 */

/** Measured medians: [shapes already present, ms for one ten-shape batch]. */
const CURVE: readonly (readonly [number, number])[] = [
  [0, 3886],
  [10, 5490],
  [35, 13995],
  [75, 18074],
];

/** Shapes per batch, matching the renderer's own batching. */
const SHAPES_PER_BATCH = 10;

/**
 * Milliseconds for one batch onto a slide already holding `present` shapes.
 *
 * Linear between measured points, flat outside them. Flat rather than
 * extrapolated on the right: past 100 shapes this project has no readings, and
 * a straight line through the last two would keep climbing on no evidence.
 */
export function batchMs(present: number): number {
  const n = Math.max(0, present);
  const first = CURVE[0];
  const last = CURVE[CURVE.length - 1];
  if (n <= first[0]) return first[1];
  if (n >= last[0]) return last[1];
  for (let i = 1; i < CURVE.length; i++) {
    const [x0, y0] = CURVE[i - 1];
    const [x1, y1] = CURVE[i];
    if (n <= x1) return y0 + ((y1 - y0) * (n - x0)) / (x1 - x0);
  }
  return last[1];
}

/**
 * Estimated milliseconds to draw `shapes` onto a slide holding `present`.
 *
 * Uses the STARTING occupancy for every batch, which understates: the slide
 * fills as we draw, so later batches are dearer than the first. Understating is
 * the right direction for a number used to warn — a warning that fires late is
 * a nuisance, one that fires early on a fast insert is noise, and this project
 * would rather be quiet than cry wolf.
 */
export function estimateInsertMs(shapes: number, present: number): number {
  if (shapes <= 0) return 0;
  return Math.ceil(shapes / SHAPES_PER_BATCH) * batchMs(present);
}

/**
 * Above this, the pane offers to put the chart on its own slide instead.
 *
 * Fourteen seconds, set by the owner. It sits between the 1-20 bucket (~5.5s
 * for a ten-shape chart) and the 21-50 one (~14s), so an ordinary chart onto a
 * lightly-used slide stays silent and one onto a loaded slide speaks up.
 */
export const SLOW_INSERT_MS = 14_000;

/** Whether this insert is slow enough to be worth telling the user about. */
export function isSlowInsert(shapes: number, present: number): boolean {
  return estimateInsertMs(shapes, present) > SLOW_INSERT_MS;
}

/**
 * Whether moving this chart to its own slide would actually help.
 *
 * SLOW IS NOT THE SAME AS WORTH MOVING, and conflating them was the first
 * version of this. A big chart takes ~15s on a completely EMPTY slide — a normal
 * insert draws shape by shape wherever it lands — so `isSlowInsert` alone fires
 * there and offers a new slide that would be exactly as slow. An offer that
 * cannot deliver is worse than silence: it spends the user's attention and their
 * trust in the next warning.
 *
 * So the offer needs the slowness to be the SLIDE'S fault. Moving must at least
 * halve the wait, which it does exactly when the target is crowded — the thing
 * the offer is actually about.
 *
 * (The 0.75s figure quoted for "a chart on a new slide" belongs to the deck
 * path, which hands the host a generated FILE and draws nothing. A single insert
 * cannot take that route, so the honest comparison is this one: the same
 * shape-by-shape draw onto an empty slide.)
 */
export function worthOwnSlide(shapes: number, present: number): boolean {
  if (!isSlowInsert(shapes, present)) return false;
  return estimateInsertMs(shapes, 0) <= estimateInsertMs(shapes, present) / 2;
}

/**
 * The sentence the offer puts in front of the user.
 *
 * A pure function rather than a template inside the pane, so what it says can be
 * asserted instead of grepped for. The first draft lived in `app.ts` and the
 * test that guarded it had to search the source for a banned phrase — which then
 * matched the comment explaining the ban. A string a test can read is worth more
 * than a comment a test can trip over.
 *
 * NO DURATIONS, AND THE HISTORY RUNS BOTH WAYS. The first draft called a new
 * slide "nearly instant" — false; that 0.75s belongs to the deck path, and a
 * single insert onto a blank slide still draws shape by shape. It was replaced
 * by the real pair from `estimateInsertMs`, so a user could judge whether a
 * slide was worth the saving.
 *
 * Those numbers came out on 2026-09-05, on the owner's call, because the curve
 * behind them had never been validated out of sample and does not survive it.
 * On the 103 rounds it was not fitted to, the two points this offer's ratio is
 * built from read 7,717 and 7,097 — where a new slide saves nothing at all.
 * Swept across chart sizes and occupancies, 18 of 36 cells flip to "should not
 * have offered". See BACKLOG item 20; the numbers can come back the moment a
 * refit is validated on rounds it was not fitted to.
 *
 * WHAT SURVIVES IS THE PART THAT IS MEASURED RATHER THAN MODELLED. The shape
 * count is read off the slide. And occupancy is real three independent ways:
 * paired inside single draws, where batch 2 beat batch 1 in 1,270 of 1,330; one
 * chart's own ten batches climbing 1,608ms to 8,293ms; and a 7-shape chart at
 * priors 0/7/14/21 costing 5,683/7,128/7,847/9,282. So "this slide is full,
 * which is what makes it slow, and a new one is quicker" is supported by
 * measurement. "By forty seconds" was not.
 */
export function offerSentence(present: number): string {
  const shapes = `${present} shape${present === 1 ? "" : "s"}`;
  return `This slide already holds ${shapes}, which is what makes adding here slow. A new slide would be quicker.`;
}

/**
 * What went wrong on the way, then what happened — one sentence for one slot.
 *
 * The pane has a single live text channel and `note()` REPLACES it, so a
 * setback posted mid-insert is destroyed: `insertSceneIntoSlide`'s first act is
 * a busy phase note into that same slot. Every setback was lost this way, and
 * the user was told only the outcome — their choice had failed and the pane
 * never said so.
 *
 * Worse, a setback SETTLES (it is not "busy"), and `guard` posts "Done." only
 * when nothing settled. So on the path where no closing message applies, the
 * setback suppressed "Done." while the last text written was the phase note —
 * an action that had finished sitting on "Working… done" in busy blue with the
 * progress bar still up. Measured, not supposed: the test "says the slide add
 * failed, and does not end on a busy note" failed on exactly that string before
 * this function existed.
 *
 * SETBACKS FIRST. The user needs to know their choice did not happen before
 * being told what happened instead; the other order buries it.
 *
 * TRANSLATE THE PARTS, NOT THE RESULT. Each piece is its own catalogue key; the
 * join is not, and `t()` would never find it. Callers pass text that has already
 * been through `t()` — this only joins. The separator is a space because these
 * are whole sentences each ending in its own stop, where the repo's "; " is for
 * problems and " · " for tallies.
 */
export function insertOutcomeSentence(setbacks: string[], outcome: string): string {
  return [...setbacks, outcome].filter(Boolean).join(" ");
}

/**
 * `describeMs` LIVED HERE AND IS GONE, 2026-09-05, with the numbers it phrased.
 *
 * It turned 19,400 into "about 20 seconds" and had four tests of its own,
 * including the one that mattered — rounding 58s to the nearest five gives 60,
 * and "about 60 seconds" is a phrase no one says, so the check had to happen
 * after the rounding that produces it rather than before.
 *
 * Deleted rather than kept warm, because `offerSentence` was its only caller
 * and a tested export with no caller is worse than nothing: the tests keep
 * passing and read as coverage of something the product does. If a validated
 * refit brings the durations back (BACKLOG item 20), it is a dozen lines and
 * this commit has them.
 */
