/**
 * An answer sheet: what THIS PowerPoint actually does, question by question.
 *
 * `CLAUDE.md` states the problem plainly — "every Office.js assertion in this
 * repo is against a fake". The fake has grown twenty-odd faults, each added
 * after a real host taught us something, and it is the reason bugs can be found
 * in CI at all. But nobody has ever checked whether it is FAITHFUL. Its faults
 * are things we learned; its happy path is a set of things we assumed. Where an
 * assumption is wrong, every test resting on it is confidently wrong with it,
 * and no amount of green tells us so.
 *
 * This is not a call log. A recording of what the add-in did is what the trace
 * already gives, and joining two of them proves nothing. This is a fixed list
 * of QUESTIONS — each one a behaviour `powerpoint.ts` actually depends on —
 * asked identically of a real host and of the fake, so the two answer sheets
 * can be diffed. Every divergence is exactly one of two things:
 *
 * - the fake lies, and the tests built on it are worth less than they look, or
 * - the host does something we did not know, which is a finding in itself.
 *
 * Both are worth a round. Neither is visible any other way.
 *
 * **Answers are small and comparable on purpose.** `"yes"`, `"no"`, `"threw"`,
 * `"silent"`, a count — never free text, because the whole point is a machine
 * diff between two hosts. Timings and error messages ride along as `detail`
 * and are deliberately NOT compared: they differ between any two runs and
 * would bury the signal.
 *
 * **Nothing here touches the user's deck.** Every probe works on one scratch
 * slide, added at the start and removed at the end. A probe that would damage a
 * real slide does not belong here.
 */
import {
  addScratchSlide,
  deadlinesFired,
  deckSlideIds,
  deleteTrailingSlides,
  deleteSlideById,
  isTimeout,
  requirementSets,
  ScratchSlideUnavailable,
  withProbeContext,
  type ProbeContext,
} from "./powerpoint";
import { trace, traceElapsed } from "../core/trace";
// @ts-expect-error — a plain .mjs table with no types, and deliberately the
// SAME one the CLI and the CI gate read. Two copies of it is how a claim
// quietly stops matching its check.
import { FAKE_BASELINE, KNOWN_DIVERGENCES, UNSTABLE_ANSWERS, diffAnswers } from "../../scripts/host-baseline.mjs";

/** One question and what this host said. */
export interface HostAnswer {
  /** Stable key. Renaming one breaks the diff against older sheets — don't. */
  id: string;
  /** What was asked, in words, for whoever reads the sheet. */
  question: string;
  /**
   * The comparable fact, from a small vocabulary. This is the ONLY field the
   * diff looks at.
   */
  answer: string;
  /** Wall clock. Context for a human; never diffed. */
  ms: number;
  /** Error text, counts, whatever helps read a surprising answer. Never diffed. */
  detail?: string;
  /**
   * Every answer this run got for this question, in the order it got them.
   *
   * A question asked ONCE has been sampled, not answered — this project has
   * written that sentence in three places and still had to learn each unstable
   * answer the expensive way, across ten rounds, because one round could only
   * ever produce one sample. `UNSTABLE_ANSWERS` in `scripts/host-baseline.mjs`
   * is the artefact of that: a table maintained by hand from whatever rounds
   * happened to be open.
   *
   * A run asks each question up to `PROBE_PASSES` times, spread across the run
   * rather than back to back, so the samples land in different minutes of the
   * host's life. `stable` is then a fact one sheet can state on its own.
   *
   * The ROW stays one per id, and `answer` keeps its meaning — the diff, the
   * contract gate, `host-diff` and `host-baseline` all read that field and none
   * of them needed to change.
   */
  samples?: ProbeSample[];
  /**
   * Whether every real sample agreed. Undefined when the run got fewer than two.
   *
   * `false` here is the finding `UNSTABLE_ANSWERS` exists to record, said by
   * the sheet itself instead of by a human comparing rounds.
   */
  stable?: boolean;
}

/** One ask of one question, and the state the host was in when it answered. */
export interface ProbeSample {
  answer: string;
  /** Which pass over the question list this came from — 1, 2, 3. */
  pass: number;
  /**
   * Milliseconds on the RUN LOG's clock, so a sample can be lined up against
   * the trace lines around it.
   *
   * It used to be milliseconds into the probe, which is a different origin —
   * the probe starts several seconds after tracing does (7.9s in the round of
   * 2026-08-11). Nothing in the file said so, so joining the two series, which
   * is how every one of these rounds gets analysed, was silently off by that
   * constant: the trace put pass 2 at 41.6s and the samples put its first
   * answer at 34.4s, i.e. a sample apparently arriving before the pass that
   * produced it. Falls back to probe-relative only when tracing is off, in
   * which case the file carries no trace to be misjoined with.
   */
  atMs: number;
  /**
   * What the run had already seen the host doing when this was asked.
   *
   * Derived from what the run has ALREADY observed — no extra call, so asking
   * cannot itself disturb the thing being measured. Recorded on every sample
   * and not only on the odd ones: a value written down only when something goes
   * wrong cannot be compared against anything, which is a mistake this project
   * has now made four times with four different fields.
   */
  regime: HostRegime;
}

/**
 * How the host was behaving, in the only terms this run can honestly claim.
 *
 * `collection-refused` is the documented regime flip — the shape collection
 * stops answering, and everything downstream of it degrades. `slide-trouble` is
 * the other known window: freshly added slides stop resolving for ~15 seconds
 * at a time. Neither is a guess about the host's insides; both are things this
 * run watched happen.
 */
export type HostRegime = "healthy" | "slide-trouble" | "collection-refused" | "unknown";

/**
 * How long an observation still describes the host.
 *
 * This host's states come and go in windows of roughly fifteen seconds — a
 * freshly added slide stops resolving, then starts again — so a refusal from a
 * minute ago is history, not a description of now. Twenty seconds is one such
 * window plus a little, and it is the number that turns this field from a latch
 * into a measurement.
 */
export const REGIME_WINDOW_MS = 20_000;

/**
 * What the host was doing, from the most RECENT thing this run watched it do.
 *
 * The first version of this was two sticky booleans: once any shape question
 * answered a refusal, every later sample was stamped `collection-refused` for
 * the rest of the run. The first real round showed what that is worth — the
 * flag latched 8.9 seconds into a 110-second probe and 55 of 65 samples came
 * back carrying the same value. A field that nearly every sample shares cannot
 * separate anything, which is the exact mistake this project has now recorded
 * against `idleMs`, `afterAnswering`, the settle's shared label and
 * `listChartsInDeck` — four fields, four times, and this was the fifth.
 *
 * So: timestamps, a window, and an explicit `unknown` for "nothing recent
 * enough to say". Refusal outranks slide trouble because a host that will not
 * list shapes is in the deeper hole, and both outrank `healthy`, which is only
 * claimed on the strength of something that actually worked.
 */
export function regimeFrom(
  o: { at: number; lastRefusalAt?: number; lastSlideTroubleAt?: number; lastGoodAt?: number },
  windowMs = REGIME_WINDOW_MS,
): HostRegime {
  const recent = (t?: number) => t !== undefined && o.at - t <= windowMs;
  if (recent(o.lastRefusalAt)) return "collection-refused";
  if (recent(o.lastSlideTroubleAt)) return "slide-trouble";
  if (recent(o.lastGoodAt)) return "healthy";
  return "unknown";
}

/** A complete sheet, plus what produced it. */
export interface HostAnswerSheet {
  kind: "powerchart-host-answers";
  /** Which host answered — a real PowerPoint, or the fake in CI. */
  source: string;
  build: string;
  /** Requirement sets this host admits to, since half the answers depend on it. */
  requirementSets: string[];
  answers: HostAnswer[];
}

/**
 * How long a single question may take before it counts as unanswered.
 *
 * Short. A wedged host must cost one question, not the sheet — the sheet's
 * value is that it comes back complete even from a host that is misbehaving,
 * which is exactly when it is most worth having.
 */
let PROBE_BUDGET_MS = 8_000;

/**
 * How long the probe will wait for a scratch slide to land.
 *
 * The per-question budget above bounds the QUESTION and nothing else, and the
 * slide a question needs is acquired outside it — so on 2026-08-10
 * `shape-resolve-held-slide-proxy` took 95.6 seconds against an eight-second
 * budget, because `addScratchSlide` defaults to `READBACK_TIMEOUT_MS`: ninety
 * seconds, sized for a twenty-slide repair page. A bound that something else
 * can walk around is not a bound, and the sheet's whole promise is that a
 * wedged host costs one question rather than the run.
 *
 * Fifteen because the two outcomes are nowhere near each other. Every scratch
 * add that worked in that run took between 0.21 and 4.0 seconds; every one that
 * failed took the full ninety. Fifteen is nearly four times the slowest success
 * and a sixth of the failure, so it cannot turn a slow add into a lost question
 * — and a probe run makes twenty-odd of these.
 */
let SCRATCH_ADD_BUDGET_MS = 15_000;

/**
 * Test-only: shorten the scratch-add budget.
 *
 * The same reason `_setProbeBudgetForTest` exists one screen up — a bound is
 * only reachable by letting an add actually miss it, and fifteen seconds of
 * that per case is longer than any suite should sit still for.
 */
export function _setScratchAddBudgetForTest(ms: number): void {
  SCRATCH_ADD_BUDGET_MS = ms;
}

/**
 * Test-only: shorten the per-question budget.
 *
 * A wedged host is only reachable in a test by letting a question actually miss
 * its deadline, and at eight seconds a full sheet of those outlasts any sane
 * test timeout. Production never touches this; the default is the eight seconds
 * a real host is given. Restore by passing 8_000 back in.
 */
export function _setProbeBudgetForTest(ms: number): void {
  PROBE_BUDGET_MS = ms;
}

/** What a probe does, given a scratch slide it is free to wreck. */
type Probe = {
  id: string;
  question: string;
  /** Answers, or throws — a throw is recorded as an answer, not a failure. */
  ask: (ctx: ProbeContext) => Promise<{ answer: string; detail?: string }>;
  /**
   * This question does not need a LIVE scratch slide, so do not spend one on it.
   *
   * `withProbeContext` checks the scratch slide's liveness before handing a
   * context to any probe, and on this host that check is the single most
   * likely thing to fail — a fresh slide's id resolves once and then stops. So
   * a question that never touches the slide inherits a failure that has nothing
   * to do with it, and the sheet records `no-scratch-slide` where a real answer
   * was available all along.
   *
   * That is not hypothetical: the 2026-08-08 `a546897` round lost
   * `untrack-available` this way, at 43 seconds, AFTER the host had recovered
   * and answered four questions in a row. `untrack` is a `typeof` on a proxy —
   * `getItemOrNullObject` builds one without a round trip — so the question was
   * answerable and the sheet said otherwise. A statement about the probe wearing
   * the clothes of a statement about the host is the error this whole file is
   * built to avoid.
   */
  noSlideNeeded?: true;
  /**
   * Asking this question leaves the scratch slide unusable, so the run takes a
   * fresh one before the NEXT question rather than handing on the wreckage.
   *
   * Earned by seven rounds of evidence and one accidental control. The question
   * after `shape-add-held-slide-proxy` has never been answered — not once in
   * seven sheets — and for six of them that was `shape-resolve-held-slide-proxy`,
   * whose 0/6 everything here (this file's own notes included) put down to its
   * reading ids back off setup shapes. Then a reorder for unrelated reasons put
   * `binding-names-shape-later` in that slot instead, and the slot failed
   * exactly the same way, in 397ms, at the liveness check, before the probe's
   * own code ran at all:
   *
   *   round 12-17  #5 shape-add-held-slide-proxy threw | #6 shape-resolve-…  never put
   *   round 18     #5 shape-add-held-slide-proxy threw | #6 binding-names-…  never put
   *
   * Two different questions, one slot, seven for seven. It is the SLOT, and the
   * cause is what the probe above it is for: `shape-add-held-slide-proxy`
   * deliberately writes through a slide proxy the host has stopped honouring,
   * and this host does not forgive that on the slide afterwards.
   *
   * It ANSWERS, so none of the not-asked machinery notices — the replacement
   * paths all trigger on a question that failed, and this one succeeds and
   * poisons the well on its way out.
   */
  burnsTheSlide?: true;
  /**
   * A second question, asked in the same run, when this one's answer admits
   * two readings.
   *
   * The repo's rule is "when two explanations fit the evidence, ask — do not
   * reason". Following it has meant me writing the partner question, the owner
   * running the probe again, and a session going by — twice, at a cost of two
   * full sheets. The reasoning was never the expensive part; the ROUND TRIP
   * was. A pair that can be decided in one run should be.
   *
   * `when` is what keeps this honest. An unconditional partner is just another
   * probe and belongs in the list; a follow-up earns its place by being worth
   * asking only in the light of a particular answer, and by SAYING which answer
   * that was — `because` goes into the sheet, so a reader sees the pair rather
   * than two unrelated rows.
   */
  follow?: {
    when: (answer: string) => boolean;
    because: string;
    probe: Probe;
  };
  /**
   * Ask this one again even when the run is short of scratch slides.
   *
   * The shortlist the later passes fall back to under pressure: the questions
   * whose answers this project does NOT yet trust — everything in
   * `UNSTABLE_ANSWERS` (answered differently on different runs of one build)
   * and everything in `PENDING_QUESTIONS` (never answered at all). Marked on
   * the question rather than kept as a list somewhere else, so it cannot drift
   * from the question it describes; `test/host-probe.test.ts` reads those two
   * tables out of `scripts/host-baseline.mjs` and fails if the marks and the
   * tables disagree, which is what keeps them honest in the other direction.
   */
  resample?: true;
};

/**
 * Answers that mean "this question was never put", not "the host said so".
 *
 * None of them is in any probe's own vocabulary, which is the point: a diff
 * that mistook one for an answer would report a host divergence that nobody
 * ever asked about. They were earned the same way, rounds apart — see
 * `ProbeSetupFailed`.
 *
 * EXPORTED because `NEVER_ASKED` in `scripts/host-baseline.mjs` is the same
 * vocabulary, read by the diff tool — and a comment promising the two are kept
 * in step is not a mechanism. `not-asked` was added here and not there, so every
 * question the mute breaker abandons was being compared against the fake and
 * reported as a real-host divergence. `test/host-probe.test.ts` asserts the two
 * sets are equal now.
 */
export const NOT_ASKED = new Set(["no-scratch-slide", "no-scratch-shape", "not-asked"]);

/**
 * How many times a run asks each question.
 *
 * Three, and the number is the whole point of the exercise: one sample cannot
 * say whether an answer is stable, two cannot say which of a disagreeing pair is
 * the odd one, and three is where a round starts producing the fact that
 * `UNSTABLE_ANSWERS` was built by hand across ten rounds to produce.
 *
 * Spread across the run, not asked back to back — three asks inside four
 * seconds sample one minute of the host's life three times, which is one
 * sample wearing a hat. Pass 1 is the whole list, exactly as before, so a host
 * that dies early still yields today's sheet and nothing regresses.
 */
export const PROBE_PASSES = 3;

/**
 * Whether a question's samples agree — `undefined` when it has fewer than two
 * real ones to compare.
 *
 * Only REAL answers count. A question put once and refused twice is not
 * unstable; it is a question this run mostly could not ask, and calling that
 * instability would manufacture exactly the noise `UNSTABLE_ANSWERS` warns
 * against. Pure, and exported, so the rule can be checked without a host.
 */
export function stabilityOf(samples: ProbeSample[] | undefined): boolean | undefined {
  const real = (samples ?? []).map((x) => x.answer).filter((a) => !NOT_ASKED.has(a));
  return real.length >= 2 ? new Set(real).size === 1 : undefined;
}

/**
 * The share of a pass that may come back never-asked before the later passes
 * stop being worth their scratch slides.
 *
 * Scratch slides are the scarcest thing on this host and the repeats spend
 * three times as many of them. When a pass is already losing a third of its
 * questions to `no-scratch-slide`, asking the settled ones twice more is
 * bidding for slides against the questions that actually need a second sample —
 * so the later passes fall back to the resample shortlist, and the run says so
 * in its trace rather than quietly doing less.
 */
export const PASS_PRESSURE_LIMIT = 1 / 3;

/**
 * Consecutive unanswered questions before the probe stops asking.
 *
 * Three, for the reason the self-test's SICK_LIMIT is three: one is often the
 * finding, and three in a row is a finding about the host rather than about a
 * question.
 */
const PROBE_MUTE_LIMIT = 3;

/**
 * Thrown when a probe could not get the shapes its question is about.
 *
 * The same lesson as `ScratchSlideUnavailable`, one layer down, and it cost a
 * second real answer sheet to learn. That guard covers a scratch slide the host
 * will not resolve; it says nothing about a slide that resolves perfectly and
 * then refuses to take a shape. PowerPoint on the web did exactly that on
 * 2026-08-04: six questions that only read were answered, and all eight that
 * needed a shape failed in setup — `GeneralException` at the add, or a sync
 * that never came back. Every one of the eight was recorded as `"threw"` or
 * `"silent"`, both legitimate answers to those questions, so `host-diff`
 * reported eight host divergences from a sheet that had asked six questions.
 *
 * A probe that never reached its question must not answer it.
 */
class ProbeSetupFailed extends Error {
  constructor(readonly why: string) {
    super(`the probe never got as far as its question: ${why}`);
    this.name = "ProbeSetupFailed";
  }
}

/**
 * Whether the sync that was to deliver a probe's setup shapes is still out.
 *
 * A setup failure that THROWS is caught where it happens; a setup failure that
 * WEDGES is caught by the probe budget, which fires in `ask` — outside any
 * knowledge of what the probe was in the middle of. Three of that real sheet's
 * eight failures were the wedging kind, and `"silent"` is a comparable answer,
 * so they read as host divergences too.
 *
 * Module-level because probes run strictly one at a time; `ask` clears it
 * before every question.
 */
let awaitingSetupShapes = false;

/**
 * A 1x1 transparent PNG — the smallest thing that is genuinely a picture.
 *
 * The probe that asks about office-js#5022 needs a real image and cares nothing
 * about what it looks like. Bare base64, no `data:` prefix, because that is what
 * `fill.setImage` takes.
 */
const ONE_PIXEL_PNG =
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==";

/** A shape's box, in points on the scratch slide. */
type ProbeBox = { left: number; top: number; width: number; height: number };

/**
 * Put shapes on the scratch slide, or admit that this host would not.
 *
 * Through a slide proxy resolved in THIS sync and used before the next one —
 * `ProbeContext.scratch` is a thunk for that reason, and holding what it
 * returns re-opens the trap it exists to close.
 *
 * `load` names a real property to read back off the shapes, and it costs a
 * SECOND sync — add, commit, then ask.
 *
 * It used to be queued in the same batch as the add, to save a round trip. That
 * saving is why five questions on the 2026-08-05 sheet were never put: the five
 * that call `idsOf`, and no others. Perfect correlation, and no other property
 * of those five explains it — batch count does not (`shapes-items-count-honest`
 * and `delete-then-lookup` each span two batches and were answered, while three
 * of the five span two and failed). The implied host fact is narrow and
 * plausible: **this PowerPoint will not name a shape in the batch that created
 * it.**
 *
 * Asking after the commit costs one round trip per id-needing question and
 * makes them answerable at all. It ages the returned proxies by a sync, which
 * is free here: every probe that asks for ids goes on to work through
 * `probeShape(ctx, id)`, so the proxies are never used again — only the ids.
 * The probes that do NOT ask for ids are untouched, and still measure exactly
 * what they measured before.
 */
async function scratchShapes(ctx: ProbeContext, boxes: ProbeBox[], load?: string): Promise<PowerPoint.Shape[]> {
  // WHICH of the two syncs failed, in the reason the sheet carries.
  //
  // Both were wrapped in one `catch` and both produced the same
  // `no-scratch-shape`, which is two different diagnoses wearing one string —
  // the mistake `ProbeSetupFailed` itself exists to stop, made one level down.
  // Adding shapes and NAMING them are separate abilities and this host has
  // them separately: `shape-add-fresh-slide-proxy` says yes to the add, and
  // `shape-proxy-survives-one-sync` says `unreadable` to reading anything back
  // through a proxy a sync later.
  //
  // It matters because six of the seven questions this host has never answered
  // are exactly the six that pass a `load` here, and every probe that calls
  // this WITHOUT one gets an answer. That correlation is the whole lead, and
  // until now it could only be read across probes rather than recorded in one.
  let phase = "adding the setup shapes";
  try {
    const shapes = ctx.scratch().shapes;
    const made = boxes.map((box) => shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle, box));
    awaitingSetupShapes = true;
    await ctx.sync();
    awaitingSetupShapes = false;
    if (load) {
      phase = `reading "${load}" back off the setup shapes`;
      for (const s of made) s.load(load);
      awaitingSetupShapes = true;
      await ctx.sync();
      awaitingSetupShapes = false;
    }
    return made;
  } catch (err) {
    awaitingSetupShapes = false;
    throw new ProbeSetupFailed(`${phase}: ${short(err)}`);
  }
}

/**
 * A shape on the scratch slide, resolved in the batch that is about to use it.
 *
 * The rule that makes the answers below mean anything: **only an id crosses a
 * sync, never a handle.** Four questions on the 2026-08-04 sheet came back with
 * answers that were about the probe rather than the host — a tag written
 * through a group proxy a sync old read back undefined, and a group whose
 * children were counted a sync after it was made reported none. Taken at face
 * value those say no chart in any deck is re-editable, which the same run
 * disproves: its repair pass landed 23 retags on grouped charts.
 *
 * A question whose setup trips the one rule this host is strictest about
 * measures the trip, not the question.
 */
/**
 * The scratch slide's shape collection, resolved in the batch about to use it.
 *
 * Same guard as `probeShape`, for the probes that reach for the collection
 * rather than one shape: a slide that has stopped answering hands back a null
 * object, and `.shapes` off one of those fails in whatever way the host
 * chooses. None of those ways is an answer to the question being asked.
 */
function probeShapes(ctx: ProbeContext): PowerPoint.ShapeCollection {
  let shapes: PowerPoint.ShapeCollection | undefined;
  try {
    shapes = ctx.scratch().shapes;
  } catch (err) {
    throw new ProbeSetupFailed(`the scratch slide stopped answering: ${short(err)}`);
  }
  // Absent, not throwing — the shape of it on a slide the host has stopped
  // resolving, and the more dangerous half: the failure then surfaces as
  // whatever the NEXT line does to `undefined`, which every probe's catch
  // faithfully records as the host having thrown.
  if (!shapes) throw new ProbeSetupFailed("the scratch slide has no shape collection");
  return shapes;
}

function probeShape(ctx: ProbeContext, id: string): PowerPoint.Shape {
  try {
    return ctx.scratch().shapes.getItemOrNullObject(id);
  } catch (err) {
    // The slide stopped answering PART WAY THROUGH a question. Not this
    // probe's answer either — `ScratchSlideUnavailable` only guards the start
    // of a probe's context, and a question that loses its slide at its third
    // batch has asked no more than one that lost it at its first.
    throw new ProbeSetupFailed(`the scratch slide stopped answering: ${short(err)}`);
  }
}

/** The 1.8 binding collection, as much of it as one probe needs. */
type BindingsLike = {
  add(shape: PowerPoint.Shape, bindingType: string, id: string): unknown;
  getItemOrNullObject(id: string): { getShape(): PowerPoint.Shape };
};

/**
 * `presentation.bindings`, or undefined on a host that does not have it.
 *
 * Feature-detected rather than read off `isSetSupported`. The requirement set a
 * host ADMITS to and the API surface it actually exposes are two facts, this
 * file exists because they have come apart before, and the sheet already
 * records the admitted list beside every answer — so a reader can tell "1.8 is
 * missing" from "1.8 is claimed and the object is not there".
 */
function bindingsOf(ctx: ProbeContext): BindingsLike | undefined {
  try {
    const b = (ctx.presentation as unknown as { bindings?: BindingsLike }).bindings;
    return b && typeof b.add === "function" && typeof b.getItemOrNullObject === "function" ? b : undefined;
  } catch {
    return undefined;
  }
}

/**
 * A probe's own catch, for the answers that are real answers.
 *
 * Every probe below turns a throw into `"threw"` or `"unreadable"`, which are
 * things this host genuinely says. A setup failure must not be dressed in
 * either — it is the whole reason `no-scratch-shape` exists — so it goes back
 * up to `ask`, where the run replaces the scratch slide and puts the question
 * again.
 */
function threw(err: unknown): { answer: string; detail?: string } {
  if (err instanceof ProbeSetupFailed) throw err;
  return { answer: "threw", detail: short(err) };
}

/** A shape's id, or undefined if the host did not answer the load. */
function readShapeId(shape: { id?: string }): string | undefined {
  try {
    const v = shape.id;
    return typeof v === "string" && v ? v : undefined;
  } catch {
    return undefined;
  }
}

/** A queued `getCount()` result, or undefined if the sync did not populate it. */
function loadedCount(result: { value: number }): number | undefined {
  try {
    const v = result.value;
    return typeof v === "number" ? v : undefined;
  } catch {
    return undefined;
  }
}

/** A slide's id, or undefined if the host did not answer the load. */
function readId(slide: PowerPoint.Slide): string | undefined {
  try {
    const v = (slide as unknown as { id?: string }).id;
    return typeof v === "string" && v ? v : undefined;
  } catch {
    return undefined;
  }
}

/** The same, for probes whose failure vocabulary is `"unreadable"`. */
function unreadable(err: unknown): { answer: string; detail?: string } {
  if (err instanceof ProbeSetupFailed) throw err;
  return { answer: "unreadable", detail: short(err) };
}

/**
 * The ids of shapes just made — the only thing allowed to outlive their sync.
 *
 * A host that will not read an id back has not answered the question either,
 * so this is a setup failure rather than a finding: `no-scratch-shape`, in the
 * vocabulary no probe can produce as its own answer.
 */
/**
 * Two shapes to group, and HOW they were obtained.
 *
 * The three group questions were unanswerable on the one host that matters, and
 * `scratchShapes`' own comment had already named the lead: the six questions
 * this host never answers are exactly the six that pass a `load`. They add
 * shapes, sync, read the ids back, and it is that read the host refuses — the
 * same refusal behind `shape-proxy-survives-one-sync: unreadable` and the empty
 * collection reads everywhere else. The questions were never reached. They died
 * in their own setup and reported `no-scratch-slide`, which is a statement about
 * the probe, not the host.
 *
 * So: try the strict route first, and fall back to grouping the proxies in the
 * batch that created them — no id, no sync in between, nothing crossing.
 *
 * The fallback is a DIFFERENT question and must never be allowed to look like
 * the same one. Grouping members resolved by id is what production does;
 * grouping same-batch proxies is the friendliest case a host could be given. An
 * answer from the second says nothing about the first. Hence `via`, which every
 * caller puts in its `detail` — a sheet that says `two (members via same-batch)`
 * and one that says `two (members via ids)` record two different facts, and the
 * diff can tell them apart.
 */
async function groupMembers(
  ctx: ProbeContext,
  boxes: ProbeBox[],
): Promise<{ members: PowerPoint.Shape[]; via: "ids" | "same-batch" }> {
  try {
    const ids = idsOf(await scratchShapes(ctx, boxes, "id"));
    if (ids.length === boxes.length) return { members: ids.map((id) => probeShape(ctx, id)), via: "ids" };
  } catch (err) {
    // Only a SETUP failure earns the fallback. Anything else is the host
    // answering the question and must travel on unchanged.
    if (!(err instanceof ProbeSetupFailed)) throw err;
  }
  // No `load`, so nothing is read back and nothing crosses a sync: the members
  // are the proxies `addGeometricShape` just returned, used in their own batch.
  return { members: await scratchShapes(ctx, boxes), via: "same-batch" };
}

function idsOf(made: PowerPoint.Shape[]): string[] {
  const ids = made.map((s) => {
    try {
      const v = (s as unknown as { id?: string }).id;
      return typeof v === "string" && v ? v : undefined;
    } catch {
      return undefined;
    }
  });
  if (ids.some((id) => !id)) throw new ProbeSetupFailed("the host would not read back the shapes' ids");
  return ids as string[];
}

/**
 * The questions.
 *
 * Each one corresponds to something `powerpoint.ts` relies on. That is the bar
 * for adding one: not "this is interesting about Office.js" but "if the answer
 * were different, code in this repo would be wrong". A probe that no code
 * depends on is a fact nobody will act on.
 */
const PROBES: Probe[] = [
  {
    id: "load-isnullobject-populates",
    question: "Does load('isNullObject') make isNullObject readable after a sync?",
    // The assumption behind `queueNullCheck`, and it is a NEGATIVE one: this
    // file loads "id" instead precisely because loading the flag by name does
    // not populate it. If a host answers yes, `queueNullCheck`'s whole comment
    // is wrong for that host.
    ask: async (ctx) => {
      const slide = ctx.slides.getItemOrNullObject(ctx.scratchId);
      slide.load("isNullObject");
      await ctx.sync();
      try {
        const v = (slide as unknown as { isNullObject: boolean }).isNullObject;
        return { answer: typeof v === "boolean" ? "yes" : "unreadable", detail: `read ${String(v)}` };
      } catch (err) {
        return unreadable(err);
      }
    },
  },
  {
    id: "load-id-populates-isnullobject",
    question: "Does loading a REAL property make isNullObject readable?",
    // The positive half, and what `queueNullCheck` actually does.
    ask: async (ctx) => {
      const slide = ctx.slides.getItemOrNullObject(ctx.scratchId);
      slide.load("id");
      await ctx.sync();
      try {
        const v = (slide as unknown as { isNullObject: boolean }).isNullObject;
        return { answer: typeof v === "boolean" ? "yes" : "unreadable", detail: `read ${String(v)}` };
      } catch (err) {
        return unreadable(err);
      }
    },
  },
  {
    id: "getitemornullobject-missing",
    // Resolves a deliberately bogus id. Needing a real slide to ask about a missing one is a contradiction.
    noSlideNeeded: true,
    question: "What does getItemOrNullObject give for an id that does not exist?",
    ask: async (ctx) => {
      const slide = ctx.slides.getItemOrNullObject("powerchart-no-such-slide");
      slide.load("id");
      await ctx.sync();
      try {
        const nul = (slide as unknown as { isNullObject: boolean }).isNullObject;
        return { answer: nul === true ? "null-object" : nul === false ? "claims-live" : "unreadable" };
      } catch (err) {
        return threw(err);
      }
    },
  },
  {
    id: "shape-add-fresh-slide-proxy",
    question: "Can a shape be added through a slide proxy resolved in THIS sync?",
    // The three questions below exist because a real host refused eight others
    // and the sheet could not say which of two things it meant: that this host
    // will not take a shape on a freshly-added slide at all, or that it will
    // not take one through a slide proxy it resolved a sync earlier. Both
    // explain every failure in that sheet; they call for different code. So ask
    // the three ways apart, and let the next sheet say which.
    ask: async (ctx) => {
      try {
        probeShapes(ctx).addGeometricShape(PowerPoint.GeometricShapeType.rectangle, {
          left: 10,
          top: 100,
          width: 20,
          height: 20,
        });
        await ctx.sync();
        return { answer: "yes" };
      } catch (err) {
        return threw(err);
      }
    },
  },
  {
    id: "shape-add-held-slide-proxy",
    resample: true,
    question: "Can a shape be added through a slide proxy resolved a sync ago?",
    // Writing through a handle the host has stopped honouring is this question,
    // and it costs the slide. Seven rounds say whoever follows never gets an
    // answer — see `Probe.burnsTheSlide`.
    burnsTheSlide: true,
    // What `withProbeContext` used to hand every probe, and what the whole
    // add-in avoids on freshly-added slides through `SlideThunk`: Office.js
    // rewrites a resolved proxy's object path to `getItem(id)`, and a new
    // slide's id does not round-trip through `getItem` on the web.
    ask: async (ctx) => {
      const held = ctx.scratch();
      held.load("id"); // a REAL property: this is the sync that resolves it
      await ctx.sync();
      try {
        held.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle, {
          left: 40,
          top: 100,
          width: 20,
          height: 20,
        });
        await ctx.sync();
        return { answer: "yes" };
      } catch (err) {
        return threw(err);
      }
    },
    // THE PARTNER: the same question again, immediately, on a second fresh slide.
    //
    // This question is the one open lead on this host. Sorted by which pass
    // asked it, the observations separate almost perfectly — pass 1 gave
    // `threw` seventeen times out of eighteen, every later pass gave `yes`
    // three times out of three — and that second row did not exist until
    // `PROBE_PASSES` shipped, so sixteen earlier rounds were sixteen samples of
    // ONE condition rather than sixteen tosses of a coin.
    //
    // Two readings fit that, and they want opposite things from this project:
    //
    //   the MOMENT   the host's answer is a property of when it is asked, so
    //                there is a state to find and `yes` is reachable on purpose.
    //   a COIN       it lands differently every time, the pass split is luck,
    //                and no amount of re-running whole rounds will ever settle it.
    //
    // A third reading — the scratch slide's own age — is already dead, and from
    // data rather than argument: the 2026-08-11 run log shows the recorded
    // answer came from a brand-new REPLACEMENT slide in all three passes (first
    // attempt `no-scratch-slide`, replacement added, retry answered), so slide
    // age was identical while the answer changed.
    //
    // Asking again at the same instant is what separates the two that are left.
    // Same second of the run, same host state, same brand-new slide, one
    // difference: this is a second toss. Agreement says the host has a definite
    // behaviour right now and the variable is the moment; disagreement says it
    // is a coin and the populations above are an artefact of when we happened to
    // look. Either way one round decides it, which is the whole point of
    // `Probe.follow` — the reasoning was never the expensive part, the round
    // trip was.
    //
    // Fires on BOTH real answers on purpose. A partner that only followed `yes`
    // would sample the coin exactly where the old design already sampled it.
    follow: {
      when: (answer) => answer === "yes" || answer === "threw",
      because: "asking the identical question a second time at the same instant is what tells a host state from a coin",
      probe: {
        id: "shape-add-held-slide-proxy-again",
        question: "Asked a second time moments later on another fresh slide, does it answer the same way?",
        // No `resample` mark, and it does not need one: a follow-up is never
        // scheduled on its own, it rides its trigger, and the trigger carries
        // the mark — which is how this pair got three paired samples in one
        // round. It carried the mark while it sat in `PENDING_QUESTIONS`; it
        // has been answered now (2026-08-11, `756682e`, stable across three
        // passes) and the shortlist invariant would fail on a mark neither
        // table asks for.
        // Same damage as its trigger — it writes through a proxy this host has
        // stopped honouring — so it gives up its slide too rather than handing
        // the wreckage to whatever runs next. See `Probe.burnsTheSlide`.
        burnsTheSlide: true,
        // Deliberately byte-for-byte the trigger's question. A partner that
        // differed in any detail would be a second question, and the answer
        // would no longer be about the host's consistency.
        ask: async (ctx) => {
          const held = ctx.scratch();
          held.load("id");
          await ctx.sync();
          try {
            held.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle, {
              left: 40,
              top: 100,
              width: 20,
              height: 20,
            });
            await ctx.sync();
            return { answer: "yes" };
          } catch (err) {
            return threw(err);
          }
        },
      },
    },
  },
  {
    id: "binding-names-shape-later",
    resample: true,
    question: "Can a binding made in a shape's CREATING sync still name that shape afterwards?",
    // ASKED SIXTH ON PURPOSE, beside the questions its control arm duplicates.
    //
    // It sat at position 20 for five rounds and was answered in none of them.
    // That is not bad luck, and six sheets say what it is. Sorting every
    // question by whether it needs a SHAPE and by where it runs:
    //
    //   needs no shape, any position  43/48  90%
    //   needs a shape, positions 1-8  23/30  77%
    //   needs a shape, positions 9+   36/77  47%
    //
    // Two effects and both are large: needing a shape costs you, and needing one
    // LATE costs you again. This question needed a shape at position 20, which
    // is the worst quadrant there is.
    //
    // Position alone was not the whole story, and the round after this moved
    // found out why in the sharpest possible way: it landed in the one slot on
    // the sheet that has never produced an answer — the one directly after
    // `shape-add-held-slide-proxy`, which wrecks the scratch slide on its way
    // out. `shape-resolve-held-slide-proxy`'s 0/6 was blamed here on its reading
    // ids back off setup shapes; it was the neighbour. See `Probe.burnsTheSlide`,
    // which this move is what uncovered, and which now takes a fresh slide in
    // between so the slot is worth having.
    //
    // The only untried route out of the failure that costs this project the most.
    //
    // `same scale across the deck` fails the same way every round: the chart
    // DRAWS, and then the host refuses every handle to it —
    // `InvalidParam passed to GetItem(id)`, code 5010, at
    // `ShapeCollection.getItem` — so there is no group, no config tag, and no id
    // to settle one onto later. On 2026-08-09 that was five charts of eight, each
    // leaving 24 shapes on a slide that is no longer a chart. Both handles the
    // settle pass has are refused: `shapes-items-count-honest` says the
    // collection reads back empty, and `shapes-items-via-positional-slide` says
    // renaming the parent does not help.
    //
    // A binding is the one reference that never crosses either. `bindings.add`
    // takes the live Shape proxy in the batch that CREATED it — no id round trip,
    // no collection read — and the document persists it, so a later, healthier
    // context can ask for the shape back by a key we chose ourselves. If that
    // works here, the repair pass gets the handle it currently lacks and a lost
    // tag becomes a repairable one rather than a chart the user cannot edit.
    //
    // PowerPointApi 1.8 (GA), and this host reports 1.10 — but feature-detected
    // rather than gated on `isSetSupported`, which is a claim about the host
    // rather than a look at it.
    //
    // Deliberately asked across a sync boundary: resolving the binding in the
    // batch that made it would answer a question nobody has, since the live
    // proxy is right there. The later sync IS the question.
    ask: async (ctx) => {
      const bindings = bindingsOf(ctx);
      if (!bindings) return { answer: "no-binding-api", detail: "presentation.bindings is absent (needs 1.8)" };
      // A fixed key on purpose: the docs say a repeat id overwrites its binding,
      // so a run cannot accumulate them, and deleting the scratch slide's shape
      // takes the binding with it.
      const key = "POWERCHART_PROBE_BINDING";
      const box = (left: number) => ({ left, top: 10, width: 20, height: 20 });
      // The CONTROL arm, and the whole reason this question is answerable.
      //
      // The binding has to be made in the batch that CREATES the shape — a
      // proxy one sync old is refused here (`shape-proxy-survives-one-sync`), so
      // binding a committed shape would measure staleness instead. That leaves
      // one batch carrying two things, and its refusal attributable to neither.
      //
      // The 2026-08-09 evening round is exactly that dead end: the commit came
      // back `UnexpectedError` in 1.3 seconds and this probe honestly reported
      // "never asked". `shape-add-fresh-slide-proxy` answered `yes` in the same
      // sheet, which points hard at the binding — but that is a different
      // question asked at a different moment on a host that flaps between
      // minutes, and inferring across two of those has cost this project two
      // full sheets already.
      //
      // So the probe carries its own control: the identical batch WITHOUT the
      // binding, on the same slide, seconds earlier. If that commits and the
      // bound one does not, the binding is the only difference and the answer
      // is about bindings. If the control fails, this host is not taking shapes
      // right now and the question was never put — which is the truth, said by
      // the probe rather than assembled by a reader.
      try {
        probeShapes(ctx).addGeometricShape(PowerPoint.GeometricShapeType.rectangle, box(290));
        await ctx.sync();
      } catch (err) {
        throw new ProbeSetupFailed(`the control shape would not commit: ${short(err)}`);
      }
      let shape: PowerPoint.Shape;
      try {
        shape = probeShapes(ctx).addGeometricShape(PowerPoint.GeometricShapeType.rectangle, box(320));
      } catch (err) {
        throw new ProbeSetupFailed(`adding the shape to bind: ${short(err)}`);
      }
      let made: unknown;
      try {
        made = bindings.add(shape, "Shape", key);
      } catch (err) {
        // Synchronous, before any round trip, so this is `bindings.add` itself
        // objecting to the call — a genuine answer about the binding API.
        return { answer: "add-threw", detail: short(err) };
      }
      try {
        await ctx.sync();
      } catch (err) {
        // An ANSWER now, not a setup failure: the same batch minus the binding
        // committed on this slide moments ago.
        return {
          answer: "commit-threw",
          detail: `the same batch without a binding committed seconds earlier: ${short(err)}`,
        };
      }
      if (!made) return { answer: "add-returned-nothing" };
      try {
        const got = bindingsOf(ctx)!.getItemOrNullObject(key).getShape();
        got.load("id");
        await ctx.sync();
        const id = readShapeId(got as unknown as { id?: string });
        return id ? { answer: "yes", detail: `shape id=${id}` } : { answer: "unreadable" };
      } catch (err) {
        return threw(err);
      }
    },
  },
  {
    id: "shape-resolve-held-slide-proxy",
    resample: true,
    question: "Can a shape be RESOLVED (not added) through a slide proxy resolved a sync ago?",
    // The half of the held-handle rule nobody has asked about, and three
    // production sites rest on it: `deleteShapesById`, `setShapeSelection` and
    // the selection path all resolve a slide, sync, and then reach through that
    // same handle for `shapes.getItemOrNullObject(...)`.
    //
    // `shape-add-held-slide-proxy` proves the host refuses a WRITE through such
    // a handle, and the error it gave — `errorLocation: SlideCollection.getItem`
    // — points at the slide lookup rather than at the add, which would mean
    // reads fail too. That is a reading, not an answer: every read the last
    // sheet got right used a handle of its own. So ask, because the answer
    // decides whether those three sites are bugs or merely untidy.
    ask: async (ctx) => {
      const [id] = idsOf(await scratchShapes(ctx, [{ left: 10, top: 140, width: 20, height: 20 }], "id"));
      const held = ctx.scratch();
      held.load("id"); // a REAL property: this is the sync that resolves it
      await ctx.sync();
      try {
        const shape = held.shapes.getItemOrNullObject(id);
        shape.load("id");
        await ctx.sync();
        const back = (shape as unknown as { id: string }).id;
        return { answer: back === id ? "yes" : "unreadable", detail: `read ${String(back)}` };
      } catch (err) {
        return threw(err);
      }
    },
  },
  {
    id: "shape-add-fresh-getitem-slide",
    question: "Can a shape be added through slides.getItem(id) on a freshly-added slide?",
    // The fourth way of naming the same slide, and the one every other question
    // here quietly avoids: all seventeen resolve through `getItemOrNullObject`,
    // so `getItem` has never been asked about in either direction.
    //
    // It is not academic. `getTargetSlide` — the only `slides.getItem(id)` in
    // `powerpoint.ts` — is how `insertSceneIntoSlide` finds the slide it draws
    // on, and it holds that handle for the whole draw. Every held-handle
    // failure this host has reported names `errorLocation:
    // SlideCollection.getItem`, which reads as though this must fail; but the
    // failures were all through handles resolved a sync earlier, and a fresh
    // `getItemOrNullObject` on the same slide works. Those are different
    // claims, and the difference decides whether that insert path is broken.
    ask: async (ctx) => {
      const slides = ctx.slides as unknown as { getItem(id: string): PowerPoint.Slide };
      try {
        const byGetItem = slides.getItem(ctx.scratchId);
        byGetItem.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle, {
          left: 100,
          top: 100,
          width: 20,
          height: 20,
        });
        await ctx.sync();
        return { answer: "yes" };
      } catch (err) {
        return threw(err);
      }
    },
    // The partner, and the one pair this project has already paid for twice.
    //
    // A refusal above has two readings and they lead opposite ways: `getItem`
    // cannot name a FRESH slide (so `insertSceneIntoSlide`'s held handle is
    // fine on the slides users actually edit, and the risk is confined to the
    // one caller that passes a new id), or `getItem` is broken on this host
    // full stop (so the everyday insert path is broken for everyone). Nothing
    // in a sheet can tell those apart, and reasoning about which cost two
    // sessions.
    //
    // Read-only, and that is not a detail. The obvious control — repeat the
    // shape ADD on a pre-existing slide — would draw on a slide in the owner's
    // own presentation, and a diagnostic that litters someone's deck is one
    // they stop clicking. Resolving the slide and reading its id back proves
    // the same thing about `getItem` and changes nothing.
    follow: {
      when: (answer) => answer !== "yes",
      because: "getItem refused a freshly-added slide, which could be about getItem or about the slide's newness",
      probe: {
        id: "getitem-durable-slide",
        question: "Does slides.getItem(id) resolve a slide that was in the deck before this run?",
        ask: async (ctx) => {
          const durable = ctx.durableSlideId;
          if (!durable) return { answer: "no-durable-slide", detail: "the deck has no slide this run did not add" };
          const slides = ctx.slides as unknown as { getItem(id: string): PowerPoint.Slide };
          try {
            const slide = slides.getItem(durable);
            slide.load("id");
            await ctx.sync();
            const v = readId(slide);
            return v === durable
              ? { answer: "yes" }
              : { answer: "unreadable", detail: `read back ${String(v)} for ${durable}` };
          } catch (err) {
            return threw(err);
          }
        },
      },
    },
  },
  {
    id: "shape-add-positional-slide-proxy",
    resample: true,
    question: "Can a shape be added through slides.getItemAt(index) rather than by id?",
    // The other half of the same fork. If by-index works where by-id does not,
    // the id is what this host will not take, and every write path that names a
    // freshly-added slide by id needs an index instead.
    ask: async (ctx) => {
      ctx.slides.load("items/id");
      await ctx.sync();
      let index: number;
      try {
        index = ctx.slides.items.findIndex((s) => s.id === ctx.scratchId);
      } catch (err) {
        return unreadable(err);
      }
      // A real fact about this host, not a failure to ask: it listed its slides
      // and the scratch slide was not among them.
      if (index < 0) return { answer: "not-listed" };
      try {
        ctx.slides.getItemAt(index).shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle, {
          left: 70,
          top: 100,
          width: 20,
          height: 20,
        });
        await ctx.sync();
        return { answer: "yes" };
      } catch (err) {
        return threw(err);
      }
    },
  },
  {
    id: "shape-proxy-survives-one-sync",
    question: "Is a shape proxy still usable one sync after it was created?",
    // office-js#2903 — the stale-proxy bug the whole `targetRef` design exists
    // for. If a host keeps proxies alive, a lot of re-fetching here is waste.
    ask: async (ctx) => {
      const [shape] = await scratchShapes(ctx, [{ left: 10, top: 10, width: 20, height: 20 }]);
      await ctx.sync(); // a second round trip: this is what ages the proxy
      try {
        shape.load("id");
        await ctx.sync();
        const id = (shape as unknown as { id: string }).id;
        return { answer: typeof id === "string" && id ? "yes" : "unreadable" };
      } catch (err) {
        return threw(err);
      }
    },
  },
  {
    id: "shapes-items-count-honest",
    resample: true,
    question: "After adding 5 shapes, how many does the collection report?",
    // `faults.hollowReads` models a host that answers SHORT without throwing —
    // observed asking about 19 shapes and being told 3. If a host is honest
    // here, the readback paging is cheaper than it needs to be.
    ask: async (ctx) => {
      await scratchShapes(
        ctx,
        Array.from({ length: 5 }, (_, i) => ({ left: i * 5, top: 40, width: 4, height: 4 })),
      );
      const shapes = probeShapes(ctx);
      shapes.load("items/id");
      await ctx.sync();
      try {
        const n = shapes.items.length;
        // Reported as "at least 5" rather than an exact count: the slide
        // carries whatever earlier probes left on it, and the question is
        // whether the host UNDER-reports, not what the total happens to be.
        return { answer: n >= 5 ? "at-least-5" : `short-${n}`, detail: `items=${n}` };
      } catch (err) {
        return unreadable(err);
      }
    },
  },
  {
    id: "shapes-items-via-positional-slide",
    resample: true,
    question: "Same collection read, but through slides.getItemAt(index) — any different?",
    // The one contaminated answer that could NOT be cleaned up by re-resolving,
    // and the reason it needs a pair instead.
    //
    // A collection read cannot avoid crossing a sync: the load is queued in one
    // batch and the answer read in the next, by definition. So when this host
    // answered the question above with `items` UNDEFINED, there was no way to
    // tell "this host under-reports collections" from "the by-id slide handle
    // the collection hangs off was spent by the very sync that answered it".
    //
    // The three shape-add questions already proved a positional handle works
    // where a held by-id one does not. If this reads back and the one above
    // does not, the collection is fine and the parent handle was the whole
    // story — which is a fact about every readback in `powerpoint.ts`, since
    // the repair pass reads by index and the probe reads by id.
    ask: async (ctx) => {
      // Its own five, not the previous question's. Probes share a scratch
      // slide, so leaning on what an earlier one left would make this answer
      // depend on whether THAT one's setup worked — and the run replaces a
      // scratch slide it loses, so it may not even be the same slide.
      await scratchShapes(
        ctx,
        Array.from({ length: 5 }, (_, i) => ({ left: i * 5, top: 120, width: 4, height: 4 })),
      );
      ctx.slides.load("items/id");
      await ctx.sync();
      let index: number;
      try {
        index = ctx.slides.items.findIndex((s) => s.id === ctx.scratchId);
      } catch (err) {
        return unreadable(err);
      }
      if (index < 0) return { answer: "not-listed" };
      const shapes = ctx.slides.getItemAt(index).shapes;
      shapes.load("items/id");
      await ctx.sync();
      try {
        const n = shapes.items.length;
        return { answer: n >= 5 ? "at-least-5" : `short-${n}`, detail: `items=${n}` };
      } catch (err) {
        return unreadable(err);
      }
    },
  },
  {
    id: "getcount-populates-same-sync",
    // `slides.getCount()` on the deck. No slide of ours involved.
    noSlideNeeded: true,
    question: "Does getCount()'s value arrive on the sync that queued it?",
    ask: async (ctx) => {
      const count = ctx.slides.getCount();
      await ctx.sync();
      try {
        const v = count.value;
        return { answer: typeof v === "number" ? "yes" : "unreadable", detail: `value=${String(v)}` };
      } catch (err) {
        return unreadable(err);
      }
    },
  },
  {
    id: "tags-add-same-key-twice",
    question: "Does writing the same tag key twice overwrite, or duplicate?",
    // Re-editing a chart rewrites POWERCHART_CONFIG on the same shape every
    // time. If a host appended instead of overwriting, a chart edited ten times
    // would carry ten configs and the reader would pick one arbitrarily.
    // Every write goes through a shape resolved in its own batch. The first
    // version held one proxy across all four syncs, and this host answered
    // "other — value=undefined": not an opinion about tag keys at all, just
    // the stale handle refusing both writes.
    ask: async (ctx) => {
      const [id] = idsOf(await scratchShapes(ctx, [{ left: 60, top: 10, width: 20, height: 20 }], "id"));
      probeShape(ctx, id).tags.add("POWERCHART_PROBE", "first");
      await ctx.sync();
      probeShape(ctx, id).tags.add("POWERCHART_PROBE", "second");
      await ctx.sync();
      const tag = probeShape(ctx, id).tags.getItemOrNullObject("POWERCHART_PROBE");
      tag.load("value");
      await ctx.sync();
      try {
        const v = (tag as unknown as { value: string }).value;
        return {
          answer: v === "second" ? "overwrites" : v === "first" ? "keeps-first" : "other",
          detail: `value=${v}`,
        };
      } catch (err) {
        return unreadable(err);
      }
    },
  },
  {
    id: "tags-on-fresh-shape",
    question: "Does a shape added THIS sync already have a .tags collection?",
    // `faults.tagsUndefinedOn` models a host handing back a shape whose `.tags`
    // is undefined — reading `.add` off it throws SYNCHRONOUSLY, which is what
    // made it so expensive: it escaped the tagging loop rather than failing one
    // chart.
    ask: async (ctx) => {
      const [shape] = await scratchShapes(ctx, [{ left: 90, top: 10, width: 20, height: 20 }]);
      const tags = (shape as unknown as { tags?: unknown }).tags;
      return { answer: tags ? "yes" : "undefined" };
    },
  },
  {
    id: "delete-then-lookup",
    question: "Right after delete()+sync, does getItemOrNullObject report it gone?",
    // `deleteSlideById` verifies from a FRESH context because of this. If a
    // host answers honestly in the same context, that re-check is unnecessary.
    ask: async (ctx) => {
      // The id is loaded in the SAME sync as the add — a second round trip for
      // it would age the proxy, and this question is not about proxy age.
      const [shape] = await scratchShapes(ctx, [{ left: 120, top: 10, width: 20, height: 20 }], "id");
      const id = (shape as unknown as { id: string }).id;
      (shape as unknown as { delete(): void }).delete();
      await ctx.sync();
      const gone = probeShapes(ctx).getItemOrNullObject(id);
      gone.load("id");
      await ctx.sync();
      try {
        const nul = (gone as unknown as { isNullObject: boolean }).isNullObject;
        return { answer: nul === true ? "reports-gone" : nul === false ? "still-there" : "unreadable" };
      } catch (err) {
        return threw(err);
      }
    },
  },
  {
    id: "addgroup-returns-usable",
    question: "Is a group proxy usable in the same sync that created it?",
    // Members resolved in the grouping batch, and the id asked for in that same
    // batch. Both were a sync late before, and this host answered "unreadable"
    // — which said only that a one-sync-old group proxy is refused, a fact
    // three other questions already establish about proxies in general.
    ask: async (ctx) => {
      const { members, via } = await groupMembers(
        ctx,
        [0, 1].map((i) => ({ left: 150 + i * 25, top: 10, width: 20, height: 20 })),
      );
      try {
        const group = (
          probeShapes(ctx) as unknown as { addGroup(shapes: unknown[]): { load(p: string): void; id: string } }
        ).addGroup(members);
        group.load("id");
        await ctx.sync();
        return {
          answer: typeof group.id === "string" && group.id ? "yes" : "unreadable",
          detail: `members via ${via}`,
        };
      } catch (err) {
        return { ...threw(err), detail: `members via ${via}` };
      }
    },
  },
  {
    id: "group-children-via-getcount",
    resample: true,
    question: "Does a group report its child COUNT, where it will not list them?",
    // The variant that separates two things `group-reports-its-children` runs
    // together, and the 2026-08-08 sheet is why it exists. That question asked
    // through `group/shapes/items/id` and the host answered `threw` — "The
    // property 'items' is not available", office-js#6363's exact signature —
    // even with the load queued in the sync that made the group.
    //
    // But the same sheet says `getcount-populates-same-sync: yes, value=9`. This
    // host COUNTS a shape collection it will not LIST. Nobody has asked whether
    // that also holds for a GROUP's collection, and the answer decides something
    // concrete: `contentShapes` returns UNKNOWN_CONTENT for every grouped slide,
    // which is what makes the reconcile report a slide complete without counting
    // it (`SlotVerdict.measured`). A count is all it needs. If this answers with
    // a number, those verdicts become measurements.
    //
    // Asked in the group's own sync for the same reason its sibling is: a proxy
    // one sync old is refused here, and that would answer a different question.
    ask: async (ctx) => {
      const { members, via } = await groupMembers(
        ctx,
        [0, 1].map((i) => ({ left: 250 + i * 25, top: 110, width: 20, height: 20 })),
      );
      try {
        const group = (
          probeShapes(ctx) as unknown as {
            addGroup(shapes: unknown[]): { group: { shapes: { getCount(): { value: number } } } };
          }
        ).addGroup(members);
        const count = group.group.shapes.getCount();
        await ctx.sync();
        const n = loadedCount(count);
        return typeof n === "number"
          ? { answer: n === 2 ? "two" : `reports-${n}`, detail: `members via ${via}` }
          : { answer: "unreadable", detail: `members via ${via}` };
      } catch (err) {
        return { ...threw(err), detail: `members via ${via}` };
      }
    },
  },
  {
    id: "group-reports-its-children",
    question: "After grouping two shapes, does the group report two children?",
    // The single most load-bearing answer here. A chart IS a group, and the
    // readback measures whether a chart survived by counting what is inside
    // it — so a host that groups successfully and then reports no children
    // makes every chart read back as wreckage, and the repair pass then
    // "fixes" charts that were never broken. The fake's own notes say a
    // version of it that put one shape there did exactly that.
    ask: async (ctx) => {
      const { members, via } = await groupMembers(
        ctx,
        [0, 1].map((i) => ({ left: 200 + i * 25, top: 60, width: 20, height: 20 })),
      );
      try {
        const group = (
          probeShapes(ctx) as unknown as {
            addGroup(shapes: unknown[]): { load(p: string): void; group: { shapes: { items: unknown[] } } };
          }
        ).addGroup(members);
        // Queued in the batch that MAKES the group, so the answer arrives on
        // the group's own sync. Asked one sync later — the first version — this
        // host answered PropertyNotLoaded, which is a statement about proxy age
        // and not about what the group contains.
        group.load("group/shapes/items/id");
        await ctx.sync();
        const n = group.group?.shapes?.items?.length;
        return {
          answer: n === 2 ? "two" : typeof n === "number" ? `reports-${n}` : "unreadable",
          detail: `n=${n}; members via ${via}`,
        };
      } catch (err) {
        return threw(err);
      }
    },
  },
  {
    id: "grouped-child-by-id-from-slide",
    question: "Can a shape INSIDE a group still be resolved by id off the slide?",
    // Sampled on every pass while it is pending: a question asked once has been
    // sampled, not answered, and this one gates a feature.
    resample: true,
    // The question that decides whether the in-place update can ever work on
    // this host, and it has never been asked.
    //
    // `tryInPlaceUpdate` writes to a chart's individual shapes, so it needs a
    // node-to-shape mapping. It gets one from CHART_PARTS_TAG, which is only
    // written for UNGROUPED charts — a grouped chart's shapes are inside the
    // group and the tag does not list them. On this host that closes both
    // doors: the `53ec985` round declined eleven times out of eleven with `the
    // chart has no parts list`, because the charts that grouped have no tag and
    // the ones that did not group had their id readback refused.
    //
    // But "the parts tag does not list them" is a fact about OUR code, not
    // about the host. If the slide's own shape collection will still resolve a
    // child by id, then a grouped chart can carry a parts list too, written
    // from the ids the grouping pass already holds — and the fast path applies
    // to the fourteen grouped charts a round produces instead of none of them.
    //
    // Asked one sync AFTER the group is made, which is how an update would ask:
    // it resolves the id a previous run wrote down, off a slide handle taken
    // now. office-js#3014 says grouped shapes come back from getItem as type
    // `unknown` and sub-shapes cannot be reached, so a no here is expected and
    // is worth recording as a measured no rather than an assumed one.
    ask: async (ctx) => {
      const { members, via } = await groupMembers(
        ctx,
        [0, 1].map((i) => ({ left: 420 + i * 25, top: 200, width: 20, height: 20 })),
      );
      try {
        const first = members[0] as unknown as { load(p: string): void; id: string };
        first.load("id");
        await ctx.sync();
        const childId = readShapeId(first);
        if (!childId) return { answer: "no-child-id", detail: `members via ${via}` };
        (probeShapes(ctx) as unknown as { addGroup(shapes: unknown[]): unknown }).addGroup(members);
        await ctx.sync();
        // A FRESH handle, the way an update asks a run later.
        const back = probeShapes(ctx).getItemOrNullObject(childId);
        back.load("isNullObject,id");
        await ctx.sync();
        let gone: boolean | undefined;
        try {
          gone = (back as unknown as { isNullObject: boolean }).isNullObject;
        } catch {
          gone = undefined;
        }
        return {
          answer: gone === false ? "yes" : gone === true ? "null-object" : "unreadable",
          detail: `childId=${childId}; members via ${via}`,
        };
      } catch (err) {
        return threw(err);
      }
    },
  },
  {
    id: "tag-on-group-survives",
    question: "Does a tag written on a GROUP read back?",
    // Where a chart's config actually lives. Tags on a plain shape are covered
    // above; if a group behaves differently, every chart in every deck is
    // un-re-editable and nothing else in the probe would say so.
    // The one whose old answer was frightening and wrong. It read back
    // `undefined` — "no chart in any deck is re-editable" — because the tag was
    // written through a group proxy a sync old. The group's id is what crosses
    // the sync now, and every use resolves its own handle, exactly as
    // `settleAndTagChart` does on the path that carries real charts.
    ask: async (ctx) => {
      const ids = idsOf(
        await scratchShapes(
          ctx,
          [0, 1].map((i) => ({ left: 260 + i * 25, top: 60, width: 20, height: 20 })),
          "id",
        ),
      );
      try {
        const group = (
          probeShapes(ctx) as unknown as {
            addGroup(shapes: unknown[]): { load(p: string): void; id: string };
          }
        ).addGroup(ids.map((id) => probeShape(ctx, id)));
        group.load("id");
        await ctx.sync();
        const groupId = group.id;
        if (typeof groupId !== "string" || !groupId)
          return { answer: "unreadable", detail: "the host would not name the group it had just made" };
        probeShape(ctx, groupId).tags.add("POWERCHART_PROBE_GROUP", "kept");
        await ctx.sync();
        const tag = probeShape(ctx, groupId).tags.getItemOrNullObject("POWERCHART_PROBE_GROUP");
        tag.load("value");
        await ctx.sync();
        const v = (tag as unknown as { value: string }).value;
        return { answer: v === "kept" ? "yes" : "no", detail: `value=${String(v)}` };
      } catch (err) {
        return threw(err);
      }
    },
  },
  {
    id: "getitemat-past-end",
    // Asks `slides.getItemAt(9999)` of the DECK collection and never touches the scratch slide. Round 15 lost it to `no-scratch-slide` — a free answer thrown away, which is the `untrack-available` incident repeating.
    noSlideNeeded: true,
    question: "What does slides.getItemAt() past the end of the deck do?",
    ask: async (ctx) => {
      try {
        const slide = ctx.slides.getItemAt(9999);
        slide.load("id");
        await ctx.sync();
        const id = (slide as unknown as { id: string }).id;
        return { answer: typeof id === "string" && id ? "returns-something" : "unreadable" };
      } catch (err) {
        return threw(err);
      }
    },
  },
  {
    id: "picture-then-shape-read",
    question: "After a picture is added, does re-reading the slide's shapes still answer?",
    // office-js#5022, open and assigned: get shapes, insert an image, delete the
    // old shapes, sync — and `context.sync()` "can run indefinitely" when the
    // shapes are read back to rename the new image. The reporter's only
    // workaround is a 1-2 second pause, and it still recurs.
    //
    // `drawDemoItem` does exactly this shape. A chart too dense to draw becomes
    // ONE picture, and `needsRefresh` is true whenever `pictureBase64` is set —
    // so the picture is added and the shape collection is re-read a sync later,
    // in the same context. Every unexplained hang this project has recorded is
    // consistent with it, and none of them can be pinned on it without asking.
    //
    // A hang here answers `silent` on the probe budget rather than taking the
    // sheet down, which is the whole reason each question gets its own bound.
    ask: async (ctx) => {
      // `requirementSets()` rather than a private `supports` — same source, and
      // the sheet already carries the list, so a reader can check the gate.
      if (!requirementSets().includes("1.8"))
        return { answer: "no-api", detail: "picture fills need PowerPointApi 1.8" };
      try {
        const [shape] = await scratchShapes(ctx, [{ left: 260, top: 120, width: 40, height: 40 }]);
        (shape.fill as unknown as { setImage(b64: string): void }).setImage(ONE_PIXEL_PNG);
        await ctx.sync();
        // The read that is said to hang — a fresh collection off a fresh slide
        // handle, one sync after the picture landed.
        const shapes = probeShapes(ctx);
        shapes.load("items/id");
        await ctx.sync();
        const items = (shapes as unknown as { items?: unknown[] }).items;
        return items ? { answer: "yes", detail: `${items.length} shape(s)` } : { answer: "unreadable" };
      } catch (err) {
        return threw(err);
      }
    },
  },
  {
    id: "group-of-existing-shape-readable",
    question: "Can a group found in a later batch report its children?",
    // office-js#5849: reading `shape.group` throws GeneralException. Distinct
    // from `group-reports-its-children`, which asks in the batch that MADE the
    // group — this asks the way `countGroupChildrenPage` asks, about a group
    // resolved from the deck afterwards.
    //
    // That pass is what decides whether a chart read back is intact or
    // wreckage, and every failure in it is swallowed per shape. So a host that
    // refuses this produces no error and no measurement, and the repair pass
    // then has nothing to go on for the charts it was meant to check.
    ask: async (ctx) => {
      const { members, via } = await groupMembers(
        ctx,
        [0, 1].map((i) => ({ left: 300 + i * 25, top: 160, width: 20, height: 20 })),
      );
      try {
        const made = (
          probeShapes(ctx) as unknown as { addGroup(shapes: unknown[]): { load(p: string): void; id: string } }
        ).addGroup(members);
        // The id in its OWN sync, never in the one that made the group.
        // That saving is what cost the 2026-08-05 sheet five of its questions:
        // this host will not name a shape in the batch that created it, and
        // `scratchShapes` was rewritten for exactly the same reason.
        await ctx.sync();
        made.load("id");
        await ctx.sync();
        const groupId = readShapeId(made);
        // An ANSWER, not a setup failure. A host that will not name a group it
        // just made cannot be asked the later-batch question at all — there is
        // no id to resolve one with — and that is a fact about the host, so it
        // belongs in the sheet. Reported as `no-scratch-slide` it read as the
        // probe having lost its slide, which is what kept this question out of
        // four consecutive rounds while the actual finding sat one line away.
        if (!groupId) return { answer: "no-group-id", detail: "the host would not name the group it just made" };
        // A LATER batch, through a fresh handle — the thing that separates this
        // question from the one above it.
        const found = probeShape(ctx, groupId) as unknown as {
          group: { shapes: { getCount(): { value: number } } };
        };
        const count = found.group.shapes.getCount();
        await ctx.sync();
        const n = loadedCount(count);
        return typeof n === "number"
          ? { answer: String(n), detail: `members via ${via}` }
          : { answer: "unreadable", detail: `members via ${via}` };
      } catch (err) {
        return { ...threw(err), detail: `members via ${via}` };
      }
    },
  },
  {
    id: "slide-layout-readable",
    // Asked of `ctx.durableSlideId` on purpose — a slide that predates the run — so demanding a live scratch slide refuses it for the absence of something it deliberately avoids.
    noSlideNeeded: true,
    question: "Can a slide's layout shapes be read?",
    // office-js#3826, open and marked a product bug: on Office on the web
    // `slide.load("layout/shapes/items")` fails the sync with GeneralException.
    // The per-slide companion to `layouts-readable`, which asks the same of the
    // deck's masters — #4906 and #2328 report the master form, #3826 the slide
    // form, and nothing says whether they are one defect or three.
    //
    // Asked of a slide that was in the deck BEFORE this run, not the scratch
    // slide, and that is the whole care taken here. A freshly-added slide's
    // handle is good for exactly one sync on this host — established, and
    // unconditional in the fake — so a load queued on it and read after its own
    // sync answers "unreadable" whatever the layout API does. That would be an
    // answer about the probe's own plumbing dressed as a fact about layouts,
    // which is the failure mode this file exists to prevent.
    //
    // Read-only. It resolves someone's own slide and asks what its layout
    // holds; nothing is created, moved or drawn.
    ask: async (ctx) => {
      const durable = ctx.durableSlideId;
      if (!durable) return { answer: "no-durable-slide", detail: "the deck has no slide this run did not add" };
      try {
        const slide = ctx.slides.getItemOrNullObject(durable) as unknown as {
          load(p: string): void;
          layout?: { shapes?: { items?: unknown[] } };
        };
        slide.load("layout/shapes/items/id");
        await ctx.sync();
        const items = slide.layout?.shapes?.items;
        return items ? { answer: "yes", detail: `${items.length} layout shape(s)` } : { answer: "unreadable" };
      } catch (err) {
        return threw(err);
      }
    },
  },
  {
    id: "layouts-readable",
    // Reads `presentation.slideMasters`. Nothing to do with any slide this run added.
    noSlideNeeded: true,
    question: "Can the deck's slide masters and their layouts be read?",
    // office-js#4906 and office-js#2328: loading shapes off `SlideLayout` or
    // `SlideMaster` throws GeneralException in PowerPoint Online — and #4906
    // reports it happening ONLY on presentations built from a custom template,
    // with `errorLocation: SlideMasterCollection.getItem`. Both are open.
    //
    // `blankLayoutId` reads exactly this, to give every slide the add-in creates
    // a BLANK layout. It is try/caught, so a host that refuses degrades quietly
    // to the inherited layout — which means the demo deck and every scratch
    // slide land on top of the previous slide's placeholders, "Click to add
    // title" showing through the chart. That is a visible defect with no error
    // anywhere, and nobody would know which of the two it was.
    //
    // The owner's decks come from a corporate template, which is precisely the
    // case #4906 singles out. Read-only, and no slide is created or touched.
    ask: async (ctx) => {
      try {
        const masters = ctx.presentation.slideMasters;
        masters.load("items/id,items/layouts/items/id,items/layouts/items/type");
        await ctx.sync();
        const items = (masters as unknown as { items?: { layouts?: { items?: unknown[] } }[] }).items;
        if (!items) return { answer: "unreadable", detail: "the host answered the sync but not the collection" };
        const layouts = items.reduce((n, m) => n + (m.layouts?.items?.length ?? 0), 0);
        // A master with no layouts is not the same as no masters, and neither is
        // the same as a throw. `blankLayoutId` returns undefined for all three
        // and the caller cannot tell them apart; the sheet can.
        return {
          answer: layouts > 0 ? "yes" : "no-layouts",
          detail: `${items.length} master(s), ${layouts} layout(s)`,
        };
      } catch (err) {
        return unreadable(err);
      }
    },
  },
  {
    id: "untrack-available",
    question: "Do proxies expose untrack()?",
    // Nothing here syncs: `getItemOrNullObject` returns a proxy without a round
    // trip, and the question is whether that object carries a method. A live
    // slide is not required and must not be demanded — see `noSlideNeeded`.
    noSlideNeeded: true,
    // `untrack` is best-effort here precisely because a null-object proxy may
    // not expose it. Worth knowing which hosts do.
    ask: async (ctx) => {
      const slide = ctx.slides.getItemOrNullObject(ctx.scratchId);
      const has = typeof (slide as unknown as { untrack?: unknown }).untrack === "function";
      return { answer: has ? "yes" : "no" };
    },
  },
];

/** An error, short enough to sit in a sheet without swamping it. */
function short(err: unknown): string {
  const text = err instanceof Error ? err.message : String(err);
  return text.slice(0, 140);
}

/**
 * Ask this host every question, and come back with a complete sheet.
 *
 * Complete is the operative word. A probe that throws, times out, or wedges
 * contributes its answer and nothing more — the sheet from a misbehaving host
 * is the one worth having, so it must survive misbehaviour. Each question gets
 * its own context and its own budget for that reason.
 */
export async function runHostProbes(source: string, build: string): Promise<HostAnswerSheet> {
  const answers: HostAnswer[] = [];
  /**
   * Every scratch slide this run has made, so all of them come back out.
   *
   * Plural because the run replaces one it has lost. A diagnostic that leaves
   * blank slides in someone's deck is one they stop clicking, and the whole
   * point of replacing a lost slide is to keep going — which means keeping a
   * list rather than overwriting the id of the thing still to be cleaned up.
   */
  /**
   * Every slide this run put in the deck, usable or not.
   *
   * The sweep clamps at "no more than this run added", so this set has to be
   * everything the run LEFT IN THE DECK — including the slides
   * `addScratchSlide` landed, refused, and then could not take back out. It
   * swallowed those, because its return value means "an id you may use" and
   * these are precisely the ids you may not. The 2026-08-11 round left exactly
   * two behind for that reason: the deck grew by 70 while this list held 68, so
   * the clamp correctly refused to touch the other two.
   *
   * Fed by `addScratchSlide`'s callback for that case, and de-duplicating,
   * because a slide it successfully removes must NOT appear here — counting it
   * would make the clean-up owe a delete for a slide that is already gone, and
   * report "some" where the truth is "all".
   */
  const scratchIds: string[] = [];
  const noteScratch = (id: string): void => {
    if (!scratchIds.includes(id)) scratchIds.push(id);
  };
  // Read BEFORE the first scratch slide is added, so it cannot pick one up.
  // Follow-up questions use it as the control for "was that about the call, or
  // about the slide being new" — the one distinction this project has paid for
  // twice — and nothing may ever write to it.
  //
  // Wrapped, and the test that made me wrap it is the one this whole file is
  // about: a host that goes silent made this throw before a single question had
  // been asked, and the run produced NO SHEET. A control nobody can name costs
  // the follow-up. It must never cost the sheet.
  let durableSlideId: string | undefined;
  try {
    durableSlideId = (await deckSlideIds())?.[0];
  } catch {
    trace("probe", "no durable slide to use as a control — the deck would not list itself", {});
  }
  // The deck's size before this run added anything — the floor the positional
  // sweep may never reach past. Read here rather than in the clean-up because
  // by then the run's own slides are already in the count.
  const deckAtStart = (await deckSlideIds().catch(() => undefined))?.length;
  let scratchId = await addScratchSlide(SCRATCH_ADD_BUDGET_MS, noteScratch);
  if (scratchId) scratchIds.push(scratchId);
  /**
   * Consecutive questions that could not get an answer out of the host at all.
   *
   * Counted on DEADLINE MISSES, not on `no-scratch-slide`, and the distinction
   * is the whole design. A round on 2026-08-08 answered `no-scratch-slide` four
   * times running — the host had stopped resolving fresh slide ids for about
   * fifteen seconds — and then recovered and answered five more questions,
   * including the two this project cares most about. A breaker keyed on that
   * signal would have thrown those away. Those questions were still being
   * ANSWERED, in two to six seconds; they just could not be set up.
   *
   * A missed deadline is different: the host took the question and never came
   * back inside its budget. Three of those in a row is not a finding about a
   * question, it is a finding about the host — the same argument, and the same
   * number, as the self-test's SICK_LIMIT.
   */
  let mute = 0;
  let abandoned: string | null = null;
  const runStarted = Date.now();
  /**
   * A sample's timestamp, on the run log's clock so it can be read beside the
   * trace lines around it. See `ProbeSample.atMs` for what that cost.
   */
  const stampNow = (): number => traceElapsed() ?? Date.now() - runStarted;
  /**
   * What the run has already watched the host do. Never a guess, never a call.
   */
  /** When this run last SAW each thing. Timestamps, not flags — see `regimeFrom`. */
  let lastRefusalAt: number | undefined;
  let lastSlideTroubleAt: number | undefined;
  let lastGoodAt: number | undefined;
  const regimeNow = (at: number): HostRegime => regimeFrom({ at, lastRefusalAt, lastSlideTroubleAt, lastGoodAt });
  /** Answers that mean the shape collection has stopped answering. */
  const REFUSAL = new Set(["unreadable", "short-0", "not-listed"]);
  /**
   * One row per question, however many times it is asked.
   *
   * `answer` keeps exactly the meaning it has always had — the first REAL
   * answer, with a never-asked sentinel standing only until a later pass
   * replaces it. That is what the second pass has always done to a lost row;
   * this generalises it to three passes without the diff, the contract gate or
   * any reader of a sheet noticing. The repeats go to `samples`.
   */
  const rows = new Map<string, HostAnswer>();
  const record = (row: HostAnswer, pass: number): HostAnswer => {
    const atMs = stampNow();
    // Each observation stamped with WHEN, so a later reading can tell a host
    // that is refusing now from one that refused a minute ago.
    const aboutTheCollection = /^shapes?-/.test(row.id);
    if (aboutTheCollection && REFUSAL.has(row.answer)) lastRefusalAt = atMs;
    else if (aboutTheCollection && !NOT_ASKED.has(row.answer)) lastGoodAt = atMs;
    if (row.answer === "no-scratch-slide") lastSlideTroubleAt = atMs;
    else if (!NOT_ASKED.has(row.answer)) lastGoodAt = atMs;
    const sample: ProbeSample = { answer: row.answer, pass, atMs, regime: regimeNow(atMs) };
    const seen = rows.get(row.id);
    if (!seen) {
      const created: HostAnswer = { ...row, samples: [sample] };
      rows.set(row.id, created);
      answers.push(created);
      return created;
    }
    seen.samples = [...(seen.samples ?? []), sample];
    // A real answer replaces a never-asked, and nothing replaces a real one:
    // the FIRST real answer is the row's answer, so a sheet means today what it
    // meant yesterday. Disagreement between samples is `stable`'s job to report,
    // not `answer`'s to hide.
    if (NOT_ASKED.has(seen.answer) && !NOT_ASKED.has(row.answer)) {
      seen.answer = row.answer;
      seen.ms = row.ms;
      seen.detail = row.detail;
    }
    return seen;
  };
  try {
    /** Questions the current pass will put. Pass 1 is always the whole list. */
    let list: Probe[] = [...PROBES];
    for (let pass = 1; pass <= PROBE_PASSES; pass++) {
      if (pass > 1) {
        // Decided from the pass just finished, not from the run's whole
        // history: pressure comes and goes here in fifteen-second windows, and
        // a run that struggled early and recovered should get its repeats.
        const put = [...rows.values()].filter((r) => r.samples?.some((x) => x.pass === pass - 1));
        const lostLast = put.filter((r) => NOT_ASKED.has(r.samples!.find((x) => x.pass === pass - 1)!.answer)).length;
        const pressured = put.length > 0 && lostLast / put.length > PASS_PRESSURE_LIMIT;
        list = pressured ? PROBES.filter((pr) => pr.resample) : [...PROBES];
        trace("probe", "starting another pass over the questions", {
          pass,
          asking: list.length,
          of: PROBES.length,
          why: pressured ? `${lostLast} of ${put.length} were never put last pass — shortlist only` : "full list",
        });
        if (!list.length) break;
      }
      for (const probe of list) {
        const started = Date.now();
        // Stop asking a host that has stopped answering.
        //
        // The probe had no rung for this at all while the self-test has had one
        // for weeks, and PowerPoint's own "Sorry, we ran into a problem" dialog is
        // what exposed the gap: the host was gone, the pane's timer kept counting,
        // and the run went on putting questions to a dead document — every one of
        // them spending its full budget to record an answer about nothing.
        if (!abandoned && mute >= PROBE_MUTE_LIMIT) {
          abandoned = `${mute} questions in a row got no answer out of the host`;
          trace("probe", "giving up on the host", { after: answers.length, why: abandoned });
        }
        if (abandoned) {
          record(
            {
              id: probe.id,
              question: probe.question,
              answer: "not-asked",
              ms: 0,
              detail: `not reached — ${abandoned}`,
            },
            pass,
          );
          continue;
        }
        const deadlinesBefore = deadlinesFired;
        trace("probe", "asking", { id: probe.id });
        let result: { answer: string; detail?: string };
        // One more attempt at a slide, every question, rather than writing the
        // rest of the sheet off.
        //
        // `scratchId` goes null when a replacement also failed to resolve, and
        // nothing used to set it back — so a single bad moment cost every
        // question after it, all of them recorded `no-scratch-slide` as though
        // the host had been asked. That is the same shape as the failure this
        // whole rung exists to prevent, one level up: an answer that describes
        // the run's own state rather than the host's.
        //
        // A host that genuinely will not keep a slide pays one `addScratchSlide`
        // per question for the honest answer, which is what the budget is for.
        if (!scratchId) {
          const recovered = await addScratchSlide(SCRATCH_ADD_BUDGET_MS, noteScratch);
          if (recovered) {
            noteScratch(recovered);
            scratchId = recovered;
            trace("probe", "took another scratch slide after giving up on the last", { id: probe.id });
          }
        }
        if (!scratchId && !probe.noSlideNeeded) {
          result = { answer: "no-scratch-slide", detail: "the host would not add a slide to work on" };
        } else if (probe.noSlideNeeded) {
          // Not gated on a scratch slide at any level — not on having one, and
          // not on its liveness. The id is passed through because `ProbeContext`
          // carries one, but nothing here resolves it. See `Probe.noSlideNeeded`.
          result = await ask(probe, scratchId ?? "", durableSlideId);
        } else {
          result = await ask(probe, scratchId as string, durableSlideId);
          // A slide that stopped resolving costs one question, not the sheet.
          //
          // It cost thirteen once: a real host lost the scratch slide after the
          // first probe and the remaining thirteen questions each reported the
          // same failure as if it were their own answer. Thirteen apparent
          // divergences, one cause, and none of the questions actually asked.
          // Replacing the slide is the difference between a sheet that reports a
          // finding and a sheet that IS the finding.
          //
          // Either kind of never-asked gets the replacement: a slide the host
          // will not resolve, and a slide that resolves and will not take a
          // shape. The second is a weaker reason to suspect the slide, but the
          // cost is one add and one question, and the alternative is a sheet
          // that gives up on eight questions because one slide went bad.
          if (NOT_ASKED.has(result.answer)) {
            const replacement = await addScratchSlide(SCRATCH_ADD_BUDGET_MS, noteScratch);
            if (replacement) {
              noteScratch(replacement);
              scratchId = replacement;
              trace("probe", "replaced the scratch slide", {
                id: probe.id,
                scratchId: replacement,
                after: result.answer,
              });
              const retry = await ask(probe, replacement, durableSlideId);
              // Only adopt a retry that actually got somewhere. A second failure
              // on a brand-new slide is a stronger statement than the first, and
              // overwriting it with a never-asked would hide that.
              if (!NOT_ASKED.has(retry.answer)) result = retry;
              // Give up on the SLIDE only when the slide is what failed. A host
              // that resolves slides and refuses shapes has nothing wrong with
              // its slides, and writing off the scratch slide there would turn
              // one refusal into "no-scratch-slide" for every question after it —
              // the very noise this whole rung exists to prevent.
              else if (retry.answer === "no-scratch-slide") scratchId = null;
            } else if (result.answer === "no-scratch-slide") {
              scratchId = null;
            }
          }
        }
        const askedMs = Date.now() - started;
        record({ id: probe.id, question: probe.question, ms: askedMs, ...result }, pass);
        // THIS PASS's answer, not the row's.
        //
        // `record` returns the accumulated ROW, whose `answer` is deliberately
        // the FIRST real answer so a sheet means today what it meant yesterday.
        // Every line below wants the opposite — what the host said just now —
        // and for three passes they all read the row instead. The 2026-08-11
        // round is the proof: the samples say `threw, yes, yes` for
        // `shape-add-held-slide-proxy` while the run log says `answered: threw`
        // three times, with an identical `ms: 637` on asks forty-five seconds
        // apart. A reader of that log concludes the opposite of what the host
        // did. Same family as `batch committed`: a line asserting something it
        // does not know.
        trace("probe", "answered", { id: probe.id, answer: result.answer, ms: askedMs, pass });
        // A question that wrecks the slide on its way out gives up the slide, so
        // the next one starts on a fresh handle instead of inheriting the damage.
        //
        // Dropped AFTER the answer is recorded, and only on `burnsTheSlide`: this
        // question succeeded, so nothing else here would have replaced anything.
        // The `if (!scratchId)` at the top of the loop takes the new one, which is
        // the same path a lost slide already uses — no second mechanism.
        if (probe.burnsTheSlide && scratchId) {
          trace("probe", "giving up the scratch slide this question wrecked", { id: probe.id, scratchId });
          scratchId = null;
        }
        // Reset on ANY question that came back inside its budget, including one
        // that could not be set up: a host still refusing questions promptly is a
        // host still talking to us.
        mute = deadlinesFired > deadlinesBefore ? mute + 1 : 0;
        // The partner question, in the same run, when the answer admits two
        // readings. Never asked when the question was never PUT: a follow-up to
        // `no-scratch-shape` would be a second question about the probe's own
        // setup, dressed as a fact about the host.
        const follow = probe.follow;
        // Gated on THIS pass's answer for the same reason. Reading the row here
        // meant a partner conditioned on a later pass's answer could never fire
        // on it, and one conditioned on the first answer fired on every pass
        // regardless of what the host had just said — a follow-up asked
        // "because X answered yes" when X had, that time, answered no.
        if (follow && !NOT_ASKED.has(result.answer) && follow.when(result.answer)) {
          const at = Date.now();
          trace("probe", "asking the partner question", { after: probe.id, id: follow.probe.id, was: result.answer });
          let r = scratchId
            ? await ask(follow.probe, scratchId, durableSlideId)
            : { answer: "no-scratch-slide", detail: "the host would not add a slide to work on" };
          // The same replacement the main loop gets, and for the same reason.
          //
          // A follow-up went unasked in three consecutive rounds:
          // `getitem-durable-slide` reads the DURABLE slide and never touches the
          // scratch one, yet `ask` liveness-checks the scratch slide first, so a
          // question with no use for that slide was refused for its absence. On a
          // host that loses a scratch slide after almost every question — this one
          // replaced seventeen in one run — that is not an edge case, it is the
          // normal path.
          if (NOT_ASKED.has(r.answer)) {
            const replacement = await addScratchSlide(SCRATCH_ADD_BUDGET_MS, noteScratch);
            if (replacement) {
              noteScratch(replacement);
              scratchId = replacement;
              trace("probe", "replaced the scratch slide for a partner question", {
                id: follow.probe.id,
                scratchId: replacement,
                after: r.answer,
              });
              const retry = await ask(follow.probe, replacement, durableSlideId);
              // Same rule as the main loop: only adopt a retry that got somewhere.
              if (!NOT_ASKED.has(retry.answer)) r = retry;
            }
          }
          record(
            {
              id: follow.probe.id,
              question: follow.probe.question,
              ms: Date.now() - at,
              ...r,
              // Which answer triggered it, in the sheet itself. Two rows that do
              // not say they are a pair are two unrelated facts to whoever reads
              // them a week later.
              detail: `asked because ${probe.id} answered "${result.answer}" — ${follow.because}${r.detail ? `; ${r.detail}` : ""}`,
            },
            pass,
          );
          trace("probe", "partner answered", { id: follow.probe.id, answer: r.answer });
        }
      }
      // The breaker is a fact about the host, not about a pass: once the run has
      // given up there is nothing for a later pass to sample.
      if (abandoned) break;
    }

    // A LAST pass over the questions this run never managed to put, at the
    // END of the run rather than beside the failure.
    //
    // The retry above already takes a fresh slide and asks again — and it is not
    // enough. Two consecutive rounds lost fourteen of twenty-seven and thirteen
    // of twenty-eight questions to `no-scratch-slide`, with twenty-one slide
    // replacements between them: the retries fired, and failed the same way.
    //
    // Time is what they were missing. This host's ability to resolve a freshly
    // added slide comes and goes in windows of roughly fifteen seconds, so a
    // retry issued immediately lands INSIDE the window that just refused. A pass
    // at the end is a different sample, and both rounds show the recovery
    // happening within the same run: the 2026-08-09 round answered questions at
    // positions 17, 18, 22, 24 and 26 after losing 10 through 16.
    //
    // Which questions get lost is close to random — the two rounds disagree on
    // six of them — so this is not a fix for one question. It is the difference
    // between a sheet that answers half its questions and one that answers as
    // many as the host will take.
    //
    // Skipped entirely when the breaker fired: `not-asked` means the run gave up
    // on a host that had stopped answering, and asking it thirteen more times is
    // the behaviour the breaker exists to stop.
    const lost = abandoned ? [] : answers.filter((a) => NOT_ASKED.has(a.answer));
    if (lost.length) {
      trace("probe", "second pass over the questions that were never put", { count: lost.length });
      let rescued = 0;
      let noSlide = 0;
      for (const entry of lost) {
        const probe = PROBE_BY_ID.get(entry.id);
        if (!probe) continue;
        if (mute >= PROBE_MUTE_LIMIT) {
          trace("probe", "stopping the second pass — the host is not answering", { at: entry.id });
          break;
        }
        const replacement = await addScratchSlide(SCRATCH_ADD_BUDGET_MS, noteScratch);
        // One refused slide is not a dead host, and treating it as one made
        // this whole pass theatre.
        //
        // It used to `break` here, silently. Round 15 is what that looks like
        // from outside: `second pass over the questions that were never put
        // {count: 10}`, one slide attempt two seconds later, `scratch slide
        // landed but its id will not resolve`, and then nothing — no rescues,
        // and no line saying why. Ten questions decided by one coin flip, on a
        // host whose slide resolution is known to flap in fifteen-second
        // windows, and a log that could not tell "tried them all and failed"
        // from "gave up on the first".
        //
        // The main loop has always taken one add attempt PER QUESTION, which is
        // the right shape against that host. The sweep took one for the whole
        // list. They match now, bounded by CONSECUTIVE failures so a genuinely
        // dead host still costs a few adds rather than one per lost question.
        if (!replacement) {
          if (++noSlide >= PROBE_MUTE_LIMIT) {
            trace("probe", "stopping the second pass — no scratch slide to be had", { at: entry.id, tried: noSlide });
            break;
          }
          continue;
        }
        noSlide = 0;
        noteScratch(replacement);
        scratchId = replacement;
        const deadlinesBefore = deadlinesFired;
        const started = Date.now();
        const retry = await ask(probe, replacement, durableSlideId);
        mute = deadlinesFired > deadlinesBefore ? mute + 1 : 0;
        // Only an answer replaces a never-asked. A second failure is the same
        // fact the first one already recorded, and overwriting the row with it
        // would lose the original `ms` for no gain.
        // Recorded as a SAMPLE like every other ask, so the rescue is visible
        // as the late sample it is rather than as an answer with no provenance.
        // `record` applies the same rule this block always did — only a real
        // answer replaces a never-asked — so the row reads exactly as before.
        const retriedAt = stampNow();
        entry.samples = [
          ...(entry.samples ?? []),
          { answer: retry.answer, pass: PROBE_PASSES + 1, atMs: retriedAt, regime: regimeNow(retriedAt) },
        ];
        if (!NOT_ASKED.has(retry.answer)) {
          entry.answer = retry.answer;
          entry.ms = Date.now() - started;
          entry.detail = `${retry.detail ? `${retry.detail}; ` : ""}answered on a second pass at the end of the run`;
          trace("probe", "second pass answered", { id: probe.id, answer: retry.answer });
          rescued++;
        }
      }
      // What the pass was WORTH, every time, including nothing. A sweep that
      // rescued none and a sweep that never ran read identically in round 15's
      // log, and the difference is the whole question of whether to keep it.
      trace("probe", "second pass finished", { of: lost.length, rescued });
    }
  } finally {
    // The scratch slides go back whatever happened. A diagnostic that litters
    // the user's deck is one they will stop running.
    //
    // And it SAYS how that went, because for as long as it did not, this
    // failing was something the owner discovered by opening a deck. One run
    // left 21 blank slides behind — a previous one, fourteen — and the sheet
    // that run produced recorded neither. `deleteSlideById` already answers
    // honestly (it re-reads the deck rather than trusting a queued delete);
    // the boolean was simply thrown away here.
    //
    // Filed as an answer rather than a trace line so it travels the same road
    // as every other fact about this host: into the sheet, through the diff,
    // and against the fake — which returns every slide it is given, so a host
    // that does not diverges and is reported without anyone remembering to
    // look.
    const cleanupStarted = Date.now();
    // COUNT the deck either side, and let the count outrank the per-slide
    // booleans — see `slidesActuallyReturned` for what that is worth.
    const idsBefore = await deckSlideIds().catch(() => undefined);
    const deckBefore = idsBefore?.length;
    // Does the deck still LIST the ids we are about to delete by?
    //
    // The one number that separates the two readings of a clean-up that
    // reports 45 deletes into a deck that does not shrink. Either the ids are
    // stable and the deletes simply fail, or the id we captured at add time is
    // not the id the deck answers to any more — in which case
    // `deleteSlideByPosition`'s `indexOf(id) < 0` reads "already gone" for a
    // slide sitting right there, and delete-by-id is not merely failing, it is
    // structurally impossible.
    //
    // 2026-08-11 made this worth asking rather than assuming: the scratch ids
    // read `4123571115#123571113` while the deck listed `256#109857222`
    // through `314#195537992`, and both come from the SAME `slideIds()`
    // projection minutes apart. Reasoning cannot choose between "the host
    // renumbers" and "two readers disagree"; one count can.
    const stillListed = idsBefore ? scratchIds.filter((id) => idsBefore.includes(id)).length : undefined;
    let returned = 0;
    for (const id of scratchIds) if (await deleteSlideById(id).catch(() => false)) returned++;
    let deckAfter = (await deckSlideIds().catch(() => undefined))?.length;
    const byId = slidesActuallyReturned({
      claimed: returned,
      added: scratchIds.length,
      deckBefore,
      deckAfter,
    });
    // The sweep, and ONLY when delete-by-id left something behind. On a host
    // where the ids work this never runs and nothing about the old path
    // changes; on this one it is the only thing that can work, because the ids
    // it would need are not ids the deck lists.
    let swept = 0;
    if (byId.left > 0) {
      const plan = positionalSweepPlan({
        deckAtStart,
        deckNow: deckAfter,
        added: scratchIds.length,
        alreadyDeleted: byId.actually,
      });
      if (plan) {
        trace("probe", "sweeping the run's own slides by position", {
          ...plan,
          deckAtStart,
          deckNow: deckAfter,
          why: "delete-by-id left slides behind and this host does not list the ids they were added under",
        });
        swept = await deleteTrailingSlides(plan.from, plan.count);
        deckAfter = (await deckSlideIds().catch(() => undefined))?.length;
        trace("probe", "swept", { swept, deckNow: deckAfter });
      } else {
        trace("probe", "no positional sweep — the deck does not support one safely", {
          deckAtStart,
          deckNow: deckAfter,
          added: scratchIds.length,
        });
      }
    }
    // Re-derived from the deck AFTER the sweep, so the number reported is what
    // the deck lost overall rather than what either mechanism claimed.
    const { actually, left, shrankBy } = slidesActuallyReturned({
      claimed: returned + swept,
      added: scratchIds.length,
      deckBefore,
      deckAfter,
    });
    answers.push({
      id: SCRATCH_CLEANUP_ID,
      question: "Does the host give back the scratch slides this probe added?",
      answer: !scratchIds.length ? "none-added" : !left ? "all" : actually ? "some" : "none",
      ms: Date.now() - cleanupStarted,
      detail:
        `${actually} of ${scratchIds.length} scratch slide(s) deleted${left ? `; ${left} left in the deck` : ""}` +
        // Only when the deck lost LESS than the deletes claimed. That clause was
        // written when the by-id path was the only one, where any disagreement
        // meant an over-claim — and the sweep made it read backwards on its
        // first outing: `the deletes reported 0 but the deck only shrank by 68`,
        // which is the sweep working, described as a shortfall. A line that says
        // the opposite of what happened is the `batch committed` mistake in a
        // new place, so it is conditioned on the direction it describes.
        (shrankBy !== undefined && shrankBy < returned
          ? ` (the deletes reported ${returned} but the deck only shrank by ${shrankBy})`
          : "") +
        (stillListed !== undefined && stillListed < scratchIds.length
          ? `; the deck still lists ${stillListed} of ${scratchIds.length} of these ids`
          : "") +
        (swept ? `; ${swept} removed by a positional sweep after delete-by-id took none` : ""),
    });
    trace("probe", "gave the scratch slides back", {
      returned,
      swept,
      left,
      deckBefore,
      deckAfter,
      shrankBy,
      stillListed,
      // A SAMPLE of both id lists, side by side. Reconstructing the id-space
      // mismatch by hand cost two rounds — the scratch ids read
      // `4123571115#123571113` while the deck listed `256#109857222`, and
      // nothing in the file put those two facts next to each other. Three of
      // each is enough to see it and small enough never to crowd the buffer.
      heldIds: scratchIds.slice(0, 3),
      deckIds: (idsBefore ?? []).slice(0, 3),
    });
  }
  // Said by the sheet, from its own samples — the fact `UNSTABLE_ANSWERS` was
  // assembled by hand across ten rounds to say. Only REAL answers count: a
  // question that was put once and refused twice is not unstable, it is a
  // question this run mostly could not ask.
  for (const row of answers) {
    const verdict = stabilityOf(row.samples);
    if (verdict !== undefined) row.stable = verdict;
  }
  return { kind: "powerchart-host-answers", source, build, requirementSets: requirementSets(), answers };
}

/**
 * One question, and whatever came back — including "I never got to ask".
 *
 * A probe that could not reach its question must not answer as though it had.
 * `"threw"` is a legitimate answer to several of these questions and the diff
 * compares it against the fake's; a scratch slide that vanished underneath a
 * probe reported as `"threw"` therefore reads as a genuine host divergence, and
 * that is precisely how thirteen of one real sheet's fourteen answers came to
 * be noise. `"no-scratch-slide"` is not in any probe's own vocabulary, so it
 * can never be mistaken for one.
 */
async function ask(
  probe: Probe,
  scratchId: string,
  durableSlideId?: string,
): Promise<{ answer: string; detail?: string }> {
  awaitingSetupShapes = false;
  try {
    return await withProbeContext(scratchId, PROBE_BUDGET_MS, probe.ask, durableSlideId, probe.noSlideNeeded);
  } catch (err) {
    if (err instanceof ScratchSlideUnavailable) return { answer: "no-scratch-slide", detail: short(err) };
    if (err instanceof ProbeSetupFailed) return { answer: "no-scratch-shape", detail: short(err) };
    // A budget that fired while the setup shapes were still out is a setup
    // failure too, and `"silent"` is a comparable answer that would read as a
    // host divergence. Which sync the deadline caught is the whole difference
    // between "this host says nothing about tag overwrites" and "this host
    // never gave the probe a shape to write a tag on".
    if (isTimeout(err) && awaitingSetupShapes)
      return { answer: "no-scratch-shape", detail: `the sync that was to add them never came back: ${short(err)}` };
    return { answer: isTimeout(err) ? "silent" : "threw", detail: short(err) };
  } finally {
    awaitingSetupShapes = false;
  }
}

/**
 * Every question this build knows how to ask — for the fake's baseline test.
 *
 * Follow-ups included, and flattened rather than listed apart: a conditional
 * question is still a question the sheet can carry, and a baseline that did not
 * know about it would report its answer as an id nobody recognises.
 */
const withFollows = (p: Probe): string[] => [p.id, ...(p.follow ? withFollows(p.follow.probe) : [])];

/**
 * Every probe by id, follow-ups included — what the end-of-run second pass
 * re-asks from. A follow-up only has a row at all when its trigger fired, so a
 * row that is there and never put is a fair thing to put again.
 */
const flatten = (p: Probe): Probe[] => [p, ...(p.follow ? flatten(p.follow.probe) : [])];
const PROBE_BY_ID = new Map(PROBES.flatMap(flatten).map((p) => [p.id, p]));

/**
 * The cleanup's own row in the sheet.
 *
 * Not a `Probe` — it has no question to put and no scratch slide to put it on,
 * because it IS what happens to the scratch slides. It is still a fact about
 * this host, it still has a fake answer to diverge from, and the sheet is the
 * only place anyone will read it. So it is listed here by hand rather than
 * derived, and the two lists below are where it joins the invariant.
 */
/**
 * Which slides a positional clean-up sweep may remove — and it is a SAFETY
 * rule before it is a clean-up one.
 *
 * Delete-by-id cannot work here: 2026-08-11 (`756682e`) measured `the deck
 * still lists 0 of 62 of these ids`, so every by-id delete takes the "already
 * gone" branch and removes nothing. What survives that finding is position: the
 * probe's slides are APPENDED by `slides.add()`, so they are the last N in the
 * deck and need no id at all.
 *
 * Position is also how an add-in destroys someone's work, so this returns a
 * plan only when every one of these holds, and each is a separate way of saying
 * "never touch a slide this run did not add":
 *
 *  - Both deck counts are known. A host that will not count its slides does not
 *    get to have slides deleted from it by arithmetic.
 *  - The deck GREW. If it did not, the run added nothing that is still there.
 *  - Never more than this run added, minus whatever already went back.
 *  - Never more than the deck grew — if something else removed slides while the
 *    probe ran, the sum no longer describes the deck and the smaller number is
 *    the honest one.
 *  - The first index to delete is at or after the deck's size when the run
 *    started. This is implied by the two rules above and asserted anyway,
 *    because it is THE property: everything before that index was the user's
 *    before this run began.
 *
 * The caller deletes from the highest index down, so removing one cannot shift
 * the index of another still to go.
 */
export function positionalSweepPlan(o: {
  deckAtStart?: number;
  deckNow?: number;
  added: number;
  alreadyDeleted: number;
}): { from: number; count: number } | null {
  if (o.deckAtStart === undefined || o.deckNow === undefined) return null;
  const grew = o.deckNow - o.deckAtStart;
  if (grew <= 0) return null;
  const count = Math.min(o.added - o.alreadyDeleted, grew);
  if (count <= 0) return null;
  const from = o.deckNow - count;
  // Belt and braces: the arithmetic above already guarantees this, and a plan
  // that reached into the user's own slides would be the one bug here that
  // cannot be apologised for.
  if (from < o.deckAtStart) return null;
  return { from, count };
}

/**
 * How many scratch slides ACTUALLY went back, from the deck rather than from
 * the deletes' own opinion of themselves.
 *
 * Extracted so it can be checked without a PowerPoint, which is the only way
 * this particular bug is checkable at all: reproducing it through the fake
 * needs three knobs turned together (a delete that no-ops, a host that refuses
 * the id, and a deck listing that id under another string), and a test built
 * out of three coincidences tests the coincidences.
 *
 * `claimed` is what `deleteSlideById` reported. The deck's own count outranks
 * it, because on 2026-08-11 the two disagreed completely: `42 of 42 deleted,
 * left: 0` into a deck that ended the round holding 56 slides it had added, 47
 * of them blank, with the owner's screenshot showing one carrying the 20x20
 * rectangle `shape-add-held-slide-proxy` draws. `deleteSlideByPosition` reads
 * `indexOf(id) < 0` as "already gone", which is sound only while the id we hold
 * and the ids the deck lists are the same strings — and there they were not:
 * scratch ids read `4123571115#123571113`, the deck listed `257#2599158489`.
 *
 * The MINIMUM of the two, never the maximum: a delete that cannot be
 * corroborated is unknown, and reporting an unknown as a success is how a run
 * comes to leave forty blank slides in someone's deck and say it left none.
 * When the deck will not give a count at all the claim is all there is and
 * stands — with `shrankBy` undefined, so nothing pretends otherwise.
 */
export function slidesActuallyReturned(o: {
  claimed: number;
  added: number;
  deckBefore?: number;
  deckAfter?: number;
}): { actually: number; left: number; shrankBy?: number } {
  const shrankBy = o.deckBefore !== undefined && o.deckAfter !== undefined ? o.deckBefore - o.deckAfter : undefined;
  const actually = shrankBy === undefined ? o.claimed : Math.min(o.claimed, Math.max(0, shrankBy));
  return { actually, left: o.added - actually, shrankBy };
}

export const SCRATCH_CLEANUP_ID = "scratch-slides-returned";

export const PROBE_IDS: readonly string[] = [...PROBES.flatMap(withFollows), SCRATCH_CLEANUP_ID];

/**
 * The questions that answer with no live scratch slide at all.
 *
 * Exported so the tests can ask the probes rather than carry a copy of this
 * list. The copy is what drifts: a hardcoded `excused` set is why marking five
 * genuinely slide-free questions turned two passing tests red for the right
 * reason, and a list that has to be edited by hand is a list that will one day
 * be edited wrong in the quiet direction.
 */
/**
 * The shortlist the later passes fall back to when scratch slides are scarce.
 *
 * Derived from the marks rather than listed again, for the same reason
 * `NO_SLIDE_NEEDED_IDS` is: a hand-kept copy is a copy that will one day be
 * edited wrong in the quiet direction.
 */
export const RESAMPLE_IDS: readonly string[] = PROBES.flatMap(flatten)
  .filter((p) => p.resample)
  .map((p) => p.id);

export const NO_SLIDE_NEEDED_IDS: readonly string[] = PROBES.flatMap(flatten)
  .filter((p) => p.noSlideNeeded)
  .map((p) => p.id);

/**
 * The questions EVERY run puts, whatever the host says.
 *
 * A follow-up is conditional by definition, so "one answer per known id" stopped
 * being the sheet's invariant the moment the first one existed. The invariant
 * that replaced it is two-sided and both halves matter: every id in this list
 * appears in every sheet, and no sheet ever carries an id outside `PROBE_IDS`.
 * A single count could not express either.
 */
export const ALWAYS_ASKED_IDS: readonly string[] = [...PROBES.map((p) => p.id), SCRATCH_CLEANUP_ID];

/**
 * What a probe run FOUND, in one line, on screen.
 *
 * Every probe run so far has cost a round trip to establish nothing: download,
 * send, wait for someone to run `host-diff`, hear that the answers are the same
 * as last time. The comparison table is plain data, so the pane can do the diff
 * itself and say whether this run is worth sending at all.
 *
 * Divergences that are already DECLARED are counted separately from new ones.
 * A run that reproduces seven known disagreements is not news; a run with one
 * undeclared answer is the only kind worth anyone's attention, and it should
 * not have to be found by eye in a seventeen-row JSON file.
 */
export function summariseHostSheet(sheet: HostAnswerSheet): {
  asked: number;
  neverPut: string[];
  known: string[];
  fresh: string[];
  /**
   * Questions that answered one way and then another INSIDE this run.
   *
   * The fact `UNSTABLE_ANSWERS` was assembled by hand across ten rounds to
   * state, now stated by the round that saw it. A sheet that knows it disagreed
   * with itself and does not say so has only moved the finding somewhere nobody
   * looks.
   */
  unstable: string[];
} {
  const answers = Object.fromEntries(sheet.answers.map((a) => [a.id, a.answer]));
  const neverPut = sheet.answers.filter((a) => NOT_ASKED.has(a.answer)).map((a) => a.id);
  const { differ } = diffAnswers(answers, FAKE_BASELINE) as { differ: { id: string }[] };
  // `UNSTABLE_ANSWERS` counts as declared too, and leaving it out made the pane
  // cry wolf on a schedule. `shape-add-positional-slide-proxy` is a coin: it has
  // answered `yes, yes, threw, yes, yes, threw` across six rounds, and the fake
  // says `yes`. So every round the coin lands `threw` the pane announced
  // "NEW: shape-add-positional-slide-proxy" — news the sixth time it was seen,
  // about the one question this repo has documented at greatest length as
  // varying run to run. A declaration is a declaration wherever it lives.
  const declared = (id: string) => id in KNOWN_DIVERGENCES || id in UNSTABLE_ANSWERS;
  const known = differ.filter((d) => declared(d.id)).map((d) => d.id);
  const fresh = differ.filter((d) => !declared(d.id)).map((d) => d.id);
  const unstable = sheet.answers.filter((a) => a.stable === false).map((a) => a.id);
  return { asked: sheet.answers.length - neverPut.length, neverPut, known, fresh, unstable };
}

/** Whether a probe run is worth sending on — anything unexplained in it. */
export function sheetNeedsAttention(sheet: HostAnswerSheet): boolean {
  const s = summariseHostSheet(sheet);
  // Instability counts. A question that answered two ways inside one round is
  // the finding this project used to need ten rounds and a hand-kept table to
  // reach, and the verdict line already says so — a round that says "worth
  // sending" in words and "nothing to see" in its return value is worse than
  // either.
  return s.fresh.length > 0 || s.neverPut.length > 0 || s.unstable.length > 0;
}

/** The same summary, in the words the pane shows. */
export function describeHostSheet(sheet: HostAnswerSheet): string {
  const { asked, neverPut, known, fresh, unstable } = summariseHostSheet(sheet);
  const parts = [`${asked} of ${sheet.answers.length} questions answered`];
  if (neverPut.length) parts.push(`${neverPut.length} never put (${neverPut.join(", ")})`);
  if (known.length) parts.push(`${known.length} known divergence${known.length === 1 ? "" : "s"}`);
  if (fresh.length) parts.push(`NEW: ${fresh.join(", ")}`);
  // A question that changed its answer mid-round is worth sending whatever the
  // diff says: it is the one thing a single sheet could never report before.
  if (unstable.length) parts.push(`CHANGED ITS ANSWER MID-ROUND: ${unstable.join(", ")}`);
  // The verdict the owner actually needs: send it, or don't.
  parts.push(
    fresh.length || neverPut.length || unstable.length
      ? "Saved — worth sending."
      : "Saved. Nothing new — no need to send it.",
  );
  return parts.join(" · ");
}
