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
  deckSlideIds,
  deleteSlideById,
  isTimeout,
  requirementSets,
  ScratchSlideUnavailable,
  withProbeContext,
  type ProbeContext,
} from "./powerpoint";
import { trace } from "../core/trace";
// @ts-expect-error — a plain .mjs table with no types, and deliberately the
// SAME one the CLI and the CI gate read. Two copies of it is how a claim
// quietly stops matching its check.
import { FAKE_BASELINE, KNOWN_DIVERGENCES, diffAnswers } from "../../scripts/host-baseline.mjs";

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
const PROBE_BUDGET_MS = 8_000;

/** What a probe does, given a scratch slide it is free to wreck. */
type Probe = {
  id: string;
  question: string;
  /** Answers, or throws — a throw is recorded as an answer, not a failure. */
  ask: (ctx: ProbeContext) => Promise<{ answer: string; detail?: string }>;
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
};

/**
 * Answers that mean "this question was never put", not "the host said so".
 *
 * None of them is in any probe's own vocabulary, which is the point: a diff
 * that mistook one for an answer would report a host divergence that nobody
 * ever asked about. Both were earned the same way, a round apart — see
 * `ProbeSetupFailed`.
 */
const NOT_ASKED = new Set(["no-scratch-slide", "no-scratch-shape"]);

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
  try {
    const shapes = ctx.scratch().shapes;
    const made = boxes.map((box) => shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle, box));
    awaitingSetupShapes = true;
    await ctx.sync();
    awaitingSetupShapes = false;
    if (load) {
      for (const s of made) s.load(load);
      awaitingSetupShapes = true;
      await ctx.sync();
      awaitingSetupShapes = false;
    }
    return made;
  } catch (err) {
    awaitingSetupShapes = false;
    throw new ProbeSetupFailed(short(err));
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
    question: "Can a shape be added through a slide proxy resolved a sync ago?",
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
  },
  {
    id: "shape-resolve-held-slide-proxy",
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
      const ids = idsOf(
        await scratchShapes(
          ctx,
          [0, 1].map((i) => ({ left: 150 + i * 25, top: 10, width: 20, height: 20 })),
          "id",
        ),
      );
      try {
        const group = (
          probeShapes(ctx) as unknown as { addGroup(shapes: unknown[]): { load(p: string): void; id: string } }
        ).addGroup(ids.map((id) => probeShape(ctx, id)));
        group.load("id");
        await ctx.sync();
        return { answer: typeof group.id === "string" && group.id ? "yes" : "unreadable" };
      } catch (err) {
        return threw(err);
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
      const ids = idsOf(
        await scratchShapes(
          ctx,
          [0, 1].map((i) => ({ left: 200 + i * 25, top: 60, width: 20, height: 20 })),
          "id",
        ),
      );
      try {
        const group = (
          probeShapes(ctx) as unknown as {
            addGroup(shapes: unknown[]): { load(p: string): void; group: { shapes: { items: unknown[] } } };
          }
        ).addGroup(ids.map((id) => probeShape(ctx, id)));
        // Queued in the batch that MAKES the group, so the answer arrives on
        // the group's own sync. Asked one sync later — the first version — this
        // host answered PropertyNotLoaded, which is a statement about proxy age
        // and not about what the group contains.
        group.load("group/shapes/items/id");
        await ctx.sync();
        const n = group.group?.shapes?.items?.length;
        return { answer: n === 2 ? "two" : typeof n === "number" ? `reports-${n}` : "unreadable", detail: `n=${n}` };
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
      const ids = idsOf(
        await scratchShapes(
          ctx,
          [0, 1].map((i) => ({ left: 300 + i * 25, top: 160, width: 20, height: 20 })),
          "id",
        ),
      );
      try {
        const made = (
          probeShapes(ctx) as unknown as { addGroup(shapes: unknown[]): { load(p: string): void; id: string } }
        ).addGroup(ids.map((id) => probeShape(ctx, id)));
        // The id in its OWN sync, never in the one that made the group.
        // That saving is what cost the 2026-08-05 sheet five of its questions:
        // this host will not name a shape in the batch that created it, and
        // `scratchShapes` was rewritten for exactly the same reason.
        await ctx.sync();
        made.load("id");
        await ctx.sync();
        const groupId = readShapeId(made);
        if (!groupId) throw new ProbeSetupFailed("the host would not name the group it just made");
        // A LATER batch, through a fresh handle — the thing that separates this
        // question from the one above it.
        const found = probeShape(ctx, groupId) as unknown as {
          group: { shapes: { getCount(): { value: number } } };
        };
        const count = found.group.shapes.getCount();
        await ctx.sync();
        const n = loadedCount(count);
        return typeof n === "number" ? { answer: String(n) } : { answer: "unreadable" };
      } catch (err) {
        return threw(err);
      }
    },
  },
  {
    id: "slide-layout-readable",
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
  const scratchIds: string[] = [];
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
  let scratchId = await addScratchSlide();
  if (scratchId) scratchIds.push(scratchId);
  try {
    for (const probe of PROBES) {
      const started = Date.now();
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
        const recovered = await addScratchSlide();
        if (recovered) {
          scratchIds.push(recovered);
          scratchId = recovered;
          trace("probe", "took another scratch slide after giving up on the last", { id: probe.id });
        }
      }
      if (!scratchId) {
        result = { answer: "no-scratch-slide", detail: "the host would not add a slide to work on" };
      } else {
        result = await ask(probe, scratchId, durableSlideId);
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
          const replacement = await addScratchSlide();
          if (replacement) {
            scratchIds.push(replacement);
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
      const entry: HostAnswer = { id: probe.id, question: probe.question, ms: Date.now() - started, ...result };
      answers.push(entry);
      trace("probe", "answered", { id: probe.id, answer: entry.answer, ms: entry.ms });
      // The partner question, in the same run, when the answer admits two
      // readings. Never asked when the question was never PUT: a follow-up to
      // `no-scratch-shape` would be a second question about the probe's own
      // setup, dressed as a fact about the host.
      const follow = probe.follow;
      if (follow && !NOT_ASKED.has(entry.answer) && follow.when(entry.answer)) {
        const at = Date.now();
        trace("probe", "asking the partner question", { after: probe.id, id: follow.probe.id, was: entry.answer });
        const r = scratchId
          ? await ask(follow.probe, scratchId, durableSlideId)
          : { answer: "no-scratch-slide", detail: "the host would not add a slide to work on" };
        answers.push({
          id: follow.probe.id,
          question: follow.probe.question,
          ms: Date.now() - at,
          ...r,
          // Which answer triggered it, in the sheet itself. Two rows that do
          // not say they are a pair are two unrelated facts to whoever reads
          // them a week later.
          detail: `asked because ${probe.id} answered "${entry.answer}" — ${follow.because}${r.detail ? `; ${r.detail}` : ""}`,
        });
        trace("probe", "partner answered", { id: follow.probe.id, answer: r.answer });
      }
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
    let returned = 0;
    for (const id of scratchIds) if (await deleteSlideById(id).catch(() => false)) returned++;
    const left = scratchIds.length - returned;
    answers.push({
      id: SCRATCH_CLEANUP_ID,
      question: "Does the host give back the scratch slides this probe added?",
      answer: !scratchIds.length ? "none-added" : !left ? "all" : returned ? "some" : "none",
      ms: Date.now() - cleanupStarted,
      detail: `${returned} of ${scratchIds.length} scratch slide(s) deleted${left ? `; ${left} left in the deck` : ""}`,
    });
    trace("probe", "gave the scratch slides back", { returned, left });
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
    return await withProbeContext(scratchId, PROBE_BUDGET_MS, probe.ask, durableSlideId);
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
 * The cleanup's own row in the sheet.
 *
 * Not a `Probe` — it has no question to put and no scratch slide to put it on,
 * because it IS what happens to the scratch slides. It is still a fact about
 * this host, it still has a fake answer to diverge from, and the sheet is the
 * only place anyone will read it. So it is listed here by hand rather than
 * derived, and the two lists below are where it joins the invariant.
 */
export const SCRATCH_CLEANUP_ID = "scratch-slides-returned";

export const PROBE_IDS: readonly string[] = [...PROBES.flatMap(withFollows), SCRATCH_CLEANUP_ID];

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
} {
  const answers = Object.fromEntries(sheet.answers.map((a) => [a.id, a.answer]));
  const neverPut = sheet.answers.filter((a) => NOT_ASKED.has(a.answer)).map((a) => a.id);
  const { differ } = diffAnswers(answers, FAKE_BASELINE) as { differ: { id: string }[] };
  const known = differ.filter((d) => d.id in KNOWN_DIVERGENCES).map((d) => d.id);
  const fresh = differ.filter((d) => !(d.id in KNOWN_DIVERGENCES)).map((d) => d.id);
  return { asked: sheet.answers.length - neverPut.length, neverPut, known, fresh };
}

/** Whether a probe run is worth sending on — anything unexplained in it. */
export function sheetNeedsAttention(sheet: HostAnswerSheet): boolean {
  const s = summariseHostSheet(sheet);
  return s.fresh.length > 0 || s.neverPut.length > 0;
}

/** The same summary, in the words the pane shows. */
export function describeHostSheet(sheet: HostAnswerSheet): string {
  const { asked, neverPut, known, fresh } = summariseHostSheet(sheet);
  const parts = [`${asked} of ${sheet.answers.length} questions answered`];
  if (neverPut.length) parts.push(`${neverPut.length} never put (${neverPut.join(", ")})`);
  if (known.length) parts.push(`${known.length} known divergence${known.length === 1 ? "" : "s"}`);
  if (fresh.length) parts.push(`NEW: ${fresh.join(", ")}`);
  // The verdict the owner actually needs: send it, or don't.
  parts.push(fresh.length || neverPut.length ? "Saved — worth sending." : "Saved. Nothing new — no need to send it.");
  return parts.join(" · ");
}
