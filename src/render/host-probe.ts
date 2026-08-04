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
  deleteSlideById,
  isTimeout,
  requirementSets,
  ScratchSlideUnavailable,
  withProbeContext,
  type ProbeContext,
} from "./powerpoint";
import { trace } from "../core/trace";

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
};

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
        return { answer: "unreadable", detail: short(err) };
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
        return { answer: "unreadable", detail: short(err) };
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
        return { answer: "threw", detail: short(err) };
      }
    },
  },
  {
    id: "shape-proxy-survives-one-sync",
    question: "Is a shape proxy still usable one sync after it was created?",
    // office-js#2903 — the stale-proxy bug the whole `targetRef` design exists
    // for. If a host keeps proxies alive, a lot of re-fetching here is waste.
    ask: async (ctx) => {
      const shape = ctx.scratch.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle, {
        left: 10,
        top: 10,
        width: 20,
        height: 20,
      });
      await ctx.sync();
      await ctx.sync(); // a second round trip: this is what ages the proxy
      try {
        shape.load("id");
        await ctx.sync();
        const id = (shape as unknown as { id: string }).id;
        return { answer: typeof id === "string" && id ? "yes" : "unreadable" };
      } catch (err) {
        return { answer: "threw", detail: short(err) };
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
      for (let i = 0; i < 5; i++) {
        ctx.scratch.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle, {
          left: i * 5,
          top: 40,
          width: 4,
          height: 4,
        });
      }
      await ctx.sync();
      ctx.scratch.shapes.load("items/id");
      await ctx.sync();
      try {
        const n = ctx.scratch.shapes.items.length;
        // Reported as "at least 5" rather than an exact count: the slide
        // carries whatever earlier probes left on it, and the question is
        // whether the host UNDER-reports, not what the total happens to be.
        return { answer: n >= 5 ? "at-least-5" : `short-${n}`, detail: `items=${n}` };
      } catch (err) {
        return { answer: "unreadable", detail: short(err) };
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
        return { answer: "unreadable", detail: short(err) };
      }
    },
  },
  {
    id: "tags-add-same-key-twice",
    question: "Does writing the same tag key twice overwrite, or duplicate?",
    // Re-editing a chart rewrites POWERCHART_CONFIG on the same shape every
    // time. If a host appended instead of overwriting, a chart edited ten times
    // would carry ten configs and the reader would pick one arbitrarily.
    ask: async (ctx) => {
      const shape = ctx.scratch.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle, {
        left: 60,
        top: 10,
        width: 20,
        height: 20,
      });
      await ctx.sync();
      shape.tags.add("POWERCHART_PROBE", "first");
      await ctx.sync();
      shape.tags.add("POWERCHART_PROBE", "second");
      await ctx.sync();
      const tag = shape.tags.getItemOrNullObject("POWERCHART_PROBE");
      tag.load("value");
      await ctx.sync();
      try {
        const v = (tag as unknown as { value: string }).value;
        return {
          answer: v === "second" ? "overwrites" : v === "first" ? "keeps-first" : "other",
          detail: `value=${v}`,
        };
      } catch (err) {
        return { answer: "unreadable", detail: short(err) };
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
      const shape = ctx.scratch.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle, {
        left: 90,
        top: 10,
        width: 20,
        height: 20,
      });
      await ctx.sync();
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
      const shape = ctx.scratch.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle, {
        left: 120,
        top: 10,
        width: 20,
        height: 20,
      });
      shape.load("id");
      await ctx.sync();
      const id = (shape as unknown as { id: string }).id;
      (shape as unknown as { delete(): void }).delete();
      await ctx.sync();
      const gone = ctx.scratch.shapes.getItemOrNullObject(id);
      gone.load("id");
      await ctx.sync();
      try {
        const nul = (gone as unknown as { isNullObject: boolean }).isNullObject;
        return { answer: nul === true ? "reports-gone" : nul === false ? "still-there" : "unreadable" };
      } catch (err) {
        return { answer: "threw", detail: short(err) };
      }
    },
  },
  {
    id: "addgroup-returns-usable",
    question: "Is a group proxy usable in the same context that created it?",
    ask: async (ctx) => {
      const made = [0, 1].map((i) =>
        ctx.scratch.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle, {
          left: 150 + i * 25,
          top: 10,
          width: 20,
          height: 20,
        }),
      );
      await ctx.sync();
      try {
        const group = (
          ctx.scratch.shapes as unknown as { addGroup(shapes: unknown[]): { load(p: string): void; id: string } }
        ).addGroup(made);
        await ctx.sync();
        group.load("id");
        await ctx.sync();
        return { answer: typeof group.id === "string" && group.id ? "yes" : "unreadable" };
      } catch (err) {
        return { answer: "threw", detail: short(err) };
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
      const made = [0, 1].map((i) =>
        ctx.scratch.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle, {
          left: 200 + i * 25,
          top: 60,
          width: 20,
          height: 20,
        }),
      );
      await ctx.sync();
      try {
        const group = (
          ctx.scratch.shapes as unknown as {
            addGroup(shapes: unknown[]): { load(p: string): void; group: { shapes: { items: unknown[] } } };
          }
        ).addGroup(made);
        await ctx.sync();
        group.load("group/shapes/items/id");
        await ctx.sync();
        const n = group.group?.shapes?.items?.length;
        return { answer: n === 2 ? "two" : typeof n === "number" ? `reports-${n}` : "unreadable", detail: `n=${n}` };
      } catch (err) {
        return { answer: "threw", detail: short(err) };
      }
    },
  },
  {
    id: "tag-on-group-survives",
    question: "Does a tag written on a GROUP read back?",
    // Where a chart's config actually lives. Tags on a plain shape are covered
    // above; if a group behaves differently, every chart in every deck is
    // un-re-editable and nothing else in the probe would say so.
    ask: async (ctx) => {
      const made = [0, 1].map((i) =>
        ctx.scratch.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle, {
          left: 260 + i * 25,
          top: 60,
          width: 20,
          height: 20,
        }),
      );
      await ctx.sync();
      try {
        const group = (
          ctx.scratch.shapes as unknown as {
            addGroup(shapes: unknown[]): {
              tags: {
                add(k: string, v: string): void;
                getItemOrNullObject(k: string): { load(p: string): void; value: string };
              };
            };
          }
        ).addGroup(made);
        await ctx.sync();
        group.tags.add("POWERCHART_PROBE_GROUP", "kept");
        await ctx.sync();
        const tag = group.tags.getItemOrNullObject("POWERCHART_PROBE_GROUP");
        tag.load("value");
        await ctx.sync();
        const v = tag.value;
        return { answer: v === "kept" ? "yes" : "no", detail: `value=${String(v)}` };
      } catch (err) {
        return { answer: "threw", detail: short(err) };
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
        return { answer: "threw", detail: short(err) };
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
  let scratchId = await addScratchSlide();
  if (scratchId) scratchIds.push(scratchId);
  try {
    for (const probe of PROBES) {
      const started = Date.now();
      trace("probe", "asking", { id: probe.id });
      let result: { answer: string; detail?: string };
      if (!scratchId) {
        result = { answer: "no-scratch-slide", detail: "the host would not add a slide to work on" };
      } else {
        result = await ask(probe, scratchId);
        // A slide that stopped resolving costs one question, not the sheet.
        //
        // It cost thirteen once: a real host lost the scratch slide after the
        // first probe and the remaining thirteen questions each reported the
        // same failure as if it were their own answer. Thirteen apparent
        // divergences, one cause, and none of the questions actually asked.
        // Replacing the slide is the difference between a sheet that reports a
        // finding and a sheet that IS the finding.
        if (result.answer === "no-scratch-slide") {
          const replacement = await addScratchSlide();
          if (replacement) {
            scratchIds.push(replacement);
            scratchId = replacement;
            trace("probe", "replaced the scratch slide", { id: probe.id, scratchId: replacement });
            const retry = await ask(probe, replacement);
            // Only adopt a retry that actually got somewhere. A second failure
            // on a brand-new slide is a stronger statement than the first, and
            // overwriting it with "no-scratch-slide" would hide that.
            if (retry.answer !== "no-scratch-slide") result = retry;
            else scratchId = null;
          } else {
            scratchId = null;
          }
        }
      }
      const entry: HostAnswer = { id: probe.id, question: probe.question, ms: Date.now() - started, ...result };
      answers.push(entry);
      trace("probe", "answered", { id: probe.id, answer: entry.answer, ms: entry.ms });
    }
  } finally {
    // The scratch slides go back whatever happened. A diagnostic that litters
    // the user's deck is one they will stop running.
    for (const id of scratchIds) await deleteSlideById(id).catch(() => false);
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
async function ask(probe: Probe, scratchId: string): Promise<{ answer: string; detail?: string }> {
  try {
    return await withProbeContext(scratchId, PROBE_BUDGET_MS, probe.ask);
  } catch (err) {
    if (err instanceof ScratchSlideUnavailable) return { answer: "no-scratch-slide", detail: short(err) };
    return { answer: isTimeout(err) ? "silent" : "threw", detail: short(err) };
  }
}

/** Every question this build knows how to ask — for the fake's baseline test. */
export const PROBE_IDS: readonly string[] = PROBES.map((p) => p.id);
