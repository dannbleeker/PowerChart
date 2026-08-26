/**
 * ONE QUESTION, ON DEMAND, WITHOUT A ROUND.
 *
 * A round costs fourteen minutes and answers thirty-odd questions at once. That
 * is the right shape for watching the product, and the wrong shape for settling
 * a single "does this host do X?" that a change depends on.
 *
 * The gap is not theoretical. `grouped-child-by-id-from-slide` sat in the probe
 * sheet for 125 rounds and was never once answered — it followed a question that
 * burns its scratch slide, so it always found the slide gone — and was retired
 * as moot. It is live again (see `docs/BACKLOG.md`), and the alternative to this
 * file was adding a speculative host call to the drawing batch, which is the
 * most trap-laden path in the renderer: loading an id on a creation handle
 * poisons it into `shapes.getItem(id)`, the mechanism behind 235 tagging
 * failures.
 *
 * So: experiments run from the pane, on their own scratch slide, and touch no
 * production path. They are not scenarios — nothing here asserts the product
 * works — and they are not probes, which run as a fixed sheet on a schedule.
 * They are the thing you reach for when a decision needs one fact.
 *
 * Rules that make them cheap enough to keep:
 *
 *   - Own scratch slide, always given back, even on the failure path.
 *   - Answer in a fixed vocabulary, like the probe sheet, so an archive of them
 *     can be compared rather than read.
 *   - Say what was actually observed in `detail`, because the vocabulary will be
 *     wrong for something eventually and the detail is what survives that.
 */
import { addScratchSlide, deleteSlideById, deleteShapesById, withProbeContext } from "./powerpoint";
import { trace } from "../core/trace";

export interface ExperimentResult {
  id: string;
  asks: string;
  /** A fixed word, so a run can be compared with the last one. */
  answer: string;
  /** What was actually seen — the half that survives a wrong vocabulary. */
  detail?: string;
  ms: number;
}

export interface Experiment {
  id: string;
  asks: string;
  /** Given a scratch slide it may use freely; it will be taken back afterwards. */
  run(slideId: string): Promise<{ answer: string; detail?: string }>;
}

/** How long one experiment may take before it counts as unanswered. */
const EXPERIMENT_BUDGET_MS = 20_000;

/**
 * Read a property that may not be loaded, without taking the experiment down.
 *
 * An unloaded read is `PropertyNotLoaded` on a real host, and that is an ANSWER
 * here rather than a crash — "the host populated nothing" is one of the things
 * these questions are asking about.
 */
const read = <T>(f: () => T): T | undefined => {
  try {
    return f();
  } catch {
    return undefined;
  }
};

/**
 * Can a shape INSIDE a group still be resolved by id off the slide?
 *
 * THE QUESTION THAT DECIDES A REAL FIX. Two of every three redraws in the
 * archive are a grouped chart the update cannot map to its shapes: the group
 * route (`shape.group.shapes`) answers `unreadable` in 34 of 40 rounds, and a
 * grouped chart carries no parts list. If a child resolves by id, those charts
 * can be updated in place instead of redrawn; if it does not, the fix is dead
 * and nobody should spend a week on it.
 *
 * Asked WITHOUT the drawing batch on purpose. Plain shapes grouped on a scratch
 * slide put the same question to the host — is a grouped child addressable by id
 * — while touching none of the renderer's trap-laden paths.
 *
 * The write matters as much as the read. An id that resolves and then refuses a
 * property change buys nothing: the update has to WRITE through these proxies.
 */
const groupedChildById: Experiment = {
  id: "grouped-child-by-id",
  asks: "Can a shape inside a group still be resolved, and written to, by id off the slide?",
  async run(slideId) {
    return withProbeContext(slideId, EXPERIMENT_BUDGET_MS, async (ctx) => {
      // THE DECK'S OWN FIRST SLIDE, not the scratch slide, and the scratch slide
      // is still held so the runner's bookkeeping is unchanged.
      //
      // Measured live on 2026-08-26, three times: on a scratch slide this host
      // names ZERO of three shapes — by `load("id")`, by reading `.id`, and by a
      // collection read alike. That is not news, it is
      // `shapes-items-count-honest` answering `unreadable` in 92% of rounds, and
      // the probe file says why in as many words: that question "is measured on
      // the SCRATCH SLIDE and is strictly worse than a real one".
      //
      // Production reads collections off real slides constantly and they answer —
      // 2,135 charts grouped by id-match. So the scratch slide cannot host this
      // question at all, and asking it there measures the slide rather than the
      // group.
      const target = () => ctx.slides.getItemAt(0);
      // Its id, captured ONCE, because the clean-up needs to name the slide and
      // a handle cannot be held across the syncs between here and there.
      const named = target();
      named.load("id");
      await ctx.sync();
      const targetSlideId = read(() => (named as unknown as { id?: string }).id);
      const made = [0, 1, 2].map((n) =>
        target().shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle, {
          left: 20 + n * 40,
          top: 20,
          width: 30,
          height: 30,
        }),
      );
      // ONE SYNC, AND NO `load("id")`. Measured live on 2026-08-26: create three
      // shapes, sync, call `load("id")`, sync again — this host names ZERO of
      // three. That is not the experiment failing, it is the wall
      // `shape-resolve-held-slide-proxy` hit for 216 rounds; `scratchShapes`
      // takes exactly that route and starved every time.
      //
      // The renderer knows the way round it and says so in as many words:
      // "`loadedValue(() => sh.id)` already answers on `it.created` — the
      // drawing batch's own sync populated them, and reading a populated
      // property issues no host call at all". Asking with `load()` is what
      // rewrites a creation proxy into `shapes.getItem(id)`, which this host
      // refuses — the mechanism behind 235 tagging failures.
      await ctx.sync();
      // A COLLECTION READ, which is the one route this host honours.
      //
      // Measured live, twice, in under a minute each: `load("id")` on the
      // creation proxies names 0 of 3, and reading `.id` off them without a load
      // names 0 of 3 as well. So the renderer's `withOwnId: 7 of 7` does not come
      // from either — it comes from the pre-grouping re-read, which asks the
      // SLIDE for its shapes rather than asking each shape for itself.
      //
      // The asymmetry is written down all over this repo: "a by-id lookup is the
      // one thing PowerPoint on the web reliably refuses, and a collection read
      // is the one thing it reliably honours". The scratch slide is empty until
      // this experiment fills it, so what comes back IS the three shapes.
      const collection = target().shapes;
      collection.load("items/id");
      await ctx.sync();
      const members = read(() => collection.items) ?? [];
      const ids = members
        .map((s) => read(() => (s as unknown as { id?: string }).id))
        .filter((id): id is string => typeof id === "string" && id.length > 0);
      if (ids.length < made.length) {
        return { answer: "no-child-ids", detail: `named ${ids.length} of ${made.length} before grouping` };
      }
      // GROUPED FROM THE COLLECTION READ, not from the creation proxies.
      //
      // The renderer does the same — `freshMembers` — and the reason is the one
      // this file keeps meeting: a creation proxy whose path Office.js has
      // rewritten to `shapes.getItem(id)` is refused by this host, and the fake
      // models it, answering `InvalidParam passed to GetItem(id)`, code 5010,
      // errorLocation `ShapeCollection.getItem`. Handing those proxies to
      // `addGroup` throws before the question is ever put.
      let group;
      try {
        group = (target().shapes as unknown as { addGroup(shapes: unknown[]): PowerPoint.Shape }).addGroup(members);
        group.load("id");
        await ctx.sync();
      } catch (err) {
        return {
          answer: "no-group",
          detail: `grouping ${ids.length} shapes threw: ${String((err as { message?: string })?.message ?? err).slice(0, 140)}`,
        };
      }
      const groupId = read(() => (group as unknown as { id?: string }).id);
      // THE QUESTION. A fresh proxy off the slide, by the id the child had
      // before the group swallowed it.
      //
      // CAUGHT, because a refusal here IS the answer. This host answers
      // `InvalidParam passed to GetItem(id)`, code 5010, errorLocation
      // `ShapeCollection.getItem` — the same 5010 that blocks the parts list —
      // and letting that escape as `threw` would file the answer under "the
      // experiment broke" instead of under "the host said no".
      let child;
      try {
        child = target().shapes.getItemOrNullObject(ids[0]);
        child.load("id");
        await ctx.sync();
      } catch (err) {
        return {
          answer: "refused-by-id",
          detail: `resolving child ${ids[0]} of group ${String(read(() => (group as unknown as { id?: string }).id))} threw: ${String((err as { message?: string })?.message ?? err).slice(0, 140)}`,
        };
      }
      const gone = read(() => (child as unknown as { isNullObject?: boolean }).isNullObject) === true;
      const back = read(() => (child as unknown as { id?: string }).id);
      if (gone || back !== ids[0]) {
        return {
          answer: gone ? "no-such-shape" : "unreadable",
          detail: `child ${ids[0]} of group ${String(groupId)} read back as ${String(back)}`,
        };
      }
      // AND CAN IT BE WRITTEN TO? An id that resolves and then refuses a
      // property change buys nothing — the update writes through these proxies.
      try {
        (child as unknown as { fill: { setSolidColor(c: string): void } }).fill.setSolidColor("FF0000");
        await ctx.sync();
        // TAKE IT BACK OFF THE USER'S SLIDE. The scratch slide is returned by the
        // runner; shapes put on a real slide are this experiment's to remove.
        if (groupId && targetSlideId) await deleteShapesById(targetSlideId, [groupId]).catch(() => 0);
        return { answer: "yes", detail: `child ${ids[0]} resolved and took a fill inside group ${String(groupId)}` };
      } catch (err) {
        return {
          answer: "reads-but-refuses-writes",
          detail: `child ${ids[0]} resolved; the write threw: ${String((err as { message?: string })?.message ?? err).slice(0, 120)}`,
        };
      }
    });
  },
};

/**
 * Does a tag written on a SLIDE read back — and does writing it twice overwrite?
 *
 * PRODUCTION DEPENDS ON THIS. `DEMO_SLOT_TAG` is written onto slides and read
 * back to find which slot a demo slide belongs to; if a slide tag does not land,
 * that path is silently broken.
 *
 * The probe sheet has asked a version of this 224 times and answered `other`
 * every time — the word this project classes as UNINFORMATIVE — with
 * `value=undefined`. Since 2026-08-26 it also says "no such tag on the slide",
 * which reads as "the write did not land".
 *
 * IT IS ASKED ON THE SCRATCH SLIDE, and that is now a known-bad environment for
 * questions like this: the same slide answers `unreadable` for its shape
 * collection in 92% of rounds, and the grouped-child question could not be
 * asked there at all. So the probe's answer may be about the slide rather than
 * about tags, and production writes its slot tags to REAL slides.
 *
 * Same question, real slide. The tag is removed afterwards.
 */
const slideTagRoundTrip: Experiment = {
  id: "slide-tag-round-trip",
  asks: "Does a tag written on a real slide read back, and does a second write overwrite it?",
  async run(slideId) {
    // SLIDELESS. This question works on the deck's first slide, so charging it
    // the scratch slide's liveness check would fail it for a reason that has
    // nothing to do with tags — the same mistake that made
    // `shape-resolve-held-slide-proxy` answer `no-scratch-slide` for years.
    return withProbeContext(
      slideId,
      EXPERIMENT_BUDGET_MS,
      async (ctx) => {
        const KEY = "POWERCHART_EXPERIMENT";
        const target = () => ctx.slides.getItemAt(0);
        const tags = () => (target() as unknown as { tags: { add(k: string, v: string): void } }).tags;
        tags().add(KEY, "first");
        await ctx.sync();
        tags().add(KEY, "second");
        await ctx.sync();
        const back = (
          target() as unknown as {
            tags: { getItemOrNullObject(k: string): { load(p: string): void; value?: string; isNullObject?: boolean } };
          }
        ).tags.getItemOrNullObject(KEY);
        back.load("value");
        await ctx.sync();
        const absent = read(() => back.isNullObject) === true;
        const value = read(() => back.value);
        // Best-effort tidy: a tag left on the user's first slide is litter, and
        // never allowed to change what this reports.
        try {
          (target() as unknown as { tags: { delete(k: string): void } }).tags.delete(KEY);
          await ctx.sync();
        } catch {
          /* the next round's sweep does not touch tags; this is a nuisance only */
        }
        if (absent) return { answer: "no-such-tag", detail: "the write did not land on a real slide either" };
        if (value === "second") return { answer: "overwrites", detail: "second write replaced the first" };
        if (value === "first") return { answer: "keeps-first", detail: "the second write did not take" };
        return { answer: "unreadable", detail: `tag present, value read back as ${String(value)}` };
      },
      undefined,
      true,
    );
  },
};

/**
 * Can a group ENUMERATE its children on a real slide?
 *
 * THE LINCHPIN OF TODAY'S CONCLUSION, and it had not been checked where it
 * matters. `grouped-child-by-id` proved a group's child is not addressable BY ID
 * on a real slide, and production reports "no readable group members" on 56 of
 * 86 redraws — but the probe's own version of this question,
 * `group-children-via-getcount`, is measured on the SCRATCH slide, which today
 * has been shown three separate times to answer differently from a real one.
 *
 * If `shape.group.shapes` DOES enumerate on a real slide, then a grouped chart's
 * update can map its nodes after all, the 56 redraws per 30 rounds are
 * avoidable, and "the redraw is the only correct behaviour available" — now
 * written into the backlog and `WHAT-WE-KNOW.md` — is wrong.
 *
 * Two seconds to check, against a conclusion resting on a slide known to
 * misreport.
 */
const groupMembersOnRealSlide: Experiment = {
  id: "group-members-real-slide",
  asks: "Can a group enumerate its own children on a real slide?",
  async run(slideId) {
    return withProbeContext(
      slideId,
      EXPERIMENT_BUDGET_MS,
      async (ctx) => {
        const target = () => ctx.slides.getItemAt(0);
        const named = target();
        named.load("id");
        await ctx.sync();
        const targetSlideId = read(() => (named as unknown as { id?: string }).id);
        for (const n of [0, 1, 2]) {
          target().shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle, {
            left: 20 + n * 40,
            top: 90,
            width: 30,
            height: 30,
          });
        }
        await ctx.sync();
        const collection = target().shapes;
        collection.load("items/id");
        await ctx.sync();
        const members = read(() => collection.items) ?? [];
        if (members.length < 3) return { answer: "no-child-ids", detail: `slide listed ${members.length} shapes` };
        let group;
        try {
          group = (target().shapes as unknown as { addGroup(shapes: unknown[]): PowerPoint.Shape }).addGroup(members);
          group.load("id");
          await ctx.sync();
        } catch (err) {
          return { answer: "no-group", detail: String((err as { message?: string })?.message ?? err).slice(0, 140) };
        }
        const groupId = read(() => (group as unknown as { id?: string }).id);
        // THE QUESTION, asked exactly as production asks it: `shape.group.shapes`
        // loaded for `items/id`, which is what `queueGroupMembers` does.
        // Declared without a placeholder value: every path below assigns both,
        // and a default here would be a value nothing can ever read — which is
        // also what the linter says.
        let answer: string;
        let detail: string;
        try {
          const inner = (group as unknown as { group: { shapes: { load(p: string): void; items?: unknown[] } } }).group
            .shapes;
          inner.load("items/id");
          await ctx.sync();
          const kids = read(() => inner.items) ?? [];
          const withIds = kids
            .map((k) => read(() => (k as { id?: string }).id))
            .filter((id): id is string => typeof id === "string" && id.length > 0);
          // THREE OUTCOMES, not two. A group that lists children without naming
          // them is not the same as one that lists none: the first is a host that
          // half-answers, the second is the flat refusal production sees.
          answer = withIds.length ? "yes" : kids.length ? "listed-but-unnamed" : "empty";
          detail = `group ${String(groupId)} listed ${kids.length} child(ren), ${withIds.length} named`;
        } catch (err) {
          answer = "threw";
          detail = String((err as { message?: string })?.message ?? err).slice(0, 140);
        }
        if (groupId && targetSlideId) await deleteShapesById(targetSlideId, [groupId]).catch(() => 0);
        return { answer, detail };
      },
      undefined,
      true,
    );
  },
};

export const EXPERIMENTS: Experiment[] = [groupedChildById, slideTagRoundTrip, groupMembersOnRealSlide];

/**
 * Run one experiment, on a slide of its own, and give the slide back.
 *
 * The slide comes back on EVERY path — answered, refused, or thrown. An
 * experiment that litters is one nobody runs twice, and this one is meant to be
 * run casually while a decision is being made.
 */
export async function runExperiment(
  id: string,
  add = addScratchSlide,
  remove = deleteSlideById,
): Promise<ExperimentResult> {
  const experiment = EXPERIMENTS.find((e) => e.id === id);
  if (!experiment) return { id, asks: "(unknown)", answer: "no-such-experiment", ms: 0 };
  const started = Date.now();
  const slideId = await add();
  if (!slideId) {
    return {
      id: experiment.id,
      asks: experiment.asks,
      answer: "no-scratch-slide",
      detail: "the host would not add a slide to work on",
      ms: Date.now() - started,
    };
  }
  try {
    const { answer, detail } = await experiment.run(slideId);
    const result = { id: experiment.id, asks: experiment.asks, answer, detail, ms: Date.now() - started };
    trace("experiment", "an experiment answered", { ...result });
    return result;
  } catch (err) {
    const detail = String((err as { message?: string })?.message ?? err).slice(0, 200);
    trace("experiment", "an experiment threw", { id: experiment.id, error: detail });
    return { id: experiment.id, asks: experiment.asks, answer: "threw", detail, ms: Date.now() - started };
  } finally {
    // Best-effort, and never allowed to mask the answer: a slide left behind is
    // a nuisance, an experiment that reports a cleanup failure as its result is
    // a lie about the host.
    try {
      await remove(slideId);
    } catch {
      /* the sweep in the next round will find it */
    }
  }
}
