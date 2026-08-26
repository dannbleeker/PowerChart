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

export const EXPERIMENTS: Experiment[] = [groupedChildById];

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
