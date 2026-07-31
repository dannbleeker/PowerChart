/**
 * Reconciliation: deciding what a demo run ACTUALLY produced, after the host
 * has stopped moving.
 *
 * The inline checks in `insertDemoDeck` all fire while the host is still
 * catching up — `waitForLateSync` gives an abandoned sync five seconds, and a
 * real PowerPoint-web run has been observed committing a batch minutes after
 * the 45s timeout that gave up on it. Every readback taken in that window is a
 * snapshot of a host mid-flight, and acting on it produces the exact damage
 * seen in `Presentation_4.pptx`:
 *
 *   - Two identical Line charts (slides 4 and 5) and two identical Gantt charts
 *     (slides 6 and 7). Attempt 1 timed out, its readback showed nothing yet,
 *     so a retry drew the chart again on a NEW slide — and then BOTH attempts'
 *     shapes committed.
 *   - A `NOT COMPLETE` banner sitting on top of a chart that is, in fact,
 *     complete. The stamp was written while the slide still looked empty.
 *   - An Agenda slide with all 13 of its shapes present, reported BLANK.
 *   - One `POWERCHART_CONFIG` tag across five charts, because the grouping sync
 *     ran at the most loaded moment of the run.
 *
 * None of that is fixable with a better timeout, because the host's commit is
 * not bounded by anything we can observe. It IS fixable afterwards: once the
 * run has finished and the host has settled, the deck itself says what
 * happened. Every demo slide carries a `POWERCHART_DEMO_SLOT` tag naming the
 * item it belongs to, so a single settled pass over the added range can pair
 * slides with intent and repair the difference.
 *
 * This module is that decision, kept pure: snapshots in, a plan of actions out.
 * No Office.js, no I/O, no clock — so every rule below is unit-testable at the
 * speed of a function call, which is the only way any of this gets tested at
 * all (the host it exists to survive cannot be run in CI). `applyReconcilePlan`
 * in `src/render/powerpoint.ts` is the thin part that talks to PowerPoint.
 */

/** The banner shape `stampSlide` writes on a slide it believes came out short. */
export const NOT_COMPLETE_NAME = "PowerChart:not-complete";

/**
 * Minimum content-vs-expected ratio for a slide to count as a real, if
 * incomplete, chart rather than wreckage. Mirrors `PARTIAL_RENDER_THRESHOLD`
 * in the renderer: at 85%, one dropped 40-shape batch out of a typical
 * ~40-shape chart is still a chart worth keeping and grouping.
 */
export const PARTIAL_CONTENT_THRESHOLD = 0.85;

/**
 * One slide in the run's added range, as read back AFTER the run — top-level
 * shape count, the slot tag it carries (if the host kept it), and the two
 * facts that decide repair: does it wear a NOT COMPLETE banner, and does it
 * carry a config tag.
 */
export interface SlideSnapshot {
  /** 0-based deck index at snapshot time. */
  index: number;
  /** Item index from `POWERCHART_DEMO_SLOT`, or null when the tag is absent. */
  slot: number | null;
  /** Item title from the slot tag, for reporting. Null when untagged. */
  title: string | null;
  /**
   * Which RUN wrote this slide — the token from the slot tag. Null when the
   * tag predates run ids or carries none.
   *
   * Slot indices restart at 0 for every run, and the demo deck's titles are
   * fixed, so slot+title identifies an item but NOT an occurrence of it. Insert
   * the demo deck twice into one presentation and every slide of run 2 has a
   * perfect twin in run 1 — which a repair pass reads as "this item landed
   * twice" and fixes by deleting one of them. A whole healthy run, deleted for
   * being a duplicate of another healthy run. PowerPoint's own Duplicate Slide
   * copies the tag too, so the same thing happens to a demo slide a user
   * duplicated to reuse.
   */
  run?: string | null;
  /** Top-level shapes on the slide, INCLUDING any banner. */
  shapes: number;
  /**
   * Children of the PowerChart group, when the item landed grouped. A grouped
   * chart is ONE top-level shape, so `shapes` alone would read a perfect
   * 36-shape chart as 1 and condemn it as wreckage. Absent when nothing is
   * grouped. Not proof of completeness on its own — `rescueGroupAndTag` groups
   * whatever landed, partial charts included — which is why it is a count and
   * not a flag.
   */
  groupChildren?: number;
  /**
   * A PowerChart group is present on the slide. Set even when its child count
   * could not be read (a host without PowerPointApi 1.8 exposes no way in),
   * because grouping only ever runs AFTER a render finished — so the group
   * itself is evidence the item landed, and a slide we cannot measure is a
   * slide we must not "repair".
   */
  grouped?: boolean;
  /** A `PowerChart:not-complete` banner is present. */
  stamped: boolean;
  /** A `POWERCHART_CONFIG` tag is present, i.e. the chart is re-editable. */
  tagged: boolean;
}

/** What the run set out to draw, one entry per demo item. */
export interface ExpectedItem {
  /** Item index — matches `SlideSnapshot.slot`. */
  slot: number;
  title: string;
  /** `estimateOfficeShapes(scene)` — what a complete render puts on the slide. */
  shapes: number;
  /** The item is a chart with config to re-edit, so it should end up tagged. */
  chart: boolean;
  /** Deliberately not rendered (over the shape budget) — its banner is honest. */
  skipped?: boolean;
}

/**
 * One repair. Every action names the slide it targets by deck index at PLAN
 * time; `applyReconcilePlan` therefore applies deletes in descending index
 * order, since deleting slide 5 renumbers everything after it.
 */
export type ReconcileAction =
  /** Remove a banner that contradicts the shapes underneath it. */
  | { kind: "unstamp"; index: number; slot: number | null; reason: string }
  /** Group + tag a chart the run left loose, making it re-editable again. */
  | { kind: "regroup"; index: number; slot: number; reason: string }
  /**
   * Write the config tag onto a chart that is already ONE object but carries
   * no tag — a degraded picture, or a group whose tagging sync was dropped.
   * There is nothing to group; only the tag is missing.
   */
  | { kind: "retag"; index: number; slot: number; reason: string }
  /** Delete a slide — only ever a provably redundant twin or an empty stray. */
  | { kind: "delete"; index: number; slot: number | null; reason: string };

/** What became of one expected item, once the deck was actually read. */
export interface SlotVerdict {
  slot: number;
  title: string;
  /**
   * - `rendered` — a slide carries at least the expected shape count.
   * - `partial` — a slide carries most of it (see PARTIAL_CONTENT_THRESHOLD).
   * - `wreckage` — a slide exists but holds less than that.
   * - `empty` — a slide exists with no content shapes at all.
   * - `lost` — no slide in the added range claims this item.
   * - `skipped` — intentionally not drawn; the banner stands.
   */
  status: "rendered" | "partial" | "wreckage" | "empty" | "lost" | "skipped";
  /** Content shapes on the kept slide (banner excluded). 0 when lost. */
  shapes: number;
  /** Expected shape count, so a report can say "31 of 36". */
  expected: number;
  /** Deck index of the slide kept for this item, or null when lost. */
  index: number | null;
  /** Re-editable: a config tag is present, or a regroup action will add one. */
  tagged: boolean;
  /** Slides that claimed this slot beyond the one kept. */
  duplicates: number;
}

/** A slide inside the added range that no expected item claims. */
export interface OrphanSlide {
  index: number;
  shapes: number;
  stamped: boolean;
  /** True when it holds no content and is safe to drop. */
  blank: boolean;
}

/** The whole verdict: what to do, and what the run really produced. */
export interface ReconcilePlan {
  /** Apply deletes in descending `index` — see `ReconcileAction`. */
  actions: ReconcileAction[];
  verdicts: SlotVerdict[];
  orphans: OrphanSlide[];
  summary: {
    items: number;
    rendered: number;
    partial: number;
    lost: number;
    skipped: number;
    /** Items whose slide holds content but too little to call a chart. */
    wreckage: number;
    /** Items whose kept slide has no content shapes at all. */
    empty: number;
    /** Slides queued for deletion as redundant twins. */
    duplicates: number;
    /** Banners queued for removal because the chart underneath is real. */
    falseStamps: number;
    /** Charts queued for group + tag, i.e. currently not re-editable. */
    untagged: number;
    /** Slides in range claimed by no item. */
    orphans: number;
  };
}

export interface ReconcileOptions {
  /**
   * Delete untagged, contentless slides inside the added range. Off by
   * default: an empty stray is still a slide the user can see, and deleting
   * something we cannot positively identify is the one mistake a repair pass
   * must never make. The pane turns it on for its own run's range, where every
   * slide is known to be ours.
   */
  dropOrphanBlanks?: boolean;
  /**
   * This run's token. When set, a slide is paired to an item only if its slot
   * tag carries the SAME token — see `SlideSnapshot.run`. Slides from an
   * earlier run, or from a slide the user duplicated, then fall through to
   * `orphans`: reported, never claimed, never deleted as somebody's twin.
   *
   * Unset means "match on slot and title alone", which is only safe when the
   * caller has already bounded the snapshots to one run's own added range.
   */
  run?: string;
}

/**
 * Shapes that are actually the item's: the group's children when it landed
 * grouped, otherwise the top-level count less the banner.
 *
 * Returns `UNKNOWN_CONTENT` for a grouped slide whose children could not be
 * counted — the one case where the honest answer is "we cannot tell", and
 * every rule below is written to leave such a slide alone rather than guess.
 */
const UNKNOWN_CONTENT = -1;

function contentShapes(s: SlideSnapshot): number {
  if (s.groupChildren !== undefined) return s.groupChildren;
  if (s.grouped) return UNKNOWN_CONTENT;
  return Math.max(0, s.shapes - (s.stamped ? 1 : 0));
}

/**
 * Pick the slide to keep when several claim one item: most content wins, then
 * an unstamped slide over a stamped one, then the earliest — so a deck's
 * reading order is preserved when two attempts landed the same chart.
 */
function pickKeeper(group: SlideSnapshot[]): SlideSnapshot {
  // A grouped slide we cannot measure still outranks a loose pile of shapes:
  // it is the one that is re-editable, and it only exists because a render
  // finished. Ranking it top keeps the repair off it.
  const rank = (s: SlideSnapshot): number => {
    const c = contentShapes(s);
    return c === UNKNOWN_CONTENT ? Number.MAX_SAFE_INTEGER : c;
  };
  return group.reduce((best, s) => {
    const bc = rank(best);
    const sc = rank(s);
    if (sc !== bc) return sc > bc ? s : best;
    if (s.stamped !== best.stamped) return best.stamped ? s : best;
    return s.index < best.index ? s : best;
  });
}

/**
 * Decide what to repair. Pure: same inputs, same plan, no host involved.
 *
 * The rules, in the order they are applied per item:
 *
 * 1. No slide claims the item → `lost`. Nothing to repair; the run has to say
 *    so honestly rather than inventing a status from a stale readback.
 * 2. Several slides claim it → keep the fullest (see `pickKeeper`) and delete
 *    the others, but ONLY when the keeper is complete AND the twin holds no
 *    more than the keeper does. A twin with content the keeper lacks is
 *    reported, never deleted: two half-charts are a puzzle for a human, and
 *    the wrong delete is unrecoverable.
 * 3. The keeper wears a banner but holds everything it should → the banner is
 *    a lie the timeout told; remove it. (An item that was deliberately skipped
 *    keeps its banner: there, the banner is the truth.)
 * 4. The keeper is a chart with at least `PARTIAL_CONTENT_THRESHOLD` of its
 *    shapes but no config tag → group + tag it, which is what makes it
 *    re-editable. This is the repair for "4 charts landed ungrouped".
 * 5. Slides in range that claim no item are reported as orphans, and deleted
 *    only if they are empty and `dropOrphanBlanks` says so.
 */
export function planReconcile(
  snapshots: SlideSnapshot[],
  expected: ExpectedItem[],
  opts: ReconcileOptions = {},
): ReconcilePlan {
  const actions: ReconcileAction[] = [];
  const verdicts: SlotVerdict[] = [];
  const orphans: OrphanSlide[] = [];

  const bySlot = new Map<number, SlideSnapshot[]>();
  const unclaimed: SlideSnapshot[] = [];
  const known = new Map(expected.map((e) => [e.slot, e]));
  for (const s of snapshots) {
    // A slot tag naming an item the run never had is not identification — it
    // is a stale slide from an EARLIER run sitting in the same range. Treat it
    // as unclaimed so it can never be deleted as somebody's twin.
    const claim = s.slot === null ? undefined : known.get(s.slot);
    // The slot INDEX is not unique across runs: each results page is inserted
    // by its own `insertDemoDeck` call and therefore lands with slot 0 — the
    // same slot the title slide carries. Matching on the index alone would
    // pair a Results slide with the Title item and, since Results holds more
    // shapes, delete the title slide as its "duplicate". The title on the tag
    // is what disambiguates them.
    // And the run token, when the caller supplied one: slot+title names an
    // ITEM, not an occurrence of it, so across two runs of the same deck every
    // slide has a perfect twin. Without this check the pass deletes one of
    // them. A slide whose token we cannot match is not ours to touch.
    const sameRun = opts.run === undefined || (s.run ?? null) === opts.run;
    if (!claim || !sameRun || (s.title !== null && s.title !== claim.title)) {
      unclaimed.push(s);
      continue;
    }
    const group = bySlot.get(claim.slot);
    if (group) group.push(s);
    else bySlot.set(claim.slot, [s]);
  }

  for (const item of expected) {
    const group = bySlot.get(item.slot) ?? [];
    if (group.length === 0) {
      verdicts.push({
        slot: item.slot,
        title: item.title,
        status: "lost",
        shapes: 0,
        expected: item.shapes,
        index: null,
        tagged: false,
        duplicates: 0,
      });
      continue;
    }

    const keeper = pickKeeper(group);
    const kept = contentShapes(keeper);
    // A grouped slide whose children we could not count. It needs no repair
    // (it is already grouped, which is the repair) and authorises none: with
    // no number to compare against, no twin can be PROVEN redundant.
    const unmeasured = kept === UNKNOWN_CONTENT;
    const complete = unmeasured || kept >= item.shapes;
    let duplicates = 0;
    for (const other of group) {
      if (other === keeper) continue;
      duplicates++;
      if (!unmeasured && complete && contentShapes(other) <= kept && contentShapes(other) !== UNKNOWN_CONTENT) {
        actions.push({
          kind: "delete",
          index: other.index,
          slot: item.slot,
          reason: `duplicate of slide ${keeper.index + 1} (${item.title})`,
        });
      }
      // else: kept as evidence — reported through `duplicates`, never deleted.
    }

    let tagged = keeper.tagged;
    // What a report should print for an unmeasurable slide: the intent, since
    // the group proves the render reached the end.
    const shown = unmeasured ? item.shapes : kept;
    if (item.skipped) {
      verdicts.push({
        slot: item.slot,
        title: item.title,
        status: "skipped",
        shapes: shown,
        expected: item.shapes,
        index: keeper.index,
        tagged,
        duplicates,
      });
      continue;
    }

    if (keeper.stamped && complete) {
      actions.push({
        kind: "unstamp",
        index: keeper.index,
        slot: item.slot,
        // `shown`, not `kept`: an unmeasured slide's `kept` is the -1 "cannot
        // tell" sentinel, and this reason is read by a human.
        reason: `${shown} of ${item.shapes} shapes present — the banner is stale`,
      });
    }

    const worthGrouping = unmeasured || (item.shapes > 0 && kept >= Math.ceil(item.shapes * PARTIAL_CONTENT_THRESHOLD));
    // Grouped and tagged are two different facts, and conflating them left a
    // whole class of chart permanently un-re-editable. A degraded picture is
    // ONE shape named `PowerChart`, so `grouped` is true and this rule used to
    // skip it — even though the same pass had just recorded `tagged: false`.
    // A real 38-item run ended with 19 charts in exactly that state: the run
    // reported them, and the repair could not touch a single one.
    //
    // So split the repair. Loose shapes need grouping AND tagging; something
    // already whole needs only the tag.
    if (item.chart && worthGrouping && !keeper.tagged) {
      actions.push(
        keeper.grouped
          ? {
              kind: "retag",
              index: keeper.index,
              slot: item.slot,
              reason: "chart is one object but carries no config tag",
            }
          : {
              kind: "regroup",
              index: keeper.index,
              slot: item.slot,
              reason: `chart is loose and untagged (${kept} shapes)`,
            },
      );
      tagged = true;
    }

    const status: SlotVerdict["status"] =
      kept === 0 ? "empty" : complete ? "rendered" : worthGrouping ? "partial" : "wreckage";
    verdicts.push({
      slot: item.slot,
      title: item.title,
      status,
      shapes: shown,
      expected: item.shapes,
      index: keeper.index,
      tagged,
      duplicates,
    });
  }

  for (const s of unclaimed) {
    const content = contentShapes(s);
    const blank = content === 0;
    // Never report the "could not measure" sentinel as a shape count.
    orphans.push({ index: s.index, shapes: Math.max(0, content), stamped: s.stamped, blank });
    // A blank slide that carries somebody else's slot tag is an earlier run's,
    // or a copy the user made. It is identified, and identified as NOT ours —
    // the one thing that must never be swept as our own litter. Only a slide
    // with no identity at all can be, and only inside our own span.
    const somebodyElses = s.slot !== null;
    if (blank && !somebodyElses && opts.dropOrphanBlanks) {
      actions.push({
        kind: "delete",
        index: s.index,
        slot: null,
        reason: "empty slide left behind by a lost add",
      });
    }
  }

  // Descending, so each delete's index is still valid when it runs — the
  // applier relies on this ordering and a test pins it.
  actions.sort((a, b) => {
    if (a.kind === "delete" && b.kind === "delete") return b.index - a.index;
    if (a.kind === "delete") return 1;
    if (b.kind === "delete") return -1;
    return a.index - b.index;
  });

  const count = (s: SlotVerdict["status"]): number => verdicts.filter((v) => v.status === s).length;
  return {
    actions,
    verdicts,
    orphans,
    summary: {
      items: expected.length,
      rendered: count("rendered"),
      partial: count("partial"),
      lost: count("lost"),
      skipped: count("skipped"),
      wreckage: count("wreckage"),
      empty: count("empty"),
      duplicates: actions.filter((a) => a.kind === "delete" && a.slot !== null).length,
      falseStamps: actions.filter((a) => a.kind === "unstamp").length,
      untagged: actions.filter((a) => a.kind === "regroup" || a.kind === "retag").length,
      orphans: orphans.length,
    },
  };
}

/**
 * One line per item, for the pane message and the run log — the honest version
 * of "Inserted 6 of 12", written from what the deck holds rather than from
 * what the run believed at the moment it gave up waiting.
 */
export function describeReconcile(plan: ReconcilePlan): string {
  const s = plan.summary;
  const bits = [`${s.rendered} of ${s.items} complete`];
  if (s.partial) bits.push(`${s.partial} partial`);
  if (s.wreckage) bits.push(`${s.wreckage} wreckage`);
  if (s.empty) bits.push(`${s.empty} empty`);
  if (s.lost) bits.push(`${s.lost} lost`);
  if (s.skipped) bits.push(`${s.skipped} skipped`);
  if (s.duplicates) bits.push(`${s.duplicates} duplicate slide${s.duplicates === 1 ? "" : "s"} removed`);
  if (s.falseStamps) bits.push(`${s.falseStamps} false NOT COMPLETE banner${s.falseStamps === 1 ? "" : "s"} cleared`);
  if (s.untagged) bits.push(`${s.untagged} chart${s.untagged === 1 ? "" : "s"} made re-editable again`);
  if (s.orphans) bits.push(`${s.orphans} orphan slide${s.orphans === 1 ? "" : "s"}`);
  return bits.join(" · ");
}
