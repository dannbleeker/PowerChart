import { describe, expect, it } from "vitest";
import {
  describeReconcile,
  planReconcile,
  type ExpectedItem,
  type ReconcileAction,
  type SlideSnapshot,
} from "../src/core/reconcile";

/** A snapshot with the boring fields filled in. */
function slide(index: number, over: Partial<SlideSnapshot> = {}): SlideSnapshot {
  return { index, slot: null, title: null, shapes: 0, stamped: false, tagged: false, ...over };
}

/** An expected chart item — `chart: true` unless a test says otherwise. */
function item(slot: number, title: string, shapes: number, over: Partial<ExpectedItem> = {}): ExpectedItem {
  return { slot, title, shapes, chart: true, ...over };
}

const kinds = (actions: ReconcileAction[], kind: ReconcileAction["kind"]): number[] =>
  actions.filter((a) => a.kind === kind).map((a) => a.index);

describe("reconcile: pairing slides back to intent", () => {
  it("reports an item no slide claims as lost, and repairs nothing", () => {
    const plan = planReconcile([], [item(0, "Stacked", 40)]);
    expect(plan.verdicts[0].status).toBe("lost");
    expect(plan.verdicts[0].index).toBeNull();
    expect(plan.actions).toEqual([]);
  });

  it("counts a grouped chart by its children, not by its one top-level shape", () => {
    // The bug this pins: a grouped 36-shape chart is ONE shape on the slide.
    // Measured top-level, a perfect chart reads as 1 of 36 — wreckage.
    const plan = planReconcile(
      [slide(1, { slot: 0, title: "Line", shapes: 1, groupChildren: 36, tagged: true })],
      [item(0, "Line", 36)],
    );
    expect(plan.verdicts[0].status).toBe("rendered");
    expect(plan.verdicts[0].shapes).toBe(36);
    expect(plan.actions).toEqual([]);
  });

  it("keeps the fullest slide and deletes a provably redundant twin", () => {
    const plan = planReconcile(
      [
        slide(1, { slot: 0, title: "Line", shapes: 36 }),
        slide(2, { slot: 0, title: "Line", shapes: 37, stamped: true }),
      ],
      [item(0, "Line", 36)],
    );
    expect(kinds(plan.actions, "delete")).toEqual([2]);
    expect(plan.verdicts[0].index).toBe(1);
    expect(plan.verdicts[0].duplicates).toBe(1);
  });

  it("prefers the unstamped slide when two twins hold the same content", () => {
    const plan = planReconcile(
      [
        slide(1, { slot: 0, title: "Line", shapes: 37, stamped: true }),
        slide(2, { slot: 0, title: "Line", shapes: 36 }),
      ],
      [item(0, "Line", 36)],
    );
    expect(plan.verdicts[0].index).toBe(2);
    expect(kinds(plan.actions, "delete")).toEqual([1]);
  });

  it("never deletes a twin that holds content the keeper lacks", () => {
    // Two half-charts: neither is complete, so neither is redundant. A human
    // decides. Deleting the wrong one is unrecoverable.
    const plan = planReconcile(
      [slide(1, { slot: 0, title: "Line", shapes: 20 }), slide(2, { slot: 0, title: "Line", shapes: 18 })],
      [item(0, "Line", 36)],
    );
    expect(kinds(plan.actions, "delete")).toEqual([]);
    expect(plan.verdicts[0].duplicates).toBe(1);
    expect(plan.summary.duplicates).toBe(0);
  });

  it("clears a NOT COMPLETE banner that sits on a complete chart", () => {
    const plan = planReconcile(
      [slide(1, { slot: 0, title: "Agenda", shapes: 14, stamped: true })],
      [item(0, "Agenda", 13, { chart: false })],
    );
    expect(kinds(plan.actions, "unstamp")).toEqual([1]);
    expect(plan.verdicts[0].status).toBe("rendered");
  });

  it("leaves the banner alone when the item was deliberately skipped", () => {
    const plan = planReconcile(
      [slide(1, { slot: 0, title: "Heatmap", shapes: 1, stamped: true })],
      [item(0, "Heatmap", 120, { skipped: true })],
    );
    expect(plan.actions).toEqual([]);
    expect(plan.verdicts[0].status).toBe("skipped");
  });

  it("re-groups a loose chart that is complete but untagged", () => {
    const plan = planReconcile([slide(1, { slot: 0, title: "Gantt", shapes: 31 })], [item(0, "Gantt", 31)]);
    expect(kinds(plan.actions, "regroup")).toEqual([1]);
    expect(plan.verdicts[0].tagged).toBe(true);
  });

  it("re-groups a partial chart above the threshold but not wreckage below it", () => {
    const good = planReconcile([slide(1, { slot: 0, title: "Line", shapes: 31 })], [item(0, "Line", 36)]);
    expect(good.verdicts[0].status).toBe("partial"); // 31/36 = 86%
    expect(kinds(good.actions, "regroup")).toEqual([1]);

    const bad = planReconcile([slide(1, { slot: 0, title: "Line", shapes: 12 })], [item(0, "Line", 36)]);
    expect(bad.verdicts[0].status).toBe("wreckage");
    expect(kinds(bad.actions, "regroup")).toEqual([]);
  });

  it("does not re-tag a chart that already carries its config", () => {
    const plan = planReconcile([slide(1, { slot: 0, title: "Pie", shapes: 50, tagged: true })], [item(0, "Pie", 50)]);
    expect(plan.actions).toEqual([]);
    expect(plan.summary.untagged).toBe(0);
  });

  it("does not group a non-chart element, which has no config to carry", () => {
    const plan = planReconcile(
      [slide(1, { slot: 0, title: "Contents", shapes: 27 })],
      [item(0, "Contents", 27, { chart: false })],
    );
    expect(plan.actions).toEqual([]);
    expect(plan.verdicts[0].status).toBe("rendered");
  });
});

describe("reconcile: a group we cannot measure", () => {
  // A host below PowerPointApi 1.8 exposes no way to count a group's children.
  // The group itself is still evidence — it is only ever made after a render
  // finished — so the pass must read it as done and touch nothing.
  const grouped = (index: number, over: Partial<SlideSnapshot> = {}) =>
    slide(index, { slot: 0, title: "Line", shapes: 1, grouped: true, tagged: true, ...over });

  it("treats it as rendered and asks for no repair", () => {
    const plan = planReconcile([grouped(1)], [item(0, "Line", 36)]);
    expect(plan.verdicts[0].status).toBe("rendered");
    expect(plan.verdicts[0].shapes).toBe(36);
    expect(plan.actions).toEqual([]);
  });

  it("outranks a loose pile of shapes when choosing what to keep", () => {
    const plan = planReconcile([slide(1, { slot: 0, title: "Line", shapes: 36 }), grouped(2)], [item(0, "Line", 36)]);
    expect(plan.verdicts[0].index).toBe(2);
  });

  it("authorises no deletion, because nothing can be proven redundant", () => {
    const plan = planReconcile([grouped(1), slide(2, { slot: 0, title: "Line", shapes: 36 })], [item(0, "Line", 36)]);
    expect(plan.actions).toEqual([]);
    expect(plan.verdicts[0].duplicates).toBe(1);
  });

  it("still clears a banner sitting on top of it", () => {
    const plan = planReconcile([grouped(1, { shapes: 2, stamped: true })], [item(0, "Line", 36)]);
    expect(kinds(plan.actions, "unstamp")).toEqual([1]);
  });

  it("tags — rather than re-groups — a chart that is whole but lost its tag", () => {
    // This used to plan NOTHING, because the rule asked `!keeper.grouped` and
    // a whole chart is grouped. Grouped and tagged are two different facts,
    // and conflating them made a chart that is visibly a chart, and provably
    // not re-editable, unreachable by every repair.
    //
    // Not hypothetical: a 38-item web run degraded to pictures and ended with
    // 19 charts in exactly this state — one shape named PowerChart, no config
    // tag. The run reported all 19 and could repair none of them.
    const plan = planReconcile([grouped(1, { tagged: false })], [item(0, "Line", 36)]);
    expect(kinds(plan.actions, "regroup")).toEqual([]); // nothing to group
    expect(kinds(plan.actions, "retag")).toEqual([1]); // just the tag
    // And the verdict says re-editable, because the repair will make it so.
    expect(plan.verdicts[0].tagged).toBe(true);
  });

  it("leaves a whole, already-tagged chart completely alone", () => {
    const plan = planReconcile([grouped(1)], [item(0, "Line", 36)]);
    expect(plan.actions).toEqual([]);
  });
});

describe("reconcile: slides nobody claims", () => {
  it("reports an orphan but leaves it alone by default", () => {
    const plan = planReconcile([slide(1, { shapes: 0 })], [item(0, "Line", 36)]);
    expect(plan.orphans).toEqual([{ index: 1, shapes: 0, stamped: false, blank: true }]);
    expect(plan.actions).toEqual([]);
  });

  it("drops an empty orphan when the caller owns the range", () => {
    const plan = planReconcile([slide(1, { shapes: 0 })], [item(0, "Line", 36)], { dropOrphanBlanks: true });
    expect(kinds(plan.actions, "delete")).toEqual([1]);
  });

  it("treats a banner-only slide as empty, since the banner is not content", () => {
    const plan = planReconcile([slide(1, { shapes: 1, stamped: true })], [item(0, "Line", 36)], {
      dropOrphanBlanks: true,
    });
    expect(kinds(plan.actions, "delete")).toEqual([1]);
  });

  it("never drops an orphan that holds real shapes", () => {
    const plan = planReconcile([slide(1, { shapes: 9 })], [item(0, "Line", 36)], { dropOrphanBlanks: true });
    expect(plan.actions).toEqual([]);
    expect(plan.orphans[0].blank).toBe(false);
  });

  it("never pairs a Results slide with the Title item they share a slot with", () => {
    // Each results page is inserted by its own insertDemoDeck call, so it
    // lands carrying slot 0 — the title slide's slot. Paired by index alone,
    // Results (more shapes) becomes the keeper and the real title slide is
    // deleted as its duplicate. The tag's title is what keeps them apart.
    const plan = planReconcile(
      [slide(1, { slot: 0, title: "Title", shapes: 5 }), slide(2, { slot: 0, title: "Results", shapes: 24 })],
      [item(0, "Title", 5, { chart: false })],
      { dropOrphanBlanks: true },
    );
    expect(plan.actions).toEqual([]);
    expect(plan.verdicts[0]).toMatchObject({ title: "Title", index: 1, status: "rendered", duplicates: 0 });
    expect(plan.orphans.map((o) => o.index)).toEqual([2]);
  });

  it("treats a slot tag from an older run as unclaimed, never as a twin", () => {
    // Slot 7 is not in this run's item list. Deleting it as "a duplicate of
    // slot 7" would destroy a slide from a previous insert.
    const plan = planReconcile(
      [slide(1, { slot: 0, title: "Line", shapes: 36 }), slide(2, { slot: 7, title: "Combo", shapes: 40 })],
      [item(0, "Line", 36)],
      { dropOrphanBlanks: true },
    );
    expect(kinds(plan.actions, "delete")).toEqual([]);
    expect(plan.orphans.map((o) => o.index)).toEqual([2]);
  });
});

describe("reconcile: plan ordering", () => {
  it("orders deletes last and descending, so each index is still valid", () => {
    const plan = planReconcile(
      [
        slide(1, { slot: 0, title: "A", shapes: 10 }),
        slide(2, { slot: 0, title: "A", shapes: 10, stamped: true }),
        slide(3, { slot: 1, title: "B", shapes: 10, stamped: true }),
        slide(4, { slot: 1, title: "B", shapes: 11, stamped: true }),
        slide(5, { shapes: 0 }),
      ],
      [item(0, "A", 10), item(1, "B", 10)],
      { dropOrphanBlanks: true },
    );
    const deletes = plan.actions.filter((a) => a.kind === "delete").map((a) => a.index);
    expect(deletes).toEqual([...deletes].sort((a, b) => b - a));
    const firstDelete = plan.actions.findIndex((a) => a.kind === "delete");
    expect(plan.actions.slice(firstDelete).every((a) => a.kind === "delete")).toBe(true);
  });
});

/**
 * The real thing: `Presentation_4.pptx`, the deck a PowerPoint-web run produced
 * on 2026-07-30, read straight out of its slide XML. The run reported
 * "Inserted 6 of 12 … Host failed on: Stacked, Scatter, Bubble, Gantt, Combo,
 * Heatmap … 3 slides came back BLANK: 9, 10, Agenda". The deck disagrees on
 * three counts, and this fixture is the disagreement:
 *
 *   - Gantt did not fail. It landed TWICE (slides 6 and 7).
 *   - The Agenda slide holds all 13 of its shapes; it was never blank.
 *   - Contents, Line, Gantt and Agenda all wear a NOT COMPLETE banner over
 *     content that is, in fact, complete.
 */
describe("reconcile: Presentation_4.pptx, the 2026-07-30 web run", () => {
  const expected: ExpectedItem[] = [
    item(0, "Title", 5, { chart: false }),
    item(1, "Contents", 27, { chart: false }),
    item(2, "Stacked", 40),
    item(3, "Line", 36),
    item(4, "Scatter", 42),
    item(5, "Bubble", 46),
    item(6, "Gantt", 31),
    item(7, "Combo", 40),
    item(8, "Pie", 50),
    item(9, "Heatmap", 67),
    item(10, "Agenda", 13, { chart: false }),
    item(11, "KPI tile", 5, { chart: false }),
  ];

  // Deck index 0 is the host's own title slide; the run added indices 1-10.
  const deck: SlideSnapshot[] = [
    slide(1, { slot: 0, title: "Title", shapes: 1, groupChildren: 5 }),
    slide(2, { slot: 1, title: "Contents", shapes: 28, stamped: true }),
    slide(3, { slot: 3, title: "Line", shapes: 36 }),
    slide(4, { slot: 3, title: "Line", shapes: 37, stamped: true }),
    slide(5, { slot: 6, title: "Gantt", shapes: 31 }),
    slide(6, { slot: 6, title: "Gantt", shapes: 32, stamped: true }),
    slide(7, { slot: 8, title: "Pie", shapes: 50, tagged: true }),
    slide(8, { shapes: 0 }),
    slide(9, { shapes: 1, stamped: true }),
    slide(10, { slot: 10, title: "Agenda", shapes: 14, stamped: true }),
  ];

  const plan = planReconcile(deck, expected, { dropOrphanBlanks: true });

  it("names the five items that were really lost — not the six the run blamed", () => {
    const lost = plan.verdicts.filter((v) => v.status === "lost").map((v) => v.title);
    expect(lost).toEqual(["Stacked", "Scatter", "Bubble", "Combo", "Heatmap", "KPI tile"]);
    expect(lost).not.toContain("Gantt");
  });

  it("finds the Agenda intact, against a run that called it blank", () => {
    const agenda = plan.verdicts.find((v) => v.title === "Agenda")!;
    expect(agenda.status).toBe("rendered");
    expect(agenda.shapes).toBe(13);
  });

  it("deletes exactly the two duplicate charts and the two empty strays", () => {
    expect(kinds(plan.actions, "delete")).toEqual([9, 8, 6, 4]);
  });

  it("clears every banner that contradicts its own slide", () => {
    // Only the two kept slides need unstamping — the other two banners are on
    // the duplicate slides, which the plan deletes outright.
    expect(kinds(plan.actions, "unstamp").sort((a, b) => a - b)).toEqual([2, 10]);
  });

  it("re-groups the two charts left loose, and only those", () => {
    expect(kinds(plan.actions, "regroup").sort((a, b) => a - b)).toEqual([3, 5]);
  });

  it("summarises the run the way the deck actually looks", () => {
    expect(plan.summary).toMatchObject({
      items: 12,
      rendered: 6,
      partial: 0,
      lost: 6,
      duplicates: 2,
      falseStamps: 2,
      untagged: 2,
      orphans: 2,
    });
    expect(describeReconcile(plan)).toBe(
      "6 of 12 complete · 6 lost · 2 duplicate slides removed · 2 false NOT COMPLETE banners cleared · 2 charts made re-editable again · 2 orphan slides",
    );
  });

  it("is idempotent: repairing the repaired deck asks for nothing more", () => {
    // Apply the plan to the fixture the way the renderer would, then re-plan.
    const deleted = new Set(kinds(plan.actions, "delete"));
    const unstamped = new Set(kinds(plan.actions, "unstamp"));
    const regrouped = new Set(kinds(plan.actions, "regroup"));
    const repaired = deck
      .filter((s) => !deleted.has(s.index))
      .map((s) => ({
        ...s,
        shapes: unstamped.has(s.index) ? s.shapes - 1 : s.shapes,
        stamped: unstamped.has(s.index) ? false : s.stamped,
        tagged: regrouped.has(s.index) ? true : s.tagged,
        groupChildren: regrouped.has(s.index) ? s.shapes : s.groupChildren,
      }));
    const second = planReconcile(repaired, expected, { dropOrphanBlanks: true });
    expect(second.actions).toEqual([]);
    expect(second.summary.rendered).toBe(6);
  });
});

describe("reconcile: a chart that is whole but carries no tag", () => {
  // The state a degraded run leaves behind: ONE shape named PowerChart, no
  // POWERCHART_CONFIG. Visibly a chart, provably not re-editable.
  const picture = (index: number, slot: number, title: string) =>
    slide(index, { slot, title, shapes: 1, grouped: true, tagged: false });

  it("plans a retag for every one of them, and no regroup", () => {
    // Reproduces the shape of a real 38-item web run: the host degraded at
    // item 2, every later chart went on as a picture, and 19 of them lost the
    // tag write. The pass reported all 19 and repaired none.
    const snaps = [picture(1, 0, "Line"), picture(2, 1, "Scatter"), picture(3, 2, "Bubble")];
    const items = [item(0, "Line", 36), item(1, "Scatter", 42), item(2, "Bubble", 46)];
    const plan = planReconcile(snaps, items);
    expect(kinds(plan.actions, "retag")).toEqual([1, 2, 3]);
    expect(kinds(plan.actions, "regroup")).toEqual([]);
    expect(plan.summary.untagged).toBe(3);
    expect(plan.verdicts.every((v) => v.tagged)).toBe(true);
  });

  it("repairs nothing when the readback could not see the slide at all", () => {
    // `tagRead: false` means the tag pass saw fewer shapes than the previous
    // pass had just counted, so it learned nothing either way. A real run read
    // 3 shapes on a page where 19 had been counted moments before, called the
    // missing ones untagged, and rewrote fourteen charts whose config was
    // already correct — one of them being wrong is a repair; fourteen of them
    // being right is damage.
    const plan = planReconcile(
      [slide(1, { slot: 0, title: "Line", shapes: 1, grouped: true, tagged: false, tagRead: false })],
      [item(0, "Line", 36)],
    );
    expect(plan.actions).toEqual([]);
    // …and says so, rather than reporting the chart as broken.
    expect(plan.summary.undetermined).toBe(1);
    expect(describeReconcile(plan)).toMatch(/would not let us check/);
  });

  it("repairs an unreadable slide anyway when the run KNOWS its tag write failed", () => {
    // Two real runs gave the readback identical evidence — "could not see this
    // slide" — with opposite right answers. The file path had written 31 tags
    // and 6 were unreadable; those were fine. The shape path had written 3 of
    // 38 and 24 were unreadable; those were ALL genuinely untagged, and
    // suppressing left 20 charts un-editable that a retag would have fixed.
    //
    // What separates them is not the readback. It is whether the run watched
    // its own write fail.
    const plan = planReconcile(
      [slide(1, { slot: 0, title: "Line", shapes: 1, grouped: true, tagged: false, tagRead: false })],
      [item(0, "Line", 36, { wroteTag: false })],
    );
    expect(kinds(plan.actions, "retag")).toEqual([1]);
    expect(plan.summary.undetermined).toBe(0);
  });

  it("leaves an unreadable slide alone when the run believes the tag landed", () => {
    const plan = planReconcile(
      [slide(1, { slot: 0, title: "Line", shapes: 1, grouped: true, tagged: false, tagRead: false })],
      [item(0, "Line", 36, { wroteTag: true })],
    );
    expect(plan.actions).toEqual([]);
    expect(plan.summary.undetermined).toBe(1);
  });

  it("still repairs when the readback DID see the slide and found no tag", () => {
    // The other half: suppressing on uncertainty must not suppress on evidence.
    const plan = planReconcile(
      [slide(1, { slot: 0, title: "Line", shapes: 1, grouped: true, tagged: false, tagRead: true })],
      [item(0, "Line", 36)],
    );
    expect(kinds(plan.actions, "retag")).toEqual([1]);
    expect(plan.summary.undetermined).toBe(0);
  });

  it("does not retag something that was never a chart", () => {
    // Contents, Agenda, KPI tile: no config exists for them, so an untagged
    // one is correct, not broken.
    const plan = planReconcile([picture(1, 0, "Contents")], [item(0, "Contents", 27, { chart: false })]);
    expect(plan.actions).toEqual([]);
  });

  it("does not retag a chart the run deliberately skipped", () => {
    // A skipped item has no chart on its slide at all — only the banner. A tag
    // there would claim an empty slide is an editable chart.
    const plan = planReconcile(
      [slide(1, { slot: 0, title: "Violin", shapes: 1, stamped: true })],
      [item(0, "Violin", 253, { skipped: true })],
    );
    expect(plan.actions).toEqual([]);
  });
});

describe("reconcile: telling one run's slides from another's", () => {
  // Slot indices restart at 0 every run and the demo titles are fixed, so
  // slot+title names an ITEM, not an occurrence of one. Insert the demo deck
  // twice into one presentation and every slide of run 2 has a perfect twin in
  // run 1 — which the pass read as "this item landed twice" and repaired by
  // deleting one of them. A whole healthy run, deleted for duplicating another
  // healthy run.
  const twoRuns = (): SlideSnapshot[] => [
    slide(0, { slot: 0, title: "Title", shapes: 4, run: "run-a", tagged: true }),
    slide(1, { slot: 1, title: "Line", shapes: 36, run: "run-a", tagged: true }),
    slide(2, { slot: 0, title: "Title", shapes: 4, run: "run-b", tagged: true }),
    slide(3, { slot: 1, title: "Line", shapes: 36, run: "run-b", tagged: true }),
  ];
  const expected = [item(0, "Title", 4, { chart: false }), item(1, "Line", 36)];

  it("deletes nothing when an earlier run's slides sit in the same span", () => {
    const plan = planReconcile(twoRuns(), expected, { run: "run-b", dropOrphanBlanks: true });
    expect(kinds(plan.actions, "delete")).toEqual([]);
    // Run B's own slides are the ones reconciled…
    expect(plan.verdicts.map((v) => v.index)).toEqual([2, 3]);
    expect(plan.verdicts.every((v) => v.status === "rendered")).toBe(true);
    // …and run A's are reported, untouched, as somebody else's.
    expect(plan.orphans.map((o) => o.index)).toEqual([0, 1]);
  });

  it("still pairs slides when no run token is supplied — an old deck stays repairable", () => {
    // The token is a new field. A slide tagged by a build that predates it
    // carries none, and a caller that has already bounded the snapshots to one
    // run's own added range does not need it.
    // 31 shapes: 30 of chart plus the banner itself.
    const plan = planReconcile(
      [slide(0, { slot: 0, title: "Line", shapes: 31, stamped: true })],
      [item(0, "Line", 30)],
    );
    expect(kinds(plan.actions, "unstamp")).toEqual([0]);
  });

  it("refuses to claim a slide whose tag carries no run token, once one is expected", () => {
    // PowerPoint's own Duplicate Slide copies the tag part. A slide the user
    // duplicated, then emptied and reused, must never be deleted as the twin
    // of the chart it was copied from.
    const plan = planReconcile(
      [
        slide(0, { slot: 0, title: "Line", shapes: 36, run: "run-b", tagged: true }),
        slide(1, { slot: 0, title: "Line", shapes: 2 }),
      ],
      [item(0, "Line", 36)],
      { run: "run-b", dropOrphanBlanks: true },
    );
    expect(kinds(plan.actions, "delete")).toEqual([]);
    expect(plan.orphans.map((o) => o.index)).toEqual([1]);
  });

  it("never sweeps a blank slide that belongs to another run", () => {
    // dropOrphanBlanks exists for OUR litter: a slide the host created and
    // never filled. A blank slide carrying somebody else's slot tag is
    // identified, and identified as not ours.
    const plan = planReconcile(
      [
        slide(0, { slot: 0, title: "Line", shapes: 36, run: "run-b", tagged: true }),
        slide(1, { slot: 3, title: "Gantt", run: "run-a" }),
        slide(2), // no tag at all — our own litter, inside our own span
      ],
      [item(0, "Line", 36)],
      { run: "run-b", dropOrphanBlanks: true },
    );
    expect(kinds(plan.actions, "delete")).toEqual([2]);
  });

  it("reports an unmeasurable orphan as 0 shapes, not as the -1 sentinel", () => {
    const plan = planReconcile([slide(9, { grouped: true })], [], { run: "run-b" });
    expect(plan.orphans[0].shapes).toBe(0);
    expect(plan.orphans[0].blank).toBe(false);
  });
});
