import { describe, it, expect } from "vitest";
import { planReconcile, type SlideSnapshot, type ExpectedItem } from "../src/core/reconcile";

/**
 * A deterministic fuzz over the repair planner.
 *
 * This is the one module in the project that decides what to DELETE from
 * someone's presentation, and its rules have grown one real-run lesson at a
 * time. The invariants below are the ones that cost a user their work if
 * broken, checked over four thousand generated decks — including combinations
 * no hand-written case would think to build, which is how the run-token hole
 * below was found.
 *
 * Seeded, so a breach is reproducible from its seed alone.
 */
const rng = (seed: number) => () => (seed = (seed * 1103515245 + 12345) & 0x7fffffff) / 0x7fffffff;

function randomCase(r: () => number) {
  const nItems = 1 + Math.floor(r() * 5);
  const run = r() < 0.8 ? "run-a" : null;
  const expected: ExpectedItem[] = Array.from({ length: nItems }, (_, i) => ({
    slot: i,
    title: `Item ${i}`,
    shapes: Math.floor(r() * 40),
    chart: r() < 0.7,
    ...(r() < 0.5 ? { wroteTag: r() < 0.5 } : {}),
  }));
  const nSlides = Math.floor(r() * 9);
  const snapshots: SlideSnapshot[] = Array.from({ length: nSlides }, (_, i) => ({
    index: i,
    slot: r() < 0.75 ? Math.floor(r() * (nItems + 1)) : null,
    title: r() < 0.8 ? `Item ${Math.floor(r() * nItems)}` : null,
    run: r() < 0.85 ? run : "run-other",
    shapes: Math.floor(r() * 40),
    ...(r() < 0.5 ? { groupChildren: Math.floor(r() * 40) } : {}),
    ...(r() < 0.5 ? { grouped: r() < 0.5 } : {}),
    stamped: r() < 0.3,
    tagged: r() < 0.5,
    ...(r() < 0.5 ? { tagRead: r() < 0.5 } : {}),
  }));
  return { expected, snapshots, run };
}

describe("the repair planner, fuzzed", () => {
  it("never proposes a destructive plan", () => {
    const breaches: string[] = [];
    for (let seed = 1; seed <= 4000; seed++) {
      const r = rng(seed);
      const { expected, snapshots, run } = randomCase(r);
      const opts = { dropOrphanBlanks: r() < 0.5, ...(r() < 0.8 && run ? { run } : {}) };
      let plan;
      try {
        plan = planReconcile(snapshots, expected, opts);
      } catch (e) {
        breaches.push(`seed ${seed}: THREW ${(e as Error).message}`);
        continue;
      }
      const deletes = plan.actions.filter((a) => a.kind === "delete").map((a) => a.index);

      // 1. Every action names a slide that exists.
      for (const a of plan.actions) {
        if (!snapshots.some((s) => s.index === a.index)) breaches.push(`seed ${seed}: acts on absent slide ${a.index}`);
      }
      // 2. Never delete a slide belonging to a different run.
      if (opts.run !== undefined) {
        for (const i of deletes) {
          const s = snapshots.find((x) => x.index === i)!;
          if (s.run != null && s.run !== opts.run) breaches.push(`seed ${seed}: deletes foreign run's slide ${i}`);
        }
      }
      // 3. Never delete every slide claiming a slot — that is the item, gone.
      const bySlot = new Map<number, SlideSnapshot[]>();
      for (const s of snapshots)
        if (s.slot !== null) (bySlot.get(s.slot) ?? bySlot.set(s.slot, []).get(s.slot)!).push(s);
      for (const [slot, group] of bySlot) {
        const mine = group.filter((s) => opts.run === undefined || s.run == null || s.run === opts.run);
        if (!mine.length) continue;
        const killed = mine.filter((s) => deletes.includes(s.index));
        const hadContent = mine.some((s) => s.shapes > 0);
        if (killed.length === mine.length && hadContent)
          breaches.push(`seed ${seed}: deletes every copy of slot ${slot} (${mine.length}), all content lost`);
      }
      // 4. A delete is never proposed twice for one slide.
      if (new Set(deletes).size !== deletes.length) breaches.push(`seed ${seed}: duplicate delete`);
      // 5. Summary counts are non-negative and bounded by the item count.
      const sm = plan.summary;
      for (const [k, v] of Object.entries(sm)) {
        if (typeof v === "number" && (!Number.isFinite(v) || v < 0)) breaches.push(`seed ${seed}: summary.${k} = ${v}`);
      }
      if (sm.items !== expected.length) breaches.push(`seed ${seed}: summary.items ${sm.items} != ${expected.length}`);
      // 6. One verdict per item, no more.
      if (plan.verdicts.length !== expected.length)
        breaches.push(`seed ${seed}: ${plan.verdicts.length} verdicts for ${expected.length} items`);
    }
    // Sliced so a systemic break prints a readable list rather than 4000 lines.
    expect(breaches.slice(0, 12)).toEqual([]);
  });

  it("leaves a blank slide alone when only its run token says whose it is", () => {
    // The hole the fuzz found. A slide's identity is its slot tag OR its run
    // token; the two travel together in a well-formed tag, so they only part
    // company when one did not survive — a half-written tag, or a slot that is
    // not a number. That is exactly when the surviving half is the only thing
    // saying whose slide this is, and reading the slot alone threw it away and
    // deleted an earlier run's slide as our own litter.
    const snapshots: SlideSnapshot[] = [
      { index: 0, slot: 0, title: "Item 0", run: "run-a", shapes: 12, stamped: false, tagged: true },
      { index: 1, slot: null, title: null, run: "run-earlier", shapes: 0, stamped: false, tagged: false },
    ];
    const expected: ExpectedItem[] = [{ slot: 0, title: "Item 0", shapes: 12, chart: true }];
    const plan = planReconcile(snapshots, expected, { dropOrphanBlanks: true, run: "run-a" });
    expect(plan.actions.filter((a) => a.kind === "delete")).toEqual([]);
    // Still reported, so nothing goes quiet about it.
    expect(plan.orphans.map((o) => o.index)).toContain(1);
  });

  it("still sweeps a blank slide that nothing claims at all", () => {
    // The negative control: an unidentified blank inside our own span is
    // exactly what `dropOrphanBlanks` exists to remove, and the fix must not
    // have turned the sweep off altogether.
    const snapshots: SlideSnapshot[] = [
      { index: 0, slot: 0, title: "Item 0", run: "run-a", shapes: 12, stamped: false, tagged: true },
      { index: 1, slot: null, title: null, run: null, shapes: 0, stamped: false, tagged: false },
    ];
    const expected: ExpectedItem[] = [{ slot: 0, title: "Item 0", shapes: 12, chart: true }];
    const plan = planReconcile(snapshots, expected, { dropOrphanBlanks: true, run: "run-a" });
    expect(plan.actions.filter((a) => a.kind === "delete").map((a) => a.index)).toEqual([1]);
  });
});
