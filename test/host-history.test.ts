import { describe, expect, it } from "vitest";
// @ts-expect-error — a plain .mjs tool with no types.
import { history } from "../scripts/host-history.mjs";

/**
 * The arithmetic behind `UNSTABLE_ANSWERS` and the fixture-swap decision.
 *
 * Both were done by hand from whatever rounds the author had open, and both
 * went stale silently: one entry read "ALTERNATES … it is a coin" about a
 * question that had by then answered the same way ten rounds running, and
 * another said "ASKED AND ANSWERED ONCE" against eight consistent samples.
 * Neither is a small error — the first tells the next reader they may not build
 * on an answer that has in fact never moved.
 */
const round = (answers: Record<string, string>) => ({ path: "x", build: "b", answers });

describe("host-history", () => {
  it("calls a question settled only when every real answer agrees", () => {
    const h = history([round({ q: "threw" }), round({ q: "threw" }), round({ q: "threw" })]);
    expect(h[0]).toMatchObject({ asked: 3, distinct: ["threw"], streak: 3, latest: "threw" });
  });

  it("does not let a question the run could not PUT break a streak", () => {
    // A run that never set the probe up says nothing about the host. Counting
    // `no-scratch-slide` as a different answer would report every question as
    // unstable after one bad night — and this host has bad nights.
    const h = history([
      round({ q: "threw" }),
      round({ q: "no-scratch-slide" }),
      round({ q: "threw" }),
      round({ q: "no-scratch-shape" }),
      round({ q: "threw" }),
    ]);
    expect(h[0]).toMatchObject({ asked: 3, distinct: ["threw"], streak: 3 });
  });

  it("counts the streak from the END, so a settled tail is visible under an unsettled history", () => {
    // The real shape of `shapes-items-count-honest`: two forms overall, but
    // `unreadable` on the last seven. Reporting only "2 faces" hides that.
    const seq = ["short-0", "unreadable", "unreadable", "unreadable"];
    const h = history(seq.map((a) => round({ q: a })));
    expect(h[0].distinct).toHaveLength(2);
    expect(h[0].streak, "a settled tail was invisible").toBe(3);
    expect(h[0].latest).toBe("unreadable");
  });

  it("reports a question no round could put as asked zero times, not as settled", () => {
    const h = history([round({ q: "no-scratch-slide" }), round({ q: "no-scratch-slide" })]);
    expect(h[0]).toMatchObject({ asked: 0, distinct: [], streak: 0 });
  });
});
