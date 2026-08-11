import { describe, expect, it } from "vitest";
// @ts-expect-error — a plain .mjs tool with no types.
import { history, byRoundTime } from "../scripts/host-history.mjs";

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

/**
 * The order rounds are read in, which the tool's whole output depends on.
 *
 * "Settled" counts from the END, `latest` is the last column, and the fixture
 * table's bottom line is read as "the newest sheet" — so every verdict here is
 * directional. The docstring tells you to pass a GLOB, and real round files are
 * named by content hash, so the shell's order is effectively random.
 *
 * The six rounds of 2026-08-11 in glob order made `shapes-items-count-honest`
 * read `steady lately — "unreadable" x 5` when the NEWEST round had said
 * `short-0`; in true order the same six say `UNSTABLE`. Opposite verdicts, same
 * data — which is exactly the staleness this tool exists to stop
 * `UNSTABLE_ANSWERS` suffering by hand.
 */
describe("rounds are ordered by their own timestamp, not by argv", () => {
  const r = (build: string) => ({ build, path: build, answers: {} });

  it("sorts oldest first however the shell listed them", () => {
    const shuffled = [
      r("c792072 · 2026-08-11 13:10Z"),
      r("3223293 · 2026-08-11 07:28Z"),
      r("756682e · 2026-08-11 12:10Z"),
      r("7027f96 · 2026-08-11 09:43Z"),
    ];
    expect([...shuffled].sort(byRoundTime).map((x) => x.build.slice(0, 7))).toEqual([
      "3223293",
      "7027f96",
      "756682e",
      "c792072",
    ]);
  });

  it("is stable and identical whatever order it is handed", () => {
    const a = [r("b · 2026-08-11 09:00Z"), r("a · 2026-08-11 08:00Z"), r("c · 2026-08-11 10:00Z")];
    const forward = [...a].sort(byRoundTime).map((x) => x.build[0]);
    const backward = [...a]
      .reverse()
      .sort(byRoundTime)
      .map((x) => x.build[0]);
    expect(forward).toEqual(["a", "b", "c"]);
    expect(backward).toEqual(forward);
  });

  it("puts a round whose stamp cannot be read LAST rather than reordering real ones around it", () => {
    // The safe direction: an unparseable stamp is most likely hand-made, and
    // putting it at the end makes it visible in `latest` instead of silently
    // displacing a real round.
    const a = [r("real · 2026-08-11 09:00Z"), r("?"), r("older · 2026-08-11 08:00Z")];
    expect([...a].sort(byRoundTime).map((x) => x.build.split(" ")[0])).toEqual(["older", "real", "?"]);
  });
});
