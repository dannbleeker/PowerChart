// @vitest-environment jsdom
import { describe, expect, it, vi } from "vitest";

/**
 * The probe does not borrow the repair page's ninety-second budget for one
 * slide.
 *
 * `PROBE_BUDGET_MS` is eight seconds and bounds the QUESTION. The slide a
 * question needs is acquired outside it, and `addScratchSlide` defaults to
 * `READBACK_TIMEOUT_MS` — ninety seconds, sized for a twenty-slide repair page.
 * So on 2026-08-10 `shape-resolve-held-slide-proxy` took **95.6 seconds**
 * against an eight-second budget, and the run burned ninety of them waiting for
 * a slide that never came. A bound something else can walk around is not a
 * bound, and the sheet's whole promise is that a wedged host costs one question
 * rather than the run.
 *
 * Asserted on the ARGUMENT rather than the clock, and the honest reason is that
 * the clock version does not work yet: `addScratchSlide` makes three round
 * trips and only the middle one takes this budget, so a fake told to wedge
 * every sync hangs in the id read before it and the test times out instead of
 * failing. That is a real latent hang — on the real host those reads answered
 * (`afterAnswering=listing the deck's slides`), which is why it has never bitten
 * — and it is not this change's to fix. What regressed here was that there was
 * no bound at all, and that is exactly what this pins.
 *
 * Its own file because the module mock is hoisted over the whole module graph,
 * and `host-probe.test.ts` drives the real thing forty-odd times.
 */
const budgets: (number | undefined)[] = [];

vi.mock("../src/render/powerpoint", async (importActual) => {
  const actual = await importActual<typeof import("../src/render/powerpoint")>();
  return {
    ...actual,
    addScratchSlide: (budgetMs?: number) => {
      budgets.push(budgetMs);
      return actual.addScratchSlide(budgetMs);
    },
  };
});

const { installHost, makeSlide } = await import("./helpers/office-host");
const { runHostProbes } = await import("../src/render/host-probe");
const { readbackTimeoutMs } = await import("../src/render/powerpoint");

describe("what the probe is willing to wait for a scratch slide", () => {
  it("asks for a bound, and one far shorter than the repair page's", async () => {
    installHost([makeSlide("s1")]);
    await runHostProbes("fake", "test");

    expect(budgets.length, "the probe never took a scratch slide at all").toBeGreaterThan(0);
    for (const ms of budgets) {
      expect(ms, "a scratch add went out with no bound — it inherits ninety seconds").toBeTypeOf("number");
      expect(ms!, "a scratch add is bounded, but by the repair page's budget").toBeLessThan(readbackTimeoutMs());
    }
    // Measured, not chosen: every scratch add that WORKED in the 2026-08-10 run
    // took 0.21s to 4.0s, and every one that failed took the full ninety. A
    // bound anywhere in that gap is right; one below the slowest success would
    // turn a slow host into lost questions.
    for (const ms of budgets) expect(ms!, "tighter than the slowest add ever observed").toBeGreaterThan(5_000);
  });
});
