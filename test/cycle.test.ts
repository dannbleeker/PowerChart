import { describe, expect, it } from "vitest";
// @ts-expect-error — plain .mjs tool, no types.
import * as cycle from "../scripts/cycle.mjs";
const { cyclePlan, nextStep, readReceipt } = cycle;

/**
 * The cycle runner's whole job is knowing when to stop, so stopping is what is
 * worth testing. A runner that carries on past a regression wastes a night; one
 * that stops on an ordinary failed scenario throws away the second half of a
 * pair, which is the only thing that separates a real fault from this host's
 * 1-in-5 mood.
 */
describe("a night's cycle", () => {
  const receipt = (over = {}) => ({
    reason: "finished",
    codes: [],
    recoverable: false,
    roundFile: "082-x.json",
    ...over,
  });

  it("runs 16:9 twice and 4:3 once, in that order", () => {
    const plan = cyclePlan();
    expect(plan.map((p: { size: string }) => p.size)).toEqual(["16:9", "16:9", "4:3"]);
    // 4:3 LAST: a night that dies early should lose the validation, not the
    // measurement.
    expect(plan[plan.length - 1].size).toBe("4:3");
    // Different decks — a 4:3 leg on the 16:9 deck would file a round under a
    // profile it was not measured at.
    expect(plan[0].deck).not.toBe(plan[2].deck);
  });

  it("carries on after a round whose scenarios failed", () => {
    // A failing scenario is the measurement WORKING. Stopping here would throw
    // away the pair.
    expect(nextStep({ exitCode: 0, receipt: receipt(), gateFailed: false }).go).toBe(true);
  });

  it("stops dead on a regression, which is the one fatal check", () => {
    const step = nextStep({ exitCode: 0, receipt: receipt(), gateFailed: true });
    expect(step.go).toBe(false);
    expect(step.why).toMatch(/WAS passing/);
  });

  it("stops on a refusal recovery does not address, and says it needs a person", () => {
    const step = nextStep({
      exitCode: 1,
      receipt: receipt({ reason: "not-ready", codes: ["wrong-size"], recoverable: false, roundFile: null }),
      gateFailed: false,
    });
    expect(step.go).toBe(false);
    expect(step.why).toMatch(/needs a person/);
    expect(step.why).toMatch(/wrong-size/);
  });

  it("does not retry what the driver already retried", () => {
    // By the time a recoverable refusal reaches here the driver has exhausted
    // its own retries. Another attempt from this layer would be a second
    // implementation of shouldRetry, and there must only be one.
    const step = nextStep({
      exitCode: 1,
      receipt: receipt({ reason: "not-ready", codes: ["pane-stale"], recoverable: true, roundFile: null }),
      gateFailed: false,
    });
    expect(step.go).toBe(false);
    expect(step.why).toMatch(/retried and still could not clear/);
  });

  it("stops when the driver left no account of itself at all", () => {
    const step = nextStep({ exitCode: 1, receipt: null, gateFailed: false });
    expect(step.go).toBe(false);
    expect(step.why).toMatch(/no \.round-outcome\.json/);
  });

  it("treats an unreadable receipt as no receipt rather than crashing on it", () => {
    expect(
      readReceipt(
        "x",
        () => true,
        () => "{ not json",
      ),
    ).toBeNull();
    expect(readReceipt("x", () => false)).toBeNull();
  });
});
