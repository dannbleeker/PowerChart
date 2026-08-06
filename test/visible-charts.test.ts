import { describe, expect, it } from "vitest";
// @ts-expect-error — a plain .mjs tool with no types, deliberately independent
// of src/ so it cannot inherit a bug from the renderer it judges.
import { judge } from "../scripts/visible-charts.mjs";

/**
 * The verdict half of the visual gate.
 *
 * The gate's own value is that it rasterises in a REAL browser, which nothing
 * here does. What these cases pin is the part that decides — because a
 * measurement is only as good as what is concluded from it, and a threshold
 * nobody checked is how a gate comes to pass everything.
 *
 * Each case is one of the failure modes the gate exists for, and every one of
 * them produces a stable, diffable, entirely correct-looking SVG that
 * `test/snapshots.test.ts` passes today.
 */
const size = { w: 400, h: 300 };
const drawn = (over: Partial<{ ink: number; colours: number; box: object }> = {}) => ({
  ink: 0.2,
  colours: 40,
  box: { minX: 10, minY: 10, maxX: 390, maxY: 290 },
  ...over,
});

describe("deciding whether a chart is visible", () => {
  it("passes a chart with ink, colour, and everything inside the frame", () => {
    expect(judge("ok", drawn(), size).ok).toBe(true);
  });

  it("fails a chart drawn in white on white", () => {
    // The SVG is perfect. Nothing is on the slide.
    const v = judge("white", { ink: 0, colours: 0, box: { minX: 400, minY: 300, maxX: -1, maxY: -1 } }, size);
    expect(v.ok).toBe(false);
    expect(v.problems.join("; ")).toMatch(/nothing was drawn at all/);
  });

  it("fails a chart whose shapes all collapsed to zero size", () => {
    const v = judge("collapsed", drawn({ ink: 0.0001 }), size);
    expect(v.ok).toBe(false);
    expect(v.problems.join("; ")).toMatch(/carries ink/);
  });

  it("fails a chart drawn outside the frame", () => {
    // The one a text snapshot is least able to see: the SVG is byte-identical
    // whether a translate is right or wildly wrong.
    const v = judge("off-canvas", drawn({ box: { minX: 500, minY: 500, maxX: 900, maxY: 900 } }), size);
    expect(v.ok).toBe(false);
    expect(v.problems.join("; ")).toMatch(/outside the frame/);
  });

  it("fails a solid block of one colour", () => {
    const v = judge("block", drawn({ ink: 0.99, colours: 1 }), size);
    expect(v.ok).toBe(false);
    expect(v.problems.join("; ")).toMatch(/block, not a chart/);
    expect(v.problems.join("; ")).toMatch(/distinct colour/);
  });

  it("fails a chart that lost most of its drawing but kept its labels", () => {
    // The case a single global floor could not see, and a sabotage proved it
    // rather than an argument: recolouring every geometric node white left the
    // text dark, so all 25 charts still scored ~1% ink and all 25 passed. Per
    // kind, at half of what that chart normally covers.
    const v = judge("gutted", drawn({ ink: 0.01 }), size, 0.31);
    expect(v.ok).toBe(false);
    expect(v.problems.join("; ")).toMatch(/over half the drawing is missing/);
  });

  it("tolerates the swing a different machine's fonts produce", () => {
    // The reason this is a ratio and not a pixel diff. Antialiasing and font
    // substitution move ink by a few percent between machines; they do not
    // halve it.
    expect(judge("other-runner", drawn({ ink: 0.26 }), size, 0.31).ok).toBe(true);
  });

  it("says every reason, not just the first", () => {
    // A gate that reports one problem per run costs a round trip per problem,
    // and this one runs in CI where a round trip is a push.
    const v = judge("bad", { ink: 0, colours: 1, box: { minX: 500, minY: 500, maxX: 900, maxY: 900 } }, size);
    expect(v.problems.length).toBeGreaterThan(2);
  });
});
