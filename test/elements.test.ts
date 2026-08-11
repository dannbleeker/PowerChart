import { describe, expect, it } from "vitest";
import { buildHarveyBall, buildKpiTile, buildProcessFlow, buildTableScene } from "../src/core/elements";
import { buildAgendaScene } from "../src/core/agenda";
import type { ArrowheadNode, Scene, TextNode } from "../src/core/scene";

/** Standalone elements — KPI tile scene layout. */

describe("KPI tile element", () => {
  it("shows caption, big value, and a colored delta with arrow", () => {
    const s = buildKpiTile({ label: "Revenue", value: "€4.2m", delta: "+12% vs LY" });
    expect((s.nodes.find((n) => n.name === "kpi-value") as TextNode).text).toBe("€4.2m");
    expect((s.nodes.find((n) => n.name === "kpi-label") as TextNode).text).toBe("Revenue");
    const arrow = s.nodes.find((n) => n.name === "kpi-arrow") as ArrowheadNode;
    expect(arrow.angle).toBe(-90); // up
    expect(arrow.fill).toBe("#0ca30c"); // up is good by default
    expect((s.nodes.find((n) => n.name === "kpi-delta") as TextNode).color).toBe("#0ca30c");
  });

  it("goodIsUp:false colors a falling metric green", () => {
    const s = buildKpiTile({ value: "2.1%", delta: "-0.4pp churn", goodIsUp: false });
    const arrow = s.nodes.find((n) => n.name === "kpi-arrow") as ArrowheadNode;
    expect(arrow.angle).toBe(90); // down
    expect(arrow.fill).toBe("#0ca30c"); // …and that's good
    const up = buildKpiTile({ value: "2.1%", delta: "+0.4pp churn", goodIsUp: false });
    expect((up.nodes.find((n) => n.name === "kpi-arrow") as ArrowheadNode).fill).toBe("#d03b3b");
  });

  it("flat or missing deltas stay neutral, long values shrink to fit", () => {
    const flat = buildKpiTile({ value: "87", delta: "unchanged", direction: "flat" });
    expect(flat.nodes.some((n) => n.name === "kpi-arrow")).toBe(false);
    expect((flat.nodes.find((n) => n.name === "kpi-delta") as TextNode).color).not.toBe("#0ca30c");
    const none = buildKpiTile({ value: "87" });
    expect(none.nodes.some((n) => n.name === "kpi-delta")).toBe(false);
    const long = buildKpiTile({ value: "€1,234,567.89 total" });
    const short = buildKpiTile({ value: "€4m" });
    const fsOf = (s: ReturnType<typeof buildKpiTile>) =>
      (s.nodes.find((n) => n.name === "kpi-value") as TextNode).fontSize;
    expect(fsOf(long)).toBeLessThan(fsOf(short));
  });
});

/**
 * The gate that protects charts did not protect elements.
 *
 * `buildChart` ends with `finiteNodes` — the engine's last stop before a scene
 * reaches three renderers, two of which write a FILE, so a bad coordinate makes
 * a .pptx PowerPoint may refuse to open. The five element builders returned
 * their nodes unfiltered, and they are public API: `src/index.ts` exports every
 * one of them.
 *
 * Nothing in this repo reaches it today — the pane passes literal sizes and a
 * range slider, and the skill renders charts only — so this is a latent
 * asymmetry in an exported surface rather than a reported crash. It is closed
 * at the same seam and for the same stated reason.
 */
describe("what an element builder promises about its coordinates", () => {
  it("never returns a node with a non-finite number", () => {
    const scenes: [string, Scene][] = [
      ["harvey", buildHarveyBall(NaN)],
      ["flow size", buildProcessFlow(["a", "b"], 0, NaN, NaN)],
      ["flow height", buildProcessFlow(["a", "b"], 0, 480, Infinity)],
      ["kpi", buildKpiTile({ label: "a", value: "1" }, NaN, NaN)],
      ["table", buildTableScene([["a", "b"]], NaN)],
    ];
    const bad: string[] = [];
    for (const [name, scene] of scenes)
      for (const n of scene.nodes)
        for (const [k, v] of Object.entries(n))
          if (typeof v === "number" && !Number.isFinite(v)) bad.push(`${name} → ${n.kind}.${k}=${v}`);
    expect(bad, "an element handed a renderer a coordinate that is not a number").toEqual([]);
  });

  it("builds an empty agenda rather than throwing when there are no chapters", () => {
    // Same defect, same class: an exported builder that trusts its parameter's
    // declared type. An agenda with no chapters is an ordinary thing to ask
    // for, and `src/index.ts` exports this to callers this repo never sees.
    for (const chapters of [undefined, null, "a,b", 42, {}] as unknown as string[][])
      expect(() => buildAgendaScene(chapters), `buildAgendaScene(${JSON.stringify(chapters)}) threw`).not.toThrow();
    expect(buildAgendaScene(["One", "Two"]).nodes.length).toBeGreaterThan(0);
  });

  it("builds an empty flow rather than throwing when there are no steps", () => {
    // A `string[]` in the types is not an array in the object someone passed.
    // These are exported, so the caller may be anyone.
    for (const steps of [undefined, null, "a,b", 42, {}] as unknown as string[][])
      expect(() => buildProcessFlow(steps), `buildProcessFlow(${JSON.stringify(steps)}) threw`).not.toThrow();
    expect(buildProcessFlow(undefined as unknown as string[]).nodes.length).toBeGreaterThanOrEqual(0);
    // And a real list still draws every step.
    expect(buildProcessFlow(["a", "b", "c"]).nodes.filter((n) => n.kind === "chevron")).toHaveLength(3);
  });
});
