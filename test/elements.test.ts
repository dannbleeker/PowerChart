import { describe, expect, it } from "vitest";
import { buildKpiTile } from "../src/core/elements";
import type { ArrowheadNode, TextNode } from "../src/core/scene";

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
