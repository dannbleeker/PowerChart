import { describe, expect, it } from "vitest";
import { buildChart, describeChart } from "../src/core/chart";
import { sceneToSvg } from "../src/render/svg";
import type { ChartConfig } from "../src/core/types";

const base: ChartConfig = {
  kind: "clustered",
  width: 480,
  height: 300,
  title: "Revenue by region",
  data: {
    categories: ["North", "South", "East"],
    series: [
      { name: "2024", values: [10, 20, 30] },
      { name: "2025", values: [12, 18, 33] },
    ],
  },
};

describe("SVG accessibility", () => {
  it("marks the root as an image and gives it a title + description", () => {
    const svg = sceneToSvg(buildChart(base));
    expect(svg).toContain('role="img"');
    expect(svg).toContain("<title>Revenue by region</title>");
    expect(svg).toContain("<desc>");
    // title/desc must be the FIRST children of the root for the SVG a11y mapping
    expect(svg.indexOf("<title>")).toBeLessThan(svg.indexOf("<rect"));
    expect(svg.indexOf("<title>")).toBeLessThan(svg.indexOf("<desc>"));
  });

  it("escapes markup in the title", () => {
    const svg = sceneToSvg(buildChart({ ...base, title: "A & B <chart>" }));
    expect(svg).toContain("<title>A &amp; B &lt;chart&gt;</title>");
  });

  it("names an untitled chart with its generated summary", () => {
    // role="img" with a <desc> but no <title> has a description and no NAME —
    // the axe-core role-img-alt failure. The summary becomes the name instead,
    // and is not then repeated as the description.
    const cfg = { ...base, title: undefined };
    const svg = sceneToSvg(buildChart(cfg));
    expect(svg).toContain('role="img"');
    expect(svg).toContain(`<title>${describeChart(cfg)}</title>`);
    expect(svg).not.toContain("<desc>");
    expect(svg.indexOf("<title>")).toBeLessThan(svg.indexOf("<rect"));
  });

  it("describeChart names the kind, series and categories", () => {
    const d = describeChart(base);
    expect(d).toContain("clustered column chart");
    expect(d).toContain("2 data series");
    expect(d).toContain("2024, 2025");
    expect(d).toContain("3 categories");
    expect(d).toContain("North, South, East");
  });

  it("describeChart notes horizontal orientation and truncates long lists", () => {
    expect(describeChart({ ...base, horizontal: true })).toContain("(horizontal)");
    const many = describeChart({
      ...base,
      data: { categories: ["a", "b", "c", "d", "e", "f"], series: [{ name: "S", values: [1, 2, 3, 4, 5, 6] }] },
    });
    expect(many).toContain("and 2 more");
  });
});
