import { describe, expect, it } from "vitest";
import { DEFAULT_SIZE, buildChart } from "../src/core/chart";
import { DEFAULT_STYLE } from "../src/core/style";
import type { LineNode, RectNode } from "../src/core/scene";
import type { ChartConfig } from "../src/core/types";

/** Candlestick / OHLC. */

describe("candlestick", () => {
  const cfg: ChartConfig = {
    kind: "candlestick",
    ...DEFAULT_SIZE,
    data: {
      categories: ["D1", "D2"],
      series: [
        { name: "Open", values: [40, 46] },
        { name: "High", values: [48, 47] },
        { name: "Low", values: [39, 42] },
        { name: "Close", values: [46, 43] }, // D1 rises, D2 falls
      ],
    },
  };
  const s = buildChart(cfg);

  it("draws a high–low wick and an open–close body per period", () => {
    expect(s.nodes.some((n): n is LineNode => n.kind === "line" && n.name === "wick-0")).toBe(true);
    expect(s.nodes.some((n): n is RectNode => n.kind === "rect" && n.name === "body-0")).toBe(true);
    expect(s.nodes.some((n) => n.name === "body-1")).toBe(true);
  });

  it("encodes rising vs falling redundantly: hollow green up, solid red down", () => {
    const up = s.nodes.find((n): n is RectNode => n.name === "body-0")!; // close 46 > open 40
    const down = s.nodes.find((n): n is RectNode => n.name === "body-1")!; // close 43 < open 46
    // Rising is HOLLOW (background fill, green outline); falling is SOLID red —
    // so direction reads without colour (greyscale / CVD safe), colour reinforces.
    expect(up.fill).toBe(DEFAULT_STYLE.background); // hollow body
    expect(up.stroke).toBe("#1a9e6e"); // green outline
    expect(down.fill).toBe(DEFAULT_STYLE.negative); // solid red
    expect(up.fill).not.toBe(down.fill);
  });

  it("the wick spans the full high–low range", () => {
    const wick = s.nodes.find((n): n is LineNode => n.name === "wick-0")!;
    // High (48) maps above Low (39): y1 (high) < y2 (low) on screen.
    expect(wick.y1).toBeLessThan(wick.y2);
  });
});
