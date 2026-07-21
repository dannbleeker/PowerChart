import { describe, expect, it } from "vitest";
import { buildAgendaScene, SLIDE } from "../src/core/agenda";
import { textWidth } from "../src/core/scene";
import type { TextNode } from "../src/core/scene";

describe("agenda slides", () => {
  const chapters = ["Intro", "Market", "Strategy"];

  it("lists every chapter with the highlighted one emphasized", () => {
    const scene = buildAgendaScene(chapters, { highlight: 1 });
    expect(scene.width).toBe(SLIDE.width);
    const items = scene.nodes.filter((n): n is TextNode => n.kind === "text" && !!n.name?.startsWith("agenda-item-"));
    expect(items.map((i) => i.text)).toEqual(chapters);
    expect(items[1].bold).toBe(true);
    expect(items[0].bold).toBe(false);
    // Highlight bar behind the active chapter only.
    expect(scene.nodes.filter((n) => n.name?.startsWith("agenda-hl-"))).toHaveLength(1);
  });

  it("renders an overview slide with no highlight", () => {
    const scene = buildAgendaScene(chapters);
    expect(scene.nodes.some((n) => n.name?.startsWith("agenda-hl-"))).toBe(false);
  });

  it("keeps a long chapter title on the slide", () => {
    // Neither PowerPoint renderer wraps the row, so an unshrunk title of this
    // length ran ~165pt past the right edge of the 960pt slide.
    const item = (scene: ReturnType<typeof buildAgendaScene>) =>
      scene.nodes.find((n): n is TextNode => n.name === "agenda-item-0")!;
    const inkRight = (t: TextNode) => t.x + textWidth(t.text, t.fontSize, t.bold);

    const title = "Financial performance and outlook for the remainder of the fiscal year and the medium-term plan";
    const long = item(buildAgendaScene([title], { highlight: 0 }));
    expect(long.fontSize).toBeLessThan(18);
    expect(long.text).toBe(title); // shrunk to fit, not truncated
    expect(inkRight(long)).toBeLessThanOrEqual(SLIDE.width);

    // Beyond what the smallest row font can hold, the title is ellipsized.
    const huge = item(buildAgendaScene(["Chapter ".repeat(60)]));
    expect(huge.text.endsWith("…")).toBe(true);
    expect(inkRight(huge)).toBeLessThanOrEqual(SLIDE.width);

    // Titles that already fit keep the full-size row font.
    expect(item(buildAgendaScene(chapters)).fontSize).toBe(18);
  });
});
