import { describe, expect, it } from "vitest";
import { buildAgendaScene, SLIDE } from "../src/core/agenda";
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
});
