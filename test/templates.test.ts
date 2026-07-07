import { describe, expect, it } from "vitest";
import { BUILTIN_TEMPLATES } from "../src/taskpane/templates";
import { buildChart, DEFAULT_SIZE } from "../src/core/chart";
import { sceneToSvg } from "../src/render/svg";

describe("built-in templates", () => {
  it("offers a non-empty, uniquely named starter set", () => {
    expect(BUILTIN_TEMPLATES.length).toBeGreaterThanOrEqual(5);
    const names = BUILTIN_TEMPLATES.map((t) => t.name);
    expect(new Set(names).size).toBe(names.length);
  });

  it.each(BUILTIN_TEMPLATES.map((t) => [t.name, t] as const))(
    "%s builds a non-empty, renderable chart",
    (_name, t) => {
      const scene = buildChart({ ...DEFAULT_SIZE, ...t.config });
      expect(scene.nodes.length).toBeGreaterThan(0);
      const svg = sceneToSvg(scene);
      expect(svg).toContain("<svg");
      expect(svg).not.toContain("NaN");
    },
  );
});
