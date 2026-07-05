import { beforeAll, describe, expect, it } from "vitest";
import { execSync } from "node:child_process";
import { existsSync, mkdtempSync, readFileSync, statSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";

/**
 * End-to-end check of the Agent Skill's pptx renderer: build the lib, render
 * a config, and assert the OOXML contains the expected native shapes.
 */
describe("skill pptx renderer", () => {
  const dir = mkdtempSync(join(tmpdir(), "pc-skill-"));
  const out = join(dir, "out.pptx");

  beforeAll(() => {
    if (!existsSync("dist-lib/powerchart.js")) {
      execSync("npx vite build --config vite.config.lib.ts", { stdio: "pipe" });
    }
    const cfg = {
      kind: "pie",
      title: "Split",
      data: { categories: ["A", "B"], series: [{ name: "S", values: [75, 25] }] },
    };
    const input = join(dir, "cfg.json");
    writeFileSync(input, JSON.stringify(cfg));
    execSync(`node skill/scripts/render-pptx.mjs ${input} ${out}`, { stdio: "pipe" });
  }, 120000);

  it("produces a non-trivial pptx", () => {
    expect(statSync(out).size).toBeGreaterThan(10000);
  });

  it("contains exact pie wedges (preset geometry with angle adjustments)", () => {
    // A .pptx is a zip; the slide XML is stored deflated but the preset names
    // survive a raw scan after inflation via python (available in CI image?).
    // Portable approach: unzip via execSync + python3 zipfile.
    const xml = execSync(
      `python3 -c "import zipfile;print(zipfile.ZipFile('${out}').read('ppt/slides/slide1.xml').decode())"`,
    ).toString();
    expect(xml.match(/prst="pie"/g)?.length).toBe(2);
    // 75% slice: scene 0→270° = OOXML 270→180.
    expect(xml).toContain('name="adj1" fmla="val 16200000"'); // 270°
    expect(xml).toContain('name="adj2" fmla="val 10800000"'); // 180°
    expect(xml).toContain("Split");
  });
});
