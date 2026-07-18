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

  it("emits solid pie wedges as custGeom, not the wrap-broken pie preset", () => {
    // A .pptx is a zip; read the slide XML via python3 zipfile.
    const xml = execSync(
      `python3 -c "import zipfile;print(zipfile.ZipFile('${out}').read('ppt/slides/slide1.xml').decode())"`,
    ).toString();
    // The OOXML "pie" preset computes swAng = end − start over two independently
    // normalized angles, so any slice crossing 3 o'clock (every pie has one)
    // renders the wrong wedge. Solid pies now take the SAME custGeom path as the
    // doughnut ring — annularSectorPoints sampled from polar(), correct across the
    // boundary. So there must be no "pie" preset, and one custGeom arc per slice.
    expect(xml).not.toContain('prst="pie"');
    expect(xml.match(/<a:custGeom>/g)?.length).toBe(2); // one filled wedge per slice
    expect(xml).not.toContain("NaN");
    // Each wedge is a real sampled arc, not a degenerate 2-point shape.
    expect(xml.match(/<a:lnTo>/g)?.length ?? 0).toBeGreaterThan(10);
    expect(xml).toContain("Split");
  });
});

describe("skill pptx renderer — annular sectors", () => {
  const dir = mkdtempSync(join(tmpdir(), "pc-rings-"));
  const out = join(dir, "rings.pptx");
  const readSlide = (n: number) =>
    execSync(
      `python3 -c "import zipfile;print(zipfile.ZipFile('${out}').read('ppt/slides/slide${n}.xml').decode())"`,
    ).toString();

  beforeAll(() => {
    if (!existsSync("dist-lib/powerchart.js")) {
      execSync("npx vite build --config vite.config.lib.ts", { stdio: "pipe" });
    }
    const cfgs = [
      {
        kind: "sunburst",
        data: {
          categories: ["A", "B"],
          series: [
            { name: "L1", values: [60, 40] },
            { name: "L2", values: [30, 30] },
          ],
        },
      },
      {
        kind: "doughnut",
        pie: { semi: true },
        data: { categories: ["X", "Y", "Z"], series: [{ name: "S", values: [50, 30, 20] }] },
      },
    ];
    const input = join(dir, "cfgs.json");
    writeFileSync(input, JSON.stringify(cfgs));
    execSync(`node skill/scripts/render-pptx.mjs ${input} ${out}`, { stdio: "pipe" });
  }, 120000);

  it("emits real filled custGeom annular sectors for sunburst rings (not center-anchored pie slices)", () => {
    const xml = readSlide(1);
    expect(xml).toContain("custGeom");
    expect(xml).not.toContain("NaN");
  });

  it("emits custGeom for the semi-doughnut gauge, honouring the inner radius", () => {
    const xml = readSlide(2);
    expect(xml.match(/custGeom/g)?.length ?? 0).toBeGreaterThanOrEqual(1);
    expect(xml).not.toContain("NaN");
  });
});

describe("skill pptx renderer — 8-digit hex colours", () => {
  const dir = mkdtempSync(join(tmpdir(), "pc-alpha-"));
  const out = join(dir, "alpha.pptx");

  beforeAll(() => {
    if (!existsSync("dist-lib/powerchart.js")) {
      execSync("npx vite build --config vite.config.lib.ts", { stdio: "pipe" });
    }
    // #RRGGBBAA is a valid, SVG-honoured colour form. A single explicit series
    // colour paints the column rects with it.
    const cfg = {
      kind: "clustered",
      data: { categories: ["A", "B"], series: [{ name: "S", values: [3, 5], color: "#4e79a780" }] },
    };
    const input = join(dir, "cfg.json");
    writeFileSync(input, JSON.stringify(cfg));
    execSync(`node skill/scripts/render-pptx.mjs ${input} ${out}`, { stdio: "pipe" });
  }, 120000);

  it("keeps the hue and folds the alpha into transparency (not solid black)", () => {
    const xml = execSync(
      `python3 -c "import zipfile;print(zipfile.ZipFile('${out}').read('ppt/slides/slide1.xml').decode())"`,
    ).toString();
    // pptxgenjs only validates 6-digit RGB: an unhandled #RRGGBBAA would have been
    // rejected and replaced with DEF_FONT_COLOR ("000000"). The hue must survive
    // as the 6-digit prefix, with the alpha byte (0x80) carried as an <a:alpha>.
    expect(xml).toMatch(/<a:srgbClr val="4E79A7"><a:alpha val="\d+"\/>/);
  });
});

describe("packaged skill layout", () => {
  beforeAll(() => {
    if (!existsSync("dist-lib/powerchart.js")) {
      execSync("npx vite build --config vite.config.lib.ts", { stdio: "pipe" });
    }
    execSync("node scripts/build-skill.mjs", { stdio: "pipe" });
  }, 120000);

  it("points the SVG renderer at the bundled lib, not the repo's dist-lib", () => {
    // render-svg.mjs is copied from the repo, where the engine lives at
    // ../dist-lib — a path that does not exist inside the package. The rewrite
    // that fixes this used to be shelled out to `node -e`, which silently did
    // nothing on Windows while still exiting 0, shipping a renderer that could
    // not start.
    const src = readFileSync("skill-dist/powerchart-charts/scripts/render-svg.mjs", "utf8");
    expect(src).toContain("../lib/powerchart.js");
    expect(src).not.toContain("../dist-lib/powerchart.js");
    // And the path it now imports has to be real.
    expect(existsSync("skill-dist/powerchart-charts/lib/powerchart.js")).toBe(true);
  });
});
