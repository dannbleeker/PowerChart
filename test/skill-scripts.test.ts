import { beforeAll, describe, expect, it } from "vitest";
import { spawnSync } from "node:child_process";
import { copyFileSync, existsSync, mkdirSync, mkdtempSync, readdirSync, readFileSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import JSZip from "jszip";

/**
 * The two shipped batch renderers are plain scripts, so they are driven here as
 * subprocesses: scripts/render-batch.mjs (copied into the skill package as
 * scripts/render-svg.mjs) and skill/scripts/render-pptx.mjs. skill.test.ts
 * covers the pptx geometry; this file covers the batch and colour behaviour
 * that used to silently lose whole charts.
 *
 * (JSZip reads the .pptx — it ships with pptxgenjs, which writes it.)
 */

const LIB = "dist-lib/powerchart.js";

function ensureLib() {
  if (existsSync(LIB)) return;
  // Run vite's entry directly rather than the node_modules/.bin shim, which is
  // not executable on every dev box (Windows AppLocker).
  const r = spawnSync(process.execPath, ["node_modules/vite/bin/vite.js", "build", "--config", "vite.config.lib.ts"], {
    encoding: "utf8",
  });
  if (r.status !== 0) throw new Error(`build:lib failed: ${r.stderr}`);
}

const run = (args: string[]) => spawnSync(process.execPath, args, { encoding: "utf8" });

const write = (dir: string, name: string, cfgs: unknown) => {
  const file = join(dir, name);
  writeFileSync(file, JSON.stringify(cfgs));
  return file;
};

describe("render-batch.mjs — the skill's SVG previewer", () => {
  const dir = mkdtempSync(join(tmpdir(), "pc-batch-"));
  beforeAll(ensureLib, 180000);

  it("renders an agenda config instead of aborting the whole batch", () => {
    // kind:"agenda" is a documented config (SKILL.md), but buildChart has no such
    // kind: routing it there threw at config #1 and left the out/ directory EMPTY,
    // silently losing the waterfall and pie previews of the same run.
    const out = join(dir, "agenda-out");
    const input = write(dir, "agenda.json", [
      { kind: "agenda", chapters: ["Context", "Findings", "Recommendation"], highlight: 0 },
      {
        kind: "waterfall",
        title: "EBITDA",
        data: { categories: ["A", "B", "C"], series: [{ name: "D", values: [50, 20, 0] }] },
        waterfall: { totalIndices: [2] },
      },
      { kind: "pie", title: "Share", data: { categories: ["A", "B"], series: [{ name: "S", values: [75, 25] }] } },
    ]);
    const r = run(["scripts/render-batch.mjs", input, out]);
    expect(r.status).toBe(0);
    expect(readdirSync(out).sort()).toEqual(["01-agenda.svg", "02-ebitda.svg", "03-share.svg"]);
    // The agenda SVG is the real chapter list, not an empty frame.
    expect(readFileSync(join(out, "01-agenda.svg"), "utf8")).toContain("Recommendation");
  });

  it("isolates a bad config so the rest of the batch still renders", () => {
    const out = join(dir, "bad-out");
    const input = write(dir, "bad.json", [
      { kind: "pie" }, // no data — throws inside buildChart
      { kind: "pie", title: "Share", data: { categories: ["A"], series: [{ name: "S", values: [1] }] } },
    ]);
    const r = run(["scripts/render-batch.mjs", input, out]);
    expect(readdirSync(out)).toEqual(["02-share.svg"]); // the good one survives
    expect(r.stderr).toContain("chart 1:");
    expect(r.status).toBe(1); // …and the failure is still reported
  });

  it("names itself from argv, so the packaged copy prints its own filename", () => {
    // build-skill.mjs copies this script in as scripts/render-svg.mjs; a usage
    // line hardcoding "render-batch.mjs" pointed the skill's user at a file that
    // does not exist in the zip. Reproduce the packaged layout in a temp dir.
    const pkg = mkdtempSync(join(tmpdir(), "pc-pkg-"));
    mkdirSync(join(pkg, "scripts"));
    mkdirSync(join(pkg, "dist-lib"));
    copyFileSync("scripts/render-batch.mjs", join(pkg, "scripts", "render-svg.mjs"));
    copyFileSync(LIB, join(pkg, "dist-lib", "powerchart.js"));
    const r = run([join(pkg, "scripts", "render-svg.mjs")]);
    expect(r.status).toBe(1);
    expect(r.stderr).toContain("usage: node scripts/render-svg.mjs");
    expect(r.stderr).not.toContain("render-batch.mjs");
  });
});

describe("render-pptx.mjs — paints", () => {
  const dir = mkdtempSync(join(tmpdir(), "pc-paint-"));
  const out = join(dir, "paints.pptx");
  const slides: string[] = [];

  beforeAll(async () => {
    ensureLib();
    const input = write(dir, "paints.json", [
      {
        kind: "clustered",
        data: {
          categories: ["Q1", "Q2"],
          series: [
            { name: "Alpha", values: [10, 20], color: "red" },
            { name: "Beta", values: [5, 8], color: "steelblue" },
          ],
        },
      },
      {
        kind: "clustered",
        data: { categories: ["Q1"], series: [{ name: "A", values: [10], color: "rgba(100, 150, 200, 50%)" }] },
      },
      {
        kind: "clustered",
        title: "Dark deck chart",
        style: { background: "#1a1a1a", text: "#ffffff", mutedText: "#cccccc", axis: "#888888" },
        data: { categories: ["Q1", "Q2"], series: [{ name: "Sales", values: [10, 20] }] },
      },
      {
        kind: "mekko",
        data: {
          categories: ["Q1", "Q2"],
          series: [
            { name: "Base", values: [10, 12], color: "transparent" },
            { name: "Range", values: [20, 18] },
          ],
        },
      },
    ]);
    const r = run(["skill/scripts/render-pptx.mjs", input, out]);
    if (r.status !== 0) throw new Error(`render-pptx failed: ${r.stderr}`);
    const zip = await JSZip.loadAsync(readFileSync(out));
    for (let i = 1; i <= 4; i++) slides.push(await zip.file(`ppt/slides/slide${i}.xml`)!.async("string"));
  }, 180000);

  it("resolves named CSS colours instead of collapsing them to one grey", () => {
    // Both series used to render 808080, so the deck could no longer tell Alpha
    // from Beta — the SVG and Office.js renderers both keep the two hues.
    expect(slides[0]).toContain('val="FF0000"'); // red
    expect(slides[0]).toContain('val="4682B4"'); // steelblue
    expect(slides[0]).not.toContain('val="808080"');
  });

  it("reads a percentage alpha without scaling the RGB channels to white", () => {
    // `/%/` over the whole argument list treated the 50% ALPHA as a percent
    // colour and multiplied every channel by 2.55 → FFFFFF.
    expect(slides[1]).toContain('val="6496C8"');
    expect(slides[1]).toMatch(/<a:srgbClr val="6496C8"><a:alpha val="50000"\/>/);
  });

  it("gives the slide the chart's own background, so dark ink stays readable", () => {
    expect(slides[2]).toContain('<p:bg><p:bgPr><a:solidFill><a:srgbClr val="1A1A1A"/>');
    expect(slides[0]).toContain('<p:bg><p:bgPr><a:solidFill><a:srgbClr val="FFFFFF"/>'); // default unchanged
  });

  it("draws a transparent series as nothing, not as a solid block", () => {
    // color:"transparent" is the documented floating-segment idiom; it has no
    // hex, and painting it as a named colour put opaque blocks on the slide.
    expect(slides[3]).not.toContain('val="808080"');
    expect(slides[3]).not.toContain('val="000000"');
  });
});
