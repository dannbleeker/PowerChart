#!/usr/bin/env node
/**
 * Rasterise every sample chart in a real browser and ask whether it is VISIBLE.
 *
 * The snapshot suite freezes each chart's SVG *text*, which is a strong check on
 * a great many things and blind to exactly one class: a chart that is
 * structurally perfect and cannot be seen. White fill on a white background, a
 * scale collapse that puts every shape at zero size, a translate that pushes the
 * drawing off the canvas — every one of those produces a stable, diffable,
 * entirely wrong SVG, and every one of them passes today.
 *
 * That class has a name in this repo already: the self-test's
 * `the chart is actually visible` asks PowerPoint the same question about the
 * live add-in, and it exists because nothing but a human looking at a deck had
 * ever asked it. This asks it of the SVG renderer, in CI, on every commit.
 *
 * **Not a pixel diff, and deliberately.** Committed baseline images would tie
 * this to one machine's fonts and antialiasing: the browser here and the browser
 * on a CI runner disagree about text rendering in ways nobody wants to arbitrate
 * every week, and a visual gate that cries wolf is a visual gate that gets
 * deleted. What is asserted instead are properties that hold on any renderer:
 * there is ink; the ink is not all one colour; it is inside the frame; and there
 * is neither almost none of it nor almost nothing but.
 */
import { accessSync, constants, readFileSync, writeFileSync, mkdirSync } from "fs";
import { chromium } from "playwright-core";
import { isMain } from "./is-main.mjs";

/**
 * Where a browser lives. The env var is what CI sets; the rest are what dev
 * boxes have.
 *
 * The Windows entries are here because the owner's box is Windows and had no
 * candidate at all: playwright-core ships no browser, so `chromium.launch({})`
 * threw, and before the entry guard was fixed the script never got that far
 * anyway. Installed Chrome is what this repo already uses for its other visual
 * checks.
 */
const BROWSER_CANDIDATES = [
  process.env.PLAYWRIGHT_CHROMIUM,
  process.env.CHROME_BIN,
  "/opt/pw-browsers/chromium",
  "/usr/bin/chromium",
  "/usr/bin/chromium-browser",
  "/usr/bin/google-chrome",
  "C:/Program Files/Google/Chrome/Application/chrome.exe",
  "C:/Program Files (x86)/Google/Chrome/Application/chrome.exe",
  "C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe",
].filter(Boolean);

/**
 * A browser path, or undefined to let playwright-core find its own.
 *
 * Undefined is the CI answer, not a fallback: `playwright-core install chromium`
 * puts the binary under its own registry (`~/.cache/ms-playwright/...`), which
 * no fixed path here would ever match, and passing a wrong `executablePath` is
 * how this gate would have failed on its first run. The candidates exist for the
 * environments that pre-install a browser somewhere known and set no env var.
 *
 * Never a skip either way. If neither a candidate nor the registry has one,
 * `chromium.launch()` throws and says so — a visual gate that quietly does
 * nothing when it cannot find a browser is worse than no gate, because CI stays
 * green and nobody knows the check stopped running.
 */
function findBrowser() {
  for (const path of BROWSER_CANDIDATES) {
    try {
      // `accessSync`, not `readFileSync`: the old check read the entire browser
      // binary into memory — a few hundred megabytes — to answer "does this file
      // exist", once per candidate, on a runner with a memory limit.
      accessSync(path, constants.R_OK);
      return path;
    } catch {
      /* next */
    }
  }
  return undefined;
}

/** The floor below which nothing is drawn at all, whatever the chart. */
const MIN_INK = 0.004;

/**
 * And a floor per KIND, at this fraction of what that chart normally covers.
 *
 * The global floor alone is too loose, and a sabotage proved it rather than an
 * argument: recolouring every geometric node white left the labels dark, so all
 * 25 charts still scored 0.7-1.0% ink and all 25 passed. A bar chart whose bars
 * have vanished and whose axis text has not is a total regression, and the gate
 * has to see it.
 *
 * Half of normal, because that tolerates the thing pixel baselines cannot — a
 * different machine's fonts and antialiasing move ink by a few percent, never by
 * half — while a chart that has lost its geometry loses far more than half.
 * `--update` rewrites the recorded coverage; the numbers are a measurement, not
 * a preference.
 */
const KIND_FLOOR = 0.5;
const COVERAGE_FILE = new URL("../test/fixtures/chart-ink.json", import.meta.url);
/** And how much is so much that the chart is a solid block rather than a chart. */
const MAX_INK = 0.9;
/** Distinct colours below this and the "chart" is one flat shape. */
const MIN_COLOURS = 3;
/** Ink this close to the background does not count as ink. */
const MIN_CONTRAST = 24;

/**
 * Read one rendered frame and say what is actually on it.
 *
 * Runs inside the page, because a canvas is the cheapest honest rasteriser
 * available and the browser has already done the hard part.
 */
export const MEASURE = `(svg, w, h) => new Promise((resolve, reject) => {
  const img = new Image();
  img.onload = () => {
    const c = document.createElement("canvas");
    c.width = w; c.height = h;
    const ctx = c.getContext("2d", { willReadFrequently: true });
    // White, because that is what a slide is. Measuring against transparent
    // would score a white-on-white chart as full of ink.
    ctx.fillStyle = "#ffffff";
    ctx.fillRect(0, 0, w, h);
    ctx.drawImage(img, 0, 0, w, h);
    const { data } = ctx.getImageData(0, 0, w, h);
    const colours = new Set();
    let ink = 0, minX = w, minY = h, maxX = -1, maxY = -1;
    for (let i = 0; i < data.length; i += 4) {
      const r = data[i], g = data[i + 1], b = data[i + 2];
      // Distance from the white background. Anti-aliased edges score low and
      // are meant to: a chart made only of edges is not a visible chart.
      const off = Math.max(255 - r, 255 - g, 255 - b);
      if (off < ${MIN_CONTRAST}) continue;
      ink++;
      colours.add((r >> 3) << 10 | (g >> 3) << 5 | (b >> 3));
      const p = i / 4, x = p % w, y = (p / w) | 0;
      if (x < minX) minX = x; if (x > maxX) maxX = x;
      if (y < minY) minY = y; if (y > maxY) maxY = y;
    }
    resolve({ ink: ink / (w * h), colours: colours.size, box: { minX, minY, maxX, maxY } });
  };
  img.onerror = () => reject(new Error("the browser would not decode the SVG"));
  img.src = "data:image/svg+xml;charset=utf-8," + encodeURIComponent(svg);
})`;

/** Turn one measurement into a verdict, with the reason a reader needs. */
export function judge(name, m, size, expectedInk) {
  const problems = [];
  if (m.ink < MIN_INK) problems.push(`only ${(m.ink * 100).toFixed(2)}% of the frame carries ink`);
  else if (typeof expectedInk === "number" && m.ink < expectedInk * KIND_FLOOR)
    problems.push(
      `${(m.ink * 100).toFixed(1)}% ink against ${(expectedInk * 100).toFixed(1)}% normally — ` +
        `over half the drawing is missing`,
    );
  if (m.ink > MAX_INK) problems.push(`${(m.ink * 100).toFixed(0)}% of the frame is ink — a block, not a chart`);
  if (m.colours < MIN_COLOURS) problems.push(`${m.colours} distinct colour(s) — nothing is distinguishable`);
  // Drawn outside the frame is the failure a text snapshot is least able to
  // see: the SVG is identical whether a translate is right or wildly wrong.
  if (m.box.maxX < 0) problems.push("nothing was drawn at all");
  else if (m.box.minX < 0 || m.box.minY < 0 || m.box.maxX >= size.w || m.box.maxY >= size.h)
    problems.push(`ink reaches outside the frame (${JSON.stringify(m.box)})`);
  return { name, ok: problems.length === 0, problems, ink: m.ink, colours: m.colours };
}

async function main() {
  // The built bundle, exactly like `build-showcase.mjs` — one artifact, one
  // entry point, and the same code the skill ships. Run `npm run build:lib`
  // first; a stale bundle here would gate the OLD renderer, which is the trap
  // `CLAUDE.md` already records for `skill-scripts.test.ts`.
  const { buildChart, sampleConfig, CHART_KINDS, sceneToSvg } = await import("../dist-lib/powerchart.js");

  const update = process.argv.includes("--update");
  const coverage = update ? {} : JSON.parse(readFileSync(COVERAGE_FILE, "utf8"));
  const size = { w: 720, h: 460 };
  const found = findBrowser();
  const browser = await chromium.launch(found ? { executablePath: found } : {});
  const page = await browser.newPage({ viewport: { width: size.w, height: size.h } });
  const results = [];
  const failing = [];
  try {
    for (const { kind } of CHART_KINDS) {
      const scene = buildChart(sampleConfig(kind));
      const svg = sceneToSvg(scene, { background: "#ffffff" });
      const m = await page.evaluate(`(${MEASURE})(${JSON.stringify(svg)}, ${size.w}, ${size.h})`);
      const verdict = judge(kind, m, size, coverage[kind]);
      if (update) coverage[kind] = Math.round(m.ink * 10000) / 10000;
      results.push(verdict);
      if (!verdict.ok) {
        failing.push(verdict);
        // The evidence, beside the verdict. A gate that says "this chart is
        // invisible" and hands over no way to look at it sends the next person
        // straight back to reproducing it by hand.
        mkdirSync("visible-charts", { recursive: true });
        writeFileSync(`visible-charts/${kind}.svg`, svg);
      }
    }
  } finally {
    await browser.close();
  }
  if (update) {
    writeFileSync(COVERAGE_FILE, `${JSON.stringify(coverage, null, 2)}\n`);
    process.stdout.write(`recorded ink coverage for ${Object.keys(coverage).length} charts\n`);
    return;
  }
  for (const r of results) {
    process.stdout.write(
      r.ok
        ? `  ok   ${r.name} — ${(r.ink * 100).toFixed(1)}% ink, ${r.colours} colours\n`
        : `  FAIL ${r.name} — ${r.problems.join("; ")}\n`,
    );
  }
  if (failing.length) {
    process.stdout.write(`\n${failing.length} of ${results.length} charts are not visibly drawn.\n`);
    process.stdout.write(`Their SVG is in visible-charts/ — open one and look at it.\n`);
    process.exitCode = 1;
    return;
  }
  process.stdout.write(`\n${results.length} charts render something a person could see.\n`);
}

if (isMain(import.meta.url, process.argv[1])) {
  await main();
}
