#!/usr/bin/env node
/**
 * Render PowerChart configs to a .pptx with NATIVE, editable shapes —
 * the same output philosophy as the live add-in, headless.
 *
 *   node scripts/render-pptx.mjs charts.json out.pptx
 *
 * charts.json holds one ChartConfig or an array (one chart per slide).
 * Requires: npm install pptxgenjs
 */
import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import pptxgen from "pptxgenjs";

// Engine location: packaged skill layout first, then repo layout.
let engine;
const failures = [];
for (const candidate of ["../lib/powerchart.js", "../../dist-lib/powerchart.js"]) {
  const href = new URL(candidate, import.meta.url).href;
  try {
    engine = await import(href);
    break;
  } catch (err) {
    // THIS candidate being absent just means "not this layout" — keep looking.
    // Anything else is the engine itself failing, and swallowing it reported a
    // corrupt-but-present lib as missing, telling the user to rebuild a file
    // that is already there (and the packaged skill has no build:lib to run).
    // Node names the missing file by path, not by href — a missing import
    // *inside* the engine names a different one, and that is a real failure.
    const missingSelf =
      err?.code === "ERR_MODULE_NOT_FOUND" &&
      String(err?.message ?? "").includes(fileURLToPath(href));
    if (!missingSelf) failures.push(`${candidate}: ${err?.message ?? err}`);
  }
}
if (!engine) {
  console.error(
    failures.length
      ? `powerchart engine failed to load:\n  ${failures.join("\n  ")}`
      : "powerchart engine not found — run `npm run build:lib` first",
  );
  process.exit(1);
}
const { buildChart, DEFAULT_SIZE, arrowheadBox, sceneToOoxmlPieAngle, annularSectorPoints } = engine;

const [, , input, output = "powerchart.pptx"] = process.argv;
if (!input) {
  console.error("usage: node scripts/render-pptx.mjs <charts.json> [out.pptx]");
  process.exit(1);
}

const raw = JSON.parse(readFileSync(input, "utf8"));
const configs = Array.isArray(raw) ? raw : [raw];

const IN = 1 / 72; // points → inches
const SLIDE = { w: 13.333, h: 7.5 };

const pres = new pptxgen();
pres.defineLayout({ name: "WIDE", width: SLIDE.w, height: SLIDE.h });
pres.layout = "WIDE";

const hex = (c) => (c ?? "000000").replace("#", "");

/** Map one scene node to PptxgenJS calls at slide offset (inches). */
function addNode(slide, n, dx, dy) {
  switch (n.kind) {
    case "rect": {
      slide.addShape("rect", {
        x: dx + n.x * IN,
        y: dy + n.y * IN,
        w: Math.max(0.003, n.w * IN),
        h: Math.max(0.003, n.h * IN),
        fill: { color: hex(n.fill) },
        line: n.stroke && (n.strokeWidth ?? 0) > 0 ? { color: hex(n.stroke), width: n.strokeWidth } : { type: "none" },
      });
      break;
    }
    case "line": {
      // PptxgenJS draws lines TL→BR of the box; flip for rising segments.
      const rising = (n.x2 - n.x1) * (n.y2 - n.y1) < 0;
      slide.addShape("line", {
        x: dx + Math.min(n.x1, n.x2) * IN,
        y: dy + Math.min(n.y1, n.y2) * IN,
        w: Math.abs(n.x2 - n.x1) * IN,
        h: Math.abs(n.y2 - n.y1) * IN,
        flipV: rising,
        line: {
          color: hex(n.stroke),
          width: n.strokeWidth ?? 1,
          ...(n.dash ? { dashType: "dash" } : {}),
        },
      });
      break;
    }
    case "text": {
      slide.addText(n.text, {
        x: dx + n.x * IN,
        y: dy + n.y * IN,
        w: Math.max(0.05, n.w * IN),
        h: Math.max(0.05, n.h * IN),
        fontSize: n.fontSize,
        color: hex(n.color),
        bold: !!n.bold,
        align: n.align,
        valign: n.valign,
        fontFace: n.fontFamily ?? "Segoe UI",
        margin: 0,
        wrap: false,
      });
      break;
    }
    case "ellipse": {
      slide.addShape("ellipse", {
        x: dx + (n.cx - n.rx) * IN,
        y: dy + (n.cy - n.ry) * IN,
        w: Math.max(0.003, n.rx * 2 * IN),
        h: Math.max(0.003, n.ry * 2 * IN),
        // fill "none" = stroke-only ring (radar circle grid).
        fill: n.fill === "none" ? { type: "none" } : { color: hex(n.fill) },
        line: n.stroke && (n.strokeWidth ?? 0) > 0 ? { color: hex(n.stroke), width: n.strokeWidth } : { type: "none" },
      });
      break;
    }
    case "polygon": {
      // Real freeform geometry: filled, optionally translucent polygons
      // (radar series and grid webs) as native editable shapes.
      const xs = n.points.map((p) => p.x);
      const ys = n.points.map((p) => p.y);
      const x0 = Math.min(...xs);
      const y0 = Math.min(...ys);
      slide.addShape("custGeom", {
        x: dx + x0 * IN,
        y: dy + y0 * IN,
        w: Math.max(0.01, (Math.max(...xs) - x0) * IN),
        h: Math.max(0.01, (Math.max(...ys) - y0) * IN),
        points: [
          ...n.points.map((p, i) => ({ x: (p.x - x0) * IN, y: (p.y - y0) * IN, moveTo: i === 0 })),
          { close: true },
        ],
        fill: n.fill
          ? { color: hex(n.fill), transparency: Math.round((1 - (n.fillOpacity ?? 1)) * 100) }
          : { type: "none" },
        line: n.stroke ? { color: hex(n.stroke), width: n.strokeWidth ?? 1 } : { type: "none" },
      });
      break;
    }
    case "wedge": {
      const span = n.endAngle - n.startAngle;
      const x0 = n.cx - n.r;
      const y0 = n.cy - n.r;
      const box = {
        x: dx + x0 * IN,
        y: dy + y0 * IN,
        w: n.r * 2 * IN,
        h: n.r * 2 * IN,
        fill: { color: hex(n.fill) },
        line: n.stroke ? { color: hex(n.stroke), width: n.strokeWidth ?? 1 } : { type: "none" },
      };
      if (n.innerR > 0) {
        // Sunburst ring / gauge: a real filled annular sector via custGeom
        // (OOXML's "pie" shape can't express an inner radius) — see
        // annularSectorPoints (outer arc forward, then inner arc back).
        const rel = (p) => ({ x: (p.x - x0) * IN, y: (p.y - y0) * IN });
        const arc = annularSectorPoints(n.cx, n.cy, n.innerR, n.r, n.startAngle, n.endAngle);
        const half = arc.length / 2; // outer points carry moveTo; inner points don't.
        const pts = arc.map((p, i) => (i < half ? { ...rel(p), moveTo: i === 0 } : rel(p)));
        pts.push({ close: true });
        slide.addShape("custGeom", { ...box, points: pts });
      } else if (span >= 359.9) {
        slide.addShape("ellipse", box);
      } else {
        // Scene angles: 0 = 12 o'clock, clockwise; OOXML pie: 0 = 3 o'clock.
        slide.addShape("pie", { ...box, angleRange: [sceneToOoxmlPieAngle(n.startAngle), sceneToOoxmlPieAngle(n.endAngle)] });
      }
      break;
    }
    case "chevron": {
      slide.addShape(n.flatLeft ? "homePlate" : "chevron", {
        x: dx + n.x * IN,
        y: dy + n.y * IN,
        w: n.w * IN,
        h: n.h * IN,
        fill: { color: hex(n.fill) },
        line: { type: "none" },
      });
      break;
    }
    case "arrowhead": {
      // Rotated triangle whose tip is offset onto (n.x, n.y) about the box
      // centre — matching the SVG renderer, which anchors the tip. See arrowheadBox.
      const box = arrowheadBox(n.x, n.y, n.size, n.angle);
      slide.addShape("triangle", {
        x: dx + box.left * IN,
        y: dy + box.top * IN,
        w: box.size * IN,
        h: box.size * IN,
        fill: { color: hex(n.fill) },
        line: { type: "none" },
        rotate: Math.round(box.rotation),
      });
      break;
    }
  }
}

for (const cfg of configs) {
  const full = { ...DEFAULT_SIZE, ...cfg };
  const scene = buildChart(full);
  const slide = pres.addSlide();
  slide.background = { color: "FFFFFF" };
  const dx = (SLIDE.w - scene.width * IN) / 2;
  const dy = (SLIDE.h - scene.height * IN) / 2;
  for (const node of scene.nodes) addNode(slide, node, dx, dy);
}

await pres.writeFile({ fileName: output });
console.log(`${output}: ${configs.length} slide(s), native shapes`);
