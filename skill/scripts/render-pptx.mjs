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
import pptxgen from "pptxgenjs";

// Engine location: packaged skill layout first, then repo layout.
let engine;
for (const candidate of ["../lib/powerchart.js", "../../dist-lib/powerchart.js"]) {
  try {
    engine = await import(new URL(candidate, import.meta.url).href);
    break;
  } catch {
    /* try next location */
  }
}
if (!engine) {
  console.error("powerchart engine not found — run `npm run build:lib` first");
  process.exit(1);
}
const { buildChart, DEFAULT_SIZE } = engine;

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
        fill: { color: hex(n.fill) },
        line: n.stroke && (n.strokeWidth ?? 0) > 0 ? { color: hex(n.stroke), width: n.strokeWidth } : { type: "none" },
      });
      break;
    }
    case "wedge": {
      const span = n.endAngle - n.startAngle;
      const box = {
        x: dx + (n.cx - n.r) * IN,
        y: dy + (n.cy - n.r) * IN,
        w: n.r * 2 * IN,
        h: n.r * 2 * IN,
        fill: { color: hex(n.fill) },
        line: n.stroke ? { color: hex(n.stroke), width: n.strokeWidth ?? 1 } : { type: "none" },
      };
      if (span >= 359.9) {
        slide.addShape("ellipse", box);
      } else {
        // Scene angles: 0 = 12 o'clock, clockwise; OOXML pie: 0 = 3 o'clock.
        const a1 = (((n.startAngle - 90) % 360) + 360) % 360;
        const a2 = (((n.endAngle - 90) % 360) + 360) % 360;
        slide.addShape("pie", { ...box, angleRange: [a1, a2] });
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
      const s = n.size * 2;
      slide.addShape("triangle", {
        x: dx + (n.x - s / 2) * IN,
        y: dy + (n.y - s / 2) * IN,
        w: s * IN,
        h: s * IN,
        fill: { color: hex(n.fill) },
        line: { type: "none" },
        rotate: Math.round((((n.angle + 90) % 360) + 360) % 360),
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
