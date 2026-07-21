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
const { buildChart, buildAgendaScene, DEFAULT_SIZE, arrowheadBox, annularSectorPoints, SYMBOL_PRESET, dashKind } =
  engine;

// A stale packaged lib (the skill ships no build step) can be missing an export,
// which otherwise blows up mid-render on the first chart that needs it. Fail
// fast with an actionable message instead.
for (const [name, fn] of Object.entries({ buildChart, buildAgendaScene, arrowheadBox, annularSectorPoints, dashKind })) {
  if (typeof fn !== "function") {
    console.error(`powerchart engine is missing export "${name}" — rebuild the lib (npm run build:lib)`);
    process.exit(1);
  }
}

const [, , input, output = "powerchart.pptx"] = process.argv;
if (!input) {
  console.error("usage: node scripts/render-pptx.mjs <charts.json> [out.pptx]");
  process.exit(1);
}

let raw;
try {
  raw = JSON.parse(readFileSync(input, "utf8"));
} catch (err) {
  console.error(`couldn't read/parse ${input}: ${err?.message ?? err}`);
  process.exit(1);
}
// Accept one config or an array of them. A wrapper object like {"charts":[…]} or
// a primitive is the common LLM mistake — name it rather than render one blank slide.
if (raw == null || typeof raw !== "object") {
  console.error("expected a ChartConfig object or an array of them (got a " + typeof raw + ")");
  process.exit(1);
}
if (!Array.isArray(raw) && !("kind" in raw) && !("data" in raw) && !("chapters" in raw)) {
  const wrapped = Object.keys(raw).find((k) => Array.isArray(raw[k]));
  console.error(
    wrapped
      ? `expected a ChartConfig or an array — did you mean the "${wrapped}" array inside this object?`
      : "expected a ChartConfig (with a kind/data) or an array of them",
  );
  process.exit(1);
}
const configs = Array.isArray(raw) ? raw : [raw];

const IN = 1 / 72; // points → inches
const SLIDE = { w: 13.333, h: 7.5 };

const pres = new pptxgen();
pres.defineLayout({ name: "WIDE", width: SLIDE.w, height: SLIDE.h });
pres.layout = "WIDE";

const to2 = (n) => Math.max(0, Math.min(255, Math.round(n))).toString(16).padStart(2, "0");
const hslToHex = (h, s, l) => {
  h = ((h % 360) + 360) % 360;
  s = Math.max(0, Math.min(1, s / 100));
  l = Math.max(0, Math.min(1, l / 100));
  const chroma = (1 - Math.abs(2 * l - 1)) * s;
  const x = chroma * (1 - Math.abs(((h / 60) % 2) - 1));
  const m = l - chroma / 2;
  const [r, g, b] = [
    [chroma, x, 0],
    [x, chroma, 0],
    [0, chroma, x],
    [0, x, chroma],
    [x, 0, chroma],
    [chroma, 0, x],
  ][Math.floor(h / 60) % 6];
  return to2((r + m) * 255) + to2((g + m) * 255) + to2((b + m) * 255);
};

// The CSS Color 4 named colours as "name rrggbb" pairs. SVG renders a bare name
// natively and Office.js passes it through to the host, but pptxgenjs takes hex
// only — without this table every named colour collapsed to one grey, so two
// series coloured "red" and "steelblue" became indistinguishable in the deck.
const CSS_NAMES = Object.fromEntries(
  `aliceblue f0f8ff, antiquewhite faebd7, aqua 00ffff, aquamarine 7fffd4, azure f0ffff, beige f5f5dc,
  bisque ffe4c4, black 000000, blanchedalmond ffebcd, blue 0000ff, blueviolet 8a2be2, brown a52a2a,
  burlywood deb887, cadetblue 5f9ea0, chartreuse 7fff00, chocolate d2691e, coral ff7f50, cornflowerblue 6495ed,
  cornsilk fff8dc, crimson dc143c, cyan 00ffff, darkblue 00008b, darkcyan 008b8b, darkgoldenrod b8860b,
  darkgray a9a9a9, darkgreen 006400, darkgrey a9a9a9, darkkhaki bdb76b, darkmagenta 8b008b,
  darkolivegreen 556b2f, darkorange ff8c00, darkorchid 9932cc, darkred 8b0000, darksalmon e9967a,
  darkseagreen 8fbc8f, darkslateblue 483d8b, darkslategray 2f4f4f, darkslategrey 2f4f4f, darkturquoise 00ced1,
  darkviolet 9400d3, deeppink ff1493, deepskyblue 00bfff, dimgray 696969, dimgrey 696969, dodgerblue 1e90ff,
  firebrick b22222, floralwhite fffaf0, forestgreen 228b22, fuchsia ff00ff, gainsboro dcdcdc,
  ghostwhite f8f8ff, gold ffd700, goldenrod daa520, gray 808080, green 008000, greenyellow adff2f, grey 808080,
  honeydew f0fff0, hotpink ff69b4, indianred cd5c5c, indigo 4b0082, ivory fffff0, khaki f0e68c,
  lavender e6e6fa, lavenderblush fff0f5, lawngreen 7cfc00, lemonchiffon fffacd, lightblue add8e6,
  lightcoral f08080, lightcyan e0ffff, lightgoldenrodyellow fafad2, lightgray d3d3d3, lightgreen 90ee90,
  lightgrey d3d3d3, lightpink ffb6c1, lightsalmon ffa07a, lightseagreen 20b2aa, lightskyblue 87cefa,
  lightslategray 778899, lightslategrey 778899, lightsteelblue b0c4de, lightyellow ffffe0, lime 00ff00,
  limegreen 32cd32, linen faf0e6, magenta ff00ff, maroon 800000, mediumaquamarine 66cdaa, mediumblue 0000cd,
  mediumorchid ba55d3, mediumpurple 9370db, mediumseagreen 3cb371, mediumslateblue 7b68ee,
  mediumspringgreen 00fa9a, mediumturquoise 48d1cc, mediumvioletred c71585, midnightblue 191970,
  mintcream f5fffa, mistyrose ffe4e1, moccasin ffe4b5, navajowhite ffdead, navy 000080, oldlace fdf5e6,
  olive 808000, olivedrab 6b8e23, orange ffa500, orangered ff4500, orchid da70d6, palegoldenrod eee8aa,
  palegreen 98fb98, paleturquoise afeeee, palevioletred db7093, papayawhip ffefd5, peachpuff ffdab9,
  peru cd853f, pink ffc0cb, plum dda0dd, powderblue b0e0e6, purple 800080, rebeccapurple 663399, red ff0000,
  rosybrown bc8f8f, royalblue 4169e1, saddlebrown 8b4513, salmon fa8072, sandybrown f4a460, seagreen 2e8b57,
  seashell fff5ee, sienna a0522d, silver c0c0c0, skyblue 87ceeb, slateblue 6a5acd, slategray 708090,
  slategrey 708090, snow fffafa, springgreen 00ff7f, steelblue 4682b4, tan d2b48c, teal 008080, thistle d8bfd8,
  tomato ff6347, turquoise 40e0d0, violet ee82ee, wheat f5deb3, white ffffff, whitesmoke f5f5f5, yellow ffff00,
  yellowgreen 9acd32`
    .split(",")
    .map((pair) => pair.trim().split(" ")),
);

// Normalise ANY paint the SVG allow-list accepts to a validated 6-digit hex, so
// the headless pptx matches the preview instead of falling back to black for
// rgb()/hsl()/named colours. SECURITY: the colour is interpolated into OOXML
// unescaped, so this MUST always return exactly six hex digits — a value like
// `000"/><a:x` could otherwise inject markup. rgb()/hsl() are parsed, named
// colours resolve through CSS_NAMES; anything unrecognised is black.
// fillOf() folds any alpha into transparency.
const hex = (c) => {
  const raw = (c ?? "").trim();
  const h = raw.replace("#", "");
  if (/^[0-9a-fA-F]{3,4}$/.test(h))
    return h
      .slice(0, 3)
      .replace(/./g, "$&$&"); // #abc / #abcd → aabbcc
  if (/^[0-9a-fA-F]{6}([0-9a-fA-F]{2})?$/.test(h)) return h.slice(0, 6);
  const nums = (str) => (str.match(/-?[\d.]+%?/g) ?? []).map((v) => parseFloat(v));
  let m;
  if ((m = /^rgba?\(([^)]*)\)$/i.exec(raw))) {
    // Only the r/g/b components decide the 0–255 vs 0–100% scale: testing the
    // whole argument list let a percentage ALPHA (the legal `rgba(r,g,b,50%)`)
    // multiply the channels by 2.55 and clip every colour to white.
    const rgb = (m[1].match(/-?[\d.]+%?/g) ?? []).slice(0, 3);
    const sc = rgb.some((v) => v.endsWith("%")) ? 2.55 : 1;
    const [r = 0, g = 0, b = 0] = rgb.map((v) => parseFloat(v));
    return to2(r * sc) + to2(g * sc) + to2(b * sc);
  }
  if ((m = /^hsla?\(([^)]*)\)$/i.exec(raw))) {
    const [hh = 0, ss = 0, ll = 0] = nums(m[1]);
    return hslToHex(hh, ss, ll);
  }
  return CSS_NAMES[raw.toLowerCase()] ?? "000000";
};

// Opacity 0..1 carried by a paint (8-digit #RRGGBBAA, rgba(), hsla(), or the
// `transparent` keyword); 1 when opaque. The SVG renderer honours the alpha
// natively; here it becomes OOXML transparency — or no paint at all.
const alphaOf = (c) => {
  const raw = (c ?? "").trim();
  if (/^transparent$/i.test(raw)) return 0;
  const h = raw.replace("#", "");
  if (/^[0-9a-fA-F]{8}$/.test(h)) return parseInt(h.slice(6), 16) / 255;
  if (/^[0-9a-fA-F]{4}$/.test(h)) return parseInt(h[3] + h[3], 16) / 255;
  const m = /^(?:rgba|hsla)\(([^)]*)\)$/i.exec(raw);
  if (m) {
    const parts = m[1].split(/[,/]/).map((s) => s.trim());
    if (parts.length >= 4) {
      const a = parts[3].endsWith("%") ? parseFloat(parts[3]) / 100 : parseFloat(parts[3]);
      return Number.isFinite(a) ? Math.max(0, Math.min(1, a)) : 1;
    }
  }
  return 1;
};

// A pptxgenjs solid fill folding an 8-digit-hex alpha and any scene fillOpacity
// into OOXML transparency (0 = opaque, 100 = clear). A zero transparency is
// dropped by pptxgenjs, so an opaque fill is byte-identical to the bare {color}.
const fillOf = (color, fillOpacity = 1) => {
  const t = Math.round((1 - alphaOf(color) * fillOpacity) * 100);
  // Fully clear: emit no fill rather than a colour. `color: "transparent"` is the
  // documented floating-segment idiom, and it has no hex — painting it as one
  // put a solid block where the preview shows nothing.
  if (t >= 100) return { type: "none" };
  return t > 0 ? { color: hex(color), transparency: t } : { color: hex(color) };
};

/** A paint that would actually draw: present, and not fully transparent. */
const visible = (paint) => !!paint && alphaOf(paint) > 0;

/** Map one scene node to PptxgenJS calls at slide offset (inches). */
function addNode(slide, n, dx, dy) {
  switch (n.kind) {
    case "rect": {
      slide.addShape("rect", {
        x: dx + n.x * IN,
        y: dy + n.y * IN,
        w: Math.max(0.003, n.w * IN),
        h: Math.max(0.003, n.h * IN),
        // A "none" fill is an outlined/hollow rect (IBCS plan/budget columns).
        fill: n.fill === "none" ? { type: "none" } : fillOf(n.fill),
        line:
          visible(n.stroke) && (n.strokeWidth ?? 0) > 0
            ? { color: hex(n.stroke), width: n.strokeWidth }
            : { type: "none" },
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
          ...(n.dash ? { dashType: dashKind(n.dash) === "dot" ? "sysDot" : "dash" } : {}),
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
        fill: n.fill === "none" ? { type: "none" } : fillOf(n.fill),
        line:
          visible(n.stroke) && (n.strokeWidth ?? 0) > 0
            ? { color: hex(n.stroke), width: n.strokeWidth }
            : { type: "none" },
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
          ? fillOf(n.fill, n.fillOpacity)
          : { type: "none" },
        line: visible(n.stroke) ? { color: hex(n.stroke), width: n.strokeWidth ?? 1 } : { type: "none" },
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
        fill: fillOf(n.fill),
        line: visible(n.stroke) ? { color: hex(n.stroke), width: n.strokeWidth ?? 1 } : { type: "none" },
      };
      if (span >= 359.9 && n.innerR <= 0) {
        slide.addShape("ellipse", box);
      } else {
        // Filled sector via custGeom for BOTH the doughnut/gauge ring (innerR>0)
        // and the solid pie wedge (innerR=0, which annularSectorPoints degenerates
        // to a fan from the centre — outer arc forward, then the centre point back).
        // The OOXML "pie" preset can't be used for the solid case: it takes two
        // independently-normalized angles and draws swAng = end − start, so any
        // slice crossing 3 o'clock (every pie has exactly one) gets a negative
        // sweep and renders the wrong wedge. custGeom samples polar() directly, so
        // it's correct across the boundary — the same reason the ring uses it.
        const rel = (p) => ({ x: (p.x - x0) * IN, y: (p.y - y0) * IN });
        const arc = annularSectorPoints(n.cx, n.cy, Math.max(0, n.innerR), n.r, n.startAngle, n.endAngle);
        const half = arc.length / 2; // outer points carry moveTo; inner points don't.
        const pts = arc.map((p, i) => (i < half ? { ...rel(p), moveTo: i === 0 } : rel(p)));
        pts.push({ close: true });
        slide.addShape("custGeom", { ...box, points: pts });
      }
      break;
    }
    case "chevron": {
      slide.addShape(n.flatLeft ? "homePlate" : "chevron", {
        x: dx + n.x * IN,
        y: dy + n.y * IN,
        w: n.w * IN,
        h: n.h * IN,
        fill: fillOf(n.fill),
        line: { type: "none" },
      });
      break;
    }
    case "symbol": {
      // Native preset geometry, so the marker stays FILLED — a custGeom polygon
      // would render here but not in the live add-in. SYMBOL_PRESET names are
      // OOXML preset names, which is exactly what addShape takes.
      slide.addShape(SYMBOL_PRESET[n.shape], {
        x: dx + (n.cx - n.size) * IN,
        y: dy + (n.cy - n.size) * IN,
        w: n.size * 2 * IN,
        h: n.size * 2 * IN,
        fill: fillOf(n.fill),
        line:
          visible(n.stroke) && (n.strokeWidth ?? 0) > 0
            ? { color: hex(n.stroke), width: n.strokeWidth ?? 1 }
            : { type: "none" },
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
        fill: fillOf(n.fill),
        line: { type: "none" },
        rotate: Math.round(box.rotation),
      });
      break;
    }
  }
}

/** Build the scene for one config: an agenda slide, or a chart. */
function sceneFor(cfg) {
  if (cfg && cfg.kind === "agenda") {
    // { kind:"agenda", chapters:[…], highlight?, title? } → a chapter-list slide
    // (highlight the current chapter, or -1 for an overview). Full-slide size.
    return buildAgendaScene(Array.isArray(cfg.chapters) ? cfg.chapters : [], {
      highlight: cfg.highlight,
      title: cfg.title,
      width: cfg.width,
      height: cfg.height,
    });
  }
  return buildChart({ ...DEFAULT_SIZE, ...cfg });
}

/** A visible error slide so a bad config in a batch surfaces instead of vanishing. */
function errorSlide(i, err) {
  const slide = pres.addSlide();
  slide.background = { color: "FFFFFF" };
  slide.addText(`Chart ${i + 1} failed: ${err?.message ?? err}`, {
    x: 0.4,
    y: 0.4,
    w: SLIDE.w - 0.8,
    h: 0.6,
    fontSize: 18,
    bold: true,
    color: "C0392B",
    fill: { color: "FBEAE8" },
    align: "left",
    valign: "middle",
  });
}

let failed = 0;
configs.forEach((cfg, i) => {
  // Isolate each config: one bad chart in a 50-slide batch must not throw away
  // the other 49 (the Office.js path isolates per item too). Stamp an error
  // slide and carry on, so partial output always survives.
  try {
    const scene = sceneFor(cfg);
    const slide = pres.addSlide();
    // The chart's own canvas colour, not a fixed white: a dark-styled config
    // paints its ink for `style.background`, so a white slide under it put white
    // labels on white. Default (no style.background) stays FFFFFF.
    slide.background = { color: hex(cfg?.style?.background ?? "#ffffff") };
    const dx = (SLIDE.w - scene.width * IN) / 2;
    const dy = (SLIDE.h - scene.height * IN) / 2;
    for (const node of scene.nodes) addNode(slide, node, dx, dy);
  } catch (err) {
    failed++;
    errorSlide(i, err);
    console.error(`chart ${i + 1}: ${err?.message ?? err}`);
  }
});

await pres.writeFile({ fileName: output });
console.log(
  `${output}: ${configs.length} slide(s), native shapes` + (failed ? ` (${failed} failed — see error slides)` : ""),
);
if (failed) process.exit(1);
