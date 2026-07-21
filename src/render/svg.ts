import { polar, symbolPoints } from "../core/geometry";
import type { Scene, SceneNode } from "../core/scene";

const esc = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");

/**
 * Colours reach the scene graph verbatim from ChartConfig (series color, custom
 * palette, per-point colors) and are interpolated straight into SVG paint
 * attributes. A value like `#000"><image href=x onerror=alert(1)>` would break
 * out of the attribute and inject nodes that execute when this SVG is assigned
 * via innerHTML (the pane preview) or opened as a standalone document — stored
 * XSS reachable cross-user through a POWERCHART_CONFIG shape tag. Allow only the
 * shapes a paint can legitimately take; anything else falls back to a safe ink.
 * Every colour the engine actually produces (hex, and internally-built url(#…))
 * passes unchanged, so valid charts render byte-identically.
 */
const PAINT_OK = /^(#[0-9a-fA-F]{3,8}|rgba?\([\d.,\s%]+\)|hsla?\([\d.,\s%]+\)|url\(#[\w.-]+\)|[a-zA-Z]{1,24})$/;
const paint = (c: string | undefined): string => (c && PAINT_OK.test(c) ? c : "#000000");

/**
 * NUMERIC attributes are the second injection surface, and the one the colour
 * allow-list above does not cover. Sizes, widths and opacities reach the markup
 * by raw interpolation, and ChartConfig's numeric fields (`style.fontSize`,
 * `decorations.fillOpacity`, …) are only `number` in TypeScript — a type erased
 * at runtime. Untrusted JSON therefore carries a STRING straight into an
 * attribute: `fontSize: '10"><image href=x onerror=alert(1) /><text x="'` closes
 * the element and injects an executing node. That JSON is genuinely untrusted —
 * a `#c=` share link, an imported config, or a POWERCHART_CONFIG shape tag
 * authored in another deck — and the pane assigns this SVG via innerHTML.
 *
 * Coerce to a finite number or fall back. Every value the engine actually
 * produces is already numeric, so valid charts render byte-identically.
 */
const num = (v: unknown, fallback: number): number => {
  const n = typeof v === "number" ? v : Number(v);
  return Number.isFinite(n) ? n : fallback;
};

/** Hatch/dot pattern tiles: series-colored base with white strokes over it. */
const PATTERN_TILE: Record<string, (id: string, color: string) => string> = {
  diagonal: (id, c) =>
    `<pattern id="${id}" width="5" height="5" patternUnits="userSpaceOnUse" patternTransform="rotate(45)"><rect width="5" height="5" fill="${c}"/><line x1="0" y1="0" x2="0" y2="5" stroke="#ffffff" stroke-width="1.4" stroke-opacity="0.75"/></pattern>`,
  crosshatch: (id, c) =>
    `<pattern id="${id}" width="6" height="6" patternUnits="userSpaceOnUse" patternTransform="rotate(45)"><rect width="6" height="6" fill="${c}"/><path d="M 0 0 V 6 M 0 0 H 6" stroke="#ffffff" stroke-width="1.1" stroke-opacity="0.7"/></pattern>`,
  dots: (id, c) =>
    `<pattern id="${id}" width="6" height="6" patternUnits="userSpaceOnUse"><rect width="6" height="6" fill="${c}"/><circle cx="3" cy="3" r="1.1" fill="#ffffff" fill-opacity="0.8"/></pattern>`,
  horizontal: (id, c) =>
    `<pattern id="${id}" width="5" height="5" patternUnits="userSpaceOnUse"><rect width="5" height="5" fill="${c}"/><line x1="0" y1="1" x2="5" y2="1" stroke="#ffffff" stroke-width="1.4" stroke-opacity="0.75"/></pattern>`,
};

// A fragment id must be a safe token. A hex paint yields one after dropping the
// "#", but an rgb()/hsl() paint (both accepted by PAINT_OK) carries "(", ")" and
// "," — which break the <pattern id> AND make url(#…) unresolvable, since the URL
// parser stops at the first ")". Mapping those chars to "-" is a no-op for hex
// (byte-identical output) and produces a resolvable id for the rest.
const patternId = (pattern: string, color: string) => `p-${pattern}-${color.replace("#", "").replace(/[^\w-]/g, "-")}`;

/** Render a scene to a standalone SVG string (1pt = 1px). */
export function sceneToSvg(scene: Scene, opts: { background?: string } = {}): string {
  const parts: string[] = [];
  parts.push(
    `<svg xmlns="http://www.w3.org/2000/svg" width="${num(scene.width, 480)}" height="${num(scene.height, 300)}" viewBox="0 0 ${num(scene.width, 480)} ${num(scene.height, 300)}" font-family="Segoe UI, Arial, sans-serif" role="img">`,
  );
  // Accessible name + text alternative as the first children of the root, so a
  // screen reader announces the chart under role="img" instead of skipping it
  // (WCAG 1.1.1). First-child <title>/<desc> are the accessible name/description
  // per the SVG mapping — no id/aria wiring needed, so multiple inline charts on
  // one page can't collide on ids.
  //
  // An UNTITLED chart (the pane's own gallery shape) has only the generated
  // summary, and <desc> maps to the description alone: role="img" with no name
  // is exactly what axe-core's role-img-alt reports as a 1.1.1 failure. So the
  // summary becomes the name in that case — announced once, not twice.
  if (scene.title) {
    parts.push(`<title>${esc(scene.title)}</title>`);
    if (scene.desc) parts.push(`<desc>${esc(scene.desc)}</desc>`);
  } else if (scene.desc) {
    parts.push(`<title>${esc(scene.desc)}</title>`);
  }
  // One <pattern> def per (pattern, color) pair used by the scene's rects.
  const defs = new Map<string, string>();
  for (const n of scene.nodes) {
    if (n.kind === "rect" && n.pattern && PATTERN_TILE[n.pattern]) {
      const fill = paint(n.fill);
      const id = patternId(n.pattern, fill);
      if (!defs.has(id)) defs.set(id, PATTERN_TILE[n.pattern](id, fill));
    }
  }
  if (defs.size) parts.push(`<defs>${[...defs.values()].join("")}</defs>`);
  if (opts.background) {
    parts.push(`<rect width="100%" height="100%" fill="${paint(opts.background)}"/>`);
  }
  for (const n of scene.nodes) parts.push(nodeToSvg(n));
  parts.push("</svg>");
  return parts.join("\n");
}

function nodeToSvg(n: SceneNode): string {
  switch (n.kind) {
    case "rect": {
      const stroke = n.stroke ? ` stroke="${paint(n.stroke)}" stroke-width="${num(n.strokeWidth, 1)}"` : "";
      const fill =
        n.pattern && PATTERN_TILE[n.pattern] ? `url(#${patternId(n.pattern, paint(n.fill))})` : paint(n.fill);
      return `<rect x="${r(n.x)}" y="${r(n.y)}" width="${r(n.w)}" height="${r(n.h)}" fill="${fill}"${stroke}${name(n)}/>`;
    }
    case "line": {
      const dash =
        Array.isArray(n.dash) && n.dash.length ? ` stroke-dasharray="${n.dash.map((d) => num(d, 0)).join(" ")}"` : "";
      return `<line x1="${r(n.x1)}" y1="${r(n.y1)}" x2="${r(n.x2)}" y2="${r(n.y2)}" stroke="${paint(n.stroke)}" stroke-width="${num(n.strokeWidth, 1)}"${dash}${name(n)}/>`;
    }
    case "text": {
      const anchor = n.align === "left" ? "start" : n.align === "right" ? "end" : "middle";
      const x = n.align === "left" ? n.x : n.align === "right" ? n.x + n.w : n.x + n.w / 2;
      // Approximate vertical alignment with baseline offsets.
      const y =
        n.valign === "top"
          ? n.y + n.fontSize
          : n.valign === "bottom"
            ? n.y + n.h - n.fontSize * 0.25
            : n.y + n.h / 2 + n.fontSize * 0.36;
      const weight = n.bold ? ` font-weight="600"` : "";
      const family = n.fontFamily ? ` font-family="${esc(n.fontFamily)}"` : "";
      return `<text x="${r(x)}" y="${r(y)}" font-size="${num(n.fontSize, 12)}" fill="${paint(n.color)}" text-anchor="${anchor}"${weight}${family}${name(n)}>${esc(n.text)}</text>`;
    }
    case "ellipse": {
      const stroke = n.stroke ? ` stroke="${paint(n.stroke)}" stroke-width="${num(n.strokeWidth, 1)}"` : "";
      return `<ellipse cx="${r(n.cx)}" cy="${r(n.cy)}" rx="${r(n.rx)}" ry="${r(n.ry)}" fill="${paint(n.fill)}"${stroke}${name(n)}/>`;
    }
    case "chevron": {
      // Fixed notch fraction. The PowerPoint renderers instead name the native
      // chevron/homePlate preset, whose default point depth differs — an
      // intentional preview-vs-deck approximation (see the parity contract in
      // scene.ts).
      const notch = n.h * 0.28;
      const left = n.flatLeft
        ? `M ${r(n.x)} ${r(n.y)}`
        : `M ${r(n.x)} ${r(n.y)} L ${r(n.x + notch)} ${r(n.y + n.h / 2)} L ${r(n.x)} ${r(n.y + n.h)}`;
      const d = n.flatLeft
        ? `M ${r(n.x)} ${r(n.y)} L ${r(n.x + n.w - notch)} ${r(n.y)} L ${r(n.x + n.w)} ${r(n.y + n.h / 2)} L ${r(n.x + n.w - notch)} ${r(n.y + n.h)} L ${r(n.x)} ${r(n.y + n.h)} Z`
        : `${left} L ${r(n.x)} ${r(n.y + n.h)} L ${r(n.x + n.w - notch)} ${r(n.y + n.h)} L ${r(n.x + n.w)} ${r(n.y + n.h / 2)} L ${r(n.x + n.w - notch)} ${r(n.y)} Z`;
      return `<path d="${d}" fill="${paint(n.fill)}"${name(n)}/>`;
    }
    case "wedge": {
      const stroke = n.stroke ? ` stroke="${paint(n.stroke)}" stroke-width="${num(n.strokeWidth, 1)}"` : "";
      // A full 360° span has coincident start/end points, so the arc collapses
      // and SVG draws nothing. Emit an explicit circle (or ring) instead.
      if (n.endAngle - n.startAngle >= 359.9) {
        const ring = (rad: number, sweep: number) => {
          const a = polar(n.cx, n.cy, rad, 0);
          const b = polar(n.cx, n.cy, rad, 180);
          return `M ${r(a.x)} ${r(a.y)} A ${r(rad)} ${r(rad)} 0 1 ${sweep} ${r(b.x)} ${r(b.y)} A ${r(rad)} ${r(rad)} 0 1 ${sweep} ${r(a.x)} ${r(a.y)} Z`;
        };
        const d = n.innerR > 0 ? `${ring(n.r, 1)} ${ring(n.innerR, 0)}` : ring(n.r, 1);
        return `<path d="${d}" fill="${paint(n.fill)}" fill-rule="evenodd"${stroke}${name(n)}/>`;
      }
      const large = n.endAngle - n.startAngle > 180 ? 1 : 0;
      const o1 = polar(n.cx, n.cy, n.r, n.startAngle);
      const o2 = polar(n.cx, n.cy, n.r, n.endAngle);
      let d: string;
      if (n.innerR > 0) {
        const i1 = polar(n.cx, n.cy, n.innerR, n.endAngle);
        const i2 = polar(n.cx, n.cy, n.innerR, n.startAngle);
        d = `M ${r(o1.x)} ${r(o1.y)} A ${r(n.r)} ${r(n.r)} 0 ${large} 1 ${r(o2.x)} ${r(o2.y)} L ${r(i1.x)} ${r(i1.y)} A ${r(n.innerR)} ${r(n.innerR)} 0 ${large} 0 ${r(i2.x)} ${r(i2.y)} Z`;
      } else {
        d = `M ${r(n.cx)} ${r(n.cy)} L ${r(o1.x)} ${r(o1.y)} A ${r(n.r)} ${r(n.r)} 0 ${large} 1 ${r(o2.x)} ${r(o2.y)} Z`;
      }
      return `<path d="${d}" fill="${paint(n.fill)}"${stroke}${name(n)}/>`;
    }
    case "polygon": {
      const pts = n.points.map((p) => `${r(p.x)},${r(p.y)}`).join(" ");
      const fill = n.fill
        ? ` fill="${paint(n.fill)}"${n.fillOpacity != null ? ` fill-opacity="${num(n.fillOpacity, 1)}"` : ""}`
        : ` fill="none"`;
      const stroke = n.stroke
        ? ` stroke="${paint(n.stroke)}" stroke-width="${num(n.strokeWidth, 1)}" stroke-linejoin="round"`
        : "";
      return `<polygon points="${pts}"${fill}${stroke}${name(n)}/>`;
    }
    case "arrowhead": {
      // Narrow isosceles triangle with tip at (x, y), pointing along angle. The
      // PowerPoint renderers name the native `triangle` preset in a 2*size box,
      // which is broader — an intentional approximation (see scene.ts parity
      // contract); only the tip anchor and angle are shared exactly.
      const s = n.size;
      return `<path d="M 0 0 L ${-s * 1.8} ${-s * 0.7} L ${-s * 1.8} ${s * 0.7} Z" fill="${paint(n.fill)}" transform="translate(${r(n.x)} ${r(n.y)}) rotate(${r(n.angle)})"${name(n)}/>`;
    }
    case "symbol": {
      // The other two renderers name a preset; here we draw its outline.
      const pts = symbolPoints(n.shape, n.cx, n.cy, n.size)
        .map((p) => `${r(p.x)},${r(p.y)}`)
        .join(" ");
      const stroke = n.stroke
        ? ` stroke="${paint(n.stroke)}" stroke-width="${num(n.strokeWidth, 1)}" stroke-linejoin="round"`
        : "";
      return `<polygon points="${pts}" fill="${paint(n.fill)}"${stroke}${name(n)}/>`;
    }
  }
}

const r = (v: number) => Math.round(v * 100) / 100;
const name = (n: SceneNode) => (n.name ? ` data-name="${esc(n.name)}"` : "");
