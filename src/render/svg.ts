import { polar, type Scene, type SceneNode } from "../core/scene";

const esc = (s: string) =>
  s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");

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

const patternId = (pattern: string, color: string) => `p-${pattern}-${color.replace("#", "")}`;

/** Render a scene to a standalone SVG string (1pt = 1px). */
export function sceneToSvg(scene: Scene, opts: { background?: string } = {}): string {
  const parts: string[] = [];
  parts.push(
    `<svg xmlns="http://www.w3.org/2000/svg" width="${scene.width}" height="${scene.height}" viewBox="0 0 ${scene.width} ${scene.height}" font-family="Segoe UI, Arial, sans-serif">`,
  );
  // One <pattern> def per (pattern, color) pair used by the scene's rects.
  const defs = new Map<string, string>();
  for (const n of scene.nodes) {
    if (n.kind === "rect" && n.pattern && PATTERN_TILE[n.pattern]) {
      const id = patternId(n.pattern, n.fill);
      if (!defs.has(id)) defs.set(id, PATTERN_TILE[n.pattern](id, n.fill));
    }
  }
  if (defs.size) parts.push(`<defs>${[...defs.values()].join("")}</defs>`);
  if (opts.background) {
    parts.push(`<rect width="100%" height="100%" fill="${opts.background}"/>`);
  }
  for (const n of scene.nodes) parts.push(nodeToSvg(n));
  parts.push("</svg>");
  return parts.join("\n");
}

function nodeToSvg(n: SceneNode): string {
  switch (n.kind) {
    case "rect": {
      const stroke = n.stroke ? ` stroke="${n.stroke}" stroke-width="${n.strokeWidth ?? 1}"` : "";
      const fill = n.pattern && PATTERN_TILE[n.pattern] ? `url(#${patternId(n.pattern, n.fill)})` : n.fill;
      return `<rect x="${r(n.x)}" y="${r(n.y)}" width="${r(n.w)}" height="${r(n.h)}" fill="${fill}"${stroke}${name(n)}/>`;
    }
    case "line": {
      const dash = n.dash ? ` stroke-dasharray="${n.dash.join(" ")}"` : "";
      return `<line x1="${r(n.x1)}" y1="${r(n.y1)}" x2="${r(n.x2)}" y2="${r(n.y2)}" stroke="${n.stroke}" stroke-width="${n.strokeWidth ?? 1}"${dash}${name(n)}/>`;
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
      return `<text x="${r(x)}" y="${r(y)}" font-size="${n.fontSize}" fill="${n.color}" text-anchor="${anchor}"${weight}${family}${name(n)}>${esc(n.text)}</text>`;
    }
    case "ellipse": {
      const stroke = n.stroke ? ` stroke="${n.stroke}" stroke-width="${n.strokeWidth ?? 1}"` : "";
      return `<ellipse cx="${r(n.cx)}" cy="${r(n.cy)}" rx="${r(n.rx)}" ry="${r(n.ry)}" fill="${n.fill}"${stroke}${name(n)}/>`;
    }
    case "chevron": {
      const notch = n.h * 0.28;
      const left = n.flatLeft ? `M ${r(n.x)} ${r(n.y)}` : `M ${r(n.x)} ${r(n.y)} L ${r(n.x + notch)} ${r(n.y + n.h / 2)} L ${r(n.x)} ${r(n.y + n.h)}`;
      const d = n.flatLeft
        ? `M ${r(n.x)} ${r(n.y)} L ${r(n.x + n.w - notch)} ${r(n.y)} L ${r(n.x + n.w)} ${r(n.y + n.h / 2)} L ${r(n.x + n.w - notch)} ${r(n.y + n.h)} L ${r(n.x)} ${r(n.y + n.h)} Z`
        : `${left} L ${r(n.x)} ${r(n.y + n.h)} L ${r(n.x + n.w - notch)} ${r(n.y + n.h)} L ${r(n.x + n.w)} ${r(n.y + n.h / 2)} L ${r(n.x + n.w - notch)} ${r(n.y)} Z`;
      return `<path d="${d}" fill="${n.fill}"${name(n)}/>`;
    }
    case "wedge": {
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
      const stroke = n.stroke ? ` stroke="${n.stroke}" stroke-width="${n.strokeWidth ?? 1}"` : "";
      return `<path d="${d}" fill="${n.fill}"${stroke}${name(n)}/>`;
    }
    case "arrowhead": {
      // Triangle with tip at (x, y), pointing along angle.
      const s = n.size;
      return `<path d="M 0 0 L ${-s * 1.8} ${-s * 0.7} L ${-s * 1.8} ${s * 0.7} Z" fill="${n.fill}" transform="translate(${r(n.x)} ${r(n.y)}) rotate(${r(n.angle)})"${name(n)}/>`;
    }
  }
}

const r = (v: number) => Math.round(v * 100) / 100;
const name = (n: SceneNode) => (n.name ? ` data-name="${esc(n.name)}"` : "");
