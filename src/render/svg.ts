import type { Scene, SceneNode } from "../core/scene";

const esc = (s: string) =>
  s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");

/** Render a scene to a standalone SVG string (1pt = 1px). */
export function sceneToSvg(scene: Scene, opts: { background?: string } = {}): string {
  const parts: string[] = [];
  parts.push(
    `<svg xmlns="http://www.w3.org/2000/svg" width="${scene.width}" height="${scene.height}" viewBox="0 0 ${scene.width} ${scene.height}" font-family="Segoe UI, Arial, sans-serif">`,
  );
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
      return `<rect x="${r(n.x)}" y="${r(n.y)}" width="${r(n.w)}" height="${r(n.h)}" fill="${n.fill}"${stroke}${name(n)}/>`;
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
    case "arrowhead": {
      // Triangle with tip at (x, y), pointing along angle.
      const s = n.size;
      return `<path d="M 0 0 L ${-s * 1.8} ${-s * 0.7} L ${-s * 1.8} ${s * 0.7} Z" fill="${n.fill}" transform="translate(${r(n.x)} ${r(n.y)}) rotate(${r(n.angle)})"${name(n)}/>`;
    }
  }
}

const r = (v: number) => Math.round(v * 100) / 100;
const name = (n: SceneNode) => (n.name ? ` data-name="${esc(n.name)}"` : "");
