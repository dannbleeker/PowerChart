import type { ChartConfig, ChartStyle, Decorations } from "../types";
import { polar, textWidth, type SceneNode } from "../scene";
import { formatNumber, resolveFormat } from "../format";
import { lerpColor } from "../color";
import { footnoteH, titleHeight, titleNode } from "./frame";
import { PALETTE } from "../style";
import type { LayoutResult } from "./column";

/**
 * Sunburst: a two-ring hierarchical pie. Categories named "Group | Item" put
 * the groups on the inner ring (angular span ∝ group total) and their items on
 * the outer ring (sub-spans within the parent). Without "|" it degrades to a
 * single-ring doughnut. Wedges render everywhere (triangle fans in the add-in).
 */
export function layoutSunburst(cfg: ChartConfig, style: ChartStyle, decor: Decorations): LayoutResult {
  const { data } = cfg;
  const fs = style.fontSize;
  const palette = cfg.style?.palette ?? PALETTE;
  const items = data.categories
    .map((c, i) => ({ label: c, value: Math.max(0, data.series[0]?.values[i] ?? 0), i }))
    .filter((r) => r.value > 0);
  const total = items.reduce((a, r) => a + r.value, 0) || 1;
  const fmt = resolveFormat(items.map((r) => r.value), cfg.numberFormat);
  const grouped = items.some((r) => r.label.includes("|"));
  const groupOf = (label: string) => (label.includes("|") ? label.split("|")[0].trim() : "");
  const labelOf = (label: string) => (label.includes("|") ? label.split("|").slice(1).join("|").trim() : label);

  const titleH = titleHeight(cfg, style);
  const footH = footnoteH(cfg, style, decor);
  const cx = cfg.width / 2;
  const cy = titleH + (cfg.height - titleH - footH) / 2;
  const r = Math.max(20, Math.min(cfg.width / 2 - fs * 4, (cfg.height - titleH - footH) / 2 - fs * 0.5));
  const rInner = r * 0.32;
  const rMid = grouped ? r * 0.6 : rInner;

  const nodes: SceneNode[] = [];
  const titleN = titleNode(cfg, style);
  if (titleN) nodes.push(titleN);

  const norm = (a: number) => ((a % 360) + 360) % 360;
  const label = (rr: number, midAngle: number, text: string, color: string, name: string, outside: boolean) => {
    if (!decor.segmentLabels) return;
    const p = polar(cx, cy, rr, midAngle);
    const w = textWidth(text, fs * 0.85) + 4;
    const rightHalf = norm(midAngle) < 180;
    nodes.push({
      kind: "text",
      x: outside ? (rightHalf ? p.x : p.x - w) : p.x - w / 2,
      y: p.y - fs * 0.7,
      w,
      h: fs * 1.4,
      text,
      fontSize: fs * 0.85,
      color,
      align: outside ? (rightHalf ? "left" : "right") : "center",
      valign: "middle",
      name,
    });
  };

  if (grouped) {
    const groups: { name: string; total: number; members: typeof items }[] = [];
    for (const it of items) {
      const g = groupOf(it.label);
      let e = groups.find((x) => x.name === g);
      if (!e) {
        e = { name: g, total: 0, members: [] };
        groups.push(e);
      }
      e.total += it.value;
      e.members.push(it);
    }
    let angle = 0;
    groups.forEach((g, gi) => {
      const span = (g.total / total) * 360;
      const gColor = palette[gi % palette.length];
      const a0 = norm(angle);
      nodes.push({ kind: "wedge", cx, cy, r: rMid, innerR: rInner, startAngle: a0, endAngle: a0 + span, fill: gColor, stroke: style.background, strokeWidth: 1, name: `group-${gi}` });
      if (span >= 16) label((rInner + rMid) / 2, angle + span / 2, g.name, "#ffffff", `group-label-${gi}`, false);
      let a2 = angle;
      g.members.forEach((m, mi) => {
        const mspan = (m.value / g.total) * span;
        const mColor = lerpColor(gColor, "#ffffff", 0.16 + 0.12 * (mi % 4));
        const ma0 = norm(a2);
        nodes.push({ kind: "wedge", cx, cy, r, innerR: rMid, startAngle: ma0, endAngle: ma0 + mspan, fill: mColor, stroke: style.background, strokeWidth: 1, name: `slice-${m.i}` });
        if (mspan >= 12) label(r + fs * 0.7, a2 + mspan / 2, labelOf(m.label), style.text, `label-${m.i}`, true);
        a2 += mspan;
      });
      angle += span;
    });
  } else {
    let angle = 0;
    items.forEach((m) => {
      const span = (m.value / total) * 360;
      const a0 = norm(angle);
      nodes.push({ kind: "wedge", cx, cy, r, innerR: rInner, startAngle: a0, endAngle: a0 + span, fill: palette[m.i % palette.length], stroke: style.background, strokeWidth: 1, name: `slice-${m.i}` });
      if (span >= 12) label(r + fs * 0.7, angle + span / 2, `${m.label} ${formatNumber(m.value, fmt)}`, style.text, `label-${m.i}`, true);
      angle += span;
    });
  }

  return {
    nodes,
    anchors: {
      categoryX: items.map(() => cx),
      categoryWidth: items.map(() => r),
      columnTop: items.map(() => cy - r),
      columnValue: data.categories.map((_, c) => data.series[0]?.values[c] ?? 0),
      baselineY: cy,
      plot: { x: cx - r, y: cy - r, w: r * 2, h: r * 2 },
    },
  };
}
