import type { ChartConfig, ChartStyle, Decorations } from "../types";
import { contrastInk, polar, textWidth, type SceneNode } from "../scene";
import { clipToWidth } from "../elements";
import { formatNumber, resolveFormat } from "../format";
import { lerpColor } from "../color";
import { fitPlot, footnoteH, titleHeight, titleNode } from "./frame";
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
  const fmt = resolveFormat(
    items.map((r) => r.value),
    cfg.numberFormat,
  );
  const grouped = items.some((r) => r.label.includes("|"));
  const groupOf = (label: string) => (label.includes("|") ? label.split("|")[0].trim() : "");
  const labelOf = (label: string) => (label.includes("|") ? label.split("|").slice(1).join("|").trim() : label);

  const titleH = titleHeight(cfg, style);
  const footH = footnoteH(cfg, style, decor);
  // Fitted so the ring's CENTRE and RADIUS both come off a positive box: on a
  // frame too short for its title and footnote the raw height goes negative,
  // which puts the centre below the bottom of the chart.
  const box = fitPlot(cfg, { x: 0, y: titleH, w: cfg.width, h: cfg.height - titleH - footH });
  const cx = box.x + box.w / 2;
  const cy = box.y + box.h / 2;
  // Reserve what the OUTER labels actually occupy, not a fixed guess. They sit
  // at radius r + fs*0.7 in a box fs*1.4 tall and as wide as their text, so the
  // old fs*0.5 vertical / fs*4 horizontal margins were always short: the stock
  // sample put its bottom label at y 295-309 on a 300pt canvas, and a long side
  // label ate the entire left margin.
  const outerLabels = decor.segmentLabels ? items.map((r0) => labelOf(r0.label)) : [];
  const labelW = outerLabels.reduce((m, t) => Math.max(m, textWidth(t, fs * 0.85) + 4), 0);
  // Only the label's own half-height clears the canvas edge; the rest of the box
  // is centred on the label anchor.
  const marginY = outerLabels.length ? fs * 1.4 : fs * 0.5;
  const marginX = outerLabels.length ? labelW + fs * 0.9 : fs * 4;
  const rWant = Math.min(box.w / 2 - marginX, box.h / 2 - marginY);
  // The 20pt floor keeps a ring visible on an ordinary chart, and like the
  // tilemap's tile floor and the radar's web it yields to the frame: honouring
  // it where `rWant` is smaller spends the margins reserved just above, and the
  // outer labels then run past the edge — 13pt below a 300x60 chart. A small
  // ring is still a ring; a label drawn off the chart is not there at all.
  const r = Math.max(1, rWant);
  // Those margins are the whole of what an OUTSIDE label has to sit in. When the
  // floor has to override the fit there is no margin left, so the outer ring of
  // labels is dropped rather than drawn off the frame. Inside labels are bounded
  // by the wedge they sit on and are unaffected.
  const outerLabelsFit = rWant >= 1;
  const rInner = r * 0.32;
  const rMid = grouped ? r * 0.6 : rInner;

  const nodes: SceneNode[] = [];
  const titleN = titleNode(cfg, style);
  if (titleN) nodes.push(titleN);

  const norm = (a: number) => ((a % 360) + 360) % 360;
  const label = (rr: number, midAngle: number, text: string, color: string, name: string, outside: boolean) => {
    if (!decor.segmentLabels) return;
    if (outside && !outerLabelsFit) return;
    const p = polar(cx, cy, rr, midAngle);
    const rightHalf = norm(midAngle) < 180;
    // An OUTSIDE label runs away from the ring edge with nothing to stop it —
    // the same defect the pie's slice labels had, in the same shape. Clipped to
    // the room actually there between the ring and the frame; an inside label
    // is bounded by the wedge it sits on and is untouched.
    const lf = fs * 0.85;
    const room = Math.max(lf, (rightHalf ? cfg.width - p.x : p.x) - 2);
    const shown = outside ? clipToWidth(text, lf, room) : text;
    const w = textWidth(shown, lf) + 4;
    nodes.push({
      kind: "text",
      x: outside ? (rightHalf ? p.x : p.x - w) : p.x - w / 2,
      y: p.y - fs * 0.7,
      w,
      h: fs * 1.4,
      text: shown,
      fontSize: lf,
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
      nodes.push({
        kind: "wedge",
        cx,
        cy,
        r: rMid,
        innerR: rInner,
        startAngle: a0,
        endAngle: a0 + span,
        fill: gColor,
        stroke: style.background,
        strokeWidth: 1,
        name: `group-${gi}`,
      });
      if (span >= 16)
        label((rInner + rMid) / 2, angle + span / 2, g.name, contrastInk(gColor), `group-label-${gi}`, false);
      let a2 = angle;
      g.members.forEach((m, mi) => {
        const mspan = (m.value / g.total) * span;
        const mColor = lerpColor(gColor, style.background, 0.16 + 0.12 * (mi % 4));
        const ma0 = norm(a2);
        nodes.push({
          kind: "wedge",
          cx,
          cy,
          r,
          innerR: rMid,
          startAngle: ma0,
          endAngle: ma0 + mspan,
          fill: mColor,
          stroke: style.background,
          strokeWidth: 1,
          name: `slice-${m.i}`,
        });
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
      nodes.push({
        kind: "wedge",
        cx,
        cy,
        r,
        innerR: rInner,
        startAngle: a0,
        endAngle: a0 + span,
        fill: palette[m.i % palette.length],
        stroke: style.background,
        strokeWidth: 1,
        name: `slice-${m.i}`,
      });
      if (span >= 12)
        label(
          r + fs * 0.7,
          angle + span / 2,
          `${m.label} ${formatNumber(m.value, fmt)}`,
          style.text,
          `label-${m.i}`,
          true,
        );
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
