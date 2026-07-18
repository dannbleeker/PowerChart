import type { ChartConfig, ChartStyle, Decorations } from "../types";
import { polar, textWidth, type SceneNode } from "../scene";
import { formatNumber, resolveFormat, segmentLabel } from "../format";
import { footnoteH, titleHeight, titleNode } from "./frame";
import type { LayoutResult } from "./column";

/**
 * Pie / doughnut chart from the first series; one slice per category,
 * colored by category. Slices start at 12 o'clock, clockwise, largest-first
 * ordering left to the data (think-cell keeps sheet order too).
 * Labels sit inside the slice when it is wide enough, otherwise outside.
 */
export function layoutPie(cfg: ChartConfig, style: ChartStyle, decor: Decorations): LayoutResult {
  const { data } = cfg;
  const fs = style.fontSize;
  const values = data.categories.map((_, c) => Math.max(0, data.series[0]?.values[c] ?? 0));
  // `total` is the honest sum shown in the doughnut hole / gauge centre;
  // `denom` guards the angle math against all-zero data (which used to display
  // the fallback "1" as the headline number).
  const total = values.reduce((a, b) => a + b, 0);
  const denom = total || 1;
  const fmt = resolveFormat(values, cfg.numberFormat);
  const doughnut = cfg.kind === "doughnut";
  // Semi-circle gauge: a half-doughnut scorecard (180° arc).
  if (doughnut && cfg.pie?.semi) return layoutGauge(cfg, style, decor, values, total, fmt);
  // Bar-of-pie breakout: these category indices collapse into one "Other"
  // slice, detailed in a stacked bar on the right (pie only).
  const breakout = !doughnut
    ? [...new Set((cfg.pie?.breakout ?? []).filter((c) => c >= 0 && c < values.length && values[c] > 0))]
    : [];
  const hasBreakout = breakout.length > 0;
  // Variable-radius pie: angle still encodes the first series, radius encodes a
  // second metric — a "Radius" datasheet row (or the second series). Pie only.
  const radiusRow = data.series.find((s) => /^radius$/i.test(s.name.trim()))?.values;
  const radiusVals = radiusRow ?? (cfg.pie?.variableRadius ? data.series[1]?.values : undefined);
  const varR = !!radiusVals && !doughnut && !hasBreakout;
  const maxRad = varR ? Math.max(1, ...radiusVals!.map((v) => Math.max(0, v ?? 0))) : 1;

  const titleH = titleHeight(cfg, style);
  const footH = footnoteH(cfg, style, decor);
  const cx = hasBreakout ? cfg.width * 0.3 : cfg.width / 2;
  const cy = titleH + (cfg.height - titleH - footH) / 2;
  // Floor at a positive radius: on a very narrow/short frame the width or height
  // term can go negative, which would mirror wedges through the centre and hand
  // the doughnut hole negative radii. Every sibling round chart (gauge, sunburst,
  // radar) clamps the same way.
  const r = Math.max(
    1,
    hasBreakout
      ? Math.min(cfg.width * 0.24, (cfg.height - titleH - footH) / 2 - fs * 2.2)
      : Math.min(cfg.width * 0.5 - fs * 7, (cfg.height - titleH - footH) / 2 - fs * 2.2),
  );

  const nodes: SceneNode[] = [];
  const titleN = titleNode(cfg, style);
  if (titleN) nodes.push(titleN);

  // Slice list: with a breakout, the collapsed categories become one muted
  // "Other" slice drawn last, and the pie rotates so Other faces the bar
  // (its midpoint at 3 o'clock).
  const otherSum = breakout.reduce((a, c) => a + values[c], 0);
  const slices: { v: number; c: number | "other" }[] = [
    ...values.map((v, c) => ({ v, c: c as number | "other" })).filter((s) => !breakout.includes(s.c as number)),
    ...(hasBreakout ? [{ v: otherSum, c: "other" as const }] : []),
  ];
  let angle = hasBreakout ? 90 + ((otherSum / denom) * 360) / 2 : 0;
  let otherStart = 0;
  slices.forEach(({ v, c }) => {
    const span = (v / denom) * 360;
    if (span <= 0) return;
    const other = c === "other";
    if (other) otherStart = angle;
    // Per-slice outer radius: full r normally, scaled by the radius metric in
    // variable-radius mode (floored at half r so small slices stay visible).
    const rr = varR && !other ? r * 0.5 + (Math.max(0, radiusVals![c as number] ?? 0) / maxRad) * (r * 0.5) : r;
    const fill = other
      ? style.neutral
      : (data.series[0]?.colors?.[c as number] ?? style.palette[(c as number) % style.palette.length]);
    // Exploding slice: offset the wedge radially to highlight it.
    const exploded = !other && (cfg.pie?.explode?.includes(c as number) ?? false);
    const off = exploded ? polar(0, 0, rr * 0.08, angle + span / 2) : { x: 0, y: 0 };
    const ecx = cx + off.x;
    const ecy = cy + off.y;
    const a0 = ((angle % 360) + 360) % 360;
    nodes.push({
      kind: "wedge",
      cx: ecx,
      cy: ecy,
      r: rr,
      innerR: 0,
      startAngle: a0,
      endAngle: a0 + span,
      fill,
      stroke: style.background,
      strokeWidth: 1,
      name: other ? "slice-other" : `slice-${c}`,
    });

    if (decor.segmentLabels) {
      const mid = angle + span / 2;
      const label = segmentLabel(decor.labelContent ?? ["category", "percent"], {
        value: v,
        fraction: v / denom,
        series: data.series[0]?.name ?? "",
        category: other ? "Other" : data.categories[c as number],
        fmt,
      });
      const inside = span >= 30 && !doughnut && !varR;
      const p = polar(ecx, ecy, inside ? rr * 0.62 : rr + fs * 0.8, mid);
      const w = textWidth(label, fs) + 4;
      const rightHalf = ((mid % 360) + 360) % 360 < 180;
      if (!inside) {
        // Leader line from the arc edge toward the label.
        const a = polar(ecx, ecy, rr + 1, mid);
        const b = polar(ecx, ecy, rr + fs * 0.65, mid);
        nodes.push({
          kind: "line",
          x1: a.x,
          y1: a.y,
          x2: b.x,
          y2: b.y,
          stroke: style.mutedText,
          strokeWidth: 0.75,
          name: other ? "leader-other" : `leader-${c}`,
        });
      }
      nodes.push({
        kind: "text",
        x: inside ? p.x - w / 2 : rightHalf ? p.x : p.x - w,
        y: p.y - fs * 0.75,
        w,
        h: fs * 1.5,
        text: label,
        fontSize: fs,
        color: inside ? "#ffffff" : style.text,
        align: inside ? "center" : rightHalf ? "left" : "right",
        valign: "middle",
        name: other ? "label-other" : `label-${c}`,
      });
    }
    angle += span;
  });

  // Detail bar for the breakout categories, joined by connector lines.
  if (hasBreakout && otherSum > 0) {
    const barW = fs * 2.8;
    const barH = Math.min(r * 1.9, cfg.height - titleH - footH - fs * 2);
    const barX = cfg.width * 0.64;
    const barY = cy - barH / 2;
    const mainsCount = slices.length - 1;
    // Other's boundary edges → bar corners. The slice is centered at
    // 3 o'clock, so its start edge (90° − span/2) is the upper one.
    const eTop = polar(cx, cy, r, otherStart);
    const eBot = polar(cx, cy, r, otherStart + (otherSum / denom) * 360);
    nodes.push(
      {
        kind: "line",
        x1: eTop.x,
        y1: eTop.y,
        x2: barX,
        y2: barY,
        stroke: style.mutedText,
        strokeWidth: 0.75,
        dash: [3, 3],
        name: "breakout-conn-0",
      },
      {
        kind: "line",
        x1: eBot.x,
        y1: eBot.y,
        x2: barX,
        y2: barY + barH,
        stroke: style.mutedText,
        strokeWidth: 0.75,
        dash: [3, 3],
        name: "breakout-conn-1",
      },
    );
    let y = barY;
    breakout.forEach((c, j) => {
      const h = (values[c] / otherSum) * barH;
      const fill = data.series[0]?.colors?.[c] ?? style.palette[(mainsCount + j) % style.palette.length];
      nodes.push({
        kind: "rect",
        x: barX,
        y,
        w: barW,
        h,
        fill,
        stroke: style.background,
        strokeWidth: 1,
        name: `breakout-seg-${c}`,
      });
      if (decor.segmentLabels) {
        const label = segmentLabel(decor.labelContent ?? ["category", "percent"], {
          value: values[c],
          fraction: values[c] / denom,
          series: data.series[0]?.name ?? "",
          category: data.categories[c],
          fmt,
        });
        nodes.push({
          kind: "text",
          x: barX + barW + 5,
          y: y + h / 2 - fs * 0.75,
          w: cfg.width - barX - barW - 7,
          h: fs * 1.5,
          text: label,
          fontSize: fs,
          color: style.text,
          align: "left",
          valign: "middle",
          name: `breakout-label-${c}`,
        });
      }
      y += h;
    });
  }

  if (doughnut) {
    nodes.push({ kind: "ellipse", cx, cy, rx: r * 0.55, ry: r * 0.55, fill: style.background, name: "hole" });
    nodes.push({
      kind: "text",
      x: cx - r * 0.5,
      y: cy - fs * 0.9,
      w: r,
      h: fs * 1.8,
      text: formatNumber(total, fmt),
      fontSize: fs * 1.3,
      bold: true,
      color: style.text,
      align: "center",
      valign: "middle",
      name: "hole-label",
    });
  }

  return {
    nodes,
    anchors: {
      categoryX: values.map(() => cx),
      categoryWidth: values.map(() => r),
      columnTop: values.map(() => cy - r),
      columnValue: values,
      baselineY: cy,
      plot: { x: cx - r, y: cy - r, w: r * 2, h: r * 2 },
    },
  };
}

/**
 * Semi-circle gauge (half doughnut): categories fill a 180° arc across the top
 * (9 o'clock → 12 → 3 o'clock), with a big total in the open centre. A
 * scorecard staple.
 */
function layoutGauge(
  cfg: ChartConfig,
  style: ChartStyle,
  decor: Decorations,
  values: number[],
  total: number,
  fmt: ReturnType<typeof resolveFormat>,
): LayoutResult {
  const { data } = cfg;
  const fs = style.fontSize;
  const denom = total || 1;
  const titleH = titleHeight(cfg, style);
  const footH = footnoteH(cfg, style, decor);
  const availH = cfg.height - titleH - footH;
  const cx = cfg.width / 2;
  const r = Math.max(20, Math.min(cfg.width / 2 - fs * 3, availH * 0.82));
  const cy = titleH + r + fs * 0.3; // arc peak at the top, flat side at cy
  const innerR = r * 0.58;

  const nodes: SceneNode[] = [];
  const titleN = titleNode(cfg, style);
  if (titleN) nodes.push(titleN);
  let angle = 270; // start at 9 o'clock, sweep clockwise over the top to 3 o'clock
  values.forEach((v, c) => {
    const span = (v / denom) * 180;
    if (span <= 0) return;
    const fill = data.series[0]?.colors?.[c] ?? style.palette[c % style.palette.length];
    const a0 = ((angle % 360) + 360) % 360;
    nodes.push({
      kind: "wedge",
      cx,
      cy,
      r,
      innerR,
      startAngle: a0,
      endAngle: a0 + span,
      fill,
      stroke: style.background,
      strokeWidth: 1,
      name: `slice-${c}`,
    });
    if (decor.segmentLabels) {
      const mid = angle + span / 2;
      const label = segmentLabel(decor.labelContent ?? ["category", "percent"], {
        value: v,
        fraction: v / denom,
        series: data.series[0]?.name ?? "",
        category: data.categories[c],
        fmt,
      });
      const p = polar(cx, cy, r + fs * 0.8, mid);
      const w = textWidth(label, fs) + 4;
      const rightHalf = ((mid % 360) + 360) % 360 < 180;
      const a = polar(cx, cy, r + 1, mid);
      const b = polar(cx, cy, r + fs * 0.65, mid);
      nodes.push({
        kind: "line",
        x1: a.x,
        y1: a.y,
        x2: b.x,
        y2: b.y,
        stroke: style.mutedText,
        strokeWidth: 0.75,
        name: `leader-${c}`,
      });
      nodes.push({
        kind: "text",
        x: rightHalf ? p.x : p.x - w,
        y: p.y - fs * 0.75,
        w,
        h: fs * 1.5,
        text: label,
        fontSize: fs,
        color: style.text,
        align: rightHalf ? "left" : "right",
        valign: "middle",
        name: `label-${c}`,
      });
    }
    angle += span;
  });
  // Big total in the open centre of the arc.
  nodes.push({
    kind: "text",
    x: cx - r,
    y: cy - fs * 1.7,
    w: r * 2,
    h: fs * 2,
    text: formatNumber(total, fmt),
    fontSize: fs * 1.7,
    bold: true,
    color: style.text,
    align: "center",
    valign: "middle",
    name: "gauge-total",
  });

  return {
    nodes,
    anchors: {
      categoryX: values.map(() => cx),
      categoryWidth: values.map(() => r),
      columnTop: values.map(() => cy - r),
      columnValue: values,
      baselineY: cy,
      plot: { x: cx - r, y: cy - r, w: r * 2, h: r },
    },
  };
}
