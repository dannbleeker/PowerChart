import type { ChartConfig, ChartStyle, Decorations } from "../types";
import { contrastInk, polar, textWidth, type SceneNode } from "../scene";
import { clipToWidth } from "../elements";
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
  const availH = cfg.height - titleH - footH;
  /** The smallest arc that still reads as a pie and not as a dot. */
  const MIN_ARC_R = 12;
  // What the arc gets when the OUTSIDE labels are given the margins they want:
  // `fs * 7` either side for the text (the breakout path reserves its side
  // differently, through `cx`, and measures its own labels 180 lines below), and
  // `fs * 2.2` above and below for the ring.
  const wantX = hasBreakout ? cfg.width * 0.24 : cfg.width * 0.5 - fs * 7;
  const wantY = availH / 2 - fs * 2.2;
  // Those margins are a flat guess, and on a small frame they are most of it: a
  // pie under ~140pt wide had nothing left and fell to the 1pt floor below, so a
  // 120x90 thumbnail was a 2pt dot — 0.1% of the frame in ink against 38% at
  // 200x150 — with four labels drawn around it as though there were a chart
  // there.
  //
  // When a reservation cannot be met the thing reserved for is DROPPED, which is
  // the answer the radar's web and the sunburst's ring already give: a label
  // ring squeezed onto the arc it is labelling is not readable either, and the
  // arc it displaced was the chart. So the outer labels come off and the arc
  // takes the whole frame. Unreachable at 200x150 and above, where the margins
  // already leave an arc — so nothing of an ordinary size moves.
  //
  // The pie's INSIDE labels are unaffected: they are bounded by the slice they
  // sit in, which is exactly what just got bigger.
  const outerLabelsFit = Math.min(wantX, wantY) >= MIN_ARC_R;
  const rWidth = outerLabelsFit || hasBreakout ? wantX : cfg.width * 0.5 - fs * 0.5;
  const rHeight = outerLabelsFit ? wantY : availH / 2 - fs * 0.5;
  // Floor at a positive radius: on a very narrow/short frame the terms above go
  // negative, which would mirror wedges through the centre and hand the doughnut
  // hole negative radii. Every sibling round chart (gauge, sunburst, radar)
  // clamps the same way.
  const r = Math.max(1, Math.min(rWidth, rHeight));

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
  /** The label a slice gets, so the fit below measures the text that is drawn. */
  const sliceLabelText = (v: number, c: number | "other") =>
    segmentLabel(decor.labelContent ?? ["category", "percent"], {
      value: v,
      fraction: v / denom,
      series: data.series[0]?.name ?? "",
      category: c === "other" ? (cfg.labels?.other ?? "Other") : data.categories[c as number],
      fmt,
    });

  /**
   * The room an INSIDE label has: the chord of its own slice at the radius the
   * label sits on. A wedge is not a box, and this is the width of the one across
   * the middle of the text — conservative for the corners, which is the right
   * direction for a label that must not cross into the next slice.
   */
  const insideChord = (spanDeg: number) => 2 * (r * 0.62) * Math.sin((Math.min(spanDeg, 180) * Math.PI) / 360);

  /**
   * One size for every inside label, small enough that each fits its own slice.
   *
   * These were drawn at the chart font and fitted to nothing: a name wider than
   * its own wedge ran across the neighbouring slices, and on a small frame past
   * the edge of the chart, where the frame clip cut it to an ellipsis. Both the
   * preview and the deck showed "mericas 38…" on a 200x150 pie.
   *
   * Shrunk TOGETHER, the way the funnel's rows and the butterfly's category names
   * are: a pie whose labels are each at a different size reads as a mistake
   * rather than as a fit. Only slices big enough to GET an inside label take
   * part, which is what bounds it — a thin slice is labelled outside and cannot
   * drag the rest down with it.
   *
   * Last resort, like its siblings: at any size where the labels already fit this
   * is `fs` and nothing moves, which is every pie at 300x200 and above.
   */
  const insideFs = (() => {
    if (doughnut || varR || !decor.segmentLabels) return fs;
    const wants = slices
      .map(({ v, c }) => ({ span: (v / denom) * 360, text: sliceLabelText(v, c) }))
      .filter((w) => w.span >= 30);
    let f = fs;
    while (f > 5 && wants.some((w) => textWidth(w.text, f) > insideChord(w.span))) f -= 0.5;
    return f;
  })();

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

    // NB: `angle` is advanced at the END of this callback, so nothing in here may
    // `return` — an early exit skips the advance and every later slice starts at
    // the same place. It is written as a condition on the block for that reason.
    const inside = span >= 30 && !doughnut && !varR;
    if (decor.segmentLabels && (inside || outerLabelsFit)) {
      const mid = angle + span / 2;
      const label = sliceLabelText(v, c);
      const p = polar(ecx, ecy, inside ? rr * 0.62 : rr + fs * 0.8, mid);
      const rightHalf = ((mid % 360) + 360) % 360 < 180;
      // An OUTSIDE label runs away from the slice edge with nothing to stop it.
      // The radius reserves a FIXED `fs * 7` for labels — a guess, where the
      // breakout path 180 lines below measures the widest label it actually has
      // — so any category name wider than that guess put ink off the chart.
      // "A very long category label indeed" on a 480pt frame reached x = 548:
      // 68pt past the right edge, and neither PowerPoint renderer wraps or clips
      // a text box, so in a deck that is a label lying across whatever sits
      // beside the chart on the slide.
      //
      // Clipped to what is actually there, which is what `clipToWidth` is for
      // and what the agenda already does with its chapter titles.
      //
      // An inside label is bounded by its own SLICE rather than by the frame, so
      // it takes the chord (see `insideChord`) and the size the whole set was
      // shrunk to. Shrink first and clip second, the same order as the funnel's
      // rows and the butterfly's category names: shrinking keeps the whole word,
      // and the clip is only there for the label no floor can fit.
      const lf = inside ? insideFs : fs;
      const room = inside ? insideChord(span) : Math.max(fs, (rightHalf ? cfg.width - p.x : p.x) - 4);
      const shown = clipToWidth(label, lf, room);
      const w = textWidth(shown, lf) + 4;
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
        y: p.y - lf * 0.75,
        w,
        h: lf * 1.5,
        text: shown,
        fontSize: lf,
        // Inside a slice the ink must contrast with THAT slice: a pale fill (the
        // default palette's #eda100, or any custom/per-point colour) printed white
        // on light — invisible. Outside, the label sits on the canvas.
        color: inside ? contrastInk(fill) : style.text,
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
  /** The label each slice will get, so the margin can be measured rather than guessed. */
  const sliceLabel = (v: number, c: number) =>
    segmentLabel(decor.labelContent ?? ["category", "percent"], {
      value: v,
      fraction: v / denom,
      series: data.series[0]?.name ?? "",
      category: data.categories[c],
      fmt,
    });
  // The side margin has to hold the LABELS. `fs * 3` is 30pt at the default
  // font, and "Others 12%" is 58 — so the outer labels of the shipped showcase
  // gauge ran 32pt past the right edge of the slide. Measure the widest label
  // and reserve that, never less than the old guess (so a gauge whose labels
  // are short is unchanged).
  const sideMargin = decor.segmentLabels
    ? Math.max(fs * 3, values.reduce((m, v, c) => Math.max(m, textWidth(sliceLabel(v, c), fs) + 4), 0) + fs * 0.8)
    : fs * 3;
  const r = Math.max(20, Math.min(cfg.width / 2 - sideMargin, availH * 0.82));
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
      const label = sliceLabel(v, c);
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
        // Clamped as a floor: the margin above sizes the arc so the labels fit,
        // but on a chart too narrow to hold them at all the arc bottoms out at
        // its 20pt minimum and the label would still leave the canvas.
        x: Math.max(0, Math.min(cfg.width - w, rightHalf ? p.x : p.x - w)),
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
