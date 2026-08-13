import type { ChartConfig, ChartStyle, Decorations } from "../types";
import { contrastInk, polar, textWidth, type SceneNode, type TextNode } from "../scene";
import { clipToWidth } from "../elements";
import { formatNumber, resolveFormat } from "../format";
import { lerpColor } from "../color";
import { fitPlot, footnoteH, titleHeight, titleNode, MIN_LABEL_FS } from "./frame";
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
  /** Every outside label, kept so each can be fitted to its neighbour after placement. */
  const outsideLabels: {
    node: TextNode;
    text: string;
    color: string;
    rightHalf: boolean;
    /** The ring point this label hangs off — its anchor, needed to re-place it. */
    px: number;
    cy: number;
    room: number;
  }[] = [];
  const label = (
    rr: number,
    midAngle: number,
    text: string,
    color: string,
    name: string,
    outside: boolean,
    /** The wedge's angular span, which is the room this label actually has. */
    spanDeg: number,
  ) => {
    if (!decor.segmentLabels) return;
    if (outside && !outerLabelsFit) return;
    const p = polar(cx, cy, rr, midAngle);
    const rightHalf = norm(midAngle) < 180;
    // An OUTSIDE label runs away from the ring edge with nothing to stop it —
    // the same defect the pie's slice labels had, in the same shape. Clipped to
    // the room actually there between the ring and the frame.
    //
    // Both kinds are also bounded by their own WEDGE, which neither was: an
    // inside label was drawn verbatim at the ring font, so a group name wider
    // than its arc ran into the next group's, and an outside label sat as tall
    // as the font whatever the arc under it could carry. Twenty of the 237
    // overlapping text pairs a sweep found were these two labels.
    //
    // Inside, the bound is the CHORD of the wedge at the label's radius, exactly
    // as the pie's inside labels take. Outside, it is the ARC each label gets at
    // its own radius, which is what separates it from its neighbours.
    const arc = (2 * Math.PI * rr * Math.min(360, Math.max(0, spanDeg))) / 360;
    const chord = 2 * rr * Math.sin((Math.min(spanDeg, 180) * Math.PI) / 360);
    const lf = outside ? Math.min(fs * 0.85, arc / 1.4) : fs * 0.85;
    const room = outside ? Math.max(lf, (rightHalf ? cfg.width - p.x : p.x) - 2) : Math.max(1, chord);
    const lfScale = lf / (fs * 0.85);
    const shown = clipToWidth(text, lf, room);
    const w = textWidth(shown, lf) + 4;
    const node: TextNode = {
      kind: "text",
      x: outside ? (rightHalf ? p.x : p.x - w) : p.x - w / 2,
      // As a RATIO of the unshrunk size, so the box and the font move together
      // and the geometry is byte-identical when nothing needed shrinking. The
      // radar's ticks were written the other way once and moved three showcase
      // slides on charts that needed no change at all.
      y: p.y - fs * 0.7 * lfScale,
      w,
      h: fs * 1.4 * lfScale,
      text: shown,
      fontSize: lf,
      color,
      align: outside ? (rightHalf ? "left" : "right") : "center",
      valign: "middle",
      name,
    };
    nodes.push(node);
    // An outside label is fitted to its own wedge's ARC above, and that says
    // nothing about where the NEIGHBOUR's midpoint falls: a wide wedge beside a
    // narrow one earns a tall label and still sits close to it. Every outside
    // label is on one circle, so the pass below can fit each to the vertical gap
    // it actually has once they are all placed.
    if (outside) outsideLabels.push({ node, text, color, rightHalf, px: p.x, cy: p.y, room });
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
        label((rInner + rMid) / 2, angle + span / 2, g.name, contrastInk(gColor), `group-label-${gi}`, false, span);
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
        if (mspan >= 12) label(r + fs * 0.7, a2 + mspan / 2, labelOf(m.label), style.text, `label-${m.i}`, true, mspan);
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
          span,
        );
      angle += span;
    });
  }

  // Fit every outside label to the vertical gap it actually has.
  //
  // Each was already bounded by its OWN wedge's arc, which is blind to where the
  // neighbour's midpoint falls — a wide wedge beside a narrow one earns a tall
  // label and still sits close to it. That left adjacent labels grazing by 0.3
  // to 1.0pt at 120x90 and 300x60: inside the frame, so no overflow gate could
  // see it, and the only overlaps left at the default font once every other
  // shape had been fixed.
  //
  // The gap between two neighbours has to carry HALF of each label's ink, so
  // sizing both to `gap / INK_RATIO` makes the two half-heights sum to exactly
  // the gap. `SNUG` takes it just inside that. Below the floor the label is
  // dropped, which is the answer the ring, radar, tilemap and pie reservations
  // already give when the room cannot be met — a label drawn through its
  // neighbour is not readable anyway.
  //
  // Sides are independent: a left label and a right label never meet, and
  // pooling them would shrink labels that have no neighbour near them.
  const INK_RATIO = 1.01; // a text node's ink height as a multiple of its font size
  const SNUG = 0.95;
  const dropped = new Set<TextNode>();
  for (const side of [true, false]) {
    const mine = outsideLabels.filter((l) => l.rightHalf === side).sort((a, b) => a.cy - b.cy);
    for (let i = 1; i < mine.length; i++) {
      const prev = mine[i - 1];
      const cur = mine[i];
      const gap = cur.cy - prev.cy;
      const need = (prev.node.fontSize + cur.node.fontSize) / 2;
      if (need * INK_RATIO <= gap) continue;
      const fit = (gap / INK_RATIO) * SNUG;
      // Past the floor, shrinking cannot separate them, so one has to go — but
      // only ONE. Dropping both loses a label the survivor's room could have
      // carried, and the smaller font is the narrower wedge, i.e. the smaller
      // share of the data. Its neighbour keeps its size: with this one gone
      // there is nothing left for it to collide with.
      if (fit < MIN_LABEL_FS) {
        dropped.add((prev.node.fontSize <= cur.node.fontSize ? prev : cur).node);
        continue;
      }
      for (const l of [prev, cur]) {
        if (l.node.fontSize <= fit) continue;
        // Re-clip at the new size: a smaller font fits more characters into the
        // same horizontal room, so keeping the old string would drop text the
        // label now has space for.
        const scale = fit / (fs * 0.85);
        l.node.fontSize = fit;
        l.node.text = clipToWidth(l.text, fit, l.room);
        l.node.w = textWidth(l.node.text, fit) + 4;
        // Re-anchored from the ring point, not nudged from the old box: a
        // right-aligned label's x is `px - w`, so a narrower box moves it.
        l.node.x = l.rightHalf ? l.px : l.px - l.node.w;
        l.node.y = l.cy - fs * 0.7 * scale;
        l.node.h = fs * 1.4 * scale;
      }
    }
  }
  const out = dropped.size ? nodes.filter((n) => !dropped.has(n as TextNode)) : nodes;

  return {
    nodes: out,
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
