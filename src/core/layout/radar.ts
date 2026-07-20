import type { ChartConfig, ChartStyle, Decorations } from "../types";
import { polar, textWidth, type SceneNode } from "../scene";
import { formatNumber, niceTicks, resolveFormat } from "../format";
import { seriesColor } from "../style";
import { footnoteH, titleHeight, titleNode } from "./frame";
import type { LayoutResult } from "./column";
import { columnPositiveTotal } from "./totals";
import { maxOf } from "../agg";

/**
 * Radar (spider) chart: categories = spokes (first at 12 o'clock,
 * clockwise), series = polygons with translucent fills. Gridlines are
 * straight polygons by default (business style); scale is shared across
 * all spokes and ticked on the 12 o'clock spoke only.
 */
export function layoutRadar(cfg: ChartConfig, style: ChartStyle, decor: Decorations): LayoutResult {
  if (cfg.radar?.bars) return layoutRadialBars(cfg, style, decor);
  const { data } = cfg;
  const n = data.categories.length;
  const fs = style.fontSize;

  const titleH = titleHeight(cfg, style);
  const footH = footnoteH(cfg, style, decor);
  const legendH = decor.seriesLabels && data.series.length > 1 ? fs * 1.6 : 0;
  const cx = cfg.width / 2;
  const cy = titleH + legendH + (cfg.height - titleH - legendH - footH) / 2;
  // Perimeter labels need a margin around the web.
  const labelW = Math.max(0, ...data.categories.map((c) => textWidth(c, fs)));
  const r = Math.max(10, Math.min(cfg.width / 2 - labelW - fs, (cfg.height - titleH - legendH - footH) / 2 - fs * 1.9));

  // Stacked radar: series stack cumulatively along each spoke, so the scale
  // must reach the per-spoke sums, not the largest single value.
  const stacked = !!cfg.radar?.stacked && data.series.length >= 2;
  const spokeSum = data.categories.map((_, c) => columnPositiveTotal(data.series, c));
  const all = data.series.flatMap((s) => s.values.filter((v): v is number => v != null));
  // `all` is cells-scaled (series x categories) so it must be FOLDED, not spread —
  // a large grid overflows the argument list. `spokeSum` is category-scaled and safe.
  const tickMax = stacked ? Math.max(1, ...spokeSum) : maxOf(all, 1);
  const ticks = niceTicks(Math.min(0, cfg.scale?.min ?? 0), Math.max(cfg.scale?.max ?? tickMax, 1), 4);
  const min = cfg.scale?.min ?? ticks[0];
  const max = cfg.scale?.max ?? ticks[ticks.length - 1];
  const fmt = resolveFormat(ticks, cfg.numberFormat);
  const toR = (v: number) => ((v - min) / (max - min || 1)) * r;
  const angle = (c: number) => (360 / Math.max(1, n)) * c;
  // Per-spoke scales: normalise each spoke to its own maximum so spokes in
  // different KPI units become comparable in shape (numeric ticks dropped).
  const perSpoke = !!cfg.radar?.perSpoke && !stacked && data.series.length >= 1;
  const spokeMax = data.categories.map((_, c) =>
    perSpoke ? Math.max(1, ...data.series.map((s) => s.values[c] ?? 0)) : max,
  );
  const toRc = (v: number, c: number) => (perSpoke ? (Math.max(0, v) / spokeMax[c]) * r : toR(Math.max(min, v)));

  const nodes: SceneNode[] = [];
  const titleN = titleNode(cfg, style);
  if (titleN) nodes.push(titleN);

  // Grid: concentric polygons (or circles). Per-spoke mode uses fraction
  // rings with no numeric labels (each spoke has its own scale); otherwise
  // rings sit at the value ticks and the 12 o'clock spoke is labelled.
  const gridShape = decor.gridShape ?? "polygon";
  const rings = perSpoke
    ? [0.25, 0.5, 0.75, 1].map((f) => ({ rr: f * r, t: f }))
    : ticks.filter((t) => t > min).map((t) => ({ rr: toR(t), t }));
  for (const { rr, t } of rings) {
    if (gridShape === "circle") {
      nodes.push({
        kind: "ellipse",
        cx,
        cy,
        rx: rr,
        ry: rr,
        fill: "none",
        stroke: style.gridline,
        strokeWidth: 0.75,
        name: `grid-${t}`,
      });
    } else {
      nodes.push({
        kind: "polygon",
        points: data.categories.map((_, c) => polar(cx, cy, rr, angle(c))),
        stroke: style.gridline,
        strokeWidth: 0.75,
        name: `grid-${t}`,
      });
    }
    if (!perSpoke) {
      // Tick label on the 12 o'clock spoke only.
      nodes.push({
        kind: "text",
        x: cx + 3,
        y: cy - rr - fs * 0.6,
        w: fs * 3.4,
        h: fs * 1.2,
        text: formatNumber(t, fmt),
        fontSize: fs * 0.85,
        color: style.mutedText,
        align: "left",
        valign: "middle",
        name: `tick-${t}`,
      });
    }
  }
  data.categories.forEach((cat, c) => {
    const end = polar(cx, cy, r, angle(c));
    nodes.push({
      kind: "line",
      x1: cx,
      y1: cy,
      x2: end.x,
      y2: end.y,
      stroke: style.gridline,
      strokeWidth: 0.75,
      name: `spoke-${c}`,
    });
    // Perimeter category label, anchored by which side of the web it sits on.
    const p = polar(cx, cy, r + fs * 0.6, angle(c));
    const a = angle(c) % 360;
    const align = a < 10 || a > 350 || Math.abs(a - 180) < 10 ? "center" : a < 180 ? "left" : "right";
    const w = textWidth(cat, fs) + 4;
    nodes.push({
      kind: "text",
      x: align === "center" ? p.x - w / 2 : align === "left" ? p.x : p.x - w,
      y: p.y - (a < 10 || a > 350 ? fs * 1.4 : Math.abs(a - 180) < 10 ? 0 : fs * 0.7),
      w,
      h: fs * 1.4,
      text: cat,
      fontSize: fs,
      color: style.text,
      align,
      valign: "middle",
      name: `category-${c}`,
    });
  });

  // Min–max band: shade the per-spoke envelope of the peer series (all but
  // the last) as an annulus of per-sector quads, then draw the last series
  // ("us") prominently on top — the "peer range + us" view.
  const band = !!decor.radarBand && !stacked && data.series.length >= 2;
  if (band) {
    const peers = data.series.slice(0, -1);
    const peerMin: number[] = [];
    const peerMax: number[] = [];
    for (let c = 0; c < n; c++) {
      const vals = peers.map((s) => s.values[c]).filter((v): v is number => v != null);
      peerMin[c] = vals.length ? Math.max(min, Math.min(...vals)) : min;
      peerMax[c] = vals.length ? Math.max(min, Math.max(...vals)) : min;
    }
    const minPts = data.categories.map((_, c) => polar(cx, cy, toRc(peerMin[c], c), angle(c)));
    const maxPts = data.categories.map((_, c) => polar(cx, cy, toRc(peerMax[c], c), angle(c)));
    for (let c = 0; c < n; c++) {
      const c2 = (c + 1) % n;
      nodes.push({
        kind: "polygon",
        points: [minPts[c], minPts[c2], maxPts[c2], maxPts[c]],
        fill: style.mutedText,
        fillOpacity: 0.16,
        name: `band-${c}`,
      });
    }
    nodes.push(
      { kind: "polygon", points: maxPts, stroke: style.mutedText, strokeWidth: 1, name: "band-max" },
      { kind: "polygon", points: minPts, stroke: style.mutedText, strokeWidth: 1, name: "band-min" },
    );
  }

  // Stacked radar: draw each series as an annular band between the cumulative
  // level below it and its own cumulative level (part-to-whole across spokes).
  if (stacked) {
    const cum = data.categories.map(() => 0);
    data.series.forEach((s, si) => {
      const innerPts = data.categories.map((_, c) => polar(cx, cy, toR(cum[c]), angle(c)));
      for (let c = 0; c < n; c++) cum[c] += Math.max(0, s.values[c] ?? 0);
      const outerPts = data.categories.map((_, c) => polar(cx, cy, toR(cum[c]), angle(c)));
      const color = seriesColor(style, si, s.color);
      nodes.push({
        kind: "polygon",
        points: [...outerPts, ...innerPts.slice().reverse()],
        fill: color,
        fillOpacity: decor.fillOpacity ?? 0.55,
        stroke: color,
        strokeWidth: 1.2,
        name: `series-${si}`,
      });
    });
    if (legendH) drawLegend();
    return {
      nodes,
      anchors: {
        categoryX: data.categories.map((_, c) => polar(cx, cy, r, angle(c)).x),
        categoryWidth: data.categories.map(() => r / 2),
        columnTop: data.categories.map((_, c) => polar(cx, cy, r, angle(c)).y),
        columnValue: data.categories.map((_, c) => data.series[0]?.values[c] ?? 0),
        baselineY: cy,
        plot: { x: cx - r, y: cy - r, w: r * 2, h: r * 2 },
      },
    };
  }

  // Series polygons: translucent fill (SVG), full-opacity outline + markers.
  const defaultOpacity = data.series.length === 1 ? 0.25 : 0.18;
  const drawIdx = band ? [data.series.length - 1] : data.series.map((_, i) => i);
  drawIdx.forEach((si) => {
    const s = data.series[si];
    const color = seriesColor(style, si, s.color);
    const pts = data.categories.map((_, c) => polar(cx, cy, toRc(s.values[c] ?? min, c), angle(c)));
    nodes.push({
      kind: "polygon",
      points: pts,
      fill: color,
      fillOpacity: decor.fillOpacity ?? defaultOpacity,
      stroke: color,
      strokeWidth: 1.6,
      name: `series-${si}`,
    });
    pts.forEach((p, c) => {
      if (s.values[c] == null) return;
      nodes.push({
        kind: "ellipse",
        cx: p.x,
        cy: p.y,
        rx: 2.4,
        ry: 2.4,
        fill: color,
        stroke: style.background,
        strokeWidth: 1,
        name: `marker-${si}-${c}`,
      });
    });
  });

  // Legend row under the title when there are multiple series. In band mode
  // it collapses the peers into one "Peer range" swatch plus the "us" series.
  if (legendH) drawLegend();

  // Hoisted so the stacked-radar early return can reuse it.
  function drawLegend() {
    let x = 0;
    const chip = fs * 0.7;
    const entries: { label: string; color: string; name: string }[] = band
      ? [
          { label: "Peer range", color: style.mutedText, name: "legend-band" },
          {
            label: data.series[data.series.length - 1].name,
            color: seriesColor(style, data.series.length - 1, data.series[data.series.length - 1].color),
            name: "legend-us",
          },
        ]
      : data.series.map((s, si) => ({ label: s.name, color: seriesColor(style, si, s.color), name: `legend-${si}` }));
    entries.forEach((e, i) => {
      nodes.push(
        {
          kind: "rect",
          x,
          y: titleH + fs * 0.35,
          w: chip,
          h: chip,
          fill: e.color,
          name: band ? e.name : `legend-chip-${i}`,
        },
        {
          kind: "text",
          x: x + chip + 3,
          y: titleH,
          w: textWidth(e.label, fs) + 6,
          h: fs * 1.4,
          text: e.label,
          fontSize: fs,
          color: style.text,
          align: "left",
          valign: "middle",
          name: e.name,
        },
      );
      x += chip + 3 + textWidth(e.label, fs) + 12;
    });
  }

  return {
    nodes,
    anchors: {
      categoryX: data.categories.map((_, c) => polar(cx, cy, r, angle(c)).x),
      categoryWidth: data.categories.map(() => r / 2),
      columnTop: data.categories.map((_, c) => polar(cx, cy, r, angle(c)).y),
      columnValue: data.categories.map((_, c) => data.series[0]?.values[c] ?? 0),
      baselineY: cy,
      plot: { x: cx - r, y: cy - r, w: r * 2, h: r * 2 },
    },
  };
}

/**
 * Radial (polar) bar chart / coxcomb: each category is an equal-angle sector
 * whose bar radius encodes its value, drawn from a small inner hole so it reads
 * as bars bent around a circle rather than a pie. A single series colours bars
 * by category (Nightingale rose); multiple series stack outward within each
 * sector. Concentric value rings give the scale.
 */
function layoutRadialBars(cfg: ChartConfig, style: ChartStyle, decor: Decorations): LayoutResult {
  const { data } = cfg;
  const n = data.categories.length;
  const fs = style.fontSize;
  const multi = data.series.length > 1;

  const titleH = titleHeight(cfg, style);
  const footH = footnoteH(cfg, style, decor);
  const legendH = decor.seriesLabels && multi ? fs * 1.6 : 0;
  const cx = cfg.width / 2;
  const cy = titleH + legendH + (cfg.height - titleH - legendH - footH) / 2;
  const labelW = Math.max(0, ...data.categories.map((c) => textWidth(c, fs)));
  const r = Math.max(10, Math.min(cfg.width / 2 - labelW - fs, (cfg.height - titleH - legendH - footH) / 2 - fs * 1.9));
  const innerR = r * 0.18;

  // Scale reaches the per-category stack sums (a single series is its own sum).
  const catSum = data.categories.map((_, c) => columnPositiveTotal(data.series, c));
  const ticks = niceTicks(0, Math.max(cfg.scale?.max ?? Math.max(1, ...catSum), 1), 4);
  const max = cfg.scale?.max ?? ticks[ticks.length - 1];
  const fmt = resolveFormat(ticks, cfg.numberFormat);
  const toR = (v: number) => innerR + (Math.max(0, v) / (max || 1)) * (r - innerR);
  const sector = 360 / Math.max(1, n);
  const angle = (c: number) => sector * c;
  const pad = sector * 0.12;

  const nodes: SceneNode[] = [];
  const titleN = titleNode(cfg, style);
  if (titleN) nodes.push(titleN);

  // Concentric value rings + tick labels on the 12 o'clock line.
  for (const t of ticks.filter((t) => t > 0)) {
    nodes.push({
      kind: "ellipse",
      cx,
      cy,
      rx: toR(t),
      ry: toR(t),
      fill: "none",
      stroke: style.gridline,
      strokeWidth: 0.75,
      name: `grid-${t}`,
    });
    nodes.push({
      kind: "text",
      x: cx + 3,
      y: cy - toR(t) - fs * 0.6,
      w: fs * 3.4,
      h: fs * 1.2,
      text: formatNumber(t, fmt),
      fontSize: fs * 0.85,
      color: style.mutedText,
      align: "left",
      valign: "middle",
      name: `tick-${t}`,
    });
  }

  data.categories.forEach((cat, c) => {
    const a0 = (((angle(c) + pad / 2) % 360) + 360) % 360;
    const aSpan = sector - pad;
    let base = 0;
    data.series.forEach((s, si) => {
      const v = Math.max(0, s.values[c] ?? 0);
      if (v <= 0) return;
      const rin = toR(base);
      const rout = toR(base + v);
      base += v;
      const color = multi
        ? seriesColor(style, si, s.color)
        : (s.colors?.[c] ?? style.palette[c % style.palette.length]);
      nodes.push({
        kind: "wedge",
        cx,
        cy,
        r: rout,
        innerR: rin,
        startAngle: a0,
        endAngle: a0 + aSpan,
        fill: color,
        stroke: style.background,
        strokeWidth: 1,
        name: multi ? `bar-${c}-${si}` : `bar-${c}`,
      });
    });
    // Perimeter category label.
    const mid = angle(c) + sector / 2;
    const p = polar(cx, cy, r + fs * 0.6, mid);
    const am = ((mid % 360) + 360) % 360;
    const align = am < 10 || am > 350 || Math.abs(am - 180) < 10 ? "center" : am < 180 ? "left" : "right";
    const w = textWidth(cat, fs) + 4;
    nodes.push({
      kind: "text",
      x: align === "center" ? p.x - w / 2 : align === "left" ? p.x : p.x - w,
      y: p.y - (am < 10 || am > 350 ? fs * 1.4 : Math.abs(am - 180) < 10 ? 0 : fs * 0.7),
      w,
      h: fs * 1.4,
      text: cat,
      fontSize: fs,
      color: style.text,
      align,
      valign: "middle",
      name: `category-${c}`,
    });
  });

  // Series legend (multi-series stacks).
  if (legendH) {
    let x = 0;
    const chip = fs * 0.7;
    data.series.forEach((s, si) => {
      const color = seriesColor(style, si, s.color);
      nodes.push(
        { kind: "rect", x, y: titleH + fs * 0.35, w: chip, h: chip, fill: color, name: `legend-chip-${si}` },
        {
          kind: "text",
          x: x + chip + 3,
          y: titleH,
          w: textWidth(s.name, fs) + 6,
          h: fs * 1.4,
          text: s.name,
          fontSize: fs,
          color: style.text,
          align: "left",
          valign: "middle",
          name: `legend-${si}`,
        },
      );
      x += chip + 3 + textWidth(s.name, fs) + 12;
    });
  }

  return {
    nodes,
    anchors: {
      categoryX: data.categories.map((_, c) => polar(cx, cy, r, angle(c) + sector / 2).x),
      categoryWidth: data.categories.map(() => (r - innerR) / 2),
      columnTop: data.categories.map((_, c) => polar(cx, cy, r, angle(c) + sector / 2).y),
      columnValue: catSum,
      baselineY: cy,
      plot: { x: cx - r, y: cy - r, w: r * 2, h: r * 2 },
    },
  };
}
