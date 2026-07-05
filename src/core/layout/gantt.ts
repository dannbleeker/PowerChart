import type { ChartConfig, ChartStyle, Decorations } from "../types";
import { textWidth, type SceneNode } from "../scene";
import { formatDay, formatNumber, monthStarts, niceTicks, resolveFormat, weekStarts } from "../format";
import { seriesColor } from "../style";
import type { LayoutResult } from "./column";

/**
 * Simplified Gantt / timeline: categories are activities; rows named
 * Start and End give each activity's span on a numeric timeline (week,
 * month index, year — any number). A row named Milestone adds a diamond
 * marker at that position. think-cell's calendar-based Gantt is richer;
 * this covers the project-on-a-slide case.
 */
export function layoutGantt(cfg: ChartConfig, style: ChartStyle, decor: Decorations): LayoutResult {
  const { data } = cfg;
  const fs = style.fontSize;
  const find = (re: RegExp) => data.series.find((s) => re.test(s.name.trim()));
  const starts = find(/^start$/i)?.values ?? [];
  const ends = find(/^end$/i)?.values ?? [];
  const milestones = find(/^milestone$/i)?.values ?? [];
  /** "After" row: 1-based predecessor index → dependency arrow. */
  const after = find(/^(after|dep(endency)?)$/i)?.values ?? [];
  /** "Today" row: a single date/number → today line. */
  const today = (find(/^today$/i)?.values ?? []).find((v): v is number => v != null);
  // "Activity | Owner | Remark" category convention; ">" prefix indents.
  const parts = data.categories.map((c) => c.split("|").map((p) => p.trim()));
  const indents = parts.map((p) => (p[0].startsWith(">") ? 1 : 0));
  const acts = parts.map((p) => p[0].replace(/^>+\s*/, ""));
  const owners = parts.map((p) => p[1] ?? "");
  const remarks = parts.map((p) => p[2] ?? "");
  const hasOwners = owners.some(Boolean);
  const hasRemarks = remarks.some(Boolean);
  // A row with no bar data at all is a section header.
  const isHeader = data.categories.map(
    (_, c) => starts[c] == null && ends[c] == null && milestones[c] == null,
  );

  const titleH = cfg.title ? fs * 1.6 + 6 : 0;
  const headerH = fs * 1.6;
  const catW = Math.min(
    cfg.width * 0.32,
    Math.max(0, ...acts.map((c) => textWidth(c, fs))) + 10,
  );
  const ownerW = hasOwners ? Math.max(0, ...owners.map((o) => textWidth(o, fs))) + 12 : 0;
  const remarkW = hasRemarks ? Math.max(0, ...remarks.map((o) => textWidth(o, fs * 0.9))) + 12 : 0;
  const bottomH = today != null ? fs * 1.6 : 6;
  const plot = { x: catW, y: titleH + headerH, w: cfg.width - catW - ownerW - remarkW - 6, h: cfg.height - titleH - headerH - bottomH };

  const dates = !!data.dates;
  const all = [...starts, ...ends, ...milestones, ...(today != null ? [today] : [])].filter(
    (v): v is number => v != null,
  );
  const lo = Math.min(...(all.length ? all : [0]));
  const hi = Math.max(...(all.length ? all : [1]));
  // Calendar granularity by span: weeks → months → quarters.
  const weeks = dates && hi - lo <= 130;
  const quarters = dates && hi - lo > 550;
  const ticks = dates
    ? weeks
      ? weekStarts(lo - 7, hi + 7)
      : monthStarts(lo - 31, hi + 31).filter((d) => !quarters || new Date(d * 86400000).getUTCMonth() % 3 === 0)
    : niceTicks(lo, hi, 6);
  const t0 = dates ? Math.min(lo, ticks[0] ?? lo) : ticks[0];
  const t1 = dates ? Math.max(hi, ticks[ticks.length - 1] ?? hi) : ticks[ticks.length - 1];
  const toX = (v: number) => plot.x + ((v - t0) / (t1 - t0 || 1)) * plot.w;
  const fmt = resolveFormat(ticks, cfg.numberFormat);
  const tickLabel = (t: number, i: number) => {
    if (!dates) return formatNumber(t, fmt);
    const d = new Date(t * 86400000);
    if (quarters) return `Q${Math.floor(d.getUTCMonth() / 3) + 1} ${String(d.getUTCFullYear()).slice(2)}`;
    if (weeks) return formatDay(t);
    return formatDay(t, i === 0 || d.getUTCMonth() === 0);
  };
  const spanLabel = (s: number, e: number) =>
    dates ? `${formatDay(s)}–${formatDay(e)}` : `${formatNumber(s, fmt)}–${formatNumber(e, fmt)}`;

  const nodes: SceneNode[] = [];
  if (cfg.title) {
    nodes.push({
      kind: "text", x: 0, y: 0, w: cfg.width, h: fs * 1.6, text: cfg.title,
      fontSize: fs * 1.2, bold: true, color: style.text, align: "left", valign: "top", name: "title",
    });
  }
  // Weekend shading in week granularity.
  if (weeks) {
    for (let d = lo - 7; d <= hi + 7; d++) {
      if (d % 7 === 2) {
        // Day ≡ 2 (mod 7) is Saturday (day 0 = Thursday); shade Sat+Sun.
        const x1 = Math.max(plot.x, toX(d));
        const x2 = Math.min(plot.x + plot.w, toX(d + 2));
        if (x2 > x1) {
          nodes.push({ kind: "rect", x: x1, y: plot.y, w: x2 - x1, h: plot.h, fill: "#f4f3f0", name: `weekend-${d}` });
        }
      }
    }
  }

  // Timeline header on top + vertical gridlines (think-cell's calendar strip).
  const minLabelGap = fs * 2.6;
  let lastLabelX = -1e9;
  ticks.forEach((t, i) => {
    const x = toX(t);
    nodes.push({ kind: "line", x1: x, y1: plot.y, x2: x, y2: plot.y + plot.h, stroke: style.gridline, strokeWidth: 0.75, name: "gridline" });
    // Thin out header labels when months are dense.
    if (x - lastLabelX >= minLabelGap) {
      nodes.push({ kind: "text", x: x - 24, y: titleH, w: 48, h: headerH, text: tickLabel(t, i), fontSize: fs * 0.9, color: style.mutedText, align: "center", valign: "middle", name: "timeline" });
      lastLabelX = x;
    }
  });

  const slotH = plot.h / Math.max(1, data.categories.length);
  const barH = Math.min(slotH * 0.55, fs * 1.4);
  const columnTop: number[] = [];

  data.categories.forEach((_, c) => {
    const cy = plot.y + slotH * (c + 0.5);
    columnTop.push(cy - barH / 2);
    // Section header rows: bold label, light band, no bar.
    if (isHeader[c]) {
      nodes.push(
        { kind: "rect", x: 0, y: cy - slotH / 2 + 1, w: cfg.width, h: slotH - 2, fill: "#f0efec", name: `section-${c}` },
        {
          kind: "text", x: 0, y: cy - fs * 0.75, w: cfg.width, h: fs * 1.5,
          text: acts[c], fontSize: fs, bold: true, color: style.text, align: "left", valign: "middle", name: `category-${c}`,
        },
      );
      return;
    }
    nodes.push({
      kind: "text", x: indents[c] * 10, y: cy - fs * 0.75, w: catW - 6 - indents[c] * 10, h: fs * 1.5,
      text: acts[c], fontSize: fs, color: style.text, align: "left", valign: "middle", name: `category-${c}`,
    });
    // Responsible + remark columns right of the timeline.
    if (hasOwners && owners[c]) {
      nodes.push({
        kind: "text", x: plot.x + plot.w + 6, y: cy - fs * 0.75, w: ownerW - 6, h: fs * 1.5,
        text: owners[c], fontSize: fs, color: style.mutedText, align: "left", valign: "middle", name: `owner-${c}`,
      });
    }
    if (hasRemarks && remarks[c]) {
      nodes.push({
        kind: "text", x: plot.x + plot.w + ownerW + 4, y: cy - fs * 0.7, w: remarkW - 4, h: fs * 1.4,
        text: remarks[c], fontSize: fs * 0.9, color: style.mutedText, align: "left", valign: "middle", name: `remark-${c}`,
      });
    }
    // Faint row separator.
    if (c > 0) {
      nodes.push({ kind: "line", x1: plot.x, y1: cy - slotH / 2, x2: plot.x + plot.w, y2: cy - slotH / 2, stroke: style.gridline, strokeWidth: 0.5, name: `row-${c}` });
    }
    const s = starts[c];
    const e = ends[c];
    if (s != null && e != null && e > s) {
      nodes.push({
        kind: "rect", x: toX(s), y: cy - barH / 2, w: toX(e) - toX(s), h: barH,
        fill: seriesColor(style, 0), name: `bar-${c}`,
      });
      if (decor.segmentLabels) {
        const label = spanLabel(s, e);
        if (toX(e) - toX(s) >= textWidth(label, fs * 0.9) + 4) {
          nodes.push({
            kind: "text", x: toX(s), y: cy - fs * 0.7, w: toX(e) - toX(s), h: fs * 1.4,
            text: label, fontSize: fs * 0.9, color: "#ffffff", align: "center", valign: "middle", name: `bar-label-${c}`,
          });
        }
      }
    }
    const m = milestones[c];
    if (m != null) {
      const r = barH * 0.45;
      // Diamond milestone marker.
      nodes.push({
        kind: "ellipse", cx: toX(m), cy, rx: r, ry: r,
        fill: style.text, name: `milestone-${c}`,
      });
    }
  });

  // Dependency arrows ("After" row): elbow from the predecessor's end down
  // to the successor's start.
  data.categories.forEach((_, c) => {
    const pred = after[c];
    if (pred == null) return;
    const p = Math.round(pred) - 1;
    if (p < 0 || p >= data.categories.length || p === c) return;
    const predEnd = ends[p];
    const succStart = starts[c];
    if (predEnd == null || succStart == null) return;
    const x1 = toX(predEnd);
    const yPred = plot.y + slotH * (p + 0.5);
    const ySucc = plot.y + slotH * (c + 0.5);
    const x2 = toX(succStart);
    nodes.push(
      { kind: "line", x1, y1: yPred, x2: x1, y2: ySucc, stroke: style.mutedText, strokeWidth: 1, name: `dep-v-${c}` },
      { kind: "line", x1, y1: ySucc, x2: x2 - 2, y2: ySucc, stroke: style.mutedText, strokeWidth: 1, name: `dep-h-${c}` },
      { kind: "arrowhead", x: x2 - 1, y: ySucc, angle: x2 >= x1 ? 0 : 180, size: 3.5, fill: style.mutedText, name: `dep-head-${c}` },
    );
  });

  // Today line.
  if (today != null && today >= t0 && today <= t1) {
    const x = toX(today);
    nodes.push(
      { kind: "line", x1: x, y1: plot.y, x2: x, y2: plot.y + plot.h, stroke: style.negative, strokeWidth: 1.25, dash: [3, 2], name: "today-line" },
      {
        kind: "text", x: x - 24, y: plot.y + plot.h + 1, w: 48, h: fs * 1.3,
        text: "Today", fontSize: fs * 0.85, color: style.negative, align: "center", valign: "top", name: "today-label",
      },
    );
  }

  return {
    nodes,
    anchors: {
      categoryX: data.categories.map((_, c) => plot.y + slotH * (c + 0.5)),
      categoryWidth: data.categories.map(() => barH),
      columnTop,
      columnValue: data.categories.map((_, c) => ends[c] ?? 0),
      baselineY: plot.x,
      plot,
    },
  };
}
