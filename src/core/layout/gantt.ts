import type { ChartConfig, ChartStyle, Decorations } from "../types";
import { textWidth, type SceneNode } from "../scene";
import { formatDay, formatNumber, monthStarts, niceTicks, resolveFormat, weekStarts } from "../format";
import { seriesColor } from "../style";
import type { LayoutResult } from "./column";
import { titleHeight, titleNode } from "./frame";

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
  /** "Holiday(s)" row: dates shaded like weekends. */
  const holidays = (find(/^holidays?$/i)?.values ?? []).filter((v): v is number => v != null);
  // Progress (0-100 or 0-1) and plan-vs-actual baseline rows.
  const completes = find(/^%?\s*complete$/i)?.values ?? [];
  const baseStarts = find(/^baseline\s*start$/i)?.values ?? [];
  const baseEnds = find(/^baseline\s*end$/i)?.values ?? [];
  /** "Bracket <label>" rows: first/last non-null values span an annotation. */
  const brackets = data.series
    .filter((s) => /^bracket\b/i.test(s.name.trim()))
    .map((s) => {
      const vals = s.values.filter((v): v is number => v != null);
      return {
        label: s.name.replace(/^bracket\s*:?\s*/i, "").trim(),
        from: Math.min(...vals),
        to: Math.max(...vals),
        ok: vals.length >= 2,
      };
    })
    .filter((b) => b.ok);
  // "Column <label>" rows: a numeric gutter column beside the task labels, the
  // MS-Project table look. Values are per-category data, so this is a datasheet
  // row like Start/End rather than config. Each column resolves its own format,
  // so a Cost column and an FTE column keep their own precision — the chart-wide
  // format is resolved over the timeline's epoch-day ticks and means nothing
  // here. suffix/forceSign are deliberately not picked up: they are value-axis
  // concerns, and units belong in the label ("Column Cost €k").
  const columns = data.series
    .filter((s) => /^column\b/i.test(s.name.trim()))
    .map((s) => {
      const nums = s.values.filter((v): v is number => v != null);
      const fmt = resolveFormat(nums, {
        decimals: cfg.numberFormat?.decimals,
        locale: cfg.numberFormat?.locale,
      });
      const label = s.name.trim().replace(/^column\s*:?\s*/i, "").trim();
      const cells = s.values.map((v) => (v == null ? "" : formatNumber(v, fmt)));
      const w = Math.min(
        cfg.width * 0.12,
        Math.max(textWidth(label, fs * 0.9), ...cells.map((t) => textWidth(t, fs))) + 10,
      );
      return { label, cells, w };
    });
  const colsW = columns.reduce((a, c) => a + c.w, 0);

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

  // Critical path: over the "After" dependency edges, find the chain with the
  // greatest cumulative duration and flag its activities + connecting arrows.
  const predOf = (c: number) => {
    const pred = after[c];
    if (pred == null) return -1;
    const p = Math.round(pred) - 1;
    return p >= 0 && p < data.categories.length && p !== c ? p : -1;
  };
  const critical = new Set<number>();
  if (decor.criticalPath && after.some((v) => v != null)) {
    const dur = (c: number) =>
      starts[c] != null && ends[c] != null ? Math.max(0, ends[c]! - starts[c]!) : 0;
    // Longest cumulative duration ending at each activity (memoized; cycle-safe).
    const cum: number[] = data.categories.map(() => -1);
    const seen = new Set<number>();
    const longest = (c: number): number => {
      if (cum[c] >= 0) return cum[c];
      if (seen.has(c)) return dur(c); // break any accidental cycle
      seen.add(c);
      const p = predOf(c);
      const v = dur(c) + (p >= 0 ? longest(p) : 0);
      seen.delete(c);
      return (cum[c] = v);
    };
    let end = -1;
    let best = -1;
    data.categories.forEach((_, c) => {
      if (!isHeader[c] && longest(c) > best) {
        best = longest(c);
        end = c;
      }
    });
    for (let c = end; c >= 0; c = predOf(c)) {
      if (critical.has(c)) break;
      critical.add(c);
    }
  }

  const titleH = titleHeight(cfg, style);
  const bracketH = brackets.length ? fs * 1.9 : 0;
  const headerH = fs * 1.6;
  const catW = Math.min(
    cfg.width * 0.32,
    Math.max(0, ...acts.map((c) => textWidth(c, fs))) + 10,
  );
  const ownerW = hasOwners ? Math.max(0, ...owners.map((o) => textWidth(o, fs))) + 12 : 0;
  const remarkW = hasRemarks ? Math.max(0, ...remarks.map((o) => textWidth(o, fs * 0.9))) + 12 : 0;
  const bottomH = today != null ? fs * 1.6 : 6;
  const plot = {
    x: catW + colsW,
    y: titleH + bracketH + headerH,
    w: cfg.width - catW - colsW - ownerW - remarkW - 6,
    h: cfg.height - titleH - bracketH - headerH - bottomH,
  };

  const dates = !!data.dates;
  const all = [
    ...starts,
    ...ends,
    ...milestones,
    ...(today != null ? [today] : []),
    ...holidays,
    ...brackets.flatMap((b) => [b.from, b.to]),
  ].filter((v): v is number => v != null);
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

  // Working-day timeline: give non-working days zero width so a bar's length
  // reads as the working days it contains. Every x in this file goes through
  // toX, so bars, milestones, dependencies, brackets, summary bars, baselines,
  // progress fills, the today line and the gridlines all follow from here.
  // Calendar only — "working day" means nothing on a numeric timeline.
  const wdCfg = cfg.gantt?.workdays;
  const workSet =
    wdCfg === true
      ? new Set([1, 2, 3, 4, 5])
      : Array.isArray(wdCfg)
        ? new Set(wdCfg.map((n) => Math.round(n)))
        : null;
  // ~55 years. A mistyped date must not allocate a giant array.
  const SPAN_CAP = 20000;
  const workdays = dates && !!workSet && workSet.size > 0 && t1 - t0 <= SPAN_CAP;
  // ISO weekday, day 0 = Thursday. The (d%7+7)%7 fold also handles pre-1970
  // days, which the weekend-shading loop below never did.
  const isoDow = (d: number) => (((((d % 7) + 7) % 7) + 3) % 7) + 1;
  const offDays = new Set(holidays.map((h) => Math.round(h)));
  const pre: number[] = [];
  if (workdays) {
    let n = 0;
    for (let d = Math.floor(t0); d <= Math.ceil(t1); d++) {
      // Working days STRICTLY BEFORE d, so a span's width is the working days
      // it contains: Mon→Fri stays 4 units (as on the elapsed scale, so the
      // "End is an instant" convention is unchanged), Mon→Mon becomes 5, not 7.
      pre.push(n);
      if (workSet!.has(isoDow(d)) && !offDays.has(d)) n++;
    }
  }
  // The last prefix entry, not the running counter: it is by construction the
  // working-day count of [t0, t1), which is what makes toX(t1) land exactly on
  // the right edge, as the linear branch does.
  const workTotal = pre.length ? pre[pre.length - 1] : 0;
  const workIndex = (v: number) =>
    pre[Math.max(0, Math.min(pre.length - 1, Math.round(v) - Math.floor(t0)))];
  const toX = (v: number) =>
    workdays
      ? plot.x + (workIndex(v) / Math.max(1, workTotal)) * plot.w
      : plot.x + ((v - t0) / (t1 - t0 || 1)) * plot.w;
  /**
   * Minimum width for a span that collapses to nothing. Only a working-day
   * scale can do that (a weekend-only task), so it stays 0 otherwise and no
   * existing output moves. A constant, so the result stays deterministic.
   */
  const minW = workdays ? 1.5 : 0;
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
  const titleN = titleNode(cfg, style);
  if (titleN) nodes.push(titleN);
  // Bracket annotations above the timeline header.
  brackets.forEach((b, i) => {
    const x1 = toX(b.from);
    const x2 = toX(b.to);
    const y = titleH + fs * 1.5;
    nodes.push(
      { kind: "line", x1, y1: y, x2, y2: y, stroke: style.text, strokeWidth: 1, name: `bracket-${i}` },
      { kind: "line", x1, y1: y, x2: x1, y2: y + 3.5, stroke: style.text, strokeWidth: 1, name: `bracket-tick-a-${i}` },
      { kind: "line", x1: x2, y1: y, x2, y2: y + 3.5, stroke: style.text, strokeWidth: 1, name: `bracket-tick-b-${i}` },
      {
        kind: "text", x: x1, y: y - fs * 1.35, w: x2 - x1, h: fs * 1.3,
        text: b.label || spanLabel(b.from, b.to), fontSize: fs * 0.9, bold: true,
        color: style.text, align: "center", valign: "middle", name: `bracket-label-${i}`,
      },
    );
  });

  // Holiday shading (any granularity). Pointless under a working-day scale:
  // a holiday has no width left to shade.
  for (const h of workdays ? [] : holidays) {
    const x1 = Math.max(plot.x, toX(h));
    const x2 = Math.min(plot.x + plot.w, toX(h + 1));
    if (x2 > x1) {
      nodes.push({ kind: "rect", x: x1, y: plot.y, w: x2 - x1, h: plot.h, fill: "#efe7e7", name: `holiday-${h}` });
    }
  }

  // Weekend shading in week granularity. Gated on !workdays explicitly rather
  // than left to the x2 > x1 guard below: that only collapses for a Mon–Fri
  // week. Under a custom workweek (say Sun–Thu) Saturday has width again, so
  // the guard passes and the block shades a Sunday — a working day there —
  // while leaving the real non-working Friday unshaded.
  if (weeks && !workdays) {
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
      nodes.push({ kind: "text", x: x - 24, y: plot.y - headerH, w: 48, h: headerH, text: tickLabel(t, i), fontSize: fs * 0.9, color: style.mutedText, align: "center", valign: "middle", name: "timeline" });
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
      // Auto-summary bar: span min(start)→max(end) of the child activities
      // (the rows below this header up to the next header), with end caps.
      if (decor.summaryBars) {
        let s = Infinity;
        let e = -Infinity;
        for (let k = c + 1; k < data.categories.length && !isHeader[k]; k++) {
          for (const v of [starts[k], ends[k], milestones[k]]) {
            if (v != null) {
              s = Math.min(s, v);
              e = Math.max(e, v);
            }
          }
        }
        if (e > s) {
          const x1 = toX(s);
          const x2 = toX(e);
          const sbH = barH * 0.4;
          nodes.push(
            { kind: "rect", x: x1, y: cy - sbH / 2, w: Math.max(x2 - x1, minW), h: sbH, fill: style.text, name: `summary-${c}` },
            { kind: "line", x1, y1: cy - sbH / 2, x2: x1, y2: cy + sbH * 1.4, stroke: style.text, strokeWidth: 1.25, name: `summary-cap-a-${c}` },
            { kind: "line", x1: x2, y1: cy - sbH / 2, x2, y2: cy + sbH * 1.4, stroke: style.text, strokeWidth: 1.25, name: `summary-cap-b-${c}` },
          );
        }
      }
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
    // Baseline ghost bar (the original plan), thin, beneath the actual bar.
    const bs = baseStarts[c];
    const be = baseEnds[c];
    if (bs != null && be != null && be > bs) {
      nodes.push({
        kind: "rect", x: toX(bs), y: cy + barH * 0.55, w: Math.max(toX(be) - toX(bs), minW), h: barH * 0.4,
        fill: "#cfcdc5", name: `gantt-baseline-${c}`,
      });
    }
    if (s != null && e != null && e > s) {
      const isCrit = critical.has(c);
      const bx = toX(s);
      // A task living entirely inside non-working days has zero working length
      // and would vanish. Keep a hairline so it is still visible and clickable.
      const bw = Math.max(toX(e) - bx, minW);
      nodes.push({
        kind: "rect", x: bx, y: cy - barH / 2, w: bw, h: barH,
        fill: seriesColor(style, 0), name: `bar-${c}`,
        ...(isCrit ? { stroke: style.negative, strokeWidth: 1.75 } : {}),
      });
      // Percent-complete fill: a darker inner bar over the elapsed share.
      const rawPct = completes[c];
      if (rawPct != null) {
        const pct = Math.max(0, Math.min(1, rawPct > 1 ? rawPct / 100 : rawPct));
        if (pct > 0) {
          nodes.push({
            kind: "rect", x: bx, y: cy - barH / 2, w: bw * pct, h: barH,
            fill: "#1b4e8a", name: `progress-${c}`,
          });
        }
      }
      if (decor.segmentLabels) {
        const label = spanLabel(s, e);
        if (bw >= textWidth(label, fs * 0.9) + 4) {
          nodes.push({
            kind: "text", x: bx, y: cy - fs * 0.7, w: bw, h: fs * 1.4,
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

  // Gutter columns, in their own pass. This has to come after the row loop:
  // the section-header band is a full-width rect (`section-*`, x: 0,
  // w: cfg.width) and would paint straight over cells emitted alongside it.
  // A header row shows whatever value it carries — nothing is auto-summed here
  // (summaryBars sums spans, not money).
  columns.forEach((col, i) => {
    let cx = catW;
    for (let k = 0; k < i; k++) cx += columns[k].w;
    nodes.push({
      kind: "text",
      x: cx,
      y: plot.y - headerH,
      w: col.w - 6,
      h: headerH,
      text: col.label,
      fontSize: fs * 0.9,
      bold: true,
      color: style.mutedText,
      align: "right",
      valign: "middle",
      name: `col-head-${i}`,
    });
    data.categories.forEach((_, c) => {
      if (!col.cells[c]) return;
      const cy = plot.y + slotH * (c + 0.5);
      nodes.push({
        kind: "text",
        x: cx,
        y: cy - fs * 0.75,
        w: col.w - 6,
        h: fs * 1.5,
        text: col.cells[c],
        fontSize: fs,
        color: isHeader[c] ? style.text : style.mutedText,
        bold: isHeader[c],
        align: "right",
        valign: "middle",
        name: `col-${i}-${c}`,
      });
    });
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
    // Edges on the critical path are drawn thicker and red.
    const critEdge = critical.has(c) && critical.has(p);
    const dcolor = critEdge ? style.negative : style.mutedText;
    const dw = critEdge ? 1.75 : 1;
    nodes.push(
      { kind: "line", x1, y1: yPred, x2: x1, y2: ySucc, stroke: dcolor, strokeWidth: dw, name: `dep-v-${c}` },
      { kind: "line", x1, y1: ySucc, x2: x2 - 2, y2: ySucc, stroke: dcolor, strokeWidth: dw, name: `dep-h-${c}` },
      { kind: "arrowhead", x: x2 - 1, y: ySucc, angle: x2 >= x1 ? 0 : 180, size: critEdge ? 4.2 : 3.5, fill: dcolor, name: `dep-head-${c}` },
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
