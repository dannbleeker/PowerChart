import type { ChartConfig, ChartStyle, Decorations } from "../types";
import { contrastInk, textWidth, type SceneNode } from "../scene";
import { formatNumber, formatPercent, resolveFormat } from "../format";
import { footnoteH, titleHeight, titleNode } from "./frame";
import type { LayoutResult } from "./column";

/**
 * Cascade / decomposition chart: each stage's bar is a subset of the
 * previous one, read left to right (Total → Answered → With a case →
 * Solved…). Bars are top-aligned on one volume scale; the complement of
 * each split ("Dropped", "Without a case") hangs as a muted box at the
 * split point, labeled with value and % of the previous stage.
 *
 * Category syntax: "Stage | Drop label | Group header" — part 2 captions
 * the remainder box of the split INTO this stage; consecutive stages
 * sharing part 3 get one spanning header band.
 */
export function layoutCascade(cfg: ChartConfig, style: ChartStyle, decor: Decorations): LayoutResult {
  const { data } = cfg;
  const fs = style.fontSize;
  const parts = data.categories.map((c) => c.split("|").map((p) => p.trim()));
  const stages = parts.map((p) => p[0] ?? "");
  const dropLabels = parts.map((p) => p[1] ?? "");
  const groups = parts.map((p) => p[2] ?? "");
  const n = stages.length;
  const values = stages.map((_, c) => Math.max(0, data.series[0]?.values[c] ?? 0));
  // Valid cascades decrease, so the first stage IS the max; scaling by the
  // max keeps malformed (growing) data inside the plot instead of overflowing.
  const v0 = Math.max(...values, 1);
  const fmt = resolveFormat(values, cfg.numberFormat);

  const titleH = titleHeight(cfg, style);
  const hasGroups = groups.some(Boolean);
  const groupH = hasGroups ? fs * 1.7 : 0;
  const plot = {
    x: 2,
    y: titleH + groupH + (hasGroups ? 4 : 0),
    w: cfg.width - 4,
    h: cfg.height - titleH - groupH - (hasGroups ? 4 : 0) - footnoteH(cfg, style, decor) - 4,
  };
  const slotW = plot.w / Math.max(1, n);
  const barW = slotW * 0.91;
  const toH = (v: number) => (v / v0) * plot.h;

  const nodes: SceneNode[] = [];
  const titleN = titleNode(cfg, style);
  if (titleN) nodes.push(titleN);

  // Spanning group header bands over consecutive same-group stages.
  if (hasGroups) {
    let start = 0;
    for (let c = 1; c <= n; c++) {
      if (c === n || groups[c] !== groups[start]) {
        if (groups[start]) {
          const x1 = plot.x + slotW * start + (slotW - barW) / 2;
          const x2 = plot.x + slotW * (c - 1) + (slotW + barW) / 2;
          nodes.push(
            { kind: "rect", x: x1, y: titleH, w: x2 - x1, h: fs * 1.5, fill: style.neutral, name: `group-${start}` },
            {
              kind: "text", x: x1, y: titleH, w: x2 - x1, h: fs * 1.5, text: groups[start],
              fontSize: fs, bold: true, color: contrastInk(style.neutral), align: "center", valign: "middle",
              name: `group-label-${start}`,
            },
          );
        }
        start = c;
      }
    }
  }

  const columnTop: number[] = [];
  const centers: number[] = [];
  values.forEach((v, c) => {
    const x = plot.x + slotW * c + (slotW - barW) / 2;
    const center = x + barW / 2;
    centers.push(center);
    columnTop.push(plot.y);
    const h = Math.max(2, toH(v));
    const fill = data.series[0]?.colors?.[c] ?? style.palette[c % style.palette.length];
    const ink = contrastInk(fill);
    nodes.push({ kind: "rect", x, y: plot.y, w: barW, h, fill, name: `stage-${c}` });

    // In-bar text: stage name near the top, value + % of previous centered.
    const pct = c > 0 && values[c - 1] > 0 ? values[c] / values[c - 1] : null;
    const lines = [
      { text: stages[c], y: plot.y + h * 0.18, bold: false, size: fs },
      { text: formatNumber(v, fmt), y: plot.y + h * 0.5 - fs * 0.75, bold: true, size: fs * 1.05 },
      ...(pct != null ? [{ text: `(${formatPercent(pct, 1)})`, y: plot.y + h * 0.5 + fs * 0.7, bold: false, size: fs }] : []),
    ];
    for (const [i, line] of lines.entries()) {
      if (h < fs * (2.2 + i * 1.4)) break; // bar too short for more lines
      nodes.push({
        kind: "text", x: x + 2, y: line.y, w: barW - 4, h: fs * 1.4, text: line.text,
        fontSize: line.size, bold: line.bold, color: ink, align: "center", valign: "middle",
        name: `stage-label-${c}-${i}`,
      });
    }

    // Remainder box: what the previous stage lost at this split.
    if (c > 0) {
      const rem = Math.max(0, values[c - 1] - v);
      if (rem > 0) {
        // The column is ONE bar split in two: the colored segment above is
        // what continues, this gray segment is what stops here. They are
        // flush, share the bar width, and their heights are exact — so the
        // block's span is identical to the previous column's continuing
        // segment, and a column can never outgrow what feeds it.
        const segY = plot.y + h;
        const segH = toH(rem);
        const remPct = values[c - 1] > 0 ? rem / values[c - 1] : null;
        const caption = dropLabels[c] || "Other";
        const numbers = `${formatNumber(rem, fmt)}${remPct != null ? ` (${formatPercent(remPct, 1)})` : ""}`;
        const oneLine = `${caption}: ${numbers}`;
        const ink = contrastInk(style.neutral);
        nodes.push({
          kind: "rect", x, y: segY, w: barW, h: segH, fill: style.neutral,
          stroke: style.background, strokeWidth: 0.75, name: `drop-${c}`,
        });
        // Labels adapt to the segment — never the other way around.
        const fitsOneLine = textWidth(oneLine, fs * 0.9) <= barW - 6;
        const outside = (text: string, name: string): SceneNode => ({
          kind: "text", x: x - slotW * 0.045, y: segY + segH + 1, w: barW + slotW * 0.09, h: fs * 1.2,
          text, fontSize: fs * 0.85, color: style.text, align: "center", valign: "top", name,
        });
        if (segH >= fs * 2.9 && !fitsOneLine) {
          // Tall enough for two lines: caption over numbers, inside.
          nodes.push(
            {
              kind: "text", x: x + 2, y: segY + segH / 2 - fs * 1.35, w: barW - 4, h: fs * 1.4, text: caption,
              fontSize: fs * 0.9, color: ink, align: "center", valign: "middle", name: `drop-label-${c}`,
            },
            {
              kind: "text", x: x + 2, y: segY + segH / 2, w: barW - 4, h: fs * 1.4, text: numbers,
              fontSize: fs * 0.9, color: ink, align: "center", valign: "middle", name: `drop-value-${c}`,
            },
          );
        } else if (segH >= fs * 1.3 && fitsOneLine) {
          // One comfortable line, inside.
          nodes.push({
            kind: "text", x: x + 2, y: segY, w: barW - 4, h: segH, text: oneLine,
            fontSize: fs * 0.9, color: ink, align: "center", valign: "middle", name: `drop-label-${c}`,
          });
        } else if (segH >= fs * 1.3) {
          // Room for one line but the caption is long: numbers inside,
          // caption just below the block.
          nodes.push(
            {
              kind: "text", x: x + 2, y: segY, w: barW - 4, h: segH, text: numbers,
              fontSize: fs * 0.9, color: ink, align: "center", valign: "middle", name: `drop-value-${c}`,
            },
            outside(caption, `drop-label-${c}`),
          );
        } else {
          // Segment too thin for any text: full label below the block.
          nodes.push(outside(oneLine, `drop-label-${c}`));
        }
      }
    }
  });

  return {
    nodes,
    anchors: {
      categoryX: centers,
      categoryWidth: values.map(() => barW),
      columnTop,
      columnValue: values,
      baselineY: plot.y + plot.h,
      plot,
    },
  };
}
