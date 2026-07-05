/**
 * think-cell's non-chart elements: harvey balls, checkboxes, process flows,
 * and a simple table — all as scenes for the same renderers as charts.
 */
import type { Scene, SceneNode } from "./scene";
import { contrastInk, textWidth } from "./scene";
import { DEFAULT_STYLE, PALETTE } from "./style";

const S = DEFAULT_STYLE;

/** Harvey ball: fraction filled clockwise from 12 o'clock (0..1). */
export function buildHarveyBall(fraction: number, size = 24): Scene {
  const f = Math.max(0, Math.min(1, fraction));
  const r = size / 2 - 1.5;
  const cx = size / 2;
  const cy = size / 2;
  const nodes: SceneNode[] = [
    { kind: "ellipse", cx, cy, rx: r, ry: r, fill: "#ffffff", stroke: S.text, strokeWidth: 1.25, name: "harvey-ring" },
  ];
  if (f >= 1) {
    nodes.push({ kind: "ellipse", cx, cy, rx: r - 1.5, ry: r - 1.5, fill: S.text, name: "harvey-fill" });
  } else if (f > 0) {
    nodes.push({
      kind: "wedge", cx, cy, r: r - 1.5, innerR: 0,
      startAngle: 0, endAngle: f * 360, fill: S.text, name: "harvey-fill",
    });
  }
  return { width: size, height: size, nodes };
}

export type CheckState = "yes" | "no" | "partial";

/** Checkbox / status mark: ✓ (good), ✗ (critical), − (neutral). */
export function buildCheckbox(state: CheckState, size = 20): Scene {
  const colors: Record<CheckState, string> = { yes: "#0ca30c", no: "#d03b3b", partial: "#898781" };
  const glyph: Record<CheckState, string> = { yes: "✓", no: "✗", partial: "–" };
  return {
    width: size,
    height: size,
    nodes: [
      { kind: "rect", x: 0.5, y: 0.5, w: size - 1, h: size - 1, fill: "#ffffff", stroke: colors[state], strokeWidth: 1.5, name: "check-box" },
      {
        kind: "text", x: 0, y: 0, w: size, h: size, text: glyph[state],
        fontSize: size * 0.62, bold: true, color: colors[state],
        align: "center", valign: "middle", name: "check-glyph",
      },
    ],
  };
}

/** Process flow: a row of chevrons with the active step highlighted. */
export function buildProcessFlow(
  steps: string[],
  highlight = -1,
  width = 480,
  height = 40,
): Scene {
  const n = Math.max(1, steps.length);
  const overlap = height * 0.28; // chevron notch overlaps the previous step
  const stepW = (width + overlap * (n - 1)) / n;
  const nodes: SceneNode[] = [];
  steps.forEach((label, i) => {
    const x = i * (stepW - overlap);
    const active = i === highlight;
    const fill = active ? PALETTE[0] : "#dcdbd2";
    nodes.push(
      { kind: "chevron", x, y: 0, w: stepW, h: height, fill, flatLeft: i === 0, name: `step-${i}` },
      {
        kind: "text",
        x: x + overlap * 0.8,
        y: 0,
        w: stepW - overlap * 1.6,
        h: height,
        text: label,
        fontSize: Math.min(11, height * 0.3),
        bold: active,
        color: contrastInk(fill),
        align: "center",
        valign: "middle",
        name: `step-label-${i}`,
      },
    );
  });
  return { width, height, nodes };
}

export interface TableOptions {
  /**
   * "rules" (default): horizontal rules only — top, under the header, and
   * bottom; never side borders, no lines between data rows. "grid": legacy
   * full grid with accent header and zebra rows.
   */
  style?: "rules" | "grid";
  /** Treat the last row as a totals row: bold, with a rule above and below. */
  totalRow?: boolean;
  /**
   * Insert a small separator gap after every N body rows so long tables stay
   * scannable (default 5; 0 disables).
   */
  groupRows?: number;
}

/** Semantic colors for in-cell effects. */
const GOOD = "#0ca30c";
const BAD = "#d03b3b";

/**
 * In-cell effects, conditional-formatting style: leading tokens turn into
 * mini harvey balls, trend arrows, or semantic font colors.
 *   "[hb:0.75] Progress"  → ◕ Progress
 *   "[up] +14%" / "[down] -3%" / "[flat] 0%"
 *   "[good] on track" / "[bad] at risk"
 */
function parseCell(raw: string): { text: string; harvey?: number; trend?: "up" | "down" | "flat"; color?: string } {
  let text = raw;
  const out: ReturnType<typeof parseCell> = { text: raw };
  let m: RegExpMatchArray | null;
  while ((m = text.match(/^\s*\[(hb:([\d.]+%?)|up|down|flat|good|bad)\]\s*/i))) {
    const tok = m[1].toLowerCase();
    if (tok.startsWith("hb:")) {
      const v = parseFloat(m[2]);
      out.harvey = Math.max(0, Math.min(1, m[2].includes("%") ? v / 100 : v));
    } else if (tok === "up" || tok === "down" || tok === "flat") out.trend = tok;
    else out.color = tok === "good" ? GOOD : BAD;
    text = text.slice(m[0].length);
  }
  out.text = text;
  return out;
}

/** Mini harvey ball + trend arrow glyph nodes for a cell, left of the text. */
function cellEffectNodes(cell: ReturnType<typeof parseCell>, x: number, cy: number, fs: number, ri: number, c: number): SceneNode[] {
  const nodes: SceneNode[] = [];
  let gx = x;
  if (cell.harvey != null) {
    const r = fs * 0.5;
    nodes.push({ kind: "ellipse", cx: gx + r, cy, rx: r, ry: r, fill: "#ffffff", stroke: S.text, strokeWidth: 1, name: `cell-hb-ring-${ri}-${c}` });
    if (cell.harvey >= 1) nodes.push({ kind: "ellipse", cx: gx + r, cy, rx: r - 1, ry: r - 1, fill: S.text, name: `cell-hb-${ri}-${c}` });
    else if (cell.harvey > 0)
      nodes.push({ kind: "wedge", cx: gx + r, cy, r: r - 1, innerR: 0, startAngle: 0, endAngle: cell.harvey * 360, fill: S.text, name: `cell-hb-${ri}-${c}` });
    gx += r * 2 + 3;
  }
  if (cell.trend) {
    const size = fs * 0.42;
    const angle = cell.trend === "up" ? -90 : cell.trend === "down" ? 90 : 0;
    const fill = cell.trend === "up" ? GOOD : cell.trend === "down" ? BAD : S.mutedText;
    nodes.push({ kind: "arrowhead", x: gx + size, y: cy, angle, size, fill, name: `cell-trend-${ri}-${c}` });
    gx += size * 2 + 3;
  }
  return nodes;
}

/** Effect glyph width reserved before the cell text. */
const effectW = (cell: ReturnType<typeof parseCell>, fs: number) =>
  (cell.harvey != null ? fs + 3 : 0) + (cell.trend ? fs * 0.84 + 3 : 0);

/**
 * Table element. Default style follows chart-design practice: rules at the
 * top, under the header, and at the bottom — never side borders or lines
 * between data rows — with a separator gap every 5 rows and an optional
 * bold totals row. Cells accept in-cell effect tokens (see parseCell).
 */
export function buildTableScene(cells: string[][], width = 480, opts: TableOptions = {}): Scene {
  const styleMode = opts.style ?? "rules";
  const groupEvery = opts.groupRows ?? 5;
  const rows = cells.length;
  const cols = Math.max(1, ...cells.map((r) => r.length));
  const fs = 10;
  const rowH = fs * 2.1;
  const gapH = rowH * 0.4;
  const parsed = cells.map((row) => row.map((c) => parseCell(c ?? "")));
  // Column widths proportional to their longest content (incl. effect glyphs).
  const widths = Array.from({ length: cols }, (_, c) =>
    Math.max(fs * 3, ...parsed.map((r) => textWidth(r[c]?.text ?? "", fs) + effectW(r[c] ?? { text: "" }, fs) + 12)),
  );
  const totalW = widths.reduce((a, b) => a + b, 0) || 1;
  const scale = width / totalW;
  const rules = styleMode === "rules";
  const rule = (y: number, weight: number, name: string): SceneNode => ({
    kind: "line", x1: 0, y1: y, x2: width, y2: y, stroke: S.text, strokeWidth: weight, name,
  });

  const nodes: SceneNode[] = [];
  let y = 0;
  cells.forEach((row, ri) => {
    const header = ri === 0;
    const total = !!opts.totalRow && ri === rows - 1;
    // Separator gap after every N body rows (rules style).
    if (rules && groupEvery > 0 && ri > 1 && !total && (ri - 1) % groupEvery === 0) y += gapH;
    if (rules && total) {
      y += gapH * 0.5;
      nodes.push(rule(y, 0.75, "rule-total"));
    }
    const fill = header ? PALETTE[0] : ri % 2 === 0 ? "#f4f3f0" : "#ffffff";
    let x = 0;
    for (let c = 0; c < cols; c++) {
      const w = widths[c] * scale;
      const cell = parsed[ri][c] ?? { text: "" };
      if (!rules) {
        nodes.push({ kind: "rect", x, y, w, h: rowH, fill, stroke: "#e1e0d9", strokeWidth: 0.75, name: `cell-${ri}-${c}` });
      }
      const ew = effectW(cell, fs);
      nodes.push(...cellEffectNodes(cell, x + 5, y + rowH / 2, fs, ri, c));
      nodes.push({
        kind: "text",
        x: x + 5 + ew,
        y,
        w: w - 10 - ew,
        h: rowH,
        text: cell.text,
        fontSize: fs,
        bold: header || total,
        color: cell.color ?? (!rules && header ? contrastInk(fill) : S.text),
        align: c === 0 ? "left" : "right",
        valign: "middle",
        name: `cell-text-${ri}-${c}`,
      });
      x += w;
    }
    y += rowH;
    if (rules && header) nodes.push(rule(y, 0.75, "rule-header"));
  });
  if (rules) {
    nodes.unshift(rule(0.5, 1.25, "rule-top"));
    nodes.push(rule(y - 0.5, 1.25, "rule-bottom"));
  }
  return { width, height: y, nodes };
}
