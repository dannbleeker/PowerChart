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

/** Simple table: first row = header (accent fill), zebra body rows. */
export function buildTableScene(cells: string[][], width = 480): Scene {
  const rows = cells.length;
  const cols = Math.max(1, ...cells.map((r) => r.length));
  const fs = 10;
  const rowH = fs * 2.1;
  const height = rows * rowH;
  // Column widths proportional to their longest content.
  const widths = Array.from({ length: cols }, (_, c) =>
    Math.max(fs * 3, ...cells.map((r) => textWidth(r[c] ?? "", fs) + 12)),
  );
  const totalW = widths.reduce((a, b) => a + b, 0) || 1;
  const scale = width / totalW;

  const nodes: SceneNode[] = [];
  let y = 0;
  cells.forEach((row, ri) => {
    const header = ri === 0;
    const fill = header ? PALETTE[0] : ri % 2 === 0 ? "#f4f3f0" : "#ffffff";
    let x = 0;
    for (let c = 0; c < cols; c++) {
      const w = widths[c] * scale;
      nodes.push(
        { kind: "rect", x, y, w, h: rowH, fill, stroke: "#e1e0d9", strokeWidth: 0.75, name: `cell-${ri}-${c}` },
        {
          kind: "text",
          x: x + 5,
          y,
          w: w - 10,
          h: rowH,
          text: row[c] ?? "",
          fontSize: fs,
          bold: header,
          color: header ? contrastInk(fill) : S.text,
          align: c === 0 ? "left" : "right",
          valign: "middle",
          name: `cell-text-${ri}-${c}`,
        },
      );
      x += w;
    }
    y += rowH;
  });
  return { width, height, nodes };
}
