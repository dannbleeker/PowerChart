import type { ChartConfig, ChartStyle, Decorations } from "../types";
import { type SceneNode } from "../scene";
import { resolveFormat } from "../format";
import { chromeNodes, computeFrame, valueScale } from "./frame";
import type { LayoutResult } from "./column";

/**
 * Candlestick / OHLC chart. Datasheet rows named Open / High / Low / Close
 * give each period's prices; categories are the periods. A thin high–low wick
 * carries a body from open to close, coloured green when the period rose and
 * red when it fell.
 */
export function layoutCandlestick(cfg: ChartConfig, style: ChartStyle, decor: Decorations): LayoutResult {
  const { data } = cfg;
  const n = data.categories.length;
  const find = (re: RegExp) => data.series.find((s) => re.test(s.name.trim()))?.values ?? [];
  const open = find(/^open$/i);
  const high = find(/^high$/i);
  const low = find(/^low$/i);
  const close = find(/^close$/i);

  const all = [...open, ...high, ...low, ...close].filter((v): v is number => v != null);
  const fmt = resolveFormat(all, cfg.numberFormat);
  const { frame } = computeFrame(cfg, style, { ...decor, seriesLabels: false }, []);
  const scale = valueScale(
    frame,
    Math.min(...(all.length ? all : [0])),
    Math.max(...(all.length ? all : [1])),
    cfg.scale,
  );
  const slotLen = frame.w / Math.max(1, n);
  const colThick = Math.min(slotLen * 0.5, 24);
  const centers = data.categories.map((_, c) => frame.x + slotLen * (c + 0.5));

  const nodes: SceneNode[] = chromeNodes(cfg, style, { ...decor, seriesLabels: false }, frame, centers, scale);
  const up = "#1a9e6e";
  const down = style.negative;
  void fmt;

  for (let c = 0; c < n; c++) {
    const hi = high[c];
    const lo = low[c];
    const x = centers[c];
    if (hi != null && lo != null) {
      nodes.push({
        kind: "line",
        x1: x,
        y1: scale.toY(hi),
        x2: x,
        y2: scale.toY(lo),
        stroke: style.text,
        strokeWidth: 1,
        name: `wick-${c}`,
      });
    }
    const o = open[c];
    const cl = close[c];
    if (o != null && cl != null) {
      const rising = cl >= o;
      const yTop = scale.toY(Math.max(o, cl));
      const yBot = scale.toY(Math.min(o, cl));
      nodes.push({
        kind: "rect",
        x: x - colThick / 2,
        y: yTop,
        w: colThick,
        h: Math.max(1, yBot - yTop),
        fill: rising ? up : down,
        stroke: style.background,
        strokeWidth: 0.5,
        name: `body-${c}`,
      });
    }
  }

  return {
    nodes,
    anchors: {
      categoryX: centers,
      categoryWidth: data.categories.map(() => colThick),
      columnTop: data.categories.map((_, c) => scale.toY(high[c] ?? 0)),
      columnValue: data.categories.map((_, c) => close[c] ?? 0),
      baselineY: frame.y + frame.h,
      plot: { x: frame.x, y: frame.y, w: frame.w, h: frame.h },
      valueToY: scale.toY,
    },
  };
}
