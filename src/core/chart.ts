import type { ChartConfig, ChartStyle, Decorations } from "./types";
import type { Scene } from "./scene";
import { DEFAULT_DECOR, DEFAULT_STYLE } from "./style";
import { layoutColumns } from "./layout/column";
import { layoutWaterfall } from "./layout/waterfall";
import { layoutMekko } from "./layout/mekko";
import { layoutLine } from "./layout/line";
import { layoutButterfly } from "./layout/butterfly";
import { decorationNodes } from "./decor";
import type { LayoutResult } from "./layout/column";

export const DEFAULT_SIZE = { width: 480, height: 300 };

/** Build a renderer-agnostic scene from a chart config. Pure and synchronous. */
export function buildChart(cfg: ChartConfig): Scene {
  const style: ChartStyle = { ...DEFAULT_STYLE, ...cfg.style };
  const decor: Decorations = { ...DEFAULT_DECOR, ...cfg.decorations };

  let result: LayoutResult;
  switch (cfg.kind) {
    case "waterfall":
      result = layoutWaterfall(cfg, style, decor);
      break;
    case "mekko":
      result = layoutMekko(cfg, style, decor);
      break;
    case "line":
    case "area":
      result = layoutLine(cfg, style, decor);
      break;
    case "butterfly":
      result = layoutButterfly(cfg, style, decor);
      break;
    default:
      result = layoutColumns(cfg, style, decor);
  }

  // Decorations assume a vertical value axis; skip them for bar orientation
  // and butterfly charts.
  const skipDecor = cfg.horizontal || cfg.kind === "butterfly";
  const nodes = skipDecor
    ? result.nodes
    : [...result.nodes, ...decorationNodes(cfg, style, decor, result.anchors)];
  return { width: cfg.width, height: cfg.height, nodes };
}

export { layoutColumns, layoutWaterfall, layoutMekko, layoutLine };
