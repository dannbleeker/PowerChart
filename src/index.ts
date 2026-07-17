/**
 * PowerChart library entry: the pure chart engine, consumable outside the
 * add-in (batch SVG rendering, .ppttc-style automation pipelines, tests).
 */
export { buildChart, valueExtent, DEFAULT_SIZE } from "./core/chart";
export { sceneToSvg } from "./render/svg";
export { buildAgendaScene, SLIDE } from "./core/agenda";
export { buildHarveyBall, buildCheckbox, buildProcessFlow, buildTableScene, buildKpiTile } from "./core/elements";
export { sampleConfig, CHART_KINDS } from "./core/samples";
export { PALETTE, PALETTES, DEFAULT_STYLE } from "./core/style";
export { formatNumber, formatPercent, parseDateToken, cagr } from "./core/format";
// Shared renderer geometry — the headless pptx renderer (skill/scripts/render-pptx.mjs)
// consumes these from the built lib instead of re-deriving them.
export { arrowheadBox, sceneToOoxmlPieAngle, annularSectorPoints, symbolPoints, SYMBOL_PRESET } from "./core/geometry";
export type { ChartConfig, ChartData, ChartKind, Decorations, Series } from "./core/types";
export type { Scene, SceneNode } from "./core/scene";
