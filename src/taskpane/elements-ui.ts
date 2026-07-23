/**
 * The "Elements" panel: harvey balls, checkboxes, process flow, and KPI tiles.
 * A self-contained widget group — it reads its own controls and paints its own
 * SVG previews, with no tie to the chart-editing state (`state`, the datasheet,
 * insert bookkeeping). Split out of app.ts so the pane controller shrinks and
 * this corner is unit-testable on its own. The table element stays in app.ts:
 * it builds from the shared datasheet, unlike these four.
 */
import { buildHarveyBall, buildCheckbox, buildProcessFlow, buildKpiTile, type CheckState } from "../core/elements";
import { sceneToSvg } from "../render/svg";
import type { Scene } from "../core/scene";

const $ = <T extends HTMLElement>(id: string) => document.getElementById(id) as T;

export function harveyScene(): Scene {
  return buildHarveyBall(Number(($("harvey-pct") as HTMLInputElement).value) / 100, 24);
}

export function checkScene(): Scene {
  return buildCheckbox(($("check-state") as HTMLSelectElement).value as CheckState, 20);
}

export function flowScene(): Scene {
  const steps = ($("flow-steps") as HTMLInputElement).value
    .split(",")
    .map((s) => s.trim())
    .filter(Boolean);
  const hl = Number(($("flow-highlight") as HTMLInputElement).value) - 1;
  return buildProcessFlow(steps, hl, 480, 40);
}

export function kpiScene(): Scene {
  return buildKpiTile({
    label: ($("kpi-label") as HTMLInputElement).value,
    value: ($("kpi-value") as HTMLInputElement).value,
    delta: ($("kpi-delta") as HTMLInputElement).value || undefined,
    goodIsUp: !($("kpi-down-good") as HTMLInputElement).checked,
  });
}

/** Repaint all four element previews from the current control values. */
export function renderElementPreviews(): void {
  $("harvey-val").textContent = `${($("harvey-pct") as HTMLInputElement).value}%`;
  $("harvey-preview").innerHTML = sceneToSvg(harveyScene());
  $("check-preview").innerHTML = sceneToSvg(checkScene());
  $("flow-preview").innerHTML = sceneToSvg(flowScene());
  $("kpi-preview").innerHTML = sceneToSvg(kpiScene());
}

/** Wire live previews to the element controls and paint the initial state. */
export function wireElementPreviews(): void {
  for (const id of [
    "harvey-pct",
    "check-state",
    "flow-steps",
    "flow-highlight",
    "kpi-label",
    "kpi-value",
    "kpi-delta",
    "kpi-down-good",
  ]) {
    $(id).addEventListener("input", renderElementPreviews);
  }
  renderElementPreviews();
}
