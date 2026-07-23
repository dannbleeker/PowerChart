/**
 * The agenda-chapters control and its live SVG preview. Like the element
 * widgets, this reads its own textarea and paints its own preview with no tie to
 * the chart-editing state — split out of app.ts to shrink the pane controller.
 * The actual "insert agenda slides" host action stays in app.ts (it needs the
 * host renderer); `agendaChapters` is exported so that handler can reuse it.
 */
import { buildAgendaScene } from "../core/agenda";
import { sceneToSvg } from "../render/svg";

const $ = <T extends HTMLElement>(id: string) => document.getElementById(id) as T;

/** The non-empty, trimmed chapter lines from the agenda textarea. */
export function agendaChapters(): string[] {
  return ($("agenda-chapters") as HTMLTextAreaElement).value
    .split("\n")
    .map((l) => l.trim())
    .filter(Boolean);
}

/** Repaint the agenda preview (the first chapter highlighted), or clear it. */
export function renderAgendaPreview(): void {
  const chapters = agendaChapters();
  const host = $("agenda-preview");
  host.innerHTML = chapters.length
    ? sceneToSvg(buildAgendaScene(chapters, { highlight: 0 }), { background: "#ffffff" })
    : "";
}

/** Wire the live preview to the chapters textarea and paint the initial state. */
export function wireAgendaPreview(): void {
  $("agenda-chapters").addEventListener("input", renderAgendaPreview);
  renderAgendaPreview();
}
