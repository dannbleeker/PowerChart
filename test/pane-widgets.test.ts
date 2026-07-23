// @vitest-environment jsdom
import { beforeEach, describe, expect, it } from "vitest";
import {
  harveyScene,
  checkScene,
  flowScene,
  kpiScene,
  renderElementPreviews,
  wireElementPreviews,
} from "../src/taskpane/elements-ui";
import { agendaChapters, renderAgendaPreview, wireAgendaPreview } from "../src/taskpane/agenda-ui";

/**
 * The self-contained widget modules split out of app.ts (Elements + Agenda
 * preview). They read their own controls and paint their own SVG previews, so
 * they can be driven directly here — no host, no app boot — which is exactly why
 * they were worth extracting from the 1.8k-line pane controller.
 */

const input = (id: string, value: string, type = "text") => {
  const el = document.createElement("input");
  el.id = id;
  el.type = type;
  if (type === "checkbox") el.checked = value === "true";
  else el.value = value;
  document.body.appendChild(el);
  return el;
};
const box = (id: string) => {
  const el = document.createElement("div");
  el.id = id;
  document.body.appendChild(el);
  return el;
};

describe("elements-ui", () => {
  beforeEach(() => {
    document.body.innerHTML = "";
    input("harvey-pct", "75", "range");
    box("harvey-val");
    box("harvey-preview");
    const sel = document.createElement("select");
    sel.id = "check-state";
    for (const v of ["yes", "no", "partial"]) {
      const o = document.createElement("option");
      o.value = v;
      sel.appendChild(o);
    }
    sel.value = "partial";
    document.body.appendChild(sel);
    box("check-preview");
    input("flow-steps", "Plan, Build, Ship");
    input("flow-highlight", "2");
    box("flow-preview");
    input("kpi-label", "Revenue");
    input("kpi-value", "€1.2M");
    input("kpi-delta", "+8%");
    input("kpi-down-good", "false", "checkbox");
    box("kpi-preview");
  });

  it("builds each element scene from its own controls", () => {
    expect(harveyScene().nodes.length).toBeGreaterThan(0);
    expect(checkScene().nodes.length).toBeGreaterThan(0);
    // Three comma-separated steps become at least three shapes.
    const flow = flowScene();
    expect(flow.nodes.filter((n) => n.kind === "text").some((n) => (n as { text?: string }).text === "Ship")).toBe(
      true,
    );
    // The KPI tile carries its label and value text.
    const kpiTexts = kpiScene()
      .nodes.filter((n) => n.kind === "text")
      .map((n) => (n as { text?: string }).text);
    expect(kpiTexts).toContain("Revenue");
    expect(kpiTexts).toContain("€1.2M");
  });

  it("omits the KPI delta when the field is blank", () => {
    (document.getElementById("kpi-delta") as HTMLInputElement).value = "";
    // Blank delta → undefined (not an empty string), so no delta chip is drawn.
    expect(() => kpiScene()).not.toThrow();
    const texts = kpiScene()
      .nodes.filter((n) => n.kind === "text")
      .map((n) => (n as { text?: string }).text);
    expect(texts).toContain("Revenue");
    expect(texts).not.toContain("");
  });

  it("paints all four previews and the harvey percentage label", () => {
    renderElementPreviews();
    expect(document.getElementById("harvey-val")!.textContent).toBe("75%");
    for (const id of ["harvey-preview", "check-preview", "flow-preview", "kpi-preview"]) {
      expect(document.getElementById(id)!.innerHTML).toContain("<svg");
    }
  });

  it("re-renders live when a control changes after wiring", () => {
    wireElementPreviews();
    const before = document.getElementById("harvey-preview")!.innerHTML;
    const pct = document.getElementById("harvey-pct") as HTMLInputElement;
    pct.value = "10";
    pct.dispatchEvent(new Event("input"));
    expect(document.getElementById("harvey-val")!.textContent).toBe("10%");
    expect(document.getElementById("harvey-preview")!.innerHTML).not.toBe(before);
  });
});

describe("agenda-ui", () => {
  beforeEach(() => {
    document.body.innerHTML = "";
    const ta = document.createElement("textarea");
    ta.id = "agenda-chapters";
    document.body.appendChild(ta);
    box("agenda-preview");
  });

  it("parses non-empty trimmed chapter lines", () => {
    (document.getElementById("agenda-chapters") as HTMLTextAreaElement).value =
      "  Context \n\n Findings \n Recommendation ";
    expect(agendaChapters()).toEqual(["Context", "Findings", "Recommendation"]);
  });

  it("paints a preview when there are chapters and clears it when empty", () => {
    const ta = document.getElementById("agenda-chapters") as HTMLTextAreaElement;
    ta.value = "One\nTwo";
    renderAgendaPreview();
    expect(document.getElementById("agenda-preview")!.innerHTML).toContain("<svg");

    ta.value = "";
    renderAgendaPreview();
    expect(document.getElementById("agenda-preview")!.innerHTML).toBe("");
  });

  it("wires live re-rendering to the textarea", () => {
    wireAgendaPreview();
    const ta = document.getElementById("agenda-chapters") as HTMLTextAreaElement;
    ta.value = "Kickoff";
    ta.dispatchEvent(new Event("input"));
    expect(document.getElementById("agenda-preview")!.innerHTML).toContain("<svg");
  });
});
