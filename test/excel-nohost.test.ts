// @vitest-environment jsdom
import { describe, expect, it, vi } from "vitest";

/** The Excel pane outside an Excel host: generate disabled, copy still works. */
describe("Excel pane without a host", () => {
  it("disables Generate, explains sideloading, and copies output", async () => {
    document.body.innerHTML = `
      <select id="kind"><option value="stacked">stacked</option></select>
      <input id="title" value="">
      <button id="generate"></button>
      <button id="copy"></button>
      <textarea id="output">{"kind":"stacked"}</textarea>
      <p id="note"></p>`;

    const writeText = vi.fn(async () => {});
    Object.defineProperty(navigator, "clipboard", { value: { writeText }, configurable: true });
    // No Office/Excel globals at all → module falls through to bare wire().
    vi.resetModules();
    await import("../src/excel/excel");

    expect((document.getElementById("generate") as HTMLButtonElement).disabled).toBe(true);
    expect(document.getElementById("note")!.textContent).toContain("manifest-excel.xml");

    document.getElementById("copy")!.click();
    expect(writeText).toHaveBeenCalledWith('{"kind":"stacked"}');

    // Copy with an empty output box is a silent no-op.
    (document.getElementById("output") as HTMLTextAreaElement).value = "";
    document.getElementById("copy")!.click();
    expect(writeText).toHaveBeenCalledTimes(1);
  });
});
