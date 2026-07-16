// @vitest-environment jsdom
import { beforeEach, describe, expect, it, vi } from "vitest";
import { readFileSync } from "fs";
import type { ChartConfig } from "../src/core/types";

/**
 * Task-pane state tests. app.ts is a side-effecting entry module: it wires
 * itself to taskpane.html on import and takes the non-Office branch when the
 * `Office` global is absent, which is exactly the case under jsdom. Boot it
 * against the real markup and drive it the way a user does — through the JSON
 * import/export box — so "load a chart" is covered end to end.
 */

const $ = <T extends HTMLElement>(id: string) => document.getElementById(id) as T;

async function bootPane() {
  // Parse rather than regex out the <script> tags: the office.js tag has no
  // business loading here, and a regex that thinks it can find "</script>"
  // misses "</script >" (CodeQL js/bad-tag-filter).
  const parsed = new DOMParser().parseFromString(
    readFileSync("src/taskpane/taskpane.html", "utf8"),
    "text/html",
  );
  parsed.querySelectorAll("script").forEach((s) => s.remove());
  document.body.innerHTML = parsed.body.innerHTML;
  // app.ts holds module-level state and element handles, so it has to re-run
  // against each fresh DOM. Without this the cached module keeps listening to
  // the previous test's detached nodes and every click silently does nothing.
  vi.resetModules();
  await import("../src/taskpane/app");
}

/** Load `cfg` through the pane's JSON import box, as a user would. */
const importConfig = (cfg: Partial<ChartConfig>) => {
  ($("json-io") as HTMLTextAreaElement).value = JSON.stringify(cfg);
  $("json-import").click();
};

/** Read the pane's current config back out of the JSON export box. */
const exportConfig = (): ChartConfig => {
  $("json-export").click();
  return JSON.parse(($("json-io") as HTMLTextAreaElement).value);
};

const cell = (r: number, c: number) =>
  document.querySelector<HTMLInputElement>(`#datasheet input[data-row="${r}"][data-col="${c}"]`)!;

const type = (r: number, c: number, value: string) => {
  const input = cell(r, c);
  input.value = value;
  input.dispatchEvent(new Event("input"));
};

const baseData = { categories: ["A", "B"], series: [{ name: "S1", values: [1, 2] }] };

describe("task pane — loading a chart config", () => {
  beforeEach(async () => {
    await bootPane();
  });

  it("syncs the size fields, so the chart keeps its own dimensions", () => {
    importConfig({ kind: "clustered", width: 720, height: 400, data: baseData });
    // currentConfig() reads the size straight off these inputs, so a stale
    // field silently resized every loaded chart back to 480x300.
    expect(($("chart-w") as HTMLInputElement).value).toBe("720");
    expect(($("chart-h") as HTMLInputElement).value).toBe("400");
    const out = exportConfig();
    expect([out.width, out.height]).toEqual([720, 400]);
  });

  it("keeps a loaded chart's style rather than falling back to pane defaults", () => {
    const style = {
      fontFamily: "Georgia",
      fontSize: 14,
      negative: "#b00020",
      neutral: "#8a8a8a",
      palette: ["#111111", "#222222", "#333333"],
    };
    importConfig({ kind: "clustered", data: baseData, style });
    expect(exportConfig().style).toMatchObject(style);
  });

  it("does not undo into the previous chart's data", () => {
    importConfig({ kind: "clustered", data: baseData });
    type(1, 1, "99"); // an edit, so there is history to undo
    importConfig({
      kind: "clustered",
      data: { categories: ["X", "Y"], series: [{ name: "T1", values: [7, 8] }] },
    });
    const loaded = exportConfig();
    document.dispatchEvent(new KeyboardEvent("keydown", { key: "z", ctrlKey: true, bubbles: true }));
    // Ctrl+Z used to replay the *previous* chart's cells into the new sheet,
    // leaving a chart whose data belonged to neither.
    expect(exportConfig().data).toEqual(loaded.data);
  });
});
