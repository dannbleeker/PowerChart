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

async function bootPane(search = "") {
  // app.ts reads ?lang= at import time, so the URL has to be set up front.
  window.history.replaceState({}, "", `/taskpane.html${search}`);
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

describe("task pane — localisation survives re-renders", () => {
  it("keeps the action button translated after a chart loads", async () => {
    await bootPane("?lang=de");
    const insert = () => $("insert").textContent;
    expect(insert()).toBe("In Folie einfügen");

    // renderActionState rewrites this label whenever the edit target changes —
    // which happens long after localizePane ran — and used to stamp the English
    // string straight back into a German pane.
    importConfig({ kind: "clustered", data: baseData });
    expect(insert()).toBe("In Folie einfügen");
  });
});

describe("task pane — status colour and headings", () => {
  beforeEach(async () => {
    await bootPane();
  });

  it("shows a failure in the error colour, not the previous success's green", () => {
    const note = () => $("host-note");
    // Stand in for a prior successful action having stamped the green class.
    note().className = "hint status-ok";
    ($("json-io") as HTMLTextAreaElement).value = "{ not json";
    $("json-import").click();
    expect(note().textContent).toMatch(/^Invalid JSON:/);
    expect(note().className).toBe("hint status-err");
  });

  it("labels a successful load as such", () => {
    importConfig({ kind: "clustered", data: baseData });
    expect($("host-note").className).toBe("hint status-ok");
  });
});

describe("task pane — accordion headings are translated", () => {
  it("translates the numbered step headings", async () => {
    await bootPane("?lang=de");
    const titles = [...document.querySelectorAll(".acc-title")].map((e) => e.textContent);
    // These live in a <span> inside the <summary>; matching the <summary> alone
    // never reached them, so the dictionary entries were dead.
    expect(titles).toContain("1 · Diagrammtyp");
    expect(titles).toContain("3 · Dekorationen");
    expect(titles).toContain("Vorschau & Größe");
    expect(titles).not.toContain("1 · Chart type");
  });
});

describe("busy-guard on host actions", () => {
  /**
   * Boot the pane down its HOST branch, with a PowerPoint.run we control: it
   * parks until we release it, so we can look at the buttons mid-flight — which
   * is the only moment the bug existed.
   */
  async function bootHost() {
    let release!: () => void;
    const parked = new Promise<void>((r) => (release = r));
    let runs = 0;
    vi.stubGlobal("Office", {
      context: { host: "PowerPoint", requirements: { isSetSupported: () => false } },
    });
    vi.stubGlobal("PowerPoint", {
      run: async (cb: (ctx: unknown) => Promise<unknown>) => {
        runs++;
        await parked;
        return cb({
          presentation: {
            slides: {
              getCount: () => ({ value: 0 }),
              add() {},
              getItemAt: () => ({ shapes: { addGeometricShape: stubShape, addLine: stubShape, addTextBox: stubShape } }),
            },
          },
          sync: async () => {},
        });
      },
      GeometricShapeType: new Proxy({}, { get: (_t, p) => String(p) }),
      ConnectorType: { straight: "straight" },
      ShapeLineDashStyle: { dash: "dash" },
      ShapeAutoSize: { autoSizeNone: "none" },
      TextVerticalAlignment: { top: "t", middle: "m", bottom: "b" },
      ParagraphHorizontalAlignment: { left: "l", center: "c", right: "r" },
    });
    const stubShape = () => ({
      fill: { setSolidColor() {}, clear() {} },
      lineFormat: {},
      textFrame: { textRange: { font: {}, paragraphFormat: {} } },
      tags: { add() {} },
    });
    await bootPane();
    return { release, runs: () => runs };
  }

  it("disables the clicked button AND the primary Insert while an action runs", async () => {
    // The bug, seen in real PowerPoint: guard disabled only the primary button,
    // so "Insert demo deck" stayed live through a multi-minute run (one more
    // click = another 35 slides), while Insert went dead WITHOUT looking dead —
    // so a stuck action read as "Insert is broken".
    const { release } = await bootHost();
    const demo = $<HTMLButtonElement>("demo-insert");
    const insert = $<HTMLButtonElement>("insert");
    expect(demo.disabled).toBe(false);
    expect(insert.disabled).toBe(false);

    demo.click();
    await Promise.resolve();
    expect(demo.disabled, "clicked button must not accept a second click").toBe(true);
    expect(insert.disabled, "primary acts on the same deck").toBe(true);

    release();
    await vi.waitFor(() => expect(demo.disabled).toBe(false));
    expect(insert.disabled).toBe(false);
  });

  it("re-enables both when the action fails", async () => {
    vi.stubGlobal("Office", { context: { host: "PowerPoint", requirements: { isSetSupported: () => false } } });
    vi.stubGlobal("PowerPoint", {
      run: async () => {
        throw new Error("host refused");
      },
    });
    await bootPane();
    const demo = $<HTMLButtonElement>("demo-insert");
    demo.click();
    await vi.waitFor(() => expect(document.getElementById("host-note")!.textContent).toMatch(/^Failed:/));
    // A failed action must not leave the pane permanently dead.
    expect(demo.disabled).toBe(false);
    expect($<HTMLButtonElement>("insert").disabled).toBe(false);
  });
});
