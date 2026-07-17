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

describe("busy-guard on host actions", () => {
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

describe("the action bar belongs to the Chart tab", () => {
  const bar = () => document.querySelector<HTMLElement>(".action-bar")!;
  const clickTab = (name: string) => document.querySelector<HTMLButtonElement>(`.tabs .tab[data-tab="${name}"]`)!.click();

  it("hides itself on every tab that is not Chart", async () => {
    // Every action in this bar reads the CHART's state: "Insert into slide"
    // inserts currentConfig(). On Elements that put a big primary button that
    // inserts a stacked column chart directly under the small "Insert" that
    // inserts the Harvey ball you are looking at — the prominent button was
    // the wrong one.
    await bootPane();
    expect(bar().hasAttribute("hidden"), "Chart tab").toBe(false);
    for (const tab of ["elements", "agenda", "automation"]) {
      clickTab(tab);
      expect(bar().hasAttribute("hidden"), tab).toBe(true);
    }
    clickTab("chart");
    expect(bar().hasAttribute("hidden"), "back on Chart").toBe(false);
  });

  it("honours the tab a deep link opens on", async () => {
    // ?tab=elements clicks the tab after wiring, so the bar must follow the
    // link rather than the markup's default active tab.
    await bootPane("?tab=elements");
    expect(bar().hasAttribute("hidden")).toBe(true);
  });

  it("actually disappears — [hidden] must beat the bar's own display:flex", async () => {
    // The trap: `.action-bar { display: flex }` is an author rule and outranks
    // the UA stylesheet's `[hidden] { display: none }`, so the attribute would
    // be set and the bar would stay on screen. Assert the CSS, not the DOM.
    const css = readFileSync("src/taskpane/taskpane.css", "utf8");
    expect(css).toMatch(/\.action-bar\[hidden\]\s*\{[^}]*display:\s*none/);
  });
});

describe("status is pane-wide, and only claims what it knows", () => {
  const strip = () => document.getElementById("status-strip")!;
  const bar = () => document.getElementById("status-bar")!;
  const noteEl = () => document.getElementById("host-note")!;

  it("lives OUTSIDE the action bar, so hiding that bar cannot silence it", async () => {
    // The regression this exists for: host-note used to live inside
    // <footer class="action-bar">, which is Chart-only. Hiding the bar took
    // every "Working…", "Failed:" and progress count on Elements / Agenda /
    // Automation down with it — including the demo deck's own counter, which
    // is on Automation. Inserting is slow enough here that silence reads as
    // broken.
    await bootPane();
    expect(document.querySelector(".action-bar #host-note"), "note must not be inside the action bar").toBeNull();
    expect(strip().contains(noteEl())).toBe(true);
    // And the strip must not be a child of the thing that gets hidden.
    expect(document.querySelector(".action-bar")!.contains(strip())).toBe(false);
  });

  it("collapses when there is nothing to say", async () => {
    await bootPane();
    expect(strip().hasAttribute("hidden")).toBe(true);
    expect(bar().hasAttribute("hidden")).toBe(true);
  });

  it("shows an INDETERMINATE bar for work whose progress we cannot know", async () => {
    // A single insert is one context.sync(); Office.js reports nothing until it
    // lands. Any percentage would be invented, and a bar stuck at 99% is a
    // worse lie than no bar.
    const { release } = await bootHost();
    $<HTMLButtonElement>("demo-insert").click();
    await Promise.resolve();
    expect(strip().hasAttribute("hidden")).toBe(false);
    expect(bar().hasAttribute("hidden")).toBe(false);
    expect(bar().classList.contains("indeterminate")).toBe(true);
    // Indeterminate means NO width claim.
    expect(bar().querySelector("i")!.style.width).toBe("");
    release();
    await vi.waitFor(() => expect(bar().hasAttribute("hidden")).toBe(true));
  });

  it("counts the seconds while the host works, and stops when it is done", async () => {
    // The only number we can honestly report mid-sync — and on a host that
    // takes 20s to draw a chart, a number that moves is the whole difference
    // between "working" and "dead".
    vi.useFakeTimers();
    try {
      const { release } = await bootHost();
      const elapsed = () => document.getElementById("status-elapsed")!.textContent;
      $<HTMLButtonElement>("demo-insert").click();
      await Promise.resolve();
      expect(elapsed()).toBe("0s");
      await vi.advanceTimersByTimeAsync(3_000);
      expect(elapsed()).toBe("3s");
      release();
      await vi.advanceTimersByTimeAsync(10);
      await vi.waitFor(() => expect(elapsed()).toBe(""));
      // The ticker must not outlive the work.
      await vi.advanceTimersByTimeAsync(5_000);
      expect(elapsed()).toBe("");
    } finally {
      vi.useRealTimers();
    }
  });

  it("clears the bar when the note turns into an error", async () => {
    vi.stubGlobal("Office", { context: { host: "PowerPoint", requirements: { isSetSupported: () => false } } });
    vi.stubGlobal("PowerPoint", { run: async () => { throw new Error("host refused"); } });
    await bootPane();
    $<HTMLButtonElement>("demo-insert").click();
    await vi.waitFor(() => expect(noteEl().textContent).toMatch(/^Failed:/));
    // The message stays; the "still working" signal must not.
    expect(strip().hasAttribute("hidden")).toBe(false);
    expect(bar().hasAttribute("hidden")).toBe(true);
  });
});

describe("element previews are sized for their own shape", () => {
  it("does not stretch the KPI tile like the process flow", async () => {
    // The flow is 480x44 and built to shrink, so it wants width:100%. The KPI
    // is 160x90 — the same rule blows it up to ~2.5x its natural height, which
    // pushed the KPI card's own Insert button out of view and left "Table from
    // datasheet"'s Insert as the nearest one. The owner clicked it and got a
    // table, exactly as the layout invited.
    await bootPane();
    expect(document.getElementById("kpi-preview")!.className).toBe("element-tile-preview");
    expect(document.getElementById("flow-preview")!.className).toBe("element-flow-preview");
    const css = readFileSync("src/taskpane/taskpane.css", "utf8");
    // A tile caps at its natural size; only the flow may stretch.
    expect(css).toMatch(/\.element-tile-preview svg \{[^}]*max-width:\s*100%/);
    expect(css).not.toMatch(/\.element-tile-preview svg \{[^}]*[^-]width:\s*100%/);
    expect(css).toMatch(/\.element-flow-preview svg \{[^}]*[^-]width:\s*100%/);
  });

  it("renders the KPI preview at the tile's real aspect, not the card's width", async () => {
    await bootPane();
    const svg = document.querySelector("#kpi-preview svg")!;
    // sceneToSvg states the scene's own size; the CSS must not override it.
    expect(svg.getAttribute("width")).toBe("160");
    expect(svg.getAttribute("height")).toBe("90");
  });
});
