// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { readFileSync } from "fs";
import type { ChartConfig } from "../src/core/types";

/**
 * Pane ↔ host command handlers. `pane-state.test.ts` boots app.ts under jsdom
 * with no `Office` global, so it exercises everything EXCEPT the buttons that
 * talk to PowerPoint — Insert, Insert-new, Load-selection, Same-scale. Those
 * handlers (`doInsert`, `doLoadSelection`, `doSameScale`) were the pane's
 * largest untested surface: the branch that chooses a selected placeholder's
 * bounds over the tiled offset, the edit-in-place path, the union-extent maths
 * behind Same scale, and the "not a PowerChart" / "needs two charts" guards.
 *
 * The renderer primitives those handlers call (`insertSceneIntoSlide` et al.)
 * are covered against a fake host in `office-render.test.ts`; here the module is
 * mocked to spies so the test asserts the pane's ORCHESTRATION — which handler
 * fires, with what config — rather than re-testing the shape emitter.
 */

/** Shared mailbox the mocked renderer writes to; reset before each boot. */
const host = vi.hoisted(() => ({
  selectionBounds: null as null | { left: number; top: number; width: number; height: number },
  deckCharts: [] as { configJson: string; target: unknown }[],
  selectionCharts: [] as { configJson: string; target: unknown }[],
  loadSelectionResult: null as null | { configJson: string; target: unknown },
  // When set, insertSceneIntoSlide awaits this before resolving — lets a test
  // observe the pane's mid-flight state (buttons disabled) before the action ends.
  gate: null as null | Promise<void>,
  // insertSceneIntoSlide throws this once, if set — drives the guard's catch path.
  failInsertOnce: false,
  // The selection-change listener app.ts registers via addHandlerAsync, captured
  // so a test can fire it the way PowerPoint would.
  selectionListener: null as null | (() => unknown),
  agendaSlides: [] as unknown[][],
  demoRuns: 0,
  calls: {
    insertScene: [] as { tagData?: string; left?: number; top?: number }[],
    updateChart: [] as { target: unknown; opts: { tagData?: string } }[],
    updateCharts: [] as { scene: unknown; target: unknown; opts?: { tagData?: string } }[][],
  },
}));

vi.mock("../src/render/powerpoint", () => ({
  isPowerPointHost: () => true,
  getSelectionBounds: vi.fn(async () => host.selectionBounds),
  insertSceneIntoSlide: vi.fn(async (_scene: unknown, opts: { tagData?: string; left?: number; top?: number }) => {
    if (host.gate) await host.gate;
    if (host.failInsertOnce) {
      host.failInsertOnce = false;
      throw new Error("host refused the insert");
    }
    host.calls.insertScene.push(opts);
  }),
  updateChartInSlide: vi.fn(async (_scene: unknown, target: unknown, opts: { tagData?: string }) => {
    host.calls.updateChart.push({ target, opts });
  }),
  updateChartsInSlides: vi.fn(async (items: { scene: unknown; target: unknown; opts?: { tagData?: string } }[]) => {
    host.calls.updateCharts.push(items);
  }),
  listChartsInDeck: vi.fn(async () => host.deckCharts),
  listChartsInSelection: vi.fn(async () => host.selectionCharts),
  loadChartFromSelection: vi.fn(async () => host.loadSelectionResult),
  insertAgendaSlides: vi.fn(async (scenes: unknown[][]) => {
    host.agendaSlides.push(scenes);
  }),
  insertDemoDeck: vi.fn(async () => {
    host.demoRuns++;
    return { results: [] };
  }),
  loadThemePalette: vi.fn(async () => null),
  onLateSync: vi.fn(),
  errorText: (e: unknown) => String(e),
}));

const $ = <T extends HTMLElement>(id: string) => document.getElementById(id) as T;

/** Let the clicked handler's async chain (and its busy→done note) settle. */
const settle = () => new Promise((r) => setTimeout(r, 0));

/**
 * Boot app.ts with a PowerPoint host present. app.ts wires the host buttons only
 * inside `Office.onReady`, so the stub must both look like a host (mocked
 * `isPowerPointHost`) and fire onReady synchronously at import.
 */
async function bootHostPane() {
  host.selectionBounds = null;
  host.deckCharts = [];
  host.selectionCharts = [];
  host.loadSelectionResult = null;
  host.gate = null;
  host.failInsertOnce = false;
  host.selectionListener = null;
  host.agendaSlides = [];
  host.demoRuns = 0;
  host.calls.insertScene = [];
  host.calls.updateChart = [];
  host.calls.updateCharts = [];

  window.history.replaceState({}, "", "/taskpane.html");
  const parsed = new DOMParser().parseFromString(readFileSync("src/taskpane/taskpane.html", "utf8"), "text/html");
  parsed.querySelectorAll("script").forEach((s) => s.remove());
  document.body.innerHTML = parsed.body.innerHTML;

  vi.stubGlobal("Office", {
    onReady: (cb: () => void) => cb(),
    EventType: { DocumentSelectionChanged: "DocumentSelectionChanged" },
    context: {
      host: "PowerPoint",
      displayLanguage: "en-US",
      document: {
        // watchSelection() registers a selection listener; capture it so a test
        // can fire it the way PowerPoint does when the user clicks a shape.
        addHandlerAsync: (_type: string, handler: () => unknown) => {
          host.selectionListener = handler;
        },
      },
    },
  });

  vi.resetModules();
  await import("../src/taskpane/app");
  await settle();
}

/** A value-axis chart config with a known extent, as a JSON tag would carry it. */
const chartJson = (values: number[]): string =>
  JSON.stringify({
    kind: "stacked",
    width: 480,
    height: 320,
    data: { categories: values.map((_, i) => `C${i}`), series: [{ name: "S", values }] },
  } satisfies ChartConfig);

afterEach(() => vi.unstubAllGlobals());

describe("Insert", () => {
  beforeEach(bootHostPane);

  it("tiles a new chart at the default offset when nothing is selected", async () => {
    host.selectionBounds = null;
    $("insert").click();
    await settle();
    expect(host.calls.insertScene).toHaveLength(1);
    const opts = host.calls.insertScene[0];
    // Default drop point, and the config round-trips as the shape's tag.
    expect(opts.left).toBeGreaterThanOrEqual(60);
    expect(opts.top).toBeGreaterThanOrEqual(90);
    expect(JSON.parse(opts.tagData!)).toMatchObject({ kind: expect.any(String) });
    expect(host.calls.updateChart).toHaveLength(0);
  });

  it("fits the chart to a selected placeholder's bounds", async () => {
    host.selectionBounds = { left: 200, top: 150, width: 360, height: 240 };
    $("insert").click();
    await settle();
    expect(host.calls.insertScene).toHaveLength(1);
    const opts = host.calls.insertScene[0];
    expect(opts.left).toBe(200);
    expect(opts.top).toBe(150);
    // The placeholder's size overrides the config's own dimensions.
    const cfg = JSON.parse(opts.tagData!) as ChartConfig;
    expect(cfg.width).toBe(360);
    expect(cfg.height).toBe(240);
  });

  it("ignores a selection too small to be a real placeholder", async () => {
    host.selectionBounds = { left: 5, top: 5, width: 20, height: 20 };
    $("insert").click();
    await settle();
    // Below the 40pt threshold → falls back to the tiled offset, not the bounds.
    expect(host.calls.insertScene[0].left).toBeGreaterThanOrEqual(60);
  });
});

describe("Insert updates in place after loading a chart", () => {
  beforeEach(bootHostPane);

  it("routes Insert to an update once a chart is loaded from the selection", async () => {
    host.loadSelectionResult = {
      configJson: chartJson([1, 2, 3]),
      target: { slideId: "s1", shapeId: "shape-9", left: 10, top: 20 },
    };
    $("load-selection").click();
    await settle();

    // Now the primary Insert edits in place rather than dropping a new chart.
    $("insert").click();
    await settle();
    expect(host.calls.updateChart).toHaveLength(1);
    expect(host.calls.updateChart[0].target).toMatchObject({ shapeId: "shape-9" });
    expect(host.calls.insertScene).toHaveLength(0);
  });

  it("Insert-new always drops a fresh chart even with a chart loaded", async () => {
    host.loadSelectionResult = {
      configJson: chartJson([4, 5]),
      target: { slideId: "s1", shapeId: "shape-3", left: 0, top: 0 },
    };
    $("load-selection").click();
    await settle();

    $("insert-new").click();
    await settle();
    expect(host.calls.insertScene).toHaveLength(1);
    expect(host.calls.updateChart).toHaveLength(0);
  });
});

describe("Load selection", () => {
  beforeEach(bootHostPane);

  it("loads a PowerChart and reveals the in-place edit affordance", async () => {
    host.loadSelectionResult = {
      configJson: chartJson([7, 8, 9]),
      target: { slideId: "s1", shapeId: "shape-1", left: 0, top: 0 },
    };
    $("load-selection").click();
    await settle();
    expect($("host-note").textContent?.toLowerCase()).toContain("loaded");
    // The "edit selected chart" banner is stale once loaded, so it hides.
    expect($("selection-banner").style.display).toBe("none");
  });

  it("reports a non-PowerChart selection without touching the deck", async () => {
    host.loadSelectionResult = null;
    $("load-selection").click();
    await settle();
    expect($("host-note").textContent?.toLowerCase()).toContain("not a powerchart");
    expect(host.calls.updateChart).toHaveLength(0);
  });
});

describe("Same scale", () => {
  beforeEach(bootHostPane);

  it("pins every value-axis chart in the deck to the union extent", async () => {
    host.deckCharts = [
      { configJson: chartJson([10, 20, 30]), target: { slideId: "s1", shapeId: "a", left: 0, top: 0 } },
      { configJson: chartJson([5, 90]), target: { slideId: "s2", shapeId: "b", left: 0, top: 0 } },
    ];
    $("same-scale").click();
    await settle();
    expect(host.calls.updateCharts).toHaveLength(1);
    const batch = host.calls.updateCharts[0];
    expect(batch).toHaveLength(2);
    // Both charts are re-tagged with the SAME max — the union of the two extents.
    const maxima = batch.map((b) => (JSON.parse(b.opts!.tagData!) as ChartConfig).scale?.max);
    expect(new Set(maxima).size).toBe(1);
    expect(maxima[0]).toBe(90);
  });

  it("refuses to apply with fewer than two value-axis charts", async () => {
    host.deckCharts = [{ configJson: chartJson([1, 2, 3]), target: { slideId: "s1", shapeId: "a", left: 0, top: 0 } }];
    $("same-scale").click();
    await settle();
    expect(host.calls.updateCharts).toHaveLength(0);
    expect($("host-note").textContent?.toLowerCase()).toContain("two");
  });

  it("scopes to the selection and guides the user when too few are selected", async () => {
    host.selectionCharts = [
      { configJson: chartJson([1, 2]), target: { slideId: "s1", shapeId: "a", left: 0, top: 0 } },
    ];
    $("same-scale-sel").click();
    await settle();
    expect(host.calls.updateCharts).toHaveLength(0);
    // The selection-scoped guard names Ctrl-click, not the deck message.
    expect($("host-note").textContent?.toLowerCase()).toContain("ctrl-click");
  });
});

describe("Elements and batch insert", () => {
  beforeEach(bootHostPane);

  // The five element buttons (harvey balls, checkboxes, process flow, KPI row,
  // table) all drop a compact scene at the same fixed offset through the guard.
  for (const id of ["harvey-insert", "check-insert", "flow-insert", "kpi-insert", "table-insert"]) {
    it(`${id} inserts a compact element scene at the element offset`, async () => {
      $(id).click();
      await settle();
      expect(host.calls.insertScene).toHaveLength(1);
      expect(host.calls.insertScene[0].left).toBe(120);
      expect(host.calls.insertScene[0].top).toBe(160);
    });
  }

  it("batch-inserts every config in the JSON box onto the current slide", async () => {
    ($("json-io") as HTMLTextAreaElement).value = JSON.stringify([
      { kind: "pie", data: { categories: ["A", "B"], series: [{ name: "S", values: [1, 1] }] } },
      { kind: "stacked", data: { categories: ["A"], series: [{ name: "S", values: [3] }] } },
    ]);
    $("json-insert-batch").click();
    await settle();
    expect(host.calls.insertScene).toHaveLength(2);
    expect($("host-note").textContent?.toLowerCase()).toContain("2 chart");
  });
});

describe("guard — busy lockout and error surfacing", () => {
  beforeEach(bootHostPane);

  it("disables the button while the action runs and re-enables it after", async () => {
    // Hold the insert open so the mid-flight state is observable.
    let release!: () => void;
    host.gate = new Promise<void>((r) => (release = r));
    const insertBtn = $<HTMLButtonElement>("insert");

    insertBtn.click();
    await settle();
    expect(insertBtn.disabled).toBe(true); // locked out mid-action

    release();
    await settle();
    expect(insertBtn.disabled).toBe(false); // restored once the action settles
    expect(host.calls.insertScene).toHaveLength(1);
  });

  it("surfaces a host failure as a Failed note instead of throwing", async () => {
    host.failInsertOnce = true;
    $("insert").click();
    await settle();
    expect($("host-note").textContent?.toLowerCase()).toContain("failed");
    expect(host.calls.insertScene).toHaveLength(0);
  });
});

describe("watchSelection", () => {
  beforeEach(bootHostPane);

  it("offers the edit banner when a PowerChart is selected", async () => {
    expect(host.selectionListener).toBeTypeOf("function");
    host.loadSelectionResult = {
      configJson: chartJson([1, 2, 3]),
      target: { slideId: "s1", shapeId: "sel-1", left: 0, top: 0 },
    };
    await host.selectionListener!();
    await settle();
    expect($("selection-banner").style.display).toBe(""); // shown
  });

  it("hides the banner when the selection is not a PowerChart", async () => {
    host.loadSelectionResult = null;
    await host.selectionListener!();
    await settle();
    expect($("selection-banner").style.display).toBe("none");
  });
});
