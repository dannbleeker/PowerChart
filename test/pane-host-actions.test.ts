// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { readFileSync } from "fs";
import type { ChartConfig } from "../src/core/types";
import { buildChart } from "../src/core/chart";
import { sceneToSvg } from "../src/render/svg";

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
  demoDeckCalls: [] as unknown[][],
  demoDeckPageFailures: new Set<number>(),
  demoDeckStatusOverride: null as null | "rendered" | "failed" | "skipped",
  /** What canInsertPicture() reports — false models a host below PowerPointApi 1.8. */
  canPicture: true,
  /**
   * What updateChartInSlide resolves to. Defaults to undefined because the
   * existing "target is gone" test depends on that; the explode tests opt in to
   * a live EditTarget, which is what a real successful update hands back.
   */
  updateResult: undefined as undefined | { slideId: string; shapeId: string; left: number; top: number },
  /** Whether the host advertises insertSlidesFromBase64 — off by default. */
  canInsertFile: false,
  slideHoldsOnlyChart: false,
  updateChartThrows: false,
  slideSwapWorks: false,
  /** How many slides the one-shot insert reports landing; null = all of them. */
  insertFileLands: null as null | number,
  insertFileError: null as null | Error,
  buildFileError: null as null | Error,
  slideCount: 1,
  reconcileOutcome: undefined as unknown,
  calls: {
    insertScene: [] as { tagData?: string; left?: number; top?: number }[],
    updateChart: [] as { target: unknown; opts: { tagData?: string; pictureBase64?: string } }[],
    updateCharts: [] as { scene: unknown; target: unknown; opts?: { tagData?: string } }[][],
    insertFile: [] as { b64: string; expected: number }[],
    deselected: [] as string[][],
  },
}));

vi.mock("../src/render/powerpoint", () => ({
  isPowerPointHost: () => true,
  canInsertPicture: vi.fn(() => host.canPicture),
  getSelectionBounds: vi.fn(async () => host.selectionBounds),
  insertSceneIntoSlide: vi.fn(async (_scene: unknown, opts: { tagData?: string; left?: number; top?: number }) => {
    if (host.gate) await host.gate;
    if (host.failInsertOnce) {
      host.failInsertOnce = false;
      throw new Error("host refused the insert");
    }
    host.calls.insertScene.push(opts);
  }),
  updateChartInSlide: vi.fn(
    async (_scene: unknown, target: unknown, opts: { tagData?: string; pictureBase64?: string }) => {
      host.calls.updateChart.push({ target, opts });
      // Models the live-canvas stall: the in-place redraw refuses, a picture
      // update (which draws one shape) does not.
      if (host.updateChartThrows && !opts.pictureBase64) throw new Error("did not respond while drawing shapes 1-10");
      return host.updateResult;
    },
  ),
  updateChartsInSlides: vi.fn(async (items: { scene: unknown; target: unknown; opts?: { tagData?: string } }[]) => {
    host.calls.updateCharts.push(items);
  }),
  listChartsInDeck: vi.fn(async () => host.deckCharts),
  listChartsInSelection: vi.fn(async () => host.selectionCharts),
  loadChartFromSelection: vi.fn(async () => host.loadSelectionResult),
  insertAgendaSlides: vi.fn(async (scenes: unknown[][]) => {
    host.agendaSlides.push(scenes);
  }),
  insertDemoDeck: vi.fn(async (items: { title?: string; scene: { nodes: unknown[] } }[]) => {
    host.demoRuns++;
    const call = host.demoRuns;
    host.demoDeckCalls.push(items);
    if (host.demoDeckPageFailures.has(call)) throw new Error(`page ${call} refused`);
    // Fabricated report — each item counted as failed so the pane's results
    // path has rows to render. Callers that need a specific status set
    // demoDeckStatusOverride.
    const status = host.demoDeckStatusOverride ?? "failed";
    const results = items.map(() => ({ created: 5, status, ms: 100 }));
    return {
      results,
      slidesAdded: items.length,
      addsIssued: items.length,
      blankSlides: [],
      blankItems: [],
      blanksRead: true,
      totalMs: items.length * 100,
    };
  }),
  loadThemePalette: vi.fn(async () => null),
  onLateSync: vi.fn(),
  errorText: (e: unknown) => String(e),
  // The one-shot deck path. Off by default so the existing cases keep
  // exercising the shape-by-shape renderer they were written for.
  canInsertSlidesFromBase64: vi.fn(() => host.canInsertFile),
  // Selection juggling around an in-place redraw. The fake host has no view,
  // so it just runs the body — with `deselected` false, which is the honest
  // answer for a host that cannot move the selection.
  withSlideDeselected: vi.fn(async (ids: string[], fn: (d: boolean) => Promise<unknown>) => {
    host.calls.deselected.push(ids);
    return fn(false);
  }),
  slideHoldsOnlyChart: vi.fn(async () => host.slideHoldsOnlyChart),
  replaceSlideWithDeck: vi.fn(async () => host.slideSwapWorks),
  OFFSCREEN_BATCH: 40,
  insertSlidesFromPptx: vi.fn(async (b64: string, expected: number) => {
    host.calls.insertFile.push({ b64, expected });
    if (host.insertFileError) throw host.insertFileError;
    return host.insertFileLands ?? expected;
  }),
  reconcileDeck: vi.fn(async () => host.reconcileOutcome),
  applyReconcilePlan: vi.fn(async () => host.reconcileOutcome),
  snapshotAddedSlides: vi.fn(async () => []),
  slideCount: vi.fn(async () => host.slideCount),
}));

// The deck builder is a real pptxgenjs run; the pane's own tests care about
// which path was taken, not about the bytes (test/pptx-deck.test.ts covers those).
vi.mock("../src/render/pptx-deck", () => ({
  buildDeckBase64: vi.fn(async (items: unknown[]) => {
    if (host.buildFileError) throw host.buildFileError;
    // The real builder reports what it actually drew per slide — the pane
    // verifies against THAT, not against the Office.js shape estimate.
    return { base64: "UEsDBBQA-fake-base64", shapesPerSlide: items.map(() => 7) };
  }),
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
  host.demoDeckCalls = [];
  host.demoDeckPageFailures = new Set();
  host.demoDeckStatusOverride = null;
  host.canInsertFile = false;
  host.slideHoldsOnlyChart = false;
  host.updateChartThrows = false;
  host.slideSwapWorks = false;
  host.insertFileLands = null;
  host.insertFileError = null;
  host.buildFileError = null;
  host.slideCount = 1;
  host.reconcileOutcome = undefined;
  host.calls.insertFile.length = 0;
  host.calls.deselected.length = 0;
  host.canPicture = true;
  host.updateResult = undefined;
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

/**
 * Stand in for the browser rasteriser. jsdom decodes no SVG, `getContext("2d")`
 * returns null and `toDataURL` returns null, so every step of `rasterizeScene`
 * needs a double. Pass `{ decode: false }` to model a browser that cannot decode
 * the SVG, or `{ encode: false }` for one whose canvas cannot produce a PNG.
 * Returns a restore function.
 */
function stubRaster({ decode = true, encode = true } = {}) {
  class FakeImage {
    onload: (() => void) | null = null;
    onerror: (() => void) | null = null;
    set src(_v: string) {
      Promise.resolve().then(() => (decode ? this.onload?.() : this.onerror?.()));
    }
  }
  vi.stubGlobal("Image", FakeImage);
  const ctx = vi.spyOn(HTMLCanvasElement.prototype, "getContext").mockReturnValue({
    scale() {},
    drawImage() {},
  } as unknown as CanvasRenderingContext2D);
  const url = vi
    .spyOn(HTMLCanvasElement.prototype, "toDataURL")
    .mockReturnValue(encode ? "data:image/png;base64,RASTER" : (null as unknown as string));
  // Capture the SVG each rasterisation was handed. Without this a test can only
  // see THAT a payload was produced, never OF WHAT — and the stub returns the
  // same bytes whatever it is given, so "a payload exists" passes even when the
  // wrong scene was rasterised.
  const svgs: Blob[] = [];
  const c = vi.spyOn(URL, "createObjectURL").mockImplementation((blob: Blob | MediaSource) => {
    if (blob instanceof Blob) svgs.push(blob);
    return "blob:x";
  });
  const r = vi.spyOn(URL, "revokeObjectURL").mockImplementation(() => {});
  return {
    rasterizedSvgs: () => Promise.all(svgs.map((b) => b.text())),
    restore: () => {
      for (const s of [ctx, url, c, r]) s.mockRestore();
    },
  };
}

/** Tick "Insert as picture", the way a user does. */
const tickPicture = () => {
  const box = $("render-image") as HTMLInputElement;
  box.checked = true;
  box.dispatchEvent(new Event("change"));
};

describe("image render mode", () => {
  beforeEach(bootHostPane);

  it("hands the insert a rasterised PNG when Insert as picture is ticked", async () => {
    const raster = stubRaster();
    try {
      tickPicture();
      $("insert").click();
      await settle();
      expect(host.calls.insertScene).toHaveLength(1);
      // The payload reached the renderer, bare — the pane strips the data: URI.
      expect((host.calls.insertScene[0] as { pictureBase64?: string }).pictureBase64).toBe("RASTER");
      // And the config it tagged says image mode, so a re-edit knows what it is.
      expect(JSON.parse(host.calls.insertScene[0].tagData!).render).toBe("image");
    } finally {
      raster.restore();
    }
  });

  it("sends no payload at all in shapes mode, so nothing pays for the canvas", async () => {
    const raster = stubRaster();
    try {
      $("insert").click();
      await settle();
      expect((host.calls.insertScene[0] as { pictureBase64?: string }).pictureBase64).toBeUndefined();
      expect(JSON.parse(host.calls.insertScene[0].tagData!).render).toBeUndefined();
    } finally {
      raster.restore();
    }
  });

  it("skips the raster and SAYS SO on a host below PowerPointApi 1.8", async () => {
    // The note must survive to the end: insertSceneIntoSlide's first act is
    // onPhase("context"), which overwrites the pane note, so a warning posted
    // BEFORE the insert would be wiped. It is posted after, on purpose.
    const raster = stubRaster();
    try {
      host.canPicture = false;
      tickPicture();
      $("insert").click();
      await settle();
      expect((host.calls.insertScene[0] as { pictureBase64?: string }).pictureBase64).toBeUndefined();
      expect($("host-note").textContent).toMatch(/can't insert pictures/i);
    } finally {
      raster.restore();
    }
  });

  it("degrades to native shapes, visibly, when the browser cannot rasterise", async () => {
    for (const mode of [{ decode: false }, { encode: false }] as const) {
      await bootHostPane();
      const raster = stubRaster(mode);
      try {
        tickPicture();
        $("insert").click();
        await settle();
        // The chart still landed — as shapes, with no payload.
        expect(host.calls.insertScene, JSON.stringify(mode)).toHaveLength(1);
        expect((host.calls.insertScene[0] as { pictureBase64?: string }).pictureBase64).toBeUndefined();
        expect($("host-note").textContent, JSON.stringify(mode)).toMatch(/rasterize/i);
      } finally {
        raster.restore();
      }
    }
  });

  it("batch insert honours render per config, not the checkbox", async () => {
    // #json-insert-batch calls insertSceneIntoSlide directly and used to ignore
    // the mode outright, so a pasted image-mode array silently drew shapes.
    const raster = stubRaster();
    try {
      ($("json-io") as HTMLTextAreaElement).value = JSON.stringify([
        { kind: "stacked", render: "image", data: { categories: ["A"], series: [{ name: "S", values: [1] }] } },
        { kind: "stacked", data: { categories: ["A"], series: [{ name: "S", values: [1] }] } },
      ]);
      $("json-insert-batch").click();
      await settle();
      expect(host.calls.insertScene).toHaveLength(2);
      const payloads = host.calls.insertScene.map((o) => (o as { pictureBase64?: string }).pictureBase64);
      expect(payloads).toEqual(["RASTER", undefined]);
    } finally {
      raster.restore();
    }
  });

  it("Same scale rasterises the RESCALED chart, not the one it loaded", async () => {
    // The scale mutation happens inside the map that builds each scene, so
    // rasterising in the same pass captures the PRE-scale chart: the picture
    // shows the old axis while the pane reports success.
    //
    // Asserting "a payload exists" does NOT catch that — the stub returns the
    // same bytes for any input, and tagData reads the same mutated object either
    // way. So compare the SVG actually handed to the rasteriser against the two
    // candidate renders. A chart whose own extent is [1,5] rescaled to max 90
    // draws visibly different bars, so the two SVGs differ.
    const raster = stubRaster();
    try {
      const own = JSON.parse(chartJson([1, 5])) as ChartConfig;
      host.deckCharts = [
        { configJson: JSON.stringify({ ...own, render: "image" }), target: { shapeId: "a" } },
        { configJson: chartJson([2, 90]), target: { shapeId: "b" } },
      ];
      $("same-scale").click();
      await settle();

      const [rasterized] = await raster.rasterizedSvgs();
      expect(rasterized, "exactly one chart was rasterised").toBeTruthy();
      const svgOf = (cfg: ChartConfig) => sceneToSvg(buildChart(cfg));
      const preScale = svgOf({ ...own, render: "image" });
      const postScale = svgOf({ ...own, render: "image", scale: { min: undefined, max: 90 } });
      expect(preScale, "the two candidates must actually differ").not.toBe(postScale);
      expect(rasterized).toBe(postScale);
      expect(rasterized).not.toBe(preScale);
    } finally {
      raster.restore();
    }
  });
});

describe("Explode to native shapes", () => {
  beforeEach(bootHostPane);

  it("re-draws the selected picture chart as shapes by OMITTING the payload", async () => {
    // The omission is the explode: no pictureBase64 means wantsPicture is false
    // in the renderer, so the node loop runs. Nothing else about the update
    // differs, which is why this needs no renderer code.
    host.loadSelectionResult = {
      configJson: JSON.stringify({ ...JSON.parse(chartJson([1, 2])), render: "image" }),
      target: { slideId: "s1", shapeId: "pic-1", left: 10, top: 20 },
    };
    host.updateResult = { slideId: "s1", shapeId: "grp-1", left: 10, top: 20 };
    $("explode").click();
    await settle();
    expect(host.calls.updateChart).toHaveLength(1);
    const { opts } = host.calls.updateChart[0];
    expect((opts as { pictureBase64?: string }).pictureBase64).toBeUndefined();
    // The re-saved config says shapes, so the chart does not turn back into a
    // picture on the next ordinary update.
    expect(JSON.parse(opts.tagData!).render).toBe("shapes");
    expect($("host-note").textContent).toMatch(/native shapes/i);
    // The checkbox follows the exploded config — the pane must not still claim
    // the next insert will be a picture.
    expect(($("render-image") as HTMLInputElement).checked).toBe(false);
  });

  it("refuses politely when the selection is not a PowerChart", async () => {
    host.loadSelectionResult = null;
    $("explode").click();
    await settle();
    expect(host.calls.updateChart).toHaveLength(0);
    expect($("host-note").textContent).toMatch(/select an inserted powerchart/i);
  });

  it("says so when the picture is already gone from the slide", async () => {
    host.loadSelectionResult = {
      configJson: JSON.stringify({ ...JSON.parse(chartJson([1, 2])), render: "image" }),
      target: { slideId: "s1", shapeId: "pic-1", left: 10, top: 20 },
    };
    host.updateResult = undefined; // the update found nothing to replace
    $("explode").click();
    await settle();
    expect($("host-note").textContent).toMatch(/no longer on the slide/i);
  });
});

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

  /**
   * `updateChartInSlide` returns null when the target slide or shape is gone —
   * deliberately, since "a chart whose slide is gone is not an error, it is
   * nothing to do". Nothing consumed the null: the guard saw an unchanged note
   * and printed "Done." in green, and the stale target kept the button reading
   * "Update chart", so every later push no-opped just as silently.
   */
  it("says so, instead of Done., when the target chart is gone", async () => {
    host.loadSelectionResult = {
      configJson: chartJson([1, 2, 3]),
      target: { slideId: "s1", shapeId: "shape-9", left: 10, top: 20 },
    };
    $("load-selection").click();
    await settle();
    expect($("insert").textContent).toBe("Update chart");

    // The mocked update returns undefined — the "nothing to do" answer.
    $("insert").click();
    await settle();
    const note = $("host-note");
    expect(note.textContent).not.toBe("Done.");
    expect(note.className).toContain("status-err");
    // And the pane falls back to inserting, rather than pushing into thin air.
    expect($("insert").textContent).toBe("Insert into slide");
    $("insert").click();
    await settle();
    expect(host.calls.insertScene).toHaveLength(1);
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

describe("demo-insert results pages", () => {
  beforeEach(bootHostPane);

  it("attempts each results page in its own insertDemoDeck call — a failing page does not drop the rest", async () => {
    // Presentation_3.pptx: one failing results page threw insertDemoDeck's
    // "everything failed" guard and NO pages landed. Fix: pane loops over
    // pages, per-page try/catch, so page 2 still gets attempted after page 1
    // throws.
    //
    // With the fake mock returning "failed" for every item and demoItems()'s
    // 32-ish failure-count on the full deck, buildResultsScenes emits multiple
    // pages (≥2). Fail the FIRST results page's insertDemoDeck call; the
    // pane must still call insertDemoDeck at least once more for page 2.
    host.demoDeckPageFailures.add(2); // main deck = call #1, results page 1 = call #2
    $("demo-insert").click();
    await settle();
    // At least 3 calls: main deck + results page 1 (thrown) + results page 2.
    // If the pane batched all results into one call (the old behavior), only
    // 2 calls would happen and the failure at #2 would drop the rest.
    expect(host.demoRuns).toBeGreaterThanOrEqual(3);
    // Each results-page call carries exactly ONE item (per-page loop), NOT
    // the whole batch — the load-bearing property of the fix.
    const resultsCalls = host.demoDeckCalls.slice(1); // drop the main deck call
    for (const call of resultsCalls) expect(call).toHaveLength(1);
    // The pane's note reflects that some pages landed and some did not.
    const noteText = $("host-note").textContent ?? "";
    expect(noteText).toMatch(/results page/i);
  });

  it("all results pages land ⇒ no 'results slide not added' warning", async () => {
    $("demo-insert").click();
    await settle();
    const noteText = $("host-note").textContent ?? "";
    expect(noteText).not.toMatch(/results slide not added/i);
    // At least the main deck + at least one results page.
    expect(host.demoRuns).toBeGreaterThanOrEqual(2);
  });
});

describe("updating a chart the live canvas will not redraw", () => {
  beforeEach(bootHostPane);

  /** Load a chart so Insert becomes Update, then push an edit. */
  async function loadThenUpdate() {
    host.loadSelectionResult = {
      configJson: chartJson([1, 2, 3]),
      target: { slideId: "s1", shapeId: "shape-9", left: 10, top: 20 },
    };
    $("load-selection").click();
    await settle();
    $("insert").click();
    await settle();
  }

  it("redraws with the slide deselected, at the off-screen batch size", async () => {
    // The whole point: a redraw is the add-in's worst case on the web, and it
    // is only bad because the slide is on screen. Looking away is free.
    await loadThenUpdate();
    expect(host.calls.deselected).toEqual([["s1"]]);
  });

  it("swaps the slide when the redraw stalls and the slide holds only the chart", async () => {
    host.updateChartThrows = true;
    host.canInsertFile = true;
    host.slideHoldsOnlyChart = true;
    host.slideSwapWorks = true;
    await loadThenUpdate();
    expect(host.calls.insertFile).toHaveLength(0); // a slide swap, not a deck insert
    expect($("host-note").textContent).toMatch(/rebuilt that slide/i);
  });

  it("will not swap a slide that holds anything besides the chart", async () => {
    // The replacement is a NEW slide: notes, transitions and any other shape
    // on the old one do not come with it. So it is only ever offered where
    // there is demonstrably nothing else to lose.
    host.updateChartThrows = true;
    host.canInsertFile = true;
    host.slideSwapWorks = true; // the swap WOULD work — the guard is the only thing stopping it
    host.slideHoldsOnlyChart = false;
    await loadThenUpdate();
    expect($("host-note").textContent).not.toMatch(/rebuilt that slide/i);
  });

  // NOT COVERED HERE: the picture floor, and the "everything failed, rethrow
  // the host's own words" path behind it. jsdom has no canvas, so
  // `rasterizeScene` neither succeeds nor fails — it simply never settles,
  // which is also why that call is bounded by a timeout in the pane. Faking
  // timers to drive it broke ten unrelated tests in this file; a test that
  // waits out a real 10s timeout is not worth 10s on every run.
});

describe("demo-insert one-shot deck insert", () => {
  beforeEach(bootHostPane);

  it("hands the host a generated file and never draws the deck shape by shape", async () => {
    host.canInsertFile = true;
    $("demo-insert").click();
    await settle();
    expect(host.calls.insertFile).toHaveLength(1);
    expect(host.calls.insertFile[0].b64).toMatch(/^UEsDBBQA/);
    // The whole point: not one slide is drawn through the shape renderer.
    expect(host.demoRuns).toBe(0);
    expect($("host-note").textContent).toMatch(/one file/i);
  });

  it("falls back to shapes when the host has no insertSlidesFromBase64", async () => {
    host.canInsertFile = false;
    $("demo-insert").click();
    await settle();
    expect(host.calls.insertFile).toHaveLength(0);
    expect(host.demoRuns).toBeGreaterThanOrEqual(1);
  });

  it("falls back when the file could not be built at all", async () => {
    host.canInsertFile = true;
    host.buildFileError = new Error("pptxgenjs blew up");
    $("demo-insert").click();
    await settle();
    expect(host.calls.insertFile).toHaveLength(0);
    expect(host.demoRuns).toBeGreaterThanOrEqual(1);
  });

  it("falls back when the insert threw and nothing landed", async () => {
    host.canInsertFile = true;
    host.insertFileError = new Error("host refused the deck");
    host.slideCount = 1; // unchanged before and after — nothing landed
    $("demo-insert").click();
    await settle();
    expect(host.calls.insertFile).toHaveLength(1);
    expect(host.demoRuns).toBeGreaterThanOrEqual(1);
  });

  it("does NOT fall back after a partial insert — that would draw the deck twice", async () => {
    // The failure this exists to prevent: some slides landed, the pane decides
    // the fast path "failed", and the shape renderer then appends the whole
    // deck a second time on top of them.
    host.canInsertFile = true;
    host.insertFileLands = 3;
    $("demo-insert").click();
    await settle();
    expect(host.calls.insertFile).toHaveLength(1);
    expect(host.demoRuns).toBe(0);
    expect($("host-note").textContent).toMatch(/host took 3 of/i);
  });

  it("appends the generated deck instead of letting the host front it", async () => {
    // The first real run put 37 generated slides AHEAD of the user's own title
    // slide, because insertSlidesFromBase64 inserts at the front unless it is
    // given a target. The pane must ask for the tail.
    host.canInsertFile = true;
    $("demo-insert").click();
    await settle();
    expect(host.calls.insertFile).toHaveLength(1);
    expect(host.calls.insertFile[0].expected).toBeGreaterThan(0);
  });

  it("respects the fast-path opt-out", async () => {
    host.canInsertFile = true;
    ($("demo-file") as HTMLInputElement).checked = false;
    $("demo-insert").click();
    await settle();
    expect(host.calls.insertFile).toHaveLength(0);
    expect(host.demoRuns).toBeGreaterThanOrEqual(1);
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
