// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import {
  _setBatchTimeoutForTest,
  _setBlankReReadDelayForTest,
  CHART_PARTS_TAG,
  CHART_TAG,
  CHART_ORIGIN_TAG,
  getSelectionBounds,
  insertAgendaSlides,
  insertDemoDeck,
  insertSceneIntoSlide,
  isPowerPointHost,
  listChartsInDeck,
  listChartsInSelection,
  loadChartFromSelection,
  onLateSync,
  READBACK_PAGE,
  wantsAutoPicture,
  DEMO_SLOT_TAG,
  DEMO_SHAPE_BUDGET,
  applyReconcilePlan,
  readAddedSlides,
  _setReadbackTimeoutForTest,
  _setDeckInsertPerSlideForTest,
  insertSlidesFromPptx,
  lastLateSyncOwner,
  lastLateSyncSeq,
  waitForLateSync,
  reconcileDeck,
  snapshotAddedSlides,
  updateChartInSlide,
  updateChartsInSlides,
  withSlideDeselected,
  slideHoldsOnlyChart,
  getSlideShapeBounds,
  _setSelectionTimeoutForTest,
  replaceSlideWithDeck,
  deleteShapesById,
  addScratchSlide,
  deleteSlideById,
  clearShapeSelection,
  MAX_ADD_RETRY_ROUNDS,
  wreckageOf,
  requestStop,
  resetStop,
  slideShots,
  isStopped,
  slideSize,
  _resetSlideSizeCache,
} from "../src/render/powerpoint";

/**
 * Dropped `slides.add()` calls needed to defeat ONE `addSlides` call outright:
 * the original plus every retry round it is allowed.
 *
 * Derived, never hardcoded. These tests used to spell it `2`, which silently
 * stopped meaning "the add and all its retries" the moment the retry bound
 * moved — the assertions still passed for the wrong reason, because a
 * recovered add and a lost one differ only in numbers the test did not check.
 */
const ADDS_TO_DEFEAT_ONE_SLIDE = 1 + MAX_ADD_RETRY_ROUNDS;
import { readFileSync } from "fs";
import { setTracing, traceLog } from "../src/core/trace";
import { planReconcile } from "../src/core/reconcile";
import { buildChart, DEFAULT_SIZE } from "../src/core/chart";
import { estimateOfficeShapes } from "../src/core/scene";
import { sampleConfig, CHART_KINDS } from "../src/core/samples";
import { buildDeckBase64 } from "../src/render/pptx-deck";
import { buildAgendaScene } from "../src/core/agenda";
import type { ChartConfig, MarkerSymbol } from "../src/core/types";
import type { DemoReport } from "../src/render/powerpoint";

/** The indices a demo run did not render as a real chart (skipped or failed). */
const failedIndices = (r: DemoReport) =>
  r.results.map((x, i) => (x.status !== "rendered" ? i : -1)).filter((i) => i >= 0);

import {
  addedWithLayout,
  blankReadbackAt,
  failSyncsOn,
  faults,
  hostSlideSize,
  installHost,
  makeShape,
  makeSlide,
  stallSyncOn,
  trips,
  unansweredNullChecks,
  untracked,
  lastShapeLoadSpec,
  type FakeShape,
  type FakeSlide,
} from "./helpers/office-host";

const config: ChartConfig = {
  kind: "stacked",
  ...DEFAULT_SIZE,
  data: {
    categories: ["A", "B"],
    series: [
      { name: "S1", values: [3, 4] },
      { name: "S2", values: [1, 2] },
    ],
  },
};

afterEach(() => vi.unstubAllGlobals());

describe("insertSceneIntoSlide", () => {
  it("creates native shapes at the requested offset, groups, and tags", async () => {
    const slide = makeSlide("s1");
    installHost([slide]);
    await insertSceneIntoSlide(buildChart(config), { left: 100, top: 50, tagData: JSON.stringify(config) });

    const rects = slide.created.filter((s) => s.geo === "rectangle");
    expect(rects.length).toBeGreaterThanOrEqual(4); // one per stacked segment
    for (const r of rects) expect(r.box.left).toBeGreaterThanOrEqual(100);
    const group = slide.created.find((s) => s.type === "group")!;
    expect(group.name).toBe("PowerChart");
    expect(group.grouped).toHaveLength(slide.created.length - 1);
    expect(group.tagStore.get(CHART_TAG)).toBe(JSON.stringify(config));
  });

  it("describes the chart group with accessible alt text", async () => {
    const slide = makeSlide("s1");
    installHost([slide]);
    const scene = buildChart(config);
    await insertSceneIntoSlide(scene, { tagData: JSON.stringify(config) });
    const group = slide.created.find((s) => s.type === "group") as FakeShape & {
      altTextDescription?: string;
      altTextTitle?: string;
    };
    expect(group.altTextDescription).toBe(scene.desc);
    if (scene.title) expect(group.altTextTitle).toBe(scene.title);
  });

  it("describes the chart even when it is NOT grouped", async () => {
    // The alt text used to be assigned only on the group object, so every
    // ungrouped chart (group:false, a refused addGroup, a one-shape chart)
    // silently lost its text alternative. It belongs on whatever shape stands
    // for the chart — the same one the config tag lands on.
    const slide = makeSlide("s1");
    installHost([slide]);
    const scene = buildChart(config);
    await insertSceneIntoSlide(scene, { tagData: "cfg", group: false });
    expect(slide.created.some((s) => s.type === "group")).toBe(false);
    const anchor = slide.created[0] as FakeShape & { altTextDescription?: string; altTextTitle?: string };
    expect(anchor.altTextDescription).toBe(scene.desc);
    expect(anchor.tagStore.get(CHART_TAG)).toBe("cfg");
  });

  it("skips alt text on a host below 1.10 instead of losing the group with it", async () => {
    // Shape.altTextDescription is PowerPointApi 1.10. Assigning it on an older
    // host is a queued command rejected at the next sync — the same sync that
    // carries the grouping, so an ungated assignment costs the group.
    const slide = makeSlide("s1");
    installHost([slide], [], slide, (v) => v !== "1.10");
    await insertSceneIntoSlide(buildChart(config), { tagData: "cfg" });
    const group = slide.created.find((s) => s.type === "group") as FakeShape & { altTextDescription?: string };
    expect(group).toBeDefined();
    expect(group.altTextDescription).toBeUndefined();
  });

  it("renders a pie as a rotated triangle fan", async () => {
    const slide = makeSlide("s1");
    installHost([slide]);
    const scene = buildChart({
      ...config,
      kind: "pie",
      data: { categories: ["A", "B"], series: [{ name: "S", values: [3, 1] }] },
    });
    await insertSceneIntoSlide(scene, {});
    const tris = slide.created.filter((s) => s.geo === "triangle" && s.name?.includes("-f"));
    expect(tris.length).toBeGreaterThan(10);
    for (const t of tris) expect(typeof t.rotation).toBe("number");
  });

  it("maps title font and alignment onto text boxes", async () => {
    const slide = makeSlide("s1");
    installHost([slide]);
    await insertSceneIntoSlide(buildChart({ ...config, title: "Hello" }), { fontFamily: "Arial" });
    const title = slide.created.find((s) => s.text === "Hello")!;
    expect(title.fillCleared).toBe(true);
    expect(title.textFrame.textRange.font).toMatchObject({ name: "Arial", bold: true });
  });

  it("draws value lines as dashed native connectors", async () => {
    const slide = makeSlide("s1");
    installHost([slide]);
    await insertSceneIntoSlide(
      buildChart({ ...config, decorations: { valueLines: [{ mode: "mean" }], segmentLabels: true } }),
      {},
    );
    const dashed = slide.created.filter((s) => s.type === "line" && s.lineFormat.dashStyle === "dash");
    expect(dashed.length).toBeGreaterThanOrEqual(1);
  });
});

describe("scene node mapping", () => {
  const insert = async (nodes: object[], opts = {}) => {
    const slide = makeSlide("s1");
    installHost([slide]);
    await insertSceneIntoSlide({ width: 200, height: 100, nodes } as never, opts);
    return slide;
  };

  it("maps ellipses with stroke or hidden outline", async () => {
    const slide = await insert([
      {
        kind: "ellipse",
        cx: 50,
        cy: 50,
        rx: 20,
        ry: 10,
        fill: "#ff0000",
        stroke: "#000000",
        strokeWidth: 2,
        name: "dot",
      },
      { kind: "ellipse", cx: 10, cy: 10, rx: 5, ry: 5, fill: "#00ff00" },
    ]);
    const [a, b] = slide.created.filter((s) => s.geo === "ellipse");
    // center − radius, plus the default 60/90pt insert offset
    expect(a.box).toEqual({ left: 90, top: 130, width: 40, height: 20 });
    expect(a.lineFormat).toMatchObject({ color: "#000000", weight: 2 });
    expect(b.lineFormat.visible).toBe(false);
  });

  it("honours an 8-digit #RRGGBBAA fill: 6-digit hue + transparency, never mis-parsed", async () => {
    // #RRGGBBAA is a valid hand-authored colour that the SVG preview and the
    // skill's pptx render translucent. Office.js setSolidColor validates 6-digit
    // hex only, so the alpha byte has to move to fill.transparency (1.4) or the
    // live add-in would mis-parse the value and lose the colour.
    const slide = await insert([{ kind: "rect", x: 0, y: 0, w: 10, h: 10, fill: "#4e79a780", name: "band" }]);
    const rect = slide.created.find((s) => s.geo === "rectangle") as (typeof slide.created)[number] & {
      fill: { transparency?: number };
    };
    expect(rect.fillColor).toBe("#4e79a7"); // hue survives, no 8-digit string reaches Office.js
    expect(rect.fill.transparency).toBeCloseTo(1 - 128 / 255, 3);
  });

  it("leaves an opaque 6-digit fill untouched (no transparency set)", async () => {
    const slide = await insert([{ kind: "rect", x: 0, y: 0, w: 10, h: 10, fill: "#4e79a7", name: "bar" }]);
    const rect = slide.created.find((s) => s.geo === "rectangle") as (typeof slide.created)[number] & {
      fill: { transparency?: number };
    };
    expect(rect.fillColor).toBe("#4e79a7");
    expect(rect.fill.transparency).toBeUndefined();
  });

  it("renders a fill:'none' rect as an outlined/hollow shape (IBCS plan columns)", async () => {
    const slide = await insert([
      { kind: "rect", x: 0, y: 0, w: 20, h: 40, fill: "none", stroke: "#3b6ea5", strokeWidth: 1.5, name: "pl" },
    ]);
    const rect = slide.created.find((s) => s.geo === "rectangle")!;
    expect(rect.fillCleared).toBe(true); // no fill
    expect(rect.fillColor).toBeNull(); // never mis-parsed as a colour
    expect(rect.lineFormat).toMatchObject({ color: "#3b6ea5", weight: 1.5 });
  });

  it("maps chevrons to chevron/homePlate geometry", async () => {
    const slide = await insert([
      { kind: "chevron", x: 0, y: 0, w: 40, h: 20, fill: "#123456", flatLeft: true },
      { kind: "chevron", x: 50, y: 0, w: 40, h: 20, fill: "#123456" },
    ]);
    expect(slide.created.filter((s) => s.type !== "group").map((s) => s.geo)).toEqual(["homePlate", "chevron"]);
  });

  it("draws axis-aligned lines with a clamped non-zero box (never a degenerate diagonal)", async () => {
    const slide = await insert([
      {
        kind: "line",
        x1: 10,
        y1: 50,
        x2: 200,
        y2: 50,
        stroke: "#333333",
        strokeWidth: 1,
        dash: [3, 2],
        name: "connector",
      },
    ]);
    const line = slide.created.find((s) => s.type === "line")!;
    // Horizontal line: width spans, height is clamped up from 0 so the web host
    // can't blow a zero-thickness box into a giant diagonal.
    expect(line.box.width).toBeGreaterThan(180);
    expect(line.box.height).toBeGreaterThanOrEqual(0.5);
    expect(line.lineFormat.dashStyle).toBe("dash");
  });

  it("honours a translucent stroke on BOTH the line and the rotated-rect branch", async () => {
    // The two branches of addSegment diverged: the axis-aligned/dashed line
    // dropped the alpha byte while the diagonal rotated-rect folded it into
    // transparency, so one series colour rendered at two opacities. Both must now
    // agree — a 50% alpha (#…80) → transparency ≈ 0.498, colour a bare 6-digit hex.
    const alpha = "#33333380";
    const horiz = await insert([{ kind: "line", x1: 10, y1: 50, x2: 200, y2: 50, stroke: alpha, name: "h" }]);
    const line = horiz.created.find((s) => s.type === "line")!;
    expect(line.lineFormat.color).toBe("#333333"); // alpha byte stripped off the colour
    expect(line.lineFormat.transparency).toBeCloseTo(1 - 0x80 / 255, 4); // …but carried here

    const diag = await insert([{ kind: "line", x1: 0, y1: 0, x2: 100, y2: 100, stroke: alpha, name: "d" }]);
    const rect = diag.created.find((s) => s.geo === "rectangle")!;
    expect(rect.fillColor).toBe("#333333");
    expect((rect.fill as unknown as { transparency?: number }).transparency).toBeCloseTo(1 - 0x80 / 255, 4);
  });

  it("draws diagonal lines as thin rotated rectangles (direction-correct on every host)", async () => {
    // Up-right and down-right diagonals a bounding box alone can't distinguish.
    const down = await insert([
      { kind: "line", x1: 0, y1: 0, x2: 100, y2: 100, stroke: "#a00000", strokeWidth: 2, name: "d" },
    ]);
    const dr = down.created.find((s) => s.geo === "rectangle")!;
    expect(dr).toBeTruthy();
    expect(dr.fillColor).toBe("#a00000");
    expect(dr.rotation).toBeCloseTo(45, 0); // down-right
    expect(down.created.some((s) => s.type === "line")).toBe(false);

    const up = await insert([
      { kind: "line", x1: 0, y1: 100, x2: 100, y2: 0, stroke: "#00a000", strokeWidth: 2, name: "u" },
    ]);
    const ur = up.created.find((s) => s.geo === "rectangle")!;
    expect(ur.rotation).toBeCloseTo(-45, 0); // up-right — the case a box would mirror
  });

  it("draws dashed diagonals as real line shapes, picking the geometry per direction", async () => {
    // A rotated rectangle carries its colour in its fill, which can't be
    // dashed — scatter trend lines and forecast segments came out solid.
    const down = await insert([
      {
        kind: "line",
        x1: 0,
        y1: 0,
        x2: 100,
        y2: 60,
        stroke: "#a00000",
        strokeWidth: 1.25,
        dash: [4, 2],
        name: "trend",
      },
    ]);
    const dl = down.created.find((s) => s.name === "trend")!;
    expect(dl.type).toBe("line"); // not a filled rectangle
    expect(dl.lineFormat.dashStyle).toBe("dash");
    expect(dl.box).toMatchObject({ width: 100, height: 60 });

    // Up-right: addLine only ever draws the box's top-left→bottom-right
    // diagonal, so this direction needs the lineInverse geometry.
    const up = await insert([
      {
        kind: "line",
        x1: 0,
        y1: 60,
        x2: 100,
        y2: 0,
        stroke: "#a00000",
        strokeWidth: 1.25,
        dash: [4, 2],
        name: "trend",
      },
    ]);
    const ul = up.created.find((s) => s.name === "trend")!;
    expect(ul.geo).toBe("lineInverse");
    expect(ul.lineFormat.dashStyle).toBe("dash");
  });

  it("renders a dotted array as roundDot, not a generic dash", async () => {
    // [1.5,1.5] is the dotted waterfall carry connector. It used to flatten to
    // the same dash enum as every other pattern, so the deck lost the dotted
    // look the SVG preview shows.
    const slide = await insert([
      {
        kind: "line",
        x1: 10,
        y1: 50,
        x2: 200,
        y2: 50,
        stroke: "#333",
        strokeWidth: 1,
        dash: [1.5, 1.5],
        name: "carry",
      },
    ]);
    const line = slide.created.find((s) => s.name === "carry")!;
    expect(line.lineFormat.dashStyle).toBe("roundDot");
  });

  it("draws polygon edges direction-correct, with no zero-thickness boxes", async () => {
    // A violin body: an up-right edge, a horizontal edge and a down-right edge.
    const slide = await insert([
      {
        kind: "polygon",
        points: [
          { x: 0, y: 40 },
          { x: 50, y: 0 },
          { x: 100, y: 40 },
          { x: 100, y: 40 },
        ],
        fill: "#eeeeee",
        stroke: "#3366cc",
        strokeWidth: 1,
        name: "violin-0",
      },
    ]);
    const edges = slide.created.filter((s) => s.name?.startsWith("violin-0-e"));
    expect(edges).toHaveLength(4);
    for (const e of edges) {
      // Every edge is a real segment: a bounding box collapsed to zero on one
      // axis let the web host blow it up into a giant diagonal.
      expect(e.box.width).toBeGreaterThanOrEqual(0.5);
      expect(e.box.height).toBeGreaterThanOrEqual(0.5);
    }
    // Edge 0 (0,40)->(50,0) rises to the right; edge 1 (50,0)->(100,40) falls.
    // Passing both bounding boxes to addLine drew them as the same diagonal.
    const [e0, e1] = edges;
    expect(e0.rotation).toBeLessThan(0);
    expect(e1.rotation).toBeGreaterThan(0);
  });

  it("maps arrowheads to rotated triangles anchored at the tip", async () => {
    const slide = await insert([{ kind: "arrowhead", x: 10, y: 10, size: 4, angle: 45, fill: "#000000", name: "ah" }]);
    const tri = slide.created[0];
    expect(tri.geo).toBe("triangle");
    expect(tri.rotation).toBe(135); // scene angle + 90
    // The triangle's tip (box top-centre, rotated θ about the box centre) must
    // land on the scene point (10,10) + the default 60/90pt insert offset.
    const s = 8; // size * 2
    const theta = (tri.rotation! * Math.PI) / 180;
    const cx = tri.box.left + s / 2;
    const cy = tri.box.top + s / 2;
    const tipX = cx + (s / 2) * Math.sin(theta);
    const tipY = cy - (s / 2) * Math.cos(theta);
    expect(tipX).toBeCloseTo(70, 4); // 60 + 10
    expect(tipY).toBeCloseTo(100, 4); // 90 + 10
  });

  it("renders an annular wedge (sunburst ring / gauge) as a rotated rectangle band", async () => {
    const slide = await insert([
      {
        kind: "wedge",
        cx: 50,
        cy: 50,
        r: 30,
        innerR: 15,
        startAngle: 0,
        endAngle: 90,
        fill: "#333333",
        stroke: "#ffffff",
        strokeWidth: 1,
        name: "ring",
      },
    ]);
    const band = slide.created.filter((s) => s.geo === "rectangle" && s.name?.includes("-f"));
    expect(band.length).toBeGreaterThan(2); // the annular band, not a triangle fan
    for (const b of band) {
      expect(b.fillColor).toBe("#333333");
      expect(typeof b.rotation).toBe("number");
    }
    // No triangles for an annular wedge (a triangle can't leave a hole).
    expect(slide.created.some((s) => s.geo === "triangle")).toBe(false);
    // Two radial separators in the stroke colour.
    const edges = slide.created.filter((s) => s.name === "ring-edge");
    expect(edges.length).toBe(2);
    for (const e of edges) expect(e.fillColor).toBe("#ffffff");
  });

  it("skips grouping when group:false or only one shape", async () => {
    const slide = await insert(
      [
        { kind: "rect", x: 0, y: 0, w: 10, h: 10, fill: "#111111" },
        { kind: "rect", x: 20, y: 0, w: 10, h: 10, fill: "#222222" },
      ],
      { group: false, tagData: "cfg" },
    );
    expect(slide.created.some((s) => s.type === "group")).toBe(false);
    // The tag falls back onto the first created shape.
    expect(slide.created[0].tagStore.get(CHART_TAG)).toBe("cfg");
  });

  // Shape.rotation is PowerPointApi 1.10 and the manifests admit hosts from 1.4.
  // A try/catch around the assignment catches NOTHING on a real host: Office.js
  // proxy setters do not throw synchronously — the host rejects the queued
  // command at the next context.sync(), which carries the whole batch. So the
  // fake here rejects at SYNC, the way PowerPoint does, not at the setter.
  it.each(["pie", "doughnut", "sunburst"])(
    "%s still inserts on a pre-1.10 host (rotation gated, not wrapped)",
    async (kind) => {
      const slide = makeSlide("s-old");
      const ctx = installHost([slide], [], slide, (v) => v !== "1.10");
      const realSync = ctx.sync;
      ctx.sync = async () => {
        // Any rotation assigned on a host without 1.10 poisons the whole batch.
        if (slide.created.some((sh) => sh.rotation !== undefined)) {
          throw new Error("PropertyNotSupported: Shape.rotation requires PowerPointApi 1.10");
        }
        return realSync();
      };
      const scene = buildChart(sampleConfig(kind as never));
      // Must not reject: the chart degrades (no wedges) instead of failing.
      let failure: unknown = null;
      try {
        await insertSceneIntoSlide(scene, { tagData: "{}" });
      } catch (e) {
        failure = e;
      }
      expect(failure, `insert rejected on a pre-1.10 host: ${String(failure)}`).toBeNull();
      // The rest of the chart — labels, leader lines — still lands on the slide.
      expect(slide.created.length).toBeGreaterThan(0);
      // …and nothing carries a rotation the host cannot accept.
      expect(slide.created.every((sh) => sh.rotation === undefined)).toBe(true);
    },
  );

  it("degrades gracefully when the host lacks grouping and rotation", async () => {
    const slide = makeSlide("s1");
    installHost([slide]);
    // Break addGroup and rotation assignment the way an old host would.
    slide.shapes.addGroup = () => {
      throw new Error("addGroup requires PowerPointApi 1.8");
    };
    const scene = {
      width: 200,
      height: 100,
      nodes: [
        { kind: "rect", x: 0, y: 0, w: 10, h: 10, fill: "#111111" },
        { kind: "wedge", cx: 50, cy: 50, r: 30, innerR: 0, startAngle: 0, endAngle: 90, fill: "#333333", name: "w" },
      ],
    };
    const realAdd = slide.shapes.addGeometricShape.bind(slide.shapes);
    slide.shapes.addGeometricShape = (geo, box) => {
      const s = realAdd(geo, box);
      if (geo === "triangle") {
        Object.defineProperty(s, "rotation", {
          set() {
            throw new Error("rotation requires PowerPointApi 1.10");
          },
        });
      }
      return s;
    };
    await insertSceneIntoSlide(scene as never, { tagData: "cfg" });
    // No group, no fan triangles survive — but the rect is inserted and tagged.
    expect(slide.created.some((s) => s.type === "group")).toBe(false);
    expect(slide.created[0].tagStore.get(CHART_TAG)).toBe("cfg");
  });

  it("still inserts (ungrouped) when the host lacks grouping — the web case", async () => {
    // PowerPoint on the web: grouping (1.8) unsupported, tags (1.3) supported.
    const slide = makeSlide("s1");
    installHost([slide], [], slide, (v) => v !== "1.8");
    await insertSceneIntoSlide(buildChart(config), { tagData: "cfg" });
    // The shapes are committed and no grouping was attempted…
    expect(slide.created.some((s) => s.type === "group")).toBe(false);
    expect(slide.created.filter((s) => s.geo === "rectangle").length).toBeGreaterThanOrEqual(4);
    // …and the config tag lands on the first shape, so the chart is re-editable.
    expect(slide.created[0].tagStore.get(CHART_TAG)).toBe("cfg");
  });

  it("skips tagging when the host lacks tags", async () => {
    const slide = makeSlide("s1");
    installHost([slide], [], slide, () => false); // nothing supported
    await insertSceneIntoSlide(buildChart(config), { tagData: "cfg" });
    expect(slide.created.some((s) => s.type === "group")).toBe(false);
    expect(slide.created[0].tagStore.get(CHART_TAG)).toBeUndefined();
    expect(slide.created.length).toBeGreaterThan(0);
  });

  it("falls back to the first slide when nothing is selected", async () => {
    const slide = makeSlide("s1");
    const ctx = installHost([slide]);
    ctx.presentation.getSelectedSlides = () => {
      throw new Error("no selection");
    };
    await insertSceneIntoSlide(
      { width: 10, height: 10, nodes: [{ kind: "rect", x: 0, y: 0, w: 5, h: 5, fill: "#111111" }] } as never,
      {},
    );
    expect(slide.created).toHaveLength(1);
  });
});

describe("looking away while a chart redraws", () => {
  /** Wire a controllable selection onto the fake host. */
  function withSelection(ctx: ReturnType<typeof installHost>, initial: string[]) {
    const state = { selected: [...initial], sets: [] as string[][] };
    const p = ctx.presentation as unknown as Record<string, unknown>;
    p.getSelectedSlides = () => ({ items: state.selected.map((id) => ({ id })), load() {} });
    p.setSelectedSlides = (ids: string[]) => {
      state.sets.push([...ids]);
      state.selected = [...ids];
    };
    return state;
  }

  it("puts the user back where they were", async () => {
    const ctx = installHost([makeSlide("s1"), makeSlide("s2")]);
    const sel = withSelection(ctx, ["s1"]);
    const saw = await withSlideDeselected(["s1"], async (deselected) => deselected);
    expect(saw).toBe(true);
    // Parked elsewhere for the redraw, then restored.
    expect(sel.sets).toEqual([["s2"], ["s1"]]);
    expect(sel.selected).toEqual(["s1"]);
  });

  it("leaves the selection alone when the user moved during the redraw", async () => {
    // An off-screen redraw runs for tens of seconds and the user is free to
    // click through the deck while it does. Restoring unconditionally snapped
    // them back to wherever they happened to be standing when it started,
    // throwing away their navigation with no notice.
    const ctx = installHost([makeSlide("s1"), makeSlide("s2"), makeSlide("s3")]);
    const sel = withSelection(ctx, ["s1"]);
    await withSlideDeselected(["s1"], async () => {
      sel.selected = ["s3"]; // the user clicks slide 3 mid-redraw
    });
    expect(sel.selected).toEqual(["s3"]);
    // Only the park was written; nothing restored over the user's own move.
    expect(sel.sets).toEqual([["s2"]]);
  });

  it("makes somewhere to look when every slide is one it is about to draw on", async () => {
    // The one-slide deck. It used to run `fn(false)` — the live canvas, batch
    // 10 — which is the exact configuration a real run died in, and it is the
    // deck a user building their first chart actually has. So make a slide to
    // look at, and take it away again.
    const deck = [makeSlide("s1")];
    const ctx = installHost(deck);
    const sel = withSelection(ctx, ["s1"]);
    const saw = await withSlideDeselected(["s1"], async (deselected) => {
      // MID-redraw: the scratch slide exists and the view is on it, which is
      // the whole point — asserting only the end state would pass against a
      // function that added and removed a slide without ever looking at it.
      expect(deck).toHaveLength(2);
      expect(sel.selected).toEqual([deck[1].id]);
      return deselected;
    });
    // True: the caller may legitimately spend the off-screen batch budget.
    expect(saw).toBe(true);
    // And the deck is exactly as the user left it.
    expect(deck).toHaveLength(1);
    expect(deck[0].id).toBe("s1");
    expect(sel.selected).toEqual(["s1"]);
  });

  it("draws on the live canvas when the scratch slide will not land", async () => {
    // The host swallows slides.add() under load — the behaviour addSlides
    // exists to survive. A scratch slide that never landed must not be
    // reported as parked: `fn` would then use the off-screen batch size on a
    // slide the user is still looking at, which is worse than not trying.
    const deck = [makeSlide("s1")];
    const ctx = installHost(deck);
    const sel = withSelection(ctx, ["s1"]);
    faults.swallowAdds = 1;
    try {
      const saw = await withSlideDeselected(["s1"], async (deselected) => deselected);
      expect(saw).toBe(false);
      expect(deck).toHaveLength(1);
      // Nothing was selected or restored — there was nowhere to go.
      expect(sel.sets).toEqual([]);
    } finally {
      faults.swallowAdds = 0;
    }
  });

  /**
   * The failure no `catch` can see: a sync that neither resolves nor rejects.
   *
   * `withSlideDeselected` was two raw `PowerPoint.run`s sandwiching draw work
   * that IS bounded per batch — so the one part of an in-place update that could
   * hang forever was the part before anything was drawn. And the second one sits
   * in a `finally`, so a sync that never settles there stops the `finally` from
   * completing, which means the CALLER's `finally` never runs either: the pane's
   * busy counter stays up for the rest of the session, the selection banner goes
   * dead, the status strip freezes, and the auto-update timer re-arms forever.
   * Stop cannot break in — it is checked at batch boundaries this never reaches.
   *
   * Raced against a timer rather than simply awaited, so the guard fails on its
   * own assertion instead of on a suite timeout.
   */
  const returnedWithin = async <T>(ms: number, work: Promise<T>): Promise<T | "never came back"> =>
    Promise.race([work, new Promise<"never came back">((r) => setTimeout(() => r("never came back"), ms))]);

  it("comes back from a redraw even when the host stops answering selection calls", async () => {
    // The selection API has to be wired, or the park throws before it ever
    // syncs and the wedge is never reached — a version of this that skipped
    // `withSelection` passed against an unbounded park, which is the whole
    // thing it was written to catch.
    const ctx = installHost([makeSlide("s1"), makeSlide("s2")]);
    withSelection(ctx, ["s1"]);
    _setSelectionTimeoutForTest(20);
    // From the first sync on, the host answers nothing at all — ever.
    faults.wedgeAfterSyncs = 0;
    try {
      const saw = await returnedWithin(
        200,
        withSlideDeselected(["s1"], async (deselected) => deselected),
      );
      expect(saw, "an in-place update never returned on a host that went quiet").not.toBe("never came back");
      // And it degraded honestly: nothing was parked, so the caller must NOT be
      // told it may use the off-screen batch size on a live canvas.
      expect(saw).toBe(false);
    } finally {
      faults.wedgeAfterSyncs = null;
      _setSelectionTimeoutForTest(4_000);
    }
  });

  it("comes back from the RESTORE too, which sits in a finally", async () => {
    // The second round trip is the worse one. A sync that never settles inside
    // a `finally` stops that `finally` from completing, so the caller's own
    // `finally` never runs either — and the pane decrements its busy counter in
    // one of those. The view not being restored is cheap; the update never
    // returning is what kills the session.
    const ctx = installHost([makeSlide("s1"), makeSlide("s2")]);
    withSelection(ctx, ["s1"]);
    _setSelectionTimeoutForTest(20);
    try {
      const saw = await returnedWithin(
        200,
        withSlideDeselected(["s1"], async (deselected) => {
          // The park has happened; go quiet from here, so only the restore hits it.
          faults.wedgeAfterSyncs = 0;
          return deselected;
        }),
      );
      expect(saw, "the redraw never returned because the restore hung").not.toBe("never came back");
      expect(saw, "did not park at all").toBe(true);
    } finally {
      faults.wedgeAfterSyncs = null;
      _setSelectionTimeoutForTest(4_000);
    }
  });

  /**
   * The draw was bounded per batch and everything AFTER the last batch was not.
   *
   * `groupAndTagAll`'s five syncs and `ungroupedFallback`'s bare one all went
   * through `step()`, which labels and adds no deadline — so a host that went
   * quiet once the shapes were on the slide left the insert with no timer, no
   * `gave up waiting`, no phase note past "grouping…", and no way for Stop to
   * break in (it is checked at batch boundaries this never reaches). That is the
   * 1819-second wedge shape, one phase further along than the one that got
   * bounded. `ungroupedFallback` matters most: grouping is refused on the web,
   * so every web-host insert goes down it.
   */
  it("comes back from grouping and tagging when the host goes quiet after the draw", async () => {
    installHost([makeSlide("s1")]);
    _setBatchTimeoutForTest(20);
    // Past the context and the first draw batches; the shapes are committed and
    // the host stops answering from here on.
    faults.wedgeAfterSyncs = 3;
    try {
      const got = await returnedWithin(400, insertSceneIntoSlide(buildChart(config), { tagData: "{}" }));
      expect(got, "the insert never returned once the host went quiet after drawing").not.toBe("never came back");
    } finally {
      faults.wedgeAfterSyncs = null;
      _setBatchTimeoutForTest(45_000);
    }
  });

  it("comes back from the Insert click's reads when the host stops answering", async () => {
    // `getSelectionBounds` was the only selection read in the file not on
    // `boundedRun`, and it is the FIRST host call the Insert button makes —
    // so a quiet host took the whole insert with it: buttons disabled,
    // "Working…" counting up, nothing drawn and nothing said. `guard()` has no
    // deadline on the action either, so there was nothing else to stop it.
    installHost([makeSlide("s1")]);
    _setSelectionTimeoutForTest(20);
    faults.wedgeAfterSyncs = 0;
    try {
      expect(await returnedWithin(200, getSelectionBounds()), "getSelectionBounds never returned").toBe(null);
      expect(await returnedWithin(200, getSlideShapeBounds()), "getSlideShapeBounds never returned").toBe(null);
    } finally {
      faults.wedgeAfterSyncs = null;
      _setSelectionTimeoutForTest(4_000);
    }
  });

  /**
   * A slide the host will not hand back by id is not necessarily gone, and
   * treating it as gone is how an add-in litters someone's deck.
   *
   * PowerPoint on the web resolved a freshly-added slide's id once and then
   * refused it — while still listing that same id in `slides.load("items/id")`.
   * `deleteSlideById` read the refusal as "already gone, nothing to do" and
   * reported success, so a host-probe run left fourteen blank slides behind and
   * said it had cleaned up. The same call cleans up after every off-screen
   * redraw, on the user's own deck.
   */
  it("takes out a slide the host will not resolve by id", async () => {
    const deck = [makeSlide("s1")];
    installHost(deck);
    const id = await addScratchSlide();
    expect(id, "the scratch slide did not land").toBeTruthy();
    expect(deck).toHaveLength(2);
    // From here on the host denies that slide exists whenever it is asked for
    // by id — while still listing it among the deck's slides.
    faults.newSlideResolvesTimes = 0;
    try {
      expect(await deleteSlideById(id!), "reported a clean-up it had not done").toBe(true);
      expect(
        deck.map((s) => s.id),
        "the slide is still in the deck",
      ).toEqual(["s1"]);
    } finally {
      faults.newSlideResolvesTimes = null;
    }
  });

  it("deletes nothing when the id is not in the deck at all, and does not claim it did", async () => {
    // The other half of the same question. A positional delete driven by an id
    // nobody can find is how an add-in destroys work, so an id the deck does
    // not list must end the search — nothing to remove, and nothing removed.
    // That half is unchanged and is the one that protects the user's deck.
    //
    // What CHANGED is the verdict it reports. "Not in the deck's list" used to
    // return true, i.e. "already gone", and 2026-08-11 (`756682e`) measured
    // that reading false on this host: of the 62 scratch slides a probe run was
    // deleting, `the deck still lists 0 of 62 of these ids` — zero — while the
    // deck stayed at 65 slides and the run reported a clean sweep. Both id
    // lists come from the same `slideIds()` projection minutes apart, so "the
    // id is not there" cannot mean the slide is not there.
    //
    // Unfindable is UNKNOWN. For an id that genuinely never existed this is now
    // a shade pessimistic, and that is the right way to be wrong: an
    // under-count costs a line in a report, while an over-count leaves sixty
    // blank slides in someone's deck and says it left none. The caller's deck
    // count (`slidesActuallyReturned`) is what turns this honest false back
    // into an honest number.
    const deck = [makeSlide("s1"), makeSlide("s2")];
    installHost(deck);
    expect(await deleteSlideById("no-such-slide"), "an id nobody can find was reported as confirmed gone").toBe(false);
    expect(
      deck.map((s) => s.id),
      "a delete driven by an unfindable id touched the deck",
    ).toEqual(["s1", "s2"]);
  });

  /**
   * The slide-swap gate authorises DELETING the user's slide. It has to be sure.
   *
   * `slideHoldsOnlyChart` read `slide.shapes.items` and answered yes for an
   * empty list — and it is consulted only AFTER the host has already stalled,
   * which is exactly the state in which this host answers shape collections
   * short (`shapesExpected=19 shapesSeen=15`). A hollow read therefore looked
   * identical to a bare slide, and the swap replaced a slide holding the user's
   * logo, title and footnote with a generated one carrying none of it and no
   * speaker notes either. Recoverable only by Ctrl-Z, and the note the user
   * gets does not mention shapes.
   */
  it("refuses the slide swap when it cannot corroborate what is on the slide", async () => {
    const slide = makeSlide("s1");
    // Two real shapes the user put there...
    slide.created.push(makeShape("geometric", "rectangle", { left: 0, top: 0, width: 10, height: 10 }));
    slide.created.push(makeShape("geometric", "rectangle", { left: 20, top: 0, width: 10, height: 10 }));
    installHost([slide]);
    // ...and a host that answers the collection empty while still counting two.
    faults.hollowNameReads = 1;
    try {
      expect(await slideHoldsOnlyChart("s1"), "believed a hollow read and offered the swap").toBe(false);
    } finally {
      faults.hollowNameReads = 0;
    }
  });

  it("still allows the swap on a slide it CAN corroborate as empty", async () => {
    // The gate has to stay usable, or the fallback it guards is dead code. A
    // bare slide, honestly read, is still a yes.
    installHost([makeSlide("s1")]);
    expect(await slideHoldsOnlyChart("s1")).toBe(true);
  });

  /**
   * A refused lookup is not "the original slide is gone".
   *
   * The swap used to open its own context, ask for the target by id, and throw
   * `original slide is gone` the moment the null flag was not `false` — on the
   * very id `insertSlidesFromBase64` had just accepted as `targetSlideId`, with
   * the count check already proving the deck had grown by one. So the original
   * was demonstrably still there, and the user was told to sort out two
   * identical slides by hand, permanently. This host does exactly that: it
   * resolves a slide id once and refuses it after, while still listing it.
   */
  it("removes the slide it replaced even when the host will not resolve its id", async () => {
    const built = await buildDeckBase64(
      [{ scene: buildChart(sampleConfig("clustered")), title: "A", configJson: "{}", slot: 0, run: "r1" }],
      { width: 720, height: 405 },
    );
    const deck = [makeSlide("s1"), makeSlide("s2")];
    installHost(deck);
    // The host takes the null check for s1 and answers nothing — neither "gone"
    // nor "there", which is the state `isLive` is deliberately pessimistic about.
    unansweredNullChecks.add("s1");
    try {
      expect(await replaceSlideWithDeck("s1", built.base64)).toBe("swapped");
      expect(
        deck.map((s) => s.id),
        "left the replaced slide in the deck",
      ).not.toContain("s1");
    } finally {
      unansweredNullChecks.clear();
    }
  });

  /**
   * A stray the host will not resolve has not been swept, and saying so matters
   * to three separate callers.
   *
   * Every id handed to this sweep names a shape that COMMITTED, so it is on the
   * slide by construction. The old code read `s.isNullObject` raw — which on an
   * unanswered proxy throws `PropertyNotLoaded`, took the whole sweep down, and
   * left the catch reporting 0 for the addressable strays as well. The picture
   * fallback then drew over the debris, the slide-swap gate saw a slide that was
   * not clean, and Stop's promise that nothing is left behind rested on a count
   * that could not tell "nothing to do" from "could not touch any of it".
   */
  it("sweeps the strays it can reach even when one of them will not answer", async () => {
    const slide = makeSlide("s1");
    installHost([slide]);
    const a = slide.shapes.addGeometricShape("rectangle", { left: 0, top: 0, width: 5, height: 5 });
    const b = slide.shapes.addGeometricShape("rectangle", { left: 6, top: 0, width: 5, height: 5 });
    unansweredNullChecks.add(a.id); // this one the host will not describe
    try {
      const swept = await deleteShapesById("s1", [a.id, b.id]);
      expect(swept, "one quiet stray took the whole sweep down").toBe(1);
      expect(b.deleted, "the reachable stray was left on the slide").toBe(true);
      expect(a.deleted, "deleted a shape the host would not confirm").toBe(false);
    } finally {
      unansweredNullChecks.clear();
    }
  });

  it("takes back a scratch slide whose id it cannot verify", async () => {
    // `addScratchSlide` refuses to hand out an id it could not resolve — but
    // the slide landed, so refusing without also removing it would leave a
    // blank slide in the deck on every attempt.
    const deck = [makeSlide("s1")];
    installHost(deck);
    // The fake names an added slide after the deck's length; the null-check on
    // it goes unanswered, which is the state a caller can least reason about.
    unansweredNullChecks.add("slide-2");
    try {
      expect(await addScratchSlide(), "handed out an id it could not resolve").toBeNull();
      expect(
        deck.map((s) => s.id),
        "left the unusable scratch slide behind",
      ).toEqual(["s1"]);
    } finally {
      unansweredNullChecks.clear();
    }
  });

  it("removes the scratch slide even when the user navigates away mid-redraw", async () => {
    // The restore is deliberately skipped when the user has moved themselves.
    // The CLEANUP is not conditional on it: a blank slide the add-in left at
    // the end of someone's deck is litter whether or not they are looking at
    // it, and tying the delete to the restore would leak one every time.
    const deck = [makeSlide("s1")];
    const ctx = installHost(deck);
    const sel = withSelection(ctx, ["s1"]);
    await withSlideDeselected(["s1"], async () => {
      sel.selected = ["s1"]; // the user clicks back to their own slide mid-redraw
    });
    expect(deck).toHaveLength(1);
    expect(sel.selected).toEqual(["s1"]);
  });
});

describe("updateChartInSlide", () => {
  it("deletes the old group and re-renders at the same position", async () => {
    const slide = makeSlide("s1");
    installHost([slide]);
    await insertSceneIntoSlide(buildChart(config), { tagData: "x" });
    const oldGroup = slide.created.find((s) => s.type === "group")!;
    const before = slide.created.length;
    await updateChartInSlide(buildChart(config), { slideId: "s1", shapeId: oldGroup.id, left: 33, top: 44 });
    expect(oldGroup.deleted).toBe(true);
    const fresh = slide.created.slice(before).filter((s) => s.type !== "group");
    expect(fresh.length).toBeGreaterThan(0);
    expect(Math.min(...fresh.map((s) => s.box.left))).toBeGreaterThanOrEqual(33);
  });

  it("deletes the WHOLE chart, not just its tagged shape, when it is ungrouped", async () => {
    // PowerPoint on the web: no grouping, so the config tag can only sit on ONE
    // of the chart's shapes. The update deleted exactly that shape and redrew
    // the chart, leaving the other twelve underneath — 13 shapes became 25, then
    // 37, on successive edits, as stacked misaligned duplicates.
    const slide = makeSlide("s1");
    installHost([slide], [], slide, (v) => v !== "1.8");
    const scene = buildChart(config);
    await insertSceneIntoSlide(scene, { tagData: "cfg" });
    expect(slide.created.some((s) => s.type === "group")).toBe(false);
    const drawn = slide.created.length;
    expect(drawn).toBeGreaterThan(1);
    // The tagged shape carries the rest of the chart with it.
    expect(JSON.parse(slide.created[0].tagStore.get(CHART_PARTS_TAG)!)).toHaveLength(drawn - 1);

    const live = () => slide.created.filter((s) => !s.deleted);
    for (const edit of [1, 2]) {
      // Same Scale: read the deck back, re-render every chart it finds.
      const found = (await listChartsInDeck()).charts.filter((c) => live().some((s) => s.id === c.target.shapeId));
      expect(found, `edit ${edit}`).toHaveLength(1);
      await updateChartsInSlides([{ scene, target: found[0].target, opts: { tagData: "cfg" } }]);
      expect(live(), `edit ${edit}`).toHaveLength(drawn);
    }
  });

  it("does not resurrect a chart whose shape the user deleted", async () => {
    // The pane still holds an editTarget for a chart the user has since removed
    // from the slide. A stale SLIDE id is already treated as nothing to do; a
    // stale SHAPE id redrew the chart at its old position instead.
    const slide = makeSlide("s1");
    installHost([slide]);
    await insertSceneIntoSlide(buildChart(config), { tagData: "cfg" });
    const group = slide.created.find((s) => s.type === "group")!;
    for (const s of slide.created) s.delete(); // the user deletes the chart
    await updateChartInSlide(buildChart(config), { slideId: "s1", shapeId: group.id, left: 10, top: 20 });
    expect(slide.created.filter((s) => !s.deleted)).toHaveLength(0);
  });

  /**
   * Grouping and tagging are best-effort by design — the shapes are on the
   * slide before either runs, and the catches around them say exactly that. But
   * the statements that QUEUE the work sat outside those catches, so a host
   * that threw while being asked to re-read a shape collection ("e.load is not
   * a function", seen on the web) rejected the whole request context and failed
   * an update that had, on the slide, already succeeded. Same Scale across the
   * deck reported one TypeError for four redrawn charts.
   */
  it("keeps a redraw that landed, when the host faults on the re-read after it", async () => {
    const slide = makeSlide("s1");
    installHost([slide]);
    const scene = buildChart(config);
    await insertSceneIntoSlide(scene, { tagData: "cfg" });
    const group = slide.created.find((s) => s.type === "group")!;
    const before = slide.created.filter((s) => !s.deleted).length;

    faults.faultShapeCollectionLoad = true;
    // Small batches, so the redraw spans several and asks for the re-read that
    // faults — the exact condition on a live canvas.
    await expect(
      updateChartInSlide(scene, { slideId: "s1", shapeId: group.id, left: 10, top: 20 }, { shapesPerSync: 2 }),
    ).resolves.not.toThrow();
    faults.faultShapeCollectionLoad = false;

    // The chart is on the slide: the old one gone, a new one drawn. What the
    // fault costs is the group and the tag, not the chart.
    expect(group.deleted).toBe(true);
    expect(slide.created.filter((s) => !s.deleted).length).toBeGreaterThanOrEqual(before - 1);
  });
});

describe("selection readers", () => {
  it("loadChartFromSelection returns the tagged config and target", async () => {
    const slide = makeSlide("s1");
    const chart = makeShape("group", undefined, { left: 10, top: 20, width: 300, height: 200 });
    chart.tagStore.set(CHART_TAG, '{"kind":"pie"}');
    const other = makeShape("geometric", "rectangle", { left: 0, top: 0, width: 5, height: 5 });
    installHost([slide], [other, chart]);
    const res = await loadChartFromSelection();
    expect(res?.configJson).toBe('{"kind":"pie"}');
    expect(res?.target).toMatchObject({ slideId: "s1", shapeId: chart.id, left: 10, top: 20 });
  });

  it("loadChartFromSelection returns null for untagged selections", async () => {
    const slide = makeSlide("s1");
    installHost([slide], [makeShape("geometric", "rectangle", { left: 0, top: 0, width: 5, height: 5 })]);
    expect(await loadChartFromSelection()).toBeNull();
  });

  it("getSelectionBounds returns plain shape bounds but skips charts and multi-selects", async () => {
    const slide = makeSlide("s1");
    const box = makeShape("geometric", "rectangle", { left: 7, top: 8, width: 100, height: 60 });
    installHost([slide], [box]);
    expect(await getSelectionBounds()).toEqual({ left: 7, top: 8, width: 100, height: 60 });

    box.tagStore.set(CHART_TAG, "{}");
    installHost([slide], [box]);
    expect(await getSelectionBounds()).toBeNull();

    installHost([slide], [box, makeShape("geometric", "rectangle", { left: 0, top: 0, width: 1, height: 1 })]);
    expect(await getSelectionBounds()).toBeNull();
  });

  it("getSelectionBounds swallows host errors", async () => {
    vi.stubGlobal("PowerPoint", {
      run: async () => {
        throw new Error("no selection");
      },
    });
    expect(await getSelectionBounds()).toBeNull();
  });

  it("listChartsInSelection filters to tagged shapes", async () => {
    const slide = makeSlide("s1");
    const a = makeShape("group", undefined, { left: 1, top: 1, width: 1, height: 1 });
    a.tagStore.set(CHART_TAG, "{}");
    const b = makeShape("geometric", "rectangle", { left: 2, top: 2, width: 1, height: 1 });
    installHost([slide], [a, b]);
    const res = await listChartsInSelection();
    expect(res).toHaveLength(1);
    expect(res[0].target.shapeId).toBe(a.id);
  });
});

describe("listChartsInDeck", () => {
  /**
   * The shapes that are NOT charts.
   *
   * The scan has always held them and always dropped them, which is the right
   * trade for rescaling and repair and the wrong one for a diagnostic: "41
   * shapes became 79", "the chart landed ungrouped", "the slide still holds what
   * was there before" are all questions about the shapes a chart scan discards,
   * and answering them has meant asking the owner to save the deck and upload
   * it.
   */
  it("hands back every shape on every slide, when asked", async () => {
    const s1 = makeSlide("s1");
    const s2 = makeSlide("s2");
    installHost([s1, s2]);
    const loose = s2.shapes.addGeometricShape("rectangle", { left: 3, top: 4, width: 1, height: 1 });
    loose.name = "not a chart";

    const scan = await listChartsInDeck({ withInventory: true });
    const second = scan.inventory?.find((s) => s.slideId === "s2");
    expect(second?.index, "the inventory has to say where in the deck a slide is").toBe(1);
    expect(second?.shapes.map((s) => s.name)).toContain("not a chart");
    expect(second?.shapes[0].left).toBe(3);
    // Charts are unaffected — the inventory rides along, it does not replace.
    expect(scan.charts).toHaveLength(0);
  });

  it("costs nothing on the paths that did not ask for it", async () => {
    // `items/name` is a per-shape string deck-wide, and the callers that scan
    // every slide on a live web host — Same Scale, the repair pass, five
    // self-test scenarios — are exactly the ones that must not pay for a
    // diagnostic. The default has to be the old request, unchanged.
    const s1 = makeSlide("s1");
    installHost([s1]);
    s1.shapes.addGeometricShape("rectangle", { left: 0, top: 0, width: 1, height: 1 });
    const scan = await listChartsInDeck();
    expect(scan.inventory, "the inventory came back unasked").toBeUndefined();
    expect(lastShapeLoadSpec(), "the default scan started asking for shape names").toBe(
      "items/id,items/left,items/top",
    );
  });

  it("finds tagged charts across all slides", async () => {
    const s1 = makeSlide("s1");
    const s2 = makeSlide("s2");
    installHost([s1, s2]);
    await insertSceneIntoSlide(buildChart(config), { tagData: '{"a":1}' });
    const g = s2.shapes.addGroup([]);
    g.tagStore.set(CHART_TAG, '{"b":2}');
    s2.shapes.addGeometricShape("rectangle", { left: 0, top: 0, width: 1, height: 1 });

    const found = await listChartsInDeck();
    expect(found.charts).toHaveLength(2);
    expect(found.charts.map((f) => f.target.slideId).sort()).toEqual(["s1", "s2"]);
    // The deck-wide sweep releases its proxy objects (one tag + one shape per
    // shape on every slide) once their values are read.
    expect(untracked.tags).toBeGreaterThan(0);
    expect(untracked.shapes).toBeGreaterThan(0);
  });
});

describe("proxy lifecycle", () => {
  it("listChartsInSelection untracks the shape and tag proxies it scans", async () => {
    const slide = makeSlide("s1");
    const a = makeShape("group", undefined, { left: 1, top: 1, width: 1, height: 1 });
    a.tagStore.set(CHART_TAG, "{}");
    const b = makeShape("geometric", "rectangle", { left: 2, top: 2, width: 1, height: 1 });
    installHost([slide], [a, b]);
    await listChartsInSelection();
    expect(untracked.tags).toBe(6); // both selected shapes' config, parts AND origin tags
    expect(untracked.shapes).toBe(2);
  });
});

describe("in-place update keeps the chart where it is", () => {
  // The tagged shape's left/top is NOT the frame origin: grouped it is the
  // group's bounding box, ungrouped it is whatever created[0] happens to be.
  // Feeding it back as the next render's origin shifted the chart by the scene's
  // ink offset — and compounded, so Same Scale walked charts off the slide.
  const cfg: ChartConfig = {
    kind: "stacked",
    ...DEFAULT_SIZE,
    data: { categories: ["A", "B"], series: [{ name: "S", values: [3, 4] }] },
  };

  const cycle = async (grouped: boolean) => {
    const slide = makeSlide("s1");
    installHost([slide], [], slide, grouped ? () => true : (v) => v !== "1.8");
    await insertSceneIntoSlide(buildChart(cfg), { tagData: JSON.stringify(cfg), left: 60, top: 90 });
    const seen: string[] = [];
    for (let i = 0; i < 3; i++) {
      const charts = (await listChartsInDeck()).charts;
      expect(charts).toHaveLength(1);
      const t = charts[0].target;
      seen.push(`${Math.round(t.left)},${Math.round(t.top)}`);
      await updateChartsInSlides([{ scene: buildChart(cfg), target: t, opts: { tagData: JSON.stringify(cfg) } }]);
    }
    return seen;
  };

  it("does not drift when the host groups (desktop)", async () => {
    const seen = await cycle(true);
    expect(new Set(seen).size, `chart moved across update cycles: ${seen.join(" -> ")}`).toBe(1);
  });

  it("follows a chart the user has DRAGGED, instead of teleporting it back", async () => {
    // Re-rendering at the tagged shape's corner drifts; re-rendering at the
    // recorded origin teleports a moved chart back to where it was first
    // inserted, silently undoing the user's drag. The origin tag therefore also
    // records the ANCHOR (where the tagged shape landed), so an update shifts the
    // origin by exactly how far the shape has moved since.
    const slide = makeSlide("s1");
    installHost([slide]);
    const cfg: ChartConfig = {
      kind: "stacked",
      ...DEFAULT_SIZE,
      data: { categories: ["A", "B"], series: [{ name: "S", values: [3, 4] }] },
    };
    await insertSceneIntoSlide(buildChart(cfg), { tagData: JSON.stringify(cfg), left: 60, top: 90 });

    const before = (await listChartsInDeck()).charts[0].target;
    // The user drags the whole chart across the slide.
    const [dx, dy] = [240, 110];
    for (const sh of slide.created) {
      sh.left += dx;
      sh.top += dy;
    }
    const moved = (await listChartsInDeck()).charts[0].target;
    expect(moved.left).toBeCloseTo(before.left + dx, 5);

    await updateChartsInSlides([{ scene: buildChart(cfg), target: moved, opts: { tagData: JSON.stringify(cfg) } }]);
    const after = (await listChartsInDeck()).charts[0].target;
    expect(after.left, "update dragged the chart back to its insert position").toBeCloseTo(moved.left, 5);
    expect(after.top, "update dragged the chart back to its insert position").toBeCloseTo(moved.top, 5);
  });

  it("follows a drag even when the caller reuses a CACHED target (what the pane does)", async () => {
    // The pane captures state.editTarget once, when the chart is loaded, and
    // re-uses it for every subsequent "Update chart". Measuring the drag against
    // that snapshot reports no movement, so the update put the chart back where
    // it was — the same teleport, just reached through the pane's real flow
    // rather than a fresh deck read.
    const slide = makeSlide("s1");
    installHost([slide]);
    const cfg: ChartConfig = {
      kind: "stacked",
      ...DEFAULT_SIZE,
      data: { categories: ["A", "B"], series: [{ name: "S", values: [3, 4] }] },
    };
    await insertSceneIntoSlide(buildChart(cfg), { tagData: JSON.stringify(cfg), left: 60, top: 90 });
    const held = (await listChartsInDeck()).charts[0].target; // captured once, then kept

    for (const sh of slide.created.filter((s) => !s.deleted)) {
      sh.left += 200;
      sh.top += 60;
    }
    const moved = (await listChartsInDeck()).charts[0].target;

    // The pane updates with the STALE target it has been holding all along.
    await updateChartInSlide(buildChart(cfg), held, { tagData: JSON.stringify(cfg) });
    const after = (await listChartsInDeck()).charts[0].target;
    expect(after.left, "update teleported the chart back").toBeCloseTo(moved.left, 5);
    expect(after.top, "update teleported the chart back").toBeCloseTo(moved.top, 5);
  });

  it("does not drift when the host cannot group (web)", async () => {
    const seen = await cycle(false);
    expect(new Set(seen).size, `chart moved across update cycles: ${seen.join(" -> ")}`).toBe(1);
  });
});

describe("repeated in-place updates keep landing", () => {
  // An update replaces every shape, so the caller's EditTarget is dead the
  // moment it returns. Before updateChartsInSlides handed back a fresh one, the
  // SECOND update named a shape that no longer existed, was filtered out as
  // "the user deleted this chart", and did nothing — with no error. Auto-update
  // (which fires 900ms after every control change) died the same way after one
  // push, so the pane looked like "Update stopped working".
  it("a second update against the returned target still lands", async () => {
    const slide = makeSlide("s1");
    installHost([slide]);
    const cfg: ChartConfig = {
      kind: "stacked",
      ...DEFAULT_SIZE,
      data: { categories: ["A", "B"], series: [{ name: "S", values: [1, 2] }] },
    };
    await insertSceneIntoSlide(buildChart(cfg), { tagData: '{"v":0}' });
    let target = (await listChartsInDeck()).charts[0].target;

    for (const v of [1, 2, 3]) {
      const next = await updateChartInSlide(buildChart(cfg), target, { tagData: `{"v":${v}}` });
      expect(next, `update ${v} returned no target`).toBeTruthy();
      target = next!;
      // The edit actually reached the slide, every time.
      const live = (await listChartsInDeck()).charts;
      expect(live, `update ${v} lost the chart`).toHaveLength(1);
      expect(JSON.parse(live[0].configJson).v, `update ${v} silently did nothing`).toBe(v);
    }
  });
});

describe("insertAgendaSlides", () => {
  it("appends one slide per chapter and renders ungrouped", async () => {
    const s1 = makeSlide("s1");
    const slides = [s1];
    installHost(slides);
    const chapters = ["Intro", "Findings", "Next steps"];
    const scenes = chapters.map((_, i) => buildAgendaScene(chapters, { highlight: i }));
    await insertAgendaSlides(scenes);
    expect(slides).toHaveLength(4);
    for (let i = 1; i < 4; i++) {
      expect(slides[i].created.length).toBeGreaterThan(0);
      expect(slides[i].created.some((s) => s.type === "group")).toBe(false);
    }
  });
});

describe("insertDemoDeck", () => {
  it("appends one slide per item and tags the charts with their config", async () => {
    const s1 = makeSlide("s1");
    const slides = [s1];
    installHost(slides);
    const items = [
      {
        scene: buildChart({
          ...DEFAULT_SIZE,
          kind: "pie" as const,
          data: { categories: ["A", "B"], series: [{ name: "S", values: [3, 1] }] },
        }),
        tagData: '{"kind":"pie"}',
      },
      {
        scene: buildChart({
          ...DEFAULT_SIZE,
          kind: "clustered" as const,
          data: { categories: ["A"], series: [{ name: "S", values: [5] }] },
        }),
        tagData: '{"kind":"clustered"}',
      },
      {
        scene: {
          width: 100,
          height: 40,
          nodes: [{ kind: "rect" as const, x: 0, y: 0, w: 10, h: 10, fill: "#111111" }],
        },
      }, // untagged element
    ];
    await insertDemoDeck(items);
    // Three slides appended after the original.
    expect(slides).toHaveLength(4);
    for (let i = 1; i < 4; i++) expect(slides[i].created.length).toBeGreaterThan(0);
    // The two chart slides carry their config tag; the element slide does not.
    expect(slides[1].created.some((s) => s.tagStore.get(CHART_TAG) === '{"kind":"pie"}')).toBe(true);
    expect(slides[3].created.every((s) => !s.tagStore.has(CHART_TAG))).toBe(true);
  });
});

describe("isPowerPointHost", () => {
  it("is false outside an Office host and true inside", () => {
    expect(isPowerPointHost()).toBe(false);
    vi.stubGlobal("PowerPoint", {});
    vi.stubGlobal("Office", { context: { host: "PowerPoint" } });
    expect(isPowerPointHost()).toBe(true);
  });
});

describe("marker symbols in the live add-in", () => {
  const markerScene = (markers: MarkerSymbol[]) =>
    buildChart({
      kind: "scatter",
      width: 480,
      height: 300,
      data: {
        categories: ["a", "b", "c"],
        series: [
          { name: "X", values: [1, 2, 3] },
          { name: "Y", values: [2, 4, 3] },
          { name: "Group", values: [1, 2, 3] },
        ],
      },
      scatter: { markers },
    });

  it("draws each symbol as native preset geometry, filled", async () => {
    const slide = makeSlide("s1");
    installHost([slide]);
    await insertSceneIntoSlide(markerScene(["diamond", "plus", "triangle"]), { left: 0, top: 0 });

    // Filled preset geometry is the whole reason a symbol is not a polygon:
    // PowerPoint can only outline a freeform, so a polygon marker would be
    // hollow here while the SVG preview showed it solid.
    for (const preset of ["diamond", "plus", "triangle"]) {
      const shapes = slide.created.filter((s) => s.geo === preset);
      expect(shapes.length, preset).toBeGreaterThan(0);
      for (const s of shapes) {
        expect(s.fillColor, preset).toMatch(/^#[0-9a-f]{6}$/i);
        expect(s.fillCleared, preset).toBe(false);
        expect(s.box.width, preset).toBeGreaterThan(0);
        expect(s.box.width, preset).toBeCloseTo(s.box.height, 9);
      }
    }
  });

  it("needs no rotation, so it works on a bare 1.4 host", async () => {
    // Arrowheads and pie fans need Shape.rotation (1.10+) and degrade without
    // it. The marker set is deliberately rotation-free: nothing here may set
    // rotation, so a 1.4 host draws the same shapes as a current one.
    const slide = makeSlide("s1");
    installHost([slide], [], slide, () => false);
    await insertSceneIntoSlide(markerScene(["diamond", "triangle", "plus"]), { left: 0, top: 0 });
    const presets = slide.created.filter((s) => ["diamond", "triangle", "plus"].includes(s.geo!));
    expect(presets.length).toBeGreaterThan(0);
    for (const s of presets) expect(s.rotation).toBeUndefined();
  });

  it("places the symbol's box centred on the point", async () => {
    const slide = makeSlide("s1");
    installHost([slide]);
    const scene = markerScene(["diamond", "diamond", "diamond"]);
    await insertSceneIntoSlide(scene, { left: 100, top: 50 });
    const nodes = scene.nodes.filter((n) => n.kind === "symbol");
    expect(nodes.length).toBeGreaterThan(0);
    for (const n of nodes) {
      if (n.kind !== "symbol") continue;
      const s = slide.created.find((c) => c.geo === "diamond" && Math.abs(c.box.left - (100 + n.cx - n.size)) < 1e-6);
      expect(s, `no shape at cx=${n.cx}`).toBeTruthy();
      expect(s!.box.top).toBeCloseTo(50 + n.cy - n.size, 9);
      expect(s!.box.width).toBeCloseTo(n.size * 2, 9);
    }
  });
});

describe("Office round-trips do not scale with the chart count", () => {
  const cfgFor = (v: number): ChartConfig => ({
    ...config,
    data: { categories: ["A", "B"], series: [{ name: "S1", values: [v, v + 1] }] },
  });
  const targetsOn = (slide: FakeSlide, n: number) =>
    Array.from({ length: n }, (_, i) => {
      // A real target names a shape that exists on the slide; make one per chart.
      const s = slide.shapes.addGeometricShape("rectangle", { left: 0, top: 0, width: 1, height: 1 });
      return {
        scene: buildChart(cfgFor(i)),
        target: { slideId: slide.id, shapeId: s.id, left: 10, top: 20 },
        opts: { tagData: `{"i":${i}}` },
      };
    });

  it("re-renders N charts in ONE context, whatever N is", async () => {
    // The defect this guards: doSameScale awaited the single-chart update in a
    // loop, so each chart opened its own PowerPoint.run — 20 contexts across a
    // 20-chart deck. That is the property worth pinning. The SYNC count is no
    // longer flat and must not be: shapes commit in batches, because a live
    // canvas will not take a whole chart at once (SHAPES_PER_SYNC).
    for (const n of [1, 2, 10, 20]) {
      const slide = makeSlide("s1");
      installHost([slide]);
      await updateChartsInSlides(targetsOn(slide, n));
      expect(trips.contexts, `${n} charts`).toBe(1);
    }
  });

  it("costs syncs per BATCH OF SHAPES, not a fixed toll per chart", async () => {
    // The two failure modes this sits between: a per-chart context (the old
    // N+1, 80 round-trips for 20 charts), and a per-chart mega-batch that the
    // host silently refuses. Syncs must track the shapes, and nothing else.
    const slide = makeSlide("s1");
    installHost([slide]);
    await updateChartsInSlides(targetsOn(slide, 1));
    const one = trips.syncs;
    installHost([makeSlide("s2")]);
    const slide2 = makeSlide("s2");
    installHost([slide2]);
    await updateChartsInSlides(targetsOn(slide2, 2));
    const two = trips.syncs;
    // Doubling the charts doubles the drawing, not a fixed per-chart overhead:
    // the growth is the extra shapes' batches, so it stays well under 2x.
    expect(two).toBeGreaterThan(one);
    expect(two).toBeLessThan(one * 2 + 2);
  });

  it("still draws every chart it batches, tagged and grouped", async () => {
    const slide = makeSlide("s1");
    installHost([slide]);
    const items = targetsOn(slide, 3);
    await updateChartsInSlides(items);
    const groups = slide.created.filter((s) => s.type === "group");
    expect(groups).toHaveLength(3);
    // Each group carries its OWN config, not the last one written.
    expect(groups.map((g) => g.tagStore.get(CHART_TAG))).toEqual(['{"i":0}', '{"i":1}', '{"i":2}']);
    // The old shape each target named is gone.
    for (const it of items) expect(slide.created.find((s) => s.id === it.target.shapeId)!.deleted).toBe(true);
    // Charts land at their target's position, not the default offset.
    for (const r of slide.created.filter((s) => s.geo === "rectangle" && s.box.width > 1)) {
      expect(r.box.left).toBeGreaterThanOrEqual(10);
    }
  });

  it("keeps the single-chart paths to ONE context each", async () => {
    // updateChartInSlide is now updateChartsInSlides([one]); the Insert button
    // opens its own. Neither may open more than one, however many shapes the
    // chart has.
    const slide = makeSlide("s1");
    installHost([slide]);
    await insertSceneIntoSlide(buildChart(config), { tagData: "{}" });
    expect(trips.contexts).toBe(1);

    const slide2 = makeSlide("s2");
    installHost([slide2]);
    const s = slide2.shapes.addGeometricShape("rectangle", { left: 0, top: 0, width: 1, height: 1 });
    await updateChartInSlide(buildChart(config), { slideId: "s2", shapeId: s.id, left: 0, top: 0 }, { tagData: "{}" });
    expect(trips.contexts).toBe(1);
  });

  it("does nothing, and opens no context, for an empty batch", async () => {
    installHost([makeSlide("s1")]);
    await updateChartsInSlides([]);
    expect([trips.syncs, trips.contexts]).toEqual([0, 0]);
  });

  it("keeps every chart re-editable when the grouping sync is refused", async () => {
    // Batching costs granularity: a refused grouping now loses grouping for the
    // whole batch, not just one chart. What must NOT be lost is the config tag —
    // the charts are already on the slide (their shapes committed a phase
    // earlier), so each must fall back to tagging its own first shape or it
    // silently stops being re-editable.
    //
    // The failure has to come from the SYNC, not from addGroup: Office.js only
    // reports queued commands there, which means every tag target has already
    // been pointed at a group that turned out not to exist. A test that throws
    // from addGroup instead never overwrites them and proves nothing.
    const slide = makeSlide("s1");
    installHost([slide]);
    // The group sync is no longer a fixed number: the shapes commit in batches
    // first, so its index depends on the chart's size. Find it rather than
    // hardcode it — a wrong number here silently tests nothing.
    const batches = Math.ceil(buildChart(cfgFor(0)).nodes.length / 10);
    // 1 resolve slides, 1 resolve old shapes, then PER CHART one delete sync
    // plus its render batches, then GROUP. The delete is per chart because a
    // shared one commits every chart's removal before any redraw runs — see
    // updateChartsInSlides.
    faults.failSyncOn = 2 + 3 /* charts */ * (1 + batches) + 1;
    try {
      const items = targetsOn(slide, 3);
      // One refreshed target per chart — the caller needs them to stay live.
      await expect(updateChartsInSlides(items)).resolves.toHaveLength(items.length);
      // Each chart's OWN config, back on each chart's OWN first shape.
      const tagged = slide.created.filter((s) => s.tagStore.has(CHART_TAG));
      expect(tagged.map((s) => s.tagStore.get(CHART_TAG))).toEqual(['{"i":0}', '{"i":1}', '{"i":2}']);
      expect(tagged.every((s) => s.type !== "group")).toBe(true);
    } finally {
      faults.failSyncOn = 0;
    }
  });

  it("does not blank the rest of the deck when one chart's redraw stalls", async () => {
    // The defect: every chart's old shapes were deleted in ONE shared sync,
    // and only then were the charts redrawn one at a time. Those deletes are
    // committed, so a single stalled redraw rejected the whole PowerPoint.run
    // and left every chart AFTER it blank — old shapes gone, new ones never
    // queued. Same Scale runs across the whole deck and necessarily includes
    // the chart on the visible slide, which is the one condition documented
    // here as reliably stalling a redraw. So one slow chart emptied the deck.
    const slide = makeSlide("s1");
    installHost([slide]);
    const batches = Math.ceil(buildChart(cfgFor(0)).nodes.length / 10);
    // 2 resolve syncs, then per chart 1 delete + `batches` renders. Fail the
    // FIRST render batch of chart 2 (0-based index 1).
    faults.failSyncOn = 2 + (1 + batches) + 1 + 1;
    const failed: string[] = [];
    try {
      const items = targetsOn(slide, 3);
      const next = await updateChartsInSlides(items, (it) => failed.push(it.opts!.tagData!));
      // Chart 2 is reported, once, by name.
      expect(failed).toEqual(['{"i":1}']);
      // Charts 1 and 3 are re-editable — they redrew and carry their OWN tags.
      const tagged = slide.created.filter((s) => s.tagStore.has(CHART_TAG));
      expect(tagged.map((s) => s.tagStore.get(CHART_TAG))).toEqual(['{"i":0}', '{"i":2}']);
      // And the caller's targets are not shifted onto each other: chart 3 must
      // NOT be handed chart 1's new shape id. Indexing `tagged` by position
      // would do exactly that once a chart drops out of the batch.
      expect(next).toHaveLength(3);
      expect(new Set(next.map((t) => t.shapeId)).size).toBe(3);
      expect(next[1].shapeId).toBe(items[1].target.shapeId); // unchanged: it failed
    } finally {
      faults.failSyncOn = 0;
    }
  });

  it("tells the caller what EVERY stalled chart destroyed, not just the first", async () => {
    // `onFailed` is the only channel a deck-wide caller has for finding out
    // that a chart went blank — and, through the wreckage on the error, what
    // was left on its slide. Same Scale sweeps that debris; it cannot sweep
    // what it is not told about.
    //
    // The bug was subtle because failure #1 looked fine: `wreck()` mutates the
    // host's error object in place, so the raw `err` handed to `onFailed`
    // happened to carry the wreckage anyway. But `wreck()` ran INSIDE the
    // `firstFailure === undefined` test, so failures #2..n were never
    // annotated at all — and a deck-wide update is exactly the caller that
    // gets more than one. Hence two stalls here: one would pass either way.
    const slide = makeSlide("s1");
    installHost([slide]);
    const batches = Math.ceil(buildChart(cfgFor(0)).nodes.length / 10);
    // Sync map: 2 resolves, then per chart 1 delete + up to `batches` renders.
    // A chart that fails on its first render batch consumes exactly 2. So
    // chart 1's first render is sync 5+batches, and — chart 1 having stopped
    // there — chart 2's is 7+batches.
    failSyncsOn.add(5 + batches);
    failSyncsOn.add(7 + batches);
    const seen: { tag: string; wreckage: ReturnType<typeof wreckageOf> }[] = [];
    try {
      const items = targetsOn(slide, 3);
      await updateChartsInSlides(items, (it, err) => seen.push({ tag: it.opts!.tagData!, wreckage: wreckageOf(err) }));
      // Charts 2 and 3 stalled; chart 1 is untouched and still redrew.
      expect(seen.map((s) => s.tag)).toEqual(['{"i":1}', '{"i":2}']);
      // BOTH carry their wreckage. Pre-fix the second was undefined, so Same
      // Scale swept the first chart's debris and left the rest on the deck.
      for (const s of seen) {
        expect(s.wreckage, `wreckage for ${s.tag}`).toBeDefined();
        expect(s.wreckage!.slideId).toBe("s1");
        expect(Array.isArray(s.wreckage!.strayIds)).toBe(true);
      }
    } finally {
      failSyncsOn.clear();
    }
  });

  it("still throws when the ONE chart it was given fails", async () => {
    // updateChartResilient catches this throw to reach its slide-swap and
    // picture fallbacks. Swallowing a total failure would strand it on layer 1
    // with a chart whose old shapes are already deleted.
    const slide = makeSlide("s1");
    installHost([slide]);
    faults.failSyncOn = 4; // 2 resolves, 1 delete, then the first render batch
    try {
      await expect(updateChartsInSlides(targetsOn(slide, 1))).rejects.toThrow();
    } finally {
      faults.failSyncOn = 0;
    }
  });

  it("renders one slide per context and reports progress per slide", async () => {
    // One PowerPoint.run per slide isolates a chart the host can't finish and
    // keeps each context light (one chart's shapes, not a chunk's four). Progress
    // is per slide, so a slow host shows slides landing instead of freezing.
    for (const n of [2, 12, 35] as const) {
      installHost([makeSlide("s1")]);
      const seen: string[] = [];
      const report = await insertDemoDeck(
        Array.from({ length: n }, (_, i) => ({ scene: buildChart(cfgFor(i)), tagData: `{"i":${i}}` })),
        (done, total) => seen.push(`${done}/${total}`),
      );
      expect(failedIndices(report), `${n} slides`).toEqual([]);
      // Self-check: the deck grew by exactly one slide per item, nothing lost.
      expect(report.slidesAdded, `${n} slides`).toBe(n);
      // One context per SLIDE, plus one more per slide for addSlides' settled
      // fresh-context verify (nothing lost, so no retry context), the two
      // settled slideCount reads (before/after), and the paged on-slide
      // readback (nothing lost, so it runs).
      expect(trips.contexts, `${n} slides`).toBe(2 * n + 2 + Math.ceil(n / READBACK_PAGE));
      expect(seen, `${n} slides`).toHaveLength(n);
      expect(seen.at(-1)).toBe(`${n}/${n}`);
      // Monotonic, never over-counting.
      expect(seen.map((x) => Number(x.split("/")[0]))).toEqual(
        [...seen.map((x) => Number(x.split("/")[0]))].sort((a, b) => a - b),
      );
    }
  });

  it("appends every demo slide, each tagged with its own config", async () => {
    const deck: FakeSlide[] = [makeSlide("s1")];
    installHost(deck);
    const n = 35;
    const report = await insertDemoDeck(
      Array.from({ length: n }, (_, i) => ({ scene: buildChart(cfgFor(i)), tagData: `{"i":${i}}` })),
    );
    expect(failedIndices(report)).toEqual([]);
    expect(report.slidesAdded).toBe(n);
    // The fake appends a slide per add(); the original + n new ones.
    expect(deck.length).toBe(1 + n);
    const tags = deck.slice(1).map((s) => s.created.map((c) => c.tagStore.get(CHART_TAG)).find(Boolean));
    expect(tags).toEqual(Array.from({ length: n }, (_, i) => `{"i":${i}}`));
  });

  it("records per-item and total wall-clock so a run's duration is on the record", async () => {
    installHost([makeSlide("s1")]);
    const n = 5;
    const report = await insertDemoDeck(Array.from({ length: n }, (_, i) => ({ scene: buildChart(cfgFor(i)) })));
    // Every item is timed; the total is present and never negative.
    expect(report.results.every((r) => typeof r.ms === "number" && r.ms >= 0)).toBe(true);
    expect(typeof report.totalMs).toBe("number");
    expect(report.totalMs).toBeGreaterThanOrEqual(0);
    // The whole run is at least as long as its slowest single item.
    expect(report.totalMs).toBeGreaterThanOrEqual(Math.max(...report.results.map((r) => r.ms)));
  });

  it("finds no blank slots and completes the readback on a clean run", async () => {
    installHost([makeSlide("s1")]);
    const n = 4;
    const report = await insertDemoDeck(Array.from({ length: n }, (_, i) => ({ scene: buildChart(cfgFor(i)) })));
    expect(report.blankSlides).toEqual([]); // every added slide read back with shapes
    expect(report.blanksRead).toBe(true); // the readback finished
    expect(report.addsIssued).toBe(n); // one add per item, no retries/fails
  });

  it("reports a host-blanked slide by DECK POSITION, not by item name", async () => {
    const deck = [makeSlide("s1")]; // pre-existing at index 0; added slides take indices 1..n
    installHost(deck);
    const n = 4;
    blankReadbackAt.add(2); // the added slide at deck index 2 reads back empty on readback
    const report = await insertDemoDeck(Array.from({ length: n }, (_, i) => ({ scene: buildChart(cfgFor(i)) })));
    // A blank slide has no content/tag to name it — reported as the 1-based deck position (index 2 → slide 3).
    expect(report.blankSlides).toEqual([3]);
    expect(report.blanksRead).toBe(true);
  });

  it("names a blank slide from its slot tag when the item carries a title", async () => {
    // A blank readback used to say only "slide 3". Every demo slide now gets a
    // POWERCHART_DEMO_SLOT tag on creation, so the readback can name the missing
    // chart by title. blankReadbackAt makes index 2 report 0 shapes; the slot
    // tag survives (we never emptied the tag store) and gives us the item name.
    _setBlankReReadDelayForTest(0); // no wall-clock sleep in the test
    try {
      const deck = [makeSlide("s1")];
      installHost(deck);
      const n = 3;
      blankReadbackAt.add(2);
      const report = await insertDemoDeck(
        Array.from({ length: n }, (_, i) => ({
          scene: buildChart(cfgFor(i)),
          title: `chart-${i}`,
        })),
      );
      expect(report.blankSlides).toEqual([3]);
      // Index 2 corresponds to item 1 (item 0 is index 1, item 1 is index 2, ...).
      expect(report.blankItems).toEqual([{ position: 3, title: "chart-1" }]);
    } finally {
      _setBlankReReadDelayForTest(200);
    }
  });

  it("marks the blank readback incomplete when it faults, not falsely clean", async () => {
    installHost([makeSlide("s1")]);
    faults.faultShapeGetCount = true; // every readback getCount throws
    const report = await insertDemoDeck(Array.from({ length: 3 }, (_, i) => ({ scene: buildChart(cfgFor(i)) })));
    // An empty list must NOT read as "no blanks" when we could not measure.
    expect(report.blanksRead).toBe(false);
    expect(report.blankSlides).toEqual([]);
    // The run itself still succeeded — a readback fault is not a render failure.
    expect(report.results.every((r) => r.status === "rendered")).toBe(true);
  });

  it("bypassBudget lets a text-heavy scene render even when its shape count is over the budget", async () => {
    // The results/contents slide bug: 32 failures pushed the results scene to
    // 135 shapes — over DEMO_SHAPE_BUDGET (90) — and the run's own summary
    // came back as a red "NOT COMPLETE" stamp. Text-only scenes don't hit the
    // wedge/polygon flood the budget guards against; they should render.
    const deck: FakeSlide[] = [makeSlide("s1")];
    installHost(deck);
    const denseTextScene = {
      width: 100,
      height: 100,
      nodes: Array.from({ length: 120 }, (_, k) => ({
        kind: "text" as const,
        x: k,
        y: 0,
        w: 40,
        h: 20,
        text: `row ${k}`,
        fontSize: 12,
        color: "#000000",
        align: "left" as const,
        valign: "top" as const,
      })),
    };
    // With bypassBudget: the scene renders as a real chart, no stamp.
    const withBypass = await insertDemoDeck([{ scene: denseTextScene, title: "Results", bypassBudget: true }]);
    expect(withBypass.results[0].status).toBe("rendered");
    expect(deck[1].created.some((s) => s.name === "PowerChart:not-complete")).toBe(false);
    expect(deck[1].created.filter((s) => s.type === "text").length).toBeGreaterThanOrEqual(120);
    // Without bypassBudget: the scene is stamped instead of drawn.
    installHost([makeSlide("s1")]);
    const withoutBypass = await insertDemoDeck([{ scene: denseTextScene, title: "Results" }]);
    expect(withoutBypass.results[0].status).toBe("skipped");
  });

  it("skips a chart too dense for the host, keeps the slide, and stamps it NOT COMPLETE", async () => {
    // The heavy charts (area ~208 shapes) will not land on web and burn the
    // timeout trying. They are skipped up-front — the slide is kept (coverage),
    // its chart NOT drawn, and a stamp makes the placeholder unmistakable.
    const deck: FakeSlide[] = [makeSlide("s1")];
    installHost(deck);
    const dense = {
      width: 100,
      height: 100,
      nodes: Array.from({ length: 120 }, (_, k) => ({
        kind: "rect" as const,
        x: k,
        y: 0,
        w: 1,
        h: 1,
        fill: "#111111",
      })),
    };
    const light = () => buildChart(cfgFor(0)); // a handful of shapes, well under budget
    const report = await insertDemoDeck([{ scene: light() }, { scene: dense }, { scene: light() }]);
    // Only the dense one is reported (as "skipped"); the deck still has all three.
    expect(failedIndices(report)).toEqual([1]);
    expect(report.results[1].status).toBe("skipped");
    expect(report.slidesAdded).toBe(3);
    expect(deck.length).toBe(1 + 3);
    // The dense slide (deck[2]) carries a stamp, NOT the 120 chart shapes.
    const denseSlide = deck[2];
    const stamp = denseSlide.created.find((s) => s.name === "PowerChart:not-complete");
    expect(stamp, "dense slide is stamped").toBeTruthy();
    expect(stamp!.text).toContain("NOT COMPLETE");
    // A top strip, not a slab over the middle — a mis-targeted stamp must not
    // obliterate a real chart under it (a 540pt-tall slide).
    expect(stamp!.top, "stamp sits at the top").toBeLessThan(80);
    expect(stamp!.height, "stamp is a strip").toBeLessThan(120);
    expect(denseSlide.created.length, "chart not drawn").toBeLessThan(120);
    // The light neighbours rendered as real charts (no stamp).
    expect(deck[1].created.some((s) => s.name === "PowerChart:not-complete")).toBe(false);
    expect(deck[3].created.length).toBeGreaterThan(1);
  });

  it("does not report a phantom lost slide for a too-dense item whose stamp was refused", async () => {
    // `addsIssued` used to be inferred as "one per item, plus one more for every
    // retried or failed item". A too-dense item never re-renders — there is
    // nothing to re-render, only a placeholder to stamp — so when its stamp sync
    // is refused it ends "failed" having issued exactly ONE add. The inference
    // then charged it a second, and `addsIssued − slidesAdded` accused the host
    // of losing a slide it had actually kept. This harness exists to stop
    // inventing failures, so a false ⚠ is the bug.
    const deck: FakeSlide[] = [makeSlide("s1")];
    installHost(deck);
    const dense = {
      width: 100,
      height: 100,
      nodes: Array.from({ length: 120 }, (_, k) => ({
        kind: "rect" as const,
        x: k,
        y: 0,
        w: 1,
        h: 1,
        fill: "#111111",
      })),
    };
    failSyncsOn.add(6); // the too-dense item's stampSlide — see the test above
    const report = await insertDemoDeck([
      { scene: dense, title: "too dense" },
      { scene: buildChart(cfgFor(1)), tagData: '{"i":1}', title: "fine" },
    ]);
    expect(report.results[0].status).toBe("failed"); // the premise: stamped, refused
    expect(report.results[0].attempts, "one add issued, not two").toBe(1);
    // Both slides are on the deck, so nothing was lost — and the report agrees.
    expect(report.slidesAdded).toBe(2);
    expect(report.addsIssued).toBe(2);
    expect(report.addsIssued - report.slidesAdded, "no phantom loss").toBe(0);
  });

  it("self-check catches a slide the host silently drops (deck grew by less than asked)", async () => {
    // The corruption a visual scan misses and today cost us 3 lost slides: an
    // add() that never lands leaves the deck one slide short with no error the
    // user sees. The report's slidesAdded is read back from the host, so the
    // shortfall is caught — and the dropped item shows up as not rendered.
    //
    // addSlides self-heals dropped adds with its own fresh-context retries, so
    // losing a slide for good takes the add AND every retry round. There is no
    // longer an outer per-item retry on top of that — it was the source of
    // every duplicate slide, and the end-of-run reconcile pass reports what
    // landed with better evidence than a mid-flight guess.
    const deck: FakeSlide[] = [makeSlide("s1")];
    installHost(deck);
    faults.swallowAdds = ADDS_TO_DEFEAT_ONE_SLIDE; // the add and all its retries → gone for good
    try {
      const report = await insertDemoDeck(
        Array.from({ length: 4 }, (_, i) => ({ scene: buildChart(cfgFor(i)), tagData: `{"i":${i}}` })),
      );
      // 4 items asked for, but the deck only grew by 3 — the lost-slide signal.
      expect(report.slidesAdded).toBe(3);
      expect(report.results).toHaveLength(4); // every item is still accounted for
      expect(failedIndices(report).length).toBeGreaterThanOrEqual(1); // the dropped one is flagged
    } finally {
      faults.swallowAdds = 0;
    }
  });

  it("addSlides self-heals one dropped add via its own fresh-context retry, and surfaces it when every retry also fails", async () => {
    // addSlides now verifies its own adds landed (a settled getCount() in a
    // FRESH context, after the existing 2 syncs) and gets MAX_ADD_RETRY_ROUNDS
    // retry rounds before giving up — the fix for the Presentation_3.pptx bug where
    // PowerPoint web silently dropped ~half of 20 issued add()s. A single
    // dropped add should never even reach insertDemoDeck's own item-level
    // retry: it should be invisible, recovered inside addSlides itself.
    const deck: FakeSlide[] = [makeSlide("s1")];
    installHost(deck);
    faults.swallowAdds = 1; // the FIRST add() call anywhere is dropped, none after
    try {
      const report = await insertDemoDeck([{ scene: buildChart(cfgFor(0)), tagData: '{"i":0}' }]);
      // The retry landed: the deck grew by exactly the one slide asked for.
      expect(report.slidesAdded).toBe(1);
      expect(report.results[0].status).toBe("rendered");
      // Nothing was lost AT COMMIT: the one drop was fully recovered.
      expect(report.addsLostAtCommit).toBe(0);
    } finally {
      faults.swallowAdds = 0;
    }

    // Second sub-case: the drop persists through every retry round too. One
    // addSlides call issues the original add plus one per retry round, so
    // defeating item 0 entirely takes exactly that many dropped adds. A second
    // item follows so the run is not a TOTAL loss (which insertDemoDeck itself
    // would throw on, per "A whole deck lost to HOST errors" below) — item 1
    // renders once faults.swallowAdds is exhausted.
    const warnSpy = vi.spyOn(console, "warn").mockImplementation(() => {});
    const deck2: FakeSlide[] = [makeSlide("s1")];
    installHost(deck2);
    faults.swallowAdds = ADDS_TO_DEFEAT_ONE_SLIDE;
    try {
      const report = await insertDemoDeck([
        { scene: buildChart(cfgFor(0)), tagData: '{"i":0}' },
        { scene: buildChart(cfgFor(1)), tagData: '{"i":1}' },
      ]);
      // Item 0 is a genuine total loss; item 1 landed once faults.swallowAdds ran out.
      expect(report.slidesAdded).toBe(1);
      expect(report.results[0].status).toBe("failed");
      expect(report.results[1].status).toBe("rendered");
      // addSlides confirmed the loss — its own retry did not recover it.
      expect(report.addsLostAtCommit).toBeGreaterThanOrEqual(1);
      expect(warnSpy).toHaveBeenCalled();
    } finally {
      faults.swallowAdds = 0;
      warnSpy.mockRestore();
    }
  });

  it("recovers a run of consecutive dropped adds that a single retry round could not", async () => {
    // The reason MAX_ADD_RETRY_ROUNDS is no longer 1. The host that motivates
    // all of this dropped ~half of 20 adds in ONE burst — drops arrive in runs,
    // not singly, and a retry issued under the same load is dropped by the same
    // load. At one round, two consecutive drops cost a slide outright; the deck
    // came back short, the item was reported failed, and the user lost a chart
    // to a condition that was transient the whole time.
    //
    // Two drops is the smallest case that separates the bounds: recoverable
    // now, a total loss before. Guarded against the constant rather than the
    // literal 2 so the case stays "one more drop than a single round can take".
    expect(MAX_ADD_RETRY_ROUNDS).toBeGreaterThan(1);
    const deck: FakeSlide[] = [makeSlide("s1")];
    installHost(deck);
    faults.swallowAdds = 2;
    try {
      const report = await insertDemoDeck([{ scene: buildChart(cfgFor(0)), tagData: '{"i":0}' }]);
      // Fully recovered: the deck grew by the one slide asked for, the chart
      // rendered, and nothing was written off at commit.
      expect(report.slidesAdded).toBe(1);
      expect(report.results[0].status).toBe("rendered");
      expect(report.addsLostAtCommit).toBe(0);
    } finally {
      faults.swallowAdds = 0;
    }
  });

  it('render:"image" draws ONE picture shape instead of the scene nodes', async () => {
    // The whole point of image mode: a dense chart becomes one shape, so the
    // PowerPoint-web dense-shape wall (office-js #4272 / #5022 / #6498) never
    // gets hit. A violin-sized scene is ~250 native shapes; here it must be 1.
    const NODES = 25;
    const scene = {
      width: 480,
      height: 300,
      nodes: Array.from({ length: NODES }, (_, k) => ({
        kind: "rect" as const,
        x: k,
        y: 0,
        w: 4,
        h: 4,
        fill: "#111111",
      })),
    };
    const slide = makeSlide("s1");
    installHost([slide]);
    await insertSceneIntoSlide(scene, { tagData: "{}", pictureBase64: "AAAA", left: 60, top: 90 });
    // Exactly one shape, geometry rectangle, sized to the FRAME (not the nodes).
    const live = slide.created.filter((s) => !s.deleted);
    expect(live).toHaveLength(1);
    expect(live[0].geo).toBe("rectangle");
    expect(live[0].box).toEqual({ left: 60, top: 90, width: 480, height: 300 });
    // The fill became a picture, carrying the payload, and the outline is off.
    expect(live[0].imageBase64).toBe("AAAA");
    expect(live[0].fillType).toBe("PictureAndTexture");
    expect(live[0].lineFormat.visible).toBe(false);
    // Named like a chart group so the Selection Pane reads the same either way,
    // and it carries the config tag — the picture IS the re-editable chart.
    expect(live[0].name).toBe("PowerChart");
    expect(live[0].tagStore.get(CHART_TAG)).toBe("{}");
  });

  it("strips every base64 spelling down to the bare payload the host wants", async () => {
    // Three forms circulate: a browser toDataURL (`data:image/png;base64,…`),
    // what render-pptx.mjs hands pptxgen (`image/png;base64,…` — NO data:
    // scheme), and already-bare. A `startsWith("data:")` guard would pass the
    // middle one through with `image/png;base64,` still glued on, and the host
    // would get a corrupt payload. Splitting on the last comma handles all three.
    const scene = { width: 100, height: 100, nodes: [{ kind: "rect" as const, x: 0, y: 0, w: 4, h: 4, fill: "#111" }] };
    for (const [input, expected] of [
      ["data:image/png;base64,PAYLOAD", "PAYLOAD"],
      ["image/png;base64,PAYLOAD", "PAYLOAD"],
      ["PAYLOAD", "PAYLOAD"],
    ] as const) {
      const slide = makeSlide("s1");
      installHost([slide]);
      await insertSceneIntoSlide(scene, { tagData: "{}", pictureBase64: input });
      expect(slide.created.filter((s) => !s.deleted)[0].imageBase64, input).toBe(expected);
    }
  });

  it("falls back to native shapes on a host without PowerPointApi 1.8", async () => {
    // setImage is 1.8 and the manifests admit hosts from 1.4, so this must be
    // GATED, not attempted-and-caught: a queued command the host rejects takes
    // the whole sync with it. The chart must still land, as shapes.
    const NODES = 12;
    const scene = {
      width: 480,
      height: 300,
      nodes: Array.from({ length: NODES }, (_, k) => ({
        kind: "rect" as const,
        x: k,
        y: 0,
        w: 4,
        h: 4,
        fill: "#111111",
      })),
    };
    const slide = makeSlide("s1");
    // Everything except 1.8 — so grouping is off too, exactly like the web host.
    installHost([slide], [], slide, (v) => v !== "1.8");
    await insertSceneIntoSlide(scene, { tagData: "{}", pictureBase64: "AAAA" });
    const live = slide.created.filter((s) => !s.deleted);
    expect(live.length).toBeGreaterThanOrEqual(NODES); // the nodes, not a picture
    expect(live.some((s) => s.imageBase64 !== undefined)).toBe(false);
    // Still re-editable: the config tag landed on the first shape (no group at 1.4).
    expect(live.some((s) => s.tagStore.get(CHART_TAG) === "{}")).toBe(true);
  });

  it("falls back to native shapes when the host refuses the picture fill", async () => {
    // A 1.8-advertising host that still rejects setImage (wrong payload format,
    // host quirk). renderPictureShape catches, best-effort deletes the rect, and
    // renderShapesChunked draws the nodes in the SAME request context.
    const NODES = 12;
    const scene = {
      width: 480,
      height: 300,
      nodes: Array.from({ length: NODES }, (_, k) => ({
        kind: "rect" as const,
        x: k,
        y: 0,
        w: 4,
        h: 4,
        fill: "#111111",
      })),
    };
    const slide = makeSlide("s1");
    installHost([slide]);
    faults.refusePictureFill = true;
    try {
      await insertSceneIntoSlide(scene, { tagData: "{}", pictureBase64: "AAAA" });
      const live = slide.created.filter((s) => !s.deleted);
      // The nodes landed, and no picture-filled shape survived.
      expect(live.some((s) => s.imageBase64 !== undefined)).toBe(false);
      expect(live.filter((s) => s.geo === "rectangle").length).toBeGreaterThanOrEqual(NODES);
      expect(live.some((s) => s.tagStore.get(CHART_TAG) === "{}")).toBe(true);
    } finally {
      faults.refusePictureFill = false;
    }
  });

  it("refuses an over-budget payload rather than burning the batch timeout", async () => {
    // MAX_PICTURE_BASE64 is a guard against a pathological custom frame size,
    // not a measured host limit — real payloads are 20-133 KB. Crossing it
    // degrades to shapes; the code that says so is PC-IMG-TOOBIG on the console,
    // because the chart still appears and the fallback is otherwise invisible.
    const scene = {
      width: 480,
      height: 300,
      nodes: Array.from({ length: 12 }, (_, k) => ({
        kind: "rect" as const,
        x: k,
        y: 0,
        w: 4,
        h: 4,
        fill: "#111111",
      })),
    };
    const slide = makeSlide("s1");
    installHost([slide]);
    const warn = vi.spyOn(console, "warn").mockImplementation(() => {});
    try {
      await insertSceneIntoSlide(scene, { tagData: "{}", pictureBase64: "A".repeat(4_000_001) });
      const live = slide.created.filter((s) => !s.deleted);
      expect(live.some((s) => s.imageBase64 !== undefined)).toBe(false);
      expect(live.length).toBeGreaterThanOrEqual(12);
      // A specific, greppable code — so a real bug report teaches us the number.
      expect(warn.mock.calls.flat().join(" ")).toContain("PC-IMG-TOOBIG");
    } finally {
      warn.mockRestore();
    }
  });

  it("costs ONE sync for the drawing, however dense the scene", async () => {
    // The reliability claim: image mode replaces N/batchSize render syncs with
    // exactly one, so a chart that could never commit as shapes lands in a
    // single round trip. Compare a 60-node scene both ways.
    const scene = {
      width: 480,
      height: 300,
      nodes: Array.from({ length: 60 }, (_, k) => ({
        kind: "rect" as const,
        x: k,
        y: 0,
        w: 4,
        h: 4,
        fill: "#111111",
      })),
    };
    installHost([makeSlide("shapes")]);
    await insertSceneIntoSlide(scene, { tagData: "{}" });
    const shapesSyncs = trips.syncs;

    installHost([makeSlide("picture")]);
    await insertSceneIntoSlide(scene, { tagData: "{}", pictureBase64: "AAAA" });
    const pictureSyncs = trips.syncs;

    // 60 nodes at SHAPES_PER_SYNC=10 is 6 render syncs; the picture is 1. Both
    // pay the same tag/group syncs afterwards, so the picture must be strictly
    // and substantially cheaper.
    expect(pictureSyncs).toBeLessThan(shapesSyncs);
    expect(shapesSyncs - pictureSyncs).toBeGreaterThanOrEqual(4);
  });

  it("re-fetches the slide's shape collection before addGroup on a multi-batch chart", async () => {
    // The real-host bug this guards: a >10-shape chart commits in multiple
    // batches, and the Shape proxies returned by earlier batches have their
    // object paths rewritten to getItem(id) by the time the group sync runs.
    // The web host silently drops addGroup(theseStaleProxies), leaving the
    // chart loose and unable to carry its POWERCHART_CONFIG tag. In the run
    // behind this fix, agenda / KPI+flow / table all landed ungrouped and
    // therefore un-re-editable. Fix: re-load slide.shapes.items right before
    // addGroup and pass those fresh proxies to addGroup.
    const NODES = 25; // 3 batches at SHAPES_PER_SYNC=10
    const bigScene = {
      width: 100,
      height: 100,
      nodes: Array.from({ length: NODES }, (_, k) => ({
        kind: "rect" as const,
        x: k,
        y: 0,
        w: 4,
        h: 4,
        fill: "#111111",
      })),
    };
    const deck: FakeSlide[] = [makeSlide("s1")];
    installHost(deck);
    faults.strictGroup = true; // web-host stale-proxy semantics — without the re-fetch, no group appears
    try {
      const report = await insertDemoDeck([{ scene: bigScene, tagData: '{"i":0}', title: "multi-batch" }]);
      expect(report.results[0].status).toBe("rendered");
      // Every appended chart is grouped — the new field flowed through.
      expect(report.results[0].grouped).toBe(true);
      // The slide holds ONE native group carrying the CHART_TAG — proving the
      // group survived and the tag landed on it (not on a stray first shape).
      const groups = deck[1].created.filter((s) => s.type === "group");
      expect(groups).toHaveLength(1);
      expect(groups[0].tagStore.get(CHART_TAG)).toBe('{"i":0}');
    } finally {
      faults.strictGroup = false;
    }
  });

  it("re-acquires each freshly-added slide per batch, so a rewritten getItemAt cannot 5010 mid-deck", async () => {
    // The real regression: HOLD one getItemAt handle to a new slide and reuse it
    // across the render's batched syncs, and once Office.js rewrites its path to
    // getItem(<web-non-round-trippable id>) the next shape throws "InvalidParam
    // passed to GetItem(id)", code 5010 — the deck dies partway through, as it did
    // on the real host. The fix re-acquires a fresh proxy each batch; the fake
    // window-limits a held one.
    //
    // Load-bearing: each slide must span MORE than one batch, because a held
    // handle only goes stale on the batch AFTER a sync. SHAPES_PER_SYNC is 10, so
    // a 25-node scene is 3 batches — a single-batch chart (e.g. cfgFor) can hold
    // its handle and never notice, which is exactly how a weaker version of this
    // test passed against the very bug it meant to guard.
    const NODES = 25;
    const bigScene = {
      width: 100,
      height: 100,
      nodes: Array.from({ length: NODES }, (_, k) => ({
        kind: "rect" as const,
        x: k,
        y: 0,
        w: 4,
        h: 4,
        fill: "#111111",
      })),
    };
    const deck: FakeSlide[] = [makeSlide("s1")];
    installHost(deck);
    const n = 6;
    const report = await insertDemoDeck(Array.from({ length: n }, () => ({ scene: bigScene })));
    expect(failedIndices(report)).toEqual([]);
    // Every appended slide got all its shapes (plus the group) — nothing stranded
    // by a mid-batch 5010.
    expect(deck.length).toBe(1 + n);
    for (let i = 1; i <= n; i++) expect(deck[i].created.length, `slide ${i}`).toBeGreaterThanOrEqual(NODES);
  });
});

describe("a stalled host is legible, and does not hang the pane", () => {
  it("reports every phase, in order, with the shape count", async () => {
    const slide = makeSlide("s1");
    installHost([slide]);
    const seen: string[] = [];
    await insertSceneIntoSlide(buildChart(config), { tagData: "{}" }, (p, d) => seen.push(d ? `${p}:${d}` : p));
    expect(seen[0]).toBe("context");
    expect(seen.at(-1)).toBe("done");
    // "commit" now repeats — once per batch — because shapes land in batches.
    expect(seen.filter((s) => s.startsWith("commit:")).length).toBeGreaterThan(1);
    expect([...new Set(seen.map((s) => s.split(":")[0]))]).toEqual(["context", "queue", "commit", "group", "done"]);
    expect(seen.find((s) => s.startsWith("queue:"))).toMatch(/^queue:\d+ nodes$/);
    // Real progress: "10 of 40 shapes", ending at the total.
    const commits = seen.filter((s) => s.startsWith("commit:"));
    expect(commits[0]).toMatch(/^commit:\d+ of \d+ shapes$/);
    const [done, total] = commits
      .at(-1)!
      .match(/(\d+) of (\d+)/)!
      .slice(1);
    expect(done).toBe(total);
  });

  it("gives up on a host that never answers, naming the phase it died in", async () => {
    // The real failure mode: Office.js does not throw when the host stops
    // answering — the sync promise simply never settles, so the pane spins for
    // ever with nothing to report. This is the only way out.
    vi.useFakeTimers();
    try {
      const slide = makeSlide("s1");
      installHost([slide]);
      // A sync that never settles, exactly like a stalled PowerPoint.ashx.
      (slide as unknown as { id: string }).id = "s1";
      const ctxSync = () => new Promise<void>(() => {});
      vi.stubGlobal("PowerPoint", {
        ...(globalThis as unknown as { PowerPoint: Record<string, unknown> }).PowerPoint,
        run: async (cb: (ctx: unknown) => Promise<unknown>) =>
          cb({
            presentation: { slides: { getItemAt: () => slide }, getSelectedSlides: () => ({ getItemAt: () => slide }) },
            sync: ctxSync,
          }),
      });
      const seen: string[] = [];
      const p = insertSceneIntoSlide(buildChart(config), {}, (ph) => seen.push(ph));
      const assertion = expect(p).rejects.toThrow(/did not respond while drawing shapes \d+-\d+ of \d+/);
      await vi.advanceTimersByTimeAsync(400_000); // past max(45s, shapes*3s)
      await assertion;
      // And it says where it stopped — "commit" is the last thing reached.
      expect(seen.at(-1)?.startsWith("commit")).toBe(true);
    } finally {
      vi.useRealTimers();
    }
  });
});

describe("a host that answers late still gets heard", () => {
  it("reports the real Office error when an abandoned sync finally rejects", async () => {
    // The evidence problem: racing a timeout throws the answer away. The
    // abandoned sync keeps running, and Office.js reports queued-command
    // failures THERE and nowhere else — so whatever it says next is the only
    // description of the bug we will ever get. Without this it is lost.
    vi.useFakeTimers();
    const heard: string[] = [];
    onLateSync((m) => heard.push(m));
    try {
      const slide = makeSlide("s1");
      installHost([slide]);
      let rejectSync!: (e: unknown) => void;
      vi.stubGlobal("PowerPoint", {
        ...(globalThis as unknown as { PowerPoint: Record<string, unknown> }).PowerPoint,
        run: async (cb: (ctx: unknown) => Promise<unknown>) =>
          cb({
            presentation: { slides: { getItemAt: () => slide }, getSelectedSlides: () => ({ getItemAt: () => slide }) },
            sync: () => new Promise<void>((_, rej) => (rejectSync = rej)),
          }),
      });
      const p = insertSceneIntoSlide(buildChart(config), {});
      const assertion = expect(p).rejects.toThrow(/did not respond/);
      await vi.advanceTimersByTimeAsync(400_000); // past max(45s, shapes*3s)
      await assertion;
      expect(heard, "nothing heard before the host answers").toHaveLength(0);

      // Now the host finally answers — with a real RichApi-shaped error.
      rejectSync({
        message: "An internal error has occurred.",
        code: "GeneralException",
        debugInfo: { errorLocation: "Shape.name" },
      });
      await vi.advanceTimersByTimeAsync(1);
      expect(heard).toHaveLength(1);
      // The generic message alone is useless; code + debugInfo name the bug.
      expect(heard[0]).toContain("the host eventually FAILED");
      expect(heard[0]).toContain("code=GeneralException");
      expect(heard[0]).toContain("Shape.name");
    } finally {
      vi.useRealTimers();
    }
  });

  it("says so when the host was merely slow, not broken", async () => {
    vi.useFakeTimers();
    const heard: string[] = [];
    onLateSync((m) => heard.push(m));
    try {
      const slide = makeSlide("s1");
      installHost([slide]);
      let finish!: () => void;
      vi.stubGlobal("PowerPoint", {
        ...(globalThis as unknown as { PowerPoint: Record<string, unknown> }).PowerPoint,
        run: async (cb: (ctx: unknown) => Promise<unknown>) =>
          cb({
            presentation: { slides: { getItemAt: () => slide }, getSelectedSlides: () => ({ getItemAt: () => slide }) },
            sync: () => new Promise<void>((res) => (finish = res)),
          }),
      });
      const p = insertSceneIntoSlide(buildChart(config), {});
      const assertion = expect(p).rejects.toThrow(/did not respond/);
      await vi.advanceTimersByTimeAsync(400_000); // past max(45s, shapes*3s)
      await assertion;
      finish();
      await vi.advanceTimersByTimeAsync(1);
      // "SUCCEEDED late" means the timeout is too short — a different bug from
      // a host that is actually broken, and the note has to distinguish them.
      expect(heard[0]).toContain("the host eventually SUCCEEDED");
    } finally {
      vi.useRealTimers();
    }
  });
});

describe("added slides use the blank layout", () => {
  it("asks for the blank layout by TYPE, not by its localised name", async () => {
    // A slide added with no layout inherits the PREVIOUS slide's — on a fresh
    // deck that is the title slide, so an agenda lands on top of "Click to add
    // title" with the placeholder showing through. We draw everything
    // ourselves and want no placeholders.
    // The master here is Danish ("Tom"), which is the point: matching the name
    // "Blank" would silently do nothing for most of the world.
    installHost([makeSlide("s1")]);
    await insertAgendaSlides([
      buildAgendaScene(["Intro", "Body"], { highlight: 0 }),
      buildAgendaScene(["Intro", "Body"], { highlight: 1 }),
    ]);
    expect(addedWithLayout).toEqual(["layout-blank", "layout-blank"]);
  });

  it("uses it for the demo deck too", async () => {
    installHost([makeSlide("s1")]);
    await insertDemoDeck([{ scene: buildChart(config), tagData: "{}" }, { scene: buildChart(config) }]);
    expect(addedWithLayout).toEqual(["layout-blank", "layout-blank"]);
  });

  it("still adds slides on a host that exposes no masters", async () => {
    // Layout choice is a nicety; inserting is not. If the host will not tell us
    // its layouts, fall back to the inherited one rather than failing.
    const ctx = installHost([makeSlide("s1")]);
    (ctx.presentation as unknown as { slideMasters: unknown }).slideMasters = {
      load() {},
      get items(): never {
        throw new Error("masters unavailable on this host");
      },
    };
    await insertAgendaSlides([buildAgendaScene(["Intro"], { highlight: 0 })]);
    expect(addedWithLayout).toEqual([undefined]);
  });
});

describe("the wait budget scales with the work", () => {
  /** Park the sync so we can watch the clock without the host ever answering. */
  const parkedHost = (slide: FakeSlide) =>
    vi.stubGlobal("PowerPoint", {
      ...(globalThis as unknown as { PowerPoint: Record<string, unknown> }).PowerPoint,
      run: async (cb: (ctx: unknown) => Promise<unknown>) =>
        cb({
          presentation: { slides: { getItemAt: () => slide }, getSelectedSlides: () => ({ getItemAt: () => slide }) },
          sync: () => new Promise<void>(() => {}),
        }),
    });

  /** A scene of `n` trivial shapes — the budget is a function of the count. */
  const sceneOf = (n: number) => ({
    width: 400,
    height: 300,
    nodes: Array.from({ length: n }, (_, i) => ({ kind: "rect" as const, x: i, y: 0, w: 4, h: 4, fill: "#111111" })),
  });

  it("never hands the host more than a batch at once — THE bug", async () => {
    // Measured against real PowerPoint on the web: ~10 shapes insert instantly,
    // the 18-shape table element works, a 30-shape butterfly NEVER commits —
    // the sync simply stops answering and nothing lands. The same shapes go
    // onto off-screen slides by the hundred, because those are not painted.
    // So the fix was never a bigger timeout; it was a smaller batch.
    const slide = makeSlide("s1");
    installHost([slide]);
    const perSync: number[] = [];
    let last = 0;
    const ctx = installHost([slide]);
    ctx.sync = async () => {
      trips.syncs++;
      perSync.push(slide.created.length - last);
      last = slide.created.length;
    };
    const scene = buildChart(config);
    expect(scene.nodes.length).toBeGreaterThan(10); // must actually span batches
    await insertSceneIntoSlide(scene, { tagData: "{}" });
    expect(Math.max(...perSync), `handed over at once: ${perSync.join(",")}`).toBeLessThanOrEqual(10);
  });

  // The batching above counts SHAPES. A node is not a shape: a wedge fans into
  // triangles and a polygon becomes one line per edge, so the kinds that flood
  // the host are exactly the ones the all-rect `stacked` config cannot exercise.
  it.each(["pie", "doughnut", "sunburst", "radar", "violin"])(
    "batches %s by shapes, not nodes — the wedge/polygon flood",
    async (kind) => {
      const slide = makeSlide(`s-${kind}`);
      const perSync: number[] = [];
      let last = 0;
      const ctx = installHost([slide]);
      ctx.sync = async () => {
        trips.syncs++;
        perSync.push(slide.created.length - last);
        last = slide.created.length;
      };
      const scene = buildChart(sampleConfig(kind as never));
      await insertSceneIntoSlide(scene, { tagData: "{}" });
      // A single indivisible node (a wedge fan) may exceed the budget on its own;
      // nothing may exceed the host's measured breaking point of ~18.
      expect(Math.max(...perSync), `${kind} handed over at once: ${perSync.join(",")}`).toBeLessThanOrEqual(18);
    },
  );

  it("still bounds a trivial insert — the floor, not zero", async () => {
    vi.useFakeTimers();
    try {
      const slide = makeSlide("s1");
      installHost([slide]);
      parkedHost(slide);
      let settled = false;
      // 1 shape: the per-shape budget is tiny, so the 45s floor is what holds —
      // and 30s is past the old flat 20s, which is the thing that broke.
      const p = insertSceneIntoSlide(sceneOf(1), {}).catch(() => void (settled = true));
      await vi.advanceTimersByTimeAsync(30_000);
      expect(settled, "the floor keeps a small insert waiting past 30s").toBe(false);
      await vi.advanceTimersByTimeAsync(200_000);
      await p;
      expect(settled).toBe(true);
    } finally {
      vi.useRealTimers();
    }
  });
});

describe("EVERY insert path batches its shapes", () => {
  /** Max shapes handed to the host in any single sync of `run`. */
  async function maxPerSync(run: () => Promise<unknown>, slides: FakeSlide[]) {
    const worst = { n: 0 };
    let last = 0;
    const ctx = installHost(slides);
    // Wrap, don't replace: addSlides now verifies its adds via a settled
    // getCount() in a fresh context (see the addSlides retry/verify test),
    // which needs the real sync's committedCount/pendingCounts bookkeeping.
    // A bare replacement (as this used to be) freezes getCount() at its
    // initial value forever, which reads as "every add() was lost" and
    // starves insertAgendaSlides/insertDemoDeck of slides to render onto.
    const realSync = ctx.sync;
    const count = () => slides.reduce((a, s) => a + s.created.length, 0);
    ctx.sync = async () => {
      worst.n = Math.max(worst.n, count() - last);
      last = count();
      await realSync();
    };
    await run();
    return worst.n;
  }

  it("insert, update, agenda AND demo deck — none may send a whole scene", async () => {
    // The omission this exists for: I chunked insertSceneIntoSlide and
    // updateChartsInSlides and forgot insertAgendaSlides and insertDemoDeck.
    // The demo deck kept handing over ~200 shapes (4 slides at once) and sat at
    // "Working… 845s" having added nothing — and reported no progress, because
    // progress only fires when a chunk COMPLETES and the first never did.
    //
    // Live-canvas paths stay at ≤10 (repaints choke past that). Off-screen
    // append paths use a larger batch — the host tolerates far more when it
    // isn't repainting — but still bounded; they must NOT hand over the whole
    // scene, so a value at or under SHAPES_PER_SYNC_OFFSCREEN (40) is the
    // invariant, not the old flat 10.
    const scene = () => buildChart(config);
    expect(scene().nodes.length).toBeGreaterThan(10); // must span batches

    const s1 = makeSlide("s1");
    expect(
      await maxPerSync(() => insertSceneIntoSlide(scene(), { tagData: "{}" }), [s1]),
      "insert",
    ).toBeLessThanOrEqual(10);

    const s2 = makeSlide("s2");
    const old = s2.shapes.addGeometricShape("rectangle", { left: 0, top: 0, width: 1, height: 1 });
    expect(
      await maxPerSync(
        () => updateChartInSlide(scene(), { slideId: "s2", shapeId: old.id, left: 0, top: 0 }, {}),
        [s2],
      ),
      "update",
    ).toBeLessThanOrEqual(11); // +1: the pre-existing shape this test planted

    const s3 = makeSlide("s3");
    expect(await maxPerSync(() => insertAgendaSlides([scene(), scene()]), [s3]), "agenda").toBeLessThanOrEqual(40);

    const s4 = makeSlide("s4");
    expect(
      await maxPerSync(
        () =>
          insertDemoDeck([
            { scene: scene() },
            { scene: scene() },
            { scene: scene() },
            { scene: scene() },
            { scene: scene() },
          ]),
        [s4],
      ),
      "demo deck",
    ).toBeLessThanOrEqual(40);
  });

  it("off-screen demo/agenda batches larger than the live canvas — cuts syncs per chart", async () => {
    // The live canvas caps at 10 shapes per batch (repaint mid-render kills the
    // host past that). Off-screen slides don't repaint, so demo/agenda push a
    // larger batch — 40 — cutting ~4x round-trips per chart. This asserts the
    // demo path ACTUALLY sends more than the live-canvas ceiling for a scene
    // that could fit in one 40-batch.
    const scene = () => ({
      width: 100,
      height: 100,
      nodes: Array.from({ length: 25 }, (_, k) => ({
        kind: "rect" as const,
        x: k,
        y: 0,
        w: 4,
        h: 4,
        fill: "#111111",
      })),
    });
    installHost([makeSlide("s-live")]);
    // Live canvas: caps at 10.
    const live = await maxPerSync(() => insertSceneIntoSlide(scene(), { tagData: "{}" }), [makeSlide("s-live")]);
    expect(live, "live canvas ≤10").toBeLessThanOrEqual(10);
    // Off-screen demo: uses ≥15 in some batch — proves the flat 10 cap is gone.
    installHost([makeSlide("s-off")]);
    const off = await maxPerSync(() => insertDemoDeck([{ scene: scene() }]), [makeSlide("s-off")]);
    expect(off, "off-screen batches larger than live").toBeGreaterThan(10);
  });
});

describe("a target whose slide is gone is nothing to do, not a crash", () => {
  it("skips a stale slideId instead of throwing InvalidParam", async () => {
    // The real error, from the real host:
    //   InvalidParam passed to GetItem(id) | code=5010
    //   errorLocation: SlideCollection.getItem
    // An EditTarget outlives the slide it names — delete the slide, undo, or
    // reopen the deck and the id is stale. getItem THROWS on that; it is a
    // normal condition wearing a crash's clothes. Same Scale over a deck would
    // take one deleted chart and lose every OTHER chart's rescale with it.
    const live = makeSlide("s-live");
    installHost([live]);
    const s = live.shapes.addGeometricShape("rectangle", { left: 0, top: 0, width: 1, height: 1 });
    await expect(
      updateChartsInSlides([
        {
          scene: buildChart(config),
          target: { slideId: "s-deleted", shapeId: "gone", left: 0, top: 0 },
          opts: { tagData: "{}" },
        },
        {
          scene: buildChart(config),
          target: { slideId: "s-live", shapeId: s.id, left: 10, top: 20 },
          opts: { tagData: '{"ok":1}' },
        },
      ]),
      // One refreshed target back: the live chart's. The dead one contributes
      // nothing, and must not take the live one down with it.
    ).resolves.toHaveLength(1);
    // The live chart still got drawn and tagged — one dead target must not take
    // the others down.
    const group = live.created.find((c) => c.type === "group");
    expect(group, "the live chart was skipped too").toBeTruthy();
    expect(group!.tagStore.get(CHART_TAG)).toBe('{"ok":1}');
  });

  it("does nothing at all when every target is stale", async () => {
    const slide = makeSlide("s1");
    installHost([slide]);
    const before = slide.created.length;
    await expect(
      updateChartsInSlides([
        { scene: buildChart(config), target: { slideId: "nope", shapeId: "nope", left: 0, top: 0 }, opts: {} },
      ]),
    ).resolves.toEqual([]);
    expect(slide.created.length).toBe(before);
  });
});

/**
 * The repair pass, against the host fake — `src/core/reconcile.ts` decides
 * what to do, this is the half that talks to PowerPoint and has to survive it.
 */
describe("charts too dense for the web to draw", () => {
  const web = { web: true, canPicture: true, alreadyPicture: false };

  it("rasterises a chart past the budget on the web", () => {
    // Violin is 253 native shapes; area 176; tile map 122; waffle 103. On the
    // host with no resource limits at all, those are the charts that take the
    // tab down rather than merely drawing slowly.
    expect(wantsAutoPicture(253, web)).toBe(true);
    expect(wantsAutoPicture(91, web)).toBe(true);
  });

  it("leaves an ordinary chart alone", () => {
    // Gantt at 31 and Heatmap at 67 have drawn as shapes on the web all along.
    expect(wantsAutoPicture(31, web)).toBe(false);
    expect(wantsAutoPicture(67, web)).toBe(false);
    expect(wantsAutoPicture(90, web)).toBe(false);
  });

  it("never overrides the user's own choice of picture mode", () => {
    expect(wantsAutoPicture(253, { ...web, alreadyPicture: true })).toBe(false);
  });

  it("draws shapes on desktop however dense the chart is", () => {
    // Desktop has the resource limits the web lacks: Office throttles or
    // restarts the ADD-IN there rather than letting the client die, so the
    // native shapes the user actually wants stay the right answer.
    expect(wantsAutoPicture(253, { ...web, web: false })).toBe(false);
  });

  it("does not promise a picture a host cannot insert", () => {
    // Below PowerPointApi 1.8 there is no setImage to call.
    expect(wantsAutoPicture(253, { ...web, canPicture: false })).toBe(false);
  });
});

describe("reading a demo deck back and repairing it", () => {
  /** A slide as a damaged run leaves it: some shapes, maybe a banner, maybe a group. */
  function demoSlide(
    id: string,
    opts: {
      slot?: { i: number; title: string };
      shapes?: number;
      stamped?: boolean;
      tagged?: boolean;
      grouped?: boolean;
      /**
       * A degraded picture: ONE shape named PowerChart that is NOT a group.
       * The readback calls it `grouped` (it matches the name) but cannot count
       * children, so the slide comes back unmeasured — exactly the state a run
       * that fell back to pictures leaves behind.
       */
      picture?: boolean;
    },
  ): FakeSlide {
    const slide = makeSlide(id);
    if (opts.slot) slide.tags.add(DEMO_SLOT_TAG, JSON.stringify(opts.slot));
    if (opts.picture) {
      const pic = slide.shapes.addGeometricShape("rectangle", { left: 0, top: 0, width: 100, height: 100 });
      pic.name = "PowerChart";
      if (opts.tagged) pic.tags.add(CHART_TAG, `{"kind":"line"}`);
      return slide;
    }
    const parts: FakeShape[] = [];
    for (let i = 0; i < (opts.shapes ?? 0); i++) {
      const shape = slide.shapes.addTextBox(`n${i}`, { left: 0, top: 0, width: 10, height: 10 });
      shape.name = `part-${i}`;
      parts.push(shape);
    }
    if (opts.grouped && parts.length) {
      const group = slide.shapes.addGroup(parts);
      group.name = "PowerChart";
      if (opts.tagged) group.tags.add(CHART_TAG, `{"kind":"line"}`);
    } else if (opts.tagged && parts.length) {
      parts[parts.length - 1].tags.add(CHART_TAG, `{"kind":"line"}`);
    }
    if (opts.stamped) {
      const banner = slide.shapes.addGeometricShape("rectangle", { left: 0, top: 0, width: 100, height: 20 });
      banner.name = "PowerChart:not-complete";
    }
    return slide;
  }

  const expect3 = (slot: number, title: string, chart = true) => ({ slot, title, shapes: 3, chart });

  /** A two-point line chart — small enough that one batch draws the whole thing. */
  const tinyChart = (): ChartConfig => ({
    ...sampleConfig("line"),
    title: "Line",
    data: { categories: ["A", "B"], series: [{ name: "S1", values: [1, 2] }] },
  });

  it("reads shape counts, banners, slot tags and group children off the deck", async () => {
    const deck = [
      demoSlide("s0", { shapes: 1 }),
      demoSlide("s1", { slot: { i: 0, title: "Line" }, shapes: 3, tagged: true }),
      demoSlide("s2", { slot: { i: 1, title: "Gantt" }, shapes: 3, stamped: true }),
      demoSlide("s3", { slot: { i: 2, title: "Pie" }, shapes: 3, grouped: true, tagged: true }),
    ];
    installHost(deck);
    const snaps = await snapshotAddedSlides(0, 4);
    expect(snaps.map((s) => s.slot)).toEqual([null, 0, 1, 2]);
    expect(snaps.map((s) => s.title)).toEqual([null, "Line", "Gantt", "Pie"]);
    expect(snaps[1]).toMatchObject({ shapes: 3, stamped: false, tagged: true });
    expect(snaps[2]).toMatchObject({ shapes: 4, stamped: true, tagged: false });
    // The group is one top-level shape; without its child count a complete
    // chart would read as 1 of 3 and be condemned as wreckage.
    expect(snaps[3]).toMatchObject({ grouped: true, groupChildren: 3, tagged: true });
  });

  /**
   * Pass C was the one pass that was not paged.
   *
   * It opened its own context with two syncs for EVERY grouped slide, which
   * made `READBACK_TIMEOUT_MS` meaningless in aggregate — the budget is per
   * slide, so sixty grouped slides is ninety minutes of rope with a full ninety
   * seconds left at every step. On a real 38-slide run it accounted for a
   * 49-second gap in the log, 43% of the post-insert wall clock, immediately
   * before the tab died.
   *
   * Counted in CONTEXTS rather than timed: the cost is round trips, and a
   * duration assertion on a fake host measures nothing.
   */
  it("counts group children a page at a time, not a context per slide", async () => {
    const n = READBACK_PAGE + 5;
    const deck = Array.from({ length: n }, (_, i) =>
      demoSlide(`s${i}`, { slot: { i, title: "Line" }, shapes: 3, grouped: true, tagged: true }),
    );
    installHost(deck);
    const before = trips.contexts;
    const snaps = await snapshotAddedSlides(0, n);
    // The measurement still happens — a pass that stopped answering would make
    // the repair treat every chart as unmeasured, which is worse than slow.
    expect(
      snaps.every((s) => s.groupChildren === 3),
      "lost the group-child counts",
    ).toBe(true);
    // Pass A + pass B + pass C, each paged. One context per grouped SLIDE would
    // be n on its own.
    expect(trips.contexts - before, "opened a context per grouped slide").toBeLessThan(n);
  });

  it("deletes a duplicate slide, clears a stale banner, and re-groups a loose chart", async () => {
    // The shape of Presentation_4.pptx: a clean chart, the same chart again
    // under a NOT COMPLETE banner, and an empty slide the host left behind.
    const deck = [
      demoSlide("title", { slot: { i: 0, title: "Title" }, shapes: 3, grouped: true }),
      demoSlide("line", { slot: { i: 1, title: "Line" }, shapes: 3 }),
      demoSlide("line-dup", { slot: { i: 1, title: "Line" }, shapes: 3, stamped: true }),
      demoSlide("stray", {}),
    ];
    installHost(deck);
    const outcome = await reconcileDeck(
      [expect3(0, "Title", false), expect3(1, "Line")],
      { before: 0, after: 4 },
      () => `{"kind":"line"}`,
      { dropOrphanBlanks: true },
    );
    expect(outcome.applied).toEqual({ unstamped: 0, regrouped: 1, deleted: 2 });
    expect(outcome.refused).toBe(0);
    expect(deck.map((s) => s.id)).toEqual(["title", "line"]);
    // The surviving chart is now a group carrying the config — re-editable.
    const group = deck[1].created.filter((s) => !s.deleted).find((s) => s.name === "PowerChart");
    expect(group?.tagStore.get(CHART_TAG)).toBe(`{"kind":"line"}`);
  });

  it("stops the readback when the user asks, and counts what it never looked at", async () => {
    // Stop used to reach exactly three places: a drawing batch boundary,
    // between demo items, and between charts in an update. The repair pass —
    // readback, tag pass, group count, the deletes — checked nothing. So a run
    // that got past drawing could not be cancelled at all: the pane switched
    // its button to "Stopping…" and the counter kept climbing. Observed at
    // 1819 seconds on a real host, and the run had to be abandoned by closing
    // the tab, which is also how its log was lost.
    //
    // What the pages it never reached are called matters as much as that it
    // stops. UNREAD, not clean: an unseen slide is never deleted by the pass
    // that follows, and "we did not look" must not read as "nothing there".
    const deck = Array.from({ length: 40 }, (_, i) =>
      demoSlide(`s${i}`, { slot: { i, title: `S${i}` }, shapes: 3, tagged: true }),
    );
    installHost(deck);
    requestStop();
    try {
      const { snapshots, unread } = await readAddedSlides(0, 40);
      expect(snapshots, "kept reading after a stop").toHaveLength(0);
      expect(unread, "a stopped readback reported slides as read").toBe(40);
    } finally {
      resetStop();
    }
  });

  it("gives up on a repair-pass page the host never answers", async () => {
    // 75 of this file's 79 syncs had no timeout. The four that did are all in
    // the insert path, so the drawing phase was bounded and everything after it
    // — group, tag, readback, repair — could wait forever on one unanswered
    // sync. That is the other half of the 1819-second run: even without a stop,
    // nothing would ever have broken the wait.
    const deck = [demoSlide("a", { slot: { i: 0, title: "A" }, shapes: 3, tagged: true })];
    installHost(deck);
    _setReadbackTimeoutForTest(5);
    stallSyncOn.add(1);
    try {
      const { snapshots, unread } = await readAddedSlides(0, 1);
      // Abandoned, and reported as unread — the same honest answer a page that
      // threw already gets. The call RETURNS, which is the whole point.
      expect(snapshots).toHaveLength(0);
      expect(unread).toBe(1);
    } finally {
      _setReadbackTimeoutForTest(90_000);
      stallSyncOn.clear();
    }
  });

  it("counts the slides a silent deck insert landed, instead of calling it a failure", async () => {
    // office-js#1650, verbatim: "the first time `context.sync()` is called the
    // promise resolves, but in subsequent calls the promise doesn't resolve,
    // although **the slide still gets added successfully**."
    //
    // Every timeout in this file used to throw. For a READ that is right — an
    // unread page is unread. For a WRITE it is wrong twice: it discards work
    // that landed, and it sends the caller off to do the work again, which is
    // how one stalled insert becomes two copies of a chart. This function
    // already measured the deck before and after; it simply never reached the
    // measurement on the runs that needed it.
    const built = await buildDeckBase64(
      [{ scene: buildChart(sampleConfig("clustered")), title: "A", configJson: "{}", slot: 0, run: "r1" }],
      { width: 720, height: 405 },
    );
    installHost([makeSlide("s1")]);
    faults.deckInsertNeverAnswers = true;
    // Generous on purpose, and this number is a FLOOR rather than a tuned
    // value. The insert's deadline fires either way — `deckInsertNeverAnswers`
    // hangs the sync forever — so nothing is weakened by giving it room. What a
    // tight budget did instead was fire during the round trip BEFORE the
    // insert, the one that reads the last slide's id to append rather than
    // front the deck: then the insert never ran, nothing landed, and the test
    // failed claiming work had been thrown away. At 40ms that lost a CI run and
    // reproduced locally 2 times in 5 while the machine was busy, and not at
    // all when it was idle — the signature of a race, not of a regression.
    _setBatchTimeoutForTest(500);
    _setDeckInsertPerSlideForTest(500);
    try {
      const landed = await insertSlidesFromPptx(built.base64, 1);
      expect(landed, "threw away a slide the host had already added").toBe(1);
    } finally {
      faults.deckInsertNeverAnswers = false;
      _setBatchTimeoutForTest(30_000);
      _setDeckInsertPerSlideForTest(5_000);
    }
  });

  it("never asks the host to clear the selection with an empty array", async () => {
    // office-js#3698, verbatim: `slide.setSelectedShapes([])` on the web "does
    // not clear the selection, causes the `PowerPoint.run` promise to never
    // resolve, and produces no error messages."
    //
    // So the call cannot succeed on the host it was written for, and is a
    // documented candidate for the silence this project measured. It is gone.
    // Re-selecting the SLIDE drops the shape selection on every host observed,
    // and was already there as the fallback — what is left is the half that
    // works.
    //
    // Asserted by counting the call rather than by inspecting its effect: the
    // effect of a call that never resolves is precisely what nobody can observe.
    installHost([makeSlide("s1")]);
    const before = trips.emptyDeselects;
    await clearShapeSelection("s1");
    expect(trips.emptyDeselects - before, "made the call office-js#3698 says never returns").toBe(0);
  });

  it("does not call a slide untagged when the host answered with no shapes", async () => {
    // The confirmed mechanism. A real 38-slide run had a readback page ask
    // about 19 slides carrying 19 shapes and see 3 — nothing threw, the
    // collection was simply short. Read naively that is "no shapes, so no
    // tags, so not re-editable", and the repair then rewrote 14 charts whose
    // config was already correct.
    const deck = [demoSlide("a", { slot: { i: 0, title: "A" }, shapes: 3, tagged: true })];
    installHost(deck);
    // Both the first pass and its re-read come back hollow.
    faults.hollowReads = 2;
    const { snapshots } = await readAddedSlides(0, 1);
    expect(snapshots[0].tagged).toBe(false); // could not see it…
    expect(snapshots[0].tagRead).toBe(false); // …and says so, which is the point
    faults.hollowReads = 0;
  });

  it("gets the right answer on the re-read when only the first look was hollow", async () => {
    const deck = [demoSlide("a", { slot: { i: 0, title: "A" }, shapes: 3, tagged: true })];
    installHost(deck);
    faults.hollowReads = 1; // first pass short, second fine
    const { snapshots } = await readAddedSlides(0, 1);
    expect(snapshots[0].tagged).toBe(true);
    expect(snapshots[0].tagRead).not.toBe(false);
    faults.hollowReads = 0;
  });

  it("traces what the tag pass asked for against what it got back", async () => {
    // The open question this exists to answer: a 39-slide run reported 20
    // tagged charts where the file provably carried 31, every miss in the
    // second page. Two candidates — collections coming back short, or
    // collections full but tag lookups resolving null — need different fixes
    // and are indistinguishable from the log as it stood. These numbers
    // separate them.
    setTracing(true);
    try {
      const deck = [
        demoSlide("a", { slot: { i: 0, title: "A" }, shapes: 3, tagged: true }),
        demoSlide("b", { slot: { i: 1, title: "B" }, shapes: 3 }),
      ];
      installHost(deck);
      await readAddedSlides(0, 2);
      const page = traceLog().entries.find((e) => e.message === "tag pass over a page");
      expect(page).toBeDefined();
      // Asked about both slides, saw every shape pass A counted, found the one
      // tag that is really there. A hollow read would show shapesSeen short of
      // shapesExpected — which is the signal the next real run has to produce.
      expect(page!.data).toMatchObject({ slides: 2, shapesExpected: 6, shapesSeen: 6, tagsFound: 1 });
    } finally {
      setTracing(false);
    }
  });

  it("counts the slides it could not read, instead of calling them lost", async () => {
    // A page whose read throws is skipped — safe, since an unseen slide is
    // never deleted. But its items then came back "lost", indistinguishable in
    // the run summary from slides the host really dropped. "Gantt: lost" when
    // Gantt rendered perfectly is exactly the report that makes someone insert
    // the deck a second time.
    // Two pages' worth, so one can be read and the other refused — a partial
    // read is the case that misreports; a total one is obvious.
    const deck = Array.from({ length: READBACK_PAGE + 5 }, (_, i) =>
      demoSlide(`s${i}`, { slot: { i, title: `T${i}` }, shapes: 3 }),
    );
    installHost(deck);
    failSyncsOn.add(trips.syncs + 2); // the SECOND page's read
    try {
      const { snapshots, unread } = await readAddedSlides(0, deck.length);
      expect(snapshots).toHaveLength(READBACK_PAGE);
      expect(unread).toBe(5);
    } finally {
      failSyncsOn.clear();
    }
  });

  it("writes the config tag onto a chart that is whole but untagged", async () => {
    // A degraded picture is one shape named PowerChart carrying no config.
    // Nothing to group — only the tag is missing — and until `retag` existed
    // there was no repair that could reach it.
    const deck = [demoSlide("pic", { slot: { i: 0, title: "Line" }, picture: true })];
    installHost(deck);
    const outcome = await reconcileDeck([expect3(0, "Line")], { before: 0, after: 1 }, () => `{"kind":"line"}`, {});
    expect(outcome.plan.actions.map((a) => a.kind)).toEqual(["retag"]);
    expect(outcome.applied.regrouped).toBe(1);
    expect(outcome.refused).toBe(0);
    // The chart object itself now carries the config — re-editable again.
    const chart = deck[0].created.filter((s) => !s.deleted).find((s) => s.name === "PowerChart");
    expect(chart?.tagStore.get(CHART_TAG)).toBe(`{"kind":"line"}`);
    // And nothing was grouped, because there was nothing to group.
    expect(deck[0].created.filter((s) => !s.deleted && s.type === "group")).toHaveLength(0);
  });

  it("never writes an origin tag when retagging — it cannot know where the chart was drawn", async () => {
    // The repair used to write `[caller.left, caller.top, shape.left,
    // shape.top]` from its DEFAULT origin of (60, 90). A generated deck
    // centres its charts, so on a real 38-slide run it rewrote 14 correct
    // origins of `[239.988, 120, 239.988, 120]` to `[60, 90, 239.988, 120]`.
    // Since an update renders at `origin + (live - anchor)`, every one of
    // those charts would have jumped ~180pt left on its first edit.
    const deck = [demoSlide("pic", { slot: { i: 0, title: "Line" }, picture: true })];
    installHost(deck);
    const chart = deck[0].created.find((s) => s.name === "PowerChart")!;
    chart.tagStore.set(CHART_ORIGIN_TAG, "[239.988,120,239.988,120]");
    const outcome = await reconcileDeck([expect3(0, "Line")], { before: 0, after: 1 }, () => `{"kind":"line"}`, {});
    expect(outcome.applied.regrouped).toBe(1);
    expect(chart.tagStore.get(CHART_TAG)).toBe(`{"kind":"line"}`); // the tag IS written
    expect(chart.tagStore.get(CHART_ORIGIN_TAG)).toBe("[239.988,120,239.988,120]"); // the origin is NOT
  });

  it("refuses the retag when there is no PowerChart object to put it on", async () => {
    const deck = [demoSlide("bare", { slot: { i: 0, title: "Line" }, shapes: 1 })];
    installHost(deck);
    // Shapes present but loose and unnamed: this is a regroup, not a retag.
    const outcome = await reconcileDeck(
      [{ slot: 0, title: "Line", shapes: 1, chart: true }],
      { before: 0, after: 1 },
      () => `{"kind":"line"}`,
      {},
    );
    expect(outcome.plan.actions.map((a) => a.kind)).toEqual(["regroup"]);
  });

  it("refuses to delete a slide that is no longer the one it profiled", async () => {
    // A plan is a list of POSITIONS, decided from a readback taken several
    // round-trips ago. Nothing locks the deck in between: the pane leaves most
    // of its buttons live during a run, and PowerPoint's own UI is always
    // there. Deleting position 1 because position 1 used to be a duplicate is
    // how a repair pass destroys whatever occupies it now.
    const deck = [
      demoSlide("keep", { slot: { i: 0, title: "Line" }, shapes: 3 }),
      demoSlide("theirs", { shapes: 9 }), // NOT what the snapshot describes
    ];
    installHost(deck);
    // What the pass believed the deck held when it decided to delete index 1.
    const snapshots = [
      { index: 0, slot: 0, title: "Line", run: null, shapes: 3, stamped: false, tagged: false },
      { index: 1, slot: 0, title: "Line", run: null, shapes: 3, stamped: false, tagged: false },
    ];
    const plan = planReconcile(snapshots, [expect3(0, "Line", false)]);
    expect(plan.actions.filter((a) => a.kind === "delete").map((a) => a.index)).toEqual([1]);
    const outcome = await applyReconcilePlan(plan, () => undefined, { left: 0, top: 0 }, snapshots);
    expect(outcome.applied.deleted).toBe(0);
    expect(outcome.refused).toBe(1);
    // Both slides still there — the stranger above all.
    expect(deck.map((s) => s.id)).toEqual(["keep", "theirs"]);
  });

  it("pulls the banner off a chart that is in fact complete", async () => {
    const deck = [demoSlide("agenda", { slot: { i: 0, title: "Agenda" }, shapes: 3, stamped: true })];
    installHost(deck);
    const outcome = await reconcileDeck([expect3(0, "Agenda", false)], { before: 0, after: 1 }, () => undefined, {});
    expect(outcome.applied.unstamped).toBe(1);
    expect(deck[0].created.filter((s) => !s.deleted).some((s) => s.name === "PowerChart:not-complete")).toBe(false);
  });

  it("deletes from the end, so an earlier delete cannot renumber a later one", async () => {
    // Ascending deletes would remove index 1, shifting B's duplicate into the
    // slot the plan meant for something else — and take a good slide with it.
    const deck = [
      demoSlide("a", { slot: { i: 0, title: "A" }, shapes: 3 }),
      demoSlide("a-dup", { slot: { i: 0, title: "A" }, shapes: 3, stamped: true }),
      demoSlide("b", { slot: { i: 1, title: "B" }, shapes: 3 }),
      demoSlide("b-dup", { slot: { i: 1, title: "B" }, shapes: 3, stamped: true }),
    ];
    installHost(deck);
    const outcome = await reconcileDeck(
      [expect3(0, "A"), expect3(1, "B")],
      { before: 0, after: 4 },
      () => undefined,
      {},
    );
    expect(outcome.applied.deleted).toBe(2);
    expect(deck.map((s) => s.id)).toEqual(["a", "b"]);
  });

  it("leaves a slide from an unrelated deck untouched", async () => {
    const deck = [demoSlide("theirs", { shapes: 2 }), demoSlide("ours", { slot: { i: 0, title: "Line" }, shapes: 3 })];
    installHost(deck);
    const outcome = await reconcileDeck([expect3(0, "Line")], { before: 0, after: 2 }, () => undefined, {});
    expect(outcome.plan.orphans.map((o) => o.index)).toEqual([0]);
    expect(deck.map((s) => s.id)).toEqual(["theirs", "ours"]);
  });

  it("never groups the NOT COMPLETE banner in with the chart", async () => {
    // A real run shipped a Line chart whose group held 37 shapes: 36 of them
    // the chart, one a red NOT COMPLETE stripe. Once inside, the banner is
    // invisible to every later repair — a snapshot reads top-level names — so
    // it rides along with the chart forever.
    // 18 of 20 shapes: enough to re-group, not enough to count as complete,
    // so the plan asks for a regroup with the banner still on the slide.
    const deck = [demoSlide("partial", { slot: { i: 0, title: "Line" }, shapes: 18, stamped: true })];
    installHost(deck);
    await reconcileDeck(
      [{ slot: 0, title: "Line", shapes: 20, chart: true }],
      { before: 0, after: 1 },
      () => undefined,
      {},
    );
    const live = deck[0].created.filter((s) => !s.deleted);
    const group = live.find((s) => s.name === "PowerChart")!;
    expect(group.grouped).toHaveLength(18);
    expect((group.grouped as { name?: string }[]).some((c) => c.name === "PowerChart:not-complete")).toBe(false);
  });

  it("clears the banner on a slide whose chart is already grouped", async () => {
    // NOTE ON WHAT THIS DOES AND DOES NOT PROVE: in this fake a grouped shape
    // stays visible at the top level, so the repair finds the banner there.
    // On a real host a group hides its children, and the deck from the
    // 2026-07-31 run has a banner buried INSIDE a chart's group — `deleteStamp`
    // reaches into the group for exactly that case, and that reach is NOT
    // exercised here. `test/reconcile.test.ts` covers the planning half
    // (a grouped, stamped slide still gets an unstamp action).
    const deck = [demoSlide("buried", { slot: { i: 0, title: "Line" }, shapes: 3, stamped: true })];
    const shapes = deck[0].created;
    const group = deck[0].shapes.addGroup(shapes.slice());
    group.name = "PowerChart";
    installHost(deck);
    const outcome = await reconcileDeck(
      [{ slot: 0, title: "Line", shapes: 3, chart: true }],
      { before: 0, after: 1 },
      () => undefined,
      {},
    );
    expect(outcome.applied.unstamped).toBe(1);
    expect(deck[0].created.filter((s) => !s.deleted).some((s) => s.name === "PowerChart:not-complete")).toBe(false);
  });

  it("stamps nothing when the item's own slide never landed", async () => {
    // `stampLastSlide` used to brand whatever was last in the deck. When the
    // host swallowed the add, that was the PREVIOUS item's slide — a real run
    // defaced a KPI tile that had rendered perfectly, because a results page
    // that never landed stamped it.
    const existing = makeSlide("theirs");
    existing.shapes.addTextBox("someone else's work", { left: 0, top: 0, width: 10, height: 10 });
    const deck: FakeSlide[] = [existing];
    installHost(deck);
    faults.swallowAdds = ADDS_TO_DEFEAT_ONE_SLIDE; // the add and every retry vanish
    _setBatchTimeoutForTest(5); // no slide to draw on — do not wait 45s for it
    try {
      // Nothing rendered, so insertDemoDeck rethrows — and it rethrows the
      // REASON. With every add dropped, `addSlides` hands back no thunk at all;
      // the destructured `getSlide` was then called unchecked and the run died
      // with "getSlide is not a function", a TypeError from renderer internals
      // standing in for a diagnosis the code had already made. Assert the
      // message, not merely that something threw — a bare toThrow() passes just
      // as happily on the TypeError.
      await expect(insertDemoDeck([{ scene: buildChart(tinyChart()), title: "Line" }])).rejects.toThrow(
        /did not add a slide/i,
      );
      expect(existing.created.some((s) => s.name === "PowerChart:not-complete")).toBe(false);
    } finally {
      _setBatchTimeoutForTest(45_000);
    }
    // Generous: an earlier test in this file can leave an abandoned sync
    // outstanding, and the failure path waits up to 5s for it to report.
  }, 20_000);

  it("stops drawing shapes and inserts pictures once the host falls behind", async () => {
    // We cannot catch the crash — the tab dies, there is no rejected promise —
    // so the run watches what comes BEFORE it. A budget of zero shapes stands
    // in for a host that has already spent its allowance.
    const deck: FakeSlide[] = [makeSlide("s1")];
    installHost(deck);
    const scene = buildChart(tinyChart());
    const report = await insertDemoDeck(
      [
        { scene, title: "One", tagData: `{"i":0}` },
        { scene, title: "Two", tagData: `{"i":1}` },
        { scene, title: "Three", tagData: `{"i":2}` },
      ],
      undefined,
      { pictureFor: async () => "iVBORw0KGgo=", shapeBudget: 0 },
    );
    expect(report.degradedAt).toBe(1);
    expect(report.degradeReason).toMatch(/shapes drawn/);
    // Item 1 drew its shapes; items 2 and 3 are one picture each.
    expect(report.results[0].created).toBeGreaterThan(1);
    expect(report.results[1].created).toBe(1);
    expect(report.results[2].created).toBe(1);
    // And the picture carries the config tag, so the chart is still editable.
    const last = deck[deck.length - 1].created.filter((s) => !s.deleted);
    expect(last.some((s) => s.imageBase64 === "iVBORw0KGgo=")).toBe(true);
  }, 20_000);

  it("reports per item whether the config tag actually committed", async () => {
    // Intent, recorded next to the settled readback's answer to the same
    // question. A real 38-slide run reported 20 tagged charts where the file
    // provably carried 31, and establishing that took unzipping the .pptx.
    const deck: FakeSlide[] = [makeSlide("s1")];
    installHost(deck);
    const report = await insertDemoDeck([
      { scene: buildChart(tinyChart()), title: "Tagged", tagData: `{"i":0}` },
      // No tagData: nothing to write, so `tagged` must be false — and that is
      // correct, not a fault. A reader has to be able to tell the two apart.
      { scene: buildChart(tinyChart()), title: "Untagged" },
    ]);
    expect(report.results.map((r) => !!r.tagged)).toEqual([true, false]);
  }, 20_000);

  it("reports tagged false when the host cannot write tags at all", async () => {
    // The case that produced 19 silently-untagged charts. `created` still
    // counts the shapes — they are on the slide — so only this field
    // distinguishes "drawn and editable" from "drawn and lost".
    const scene = buildChart(tinyChart());
    const item = { scene, title: "One", tagData: `{"i":0}` };

    // Control: a host that can tag reports true.
    installHost([makeSlide("ok")]);
    expect((await insertDemoDeck([item])).results[0].tagged).toBe(true);

    // A host below PowerPointApi 1.3 has no shape tags at all, so the chart is
    // drawn and is NOT re-editable. `created` still counts the shapes — they
    // are on the slide — so this field is the only thing that tells the two
    // states apart, which is exactly what a run of 19 silently-untagged charts
    // had no way to say.
    const deck: FakeSlide[] = [makeSlide("s1")];
    installHost(deck, [], undefined, (v) => v !== "1.3");
    const report = await insertDemoDeck([item]);
    expect(report.results[0].created).toBeGreaterThan(1);
    expect(report.results[0].tagged).toBe(false);
  }, 20_000);

  it("does not degrade a healthy run because somebody ELSE's stalled call answered", async () => {
    // The degrade signal was a global counter with no owner, so ANY abandoned
    // promise settling during a run bumped it. A chart the user edited before
    // the run started, whose stalled sync happened to answer during item 0,
    // degraded a perfectly healthy deck to rasters from item 1 — and reported
    // "the host answered after we gave up waiting" about an operation the run
    // never made. Ownership is captured when a call is ISSUED, which is the
    // only moment that distinguishes the two.
    const slide = makeSlide("s1");
    installHost([slide]);
    const realPowerPoint = (globalThis as unknown as { PowerPoint: Record<string, unknown> }).PowerPoint;
    _setBatchTimeoutForTest(5);
    let finishForeign!: () => void;
    // A call issued OUTSIDE any run, on a host that never answers it.
    vi.stubGlobal("PowerPoint", {
      ...realPowerPoint,
      run: async (cb: (ctx: unknown) => Promise<unknown>) =>
        cb({
          presentation: { slides: { getItemAt: () => slide }, getSelectedSlides: () => ({ getItemAt: () => slide }) },
          sync: () => new Promise<void>((res) => (finishForeign = res)),
        }),
    });
    await expect(insertSceneIntoSlide(buildChart(tinyChart()), {})).rejects.toThrow(/did not respond/);
    const seqBefore = lastLateSyncSeq;
    // Back to a healthy host for the run itself.
    installHost([makeSlide("s2")]);
    _setBatchTimeoutForTest(45_000);
    let released = false;
    const report = await insertDemoDeck(
      Array.from({ length: 4 }, (_, i) => ({ scene: buildChart(tinyChart()), title: `C${i}`, tagData: `{"i":${i}}` })),
      () => {
        // Mid-run, exactly as it happened: the foreign call finally answers.
        if (!released) {
          released = true;
          finishForeign();
        }
      },
      { pictureFor: async () => "iVBORw0KGgo=", shapeBudget: 10_000, runId: "the-run" },
    );
    // Non-vacuity: the foreign call really did answer late, during the run.
    expect(lastLateSyncSeq).toBeGreaterThan(seqBefore);
    expect(lastLateSyncOwner).toBeNull();
    // And the run, which never stalled, drew shapes throughout.
    expect(report.degradeReason).toBeUndefined();
    expect(report.degradedAt).toBeUndefined();
    // Same ownership bug, one level down and unfixed for longer: `lateOutcome`
    // read the global late-sync string with no owner check at all, so the item
    // that happened to be drawing when the foreign call answered was recorded
    // as having stalled. Not one of these four gave up on anything.
    expect(report.results.map((r) => r.lateOutcome)).toEqual([undefined, undefined, undefined, undefined]);
    expect(report.results.some((r) => r.abandoned)).toBe(false);
  }, 20_000);

  it("attributes a stall to the run that issued it", async () => {
    // The other half: ownership must not cost the signal it exists to sharpen.
    // A call THIS run made, answering after it gave up, is exactly the
    // host-drowning evidence the degrade path was built for — and it must
    // still be recognised as the run's own.
    installHost([makeSlide("s1")]);
    _setBatchTimeoutForTest(5);
    try {
      // Stall whatever sync the run reaches next — which one is wrapped in
      // withTimeout is an implementation detail, and pinning an index here
      // would silently stop testing anything the day it moves.
      for (let k = 1; k <= 40; k++) stallSyncOn.add(trips.syncs + k);
      await insertDemoDeck([{ scene: buildChart(tinyChart()), title: "One", tagData: `{"i":0}` }], undefined, {
        runId: "the-run",
      }).catch(() => {});
      stallSyncOn.clear();
      await waitForLateSync(500);
      expect(lastLateSyncOwner).toBe("the-run");
    } finally {
      _setBatchTimeoutForTest(45_000);
      stallSyncOn.clear();
    }
  }, 20_000);

  it("says which item gave up on a call, and pairs the late answer with it", async () => {
    // The other half of the ownership fix above. `abandoned` is the part that
    // is knowable the moment an item ends — a deadline fired inside it — and
    // it is the part a run log can always carry, because the host's eventual
    // answer routinely arrives minutes later, long after any wait a run can
    // afford. `lateOutcome` is then only ever set on an item that abandoned
    // something, so a healthy item can no longer inherit one.
    installHost([makeSlide("s1")]);
    _setBatchTimeoutForTest(20);
    try {
      // Stall everything item 0 reaches. The fake settles a stalled sync 40ms
      // in, well past the 20ms deadline, so item 0 is guaranteed to give up
      // and then be told how it went.
      for (let k = 1; k <= 40; k++) stallSyncOn.add(trips.syncs + k);
      const report = await insertDemoDeck(
        [
          { scene: buildChart(tinyChart()), title: "One", tagData: `{"i":0}` },
          { scene: buildChart(tinyChart()), title: "Two", tagData: `{"i":1}` },
        ],
        (done) => {
          // Item 0 is over: let the host be healthy again, on a deadline no
          // ordinary sync can trip. Item 1 must show a clean sheet.
          if (done === 1) {
            stallSyncOn.clear();
            _setBatchTimeoutForTest(45_000);
          }
        },
        { runId: "the-run" },
      );
      // The bug, first: the late answer belongs to the item that gave up.
      expect(report.results[0].lateOutcome, "item 0 abandoned a call and was never told how it ended").toMatch(
        /SUCCEEDED/,
      );
      // …and it is not handed on to the next item, which stalled on nothing.
      expect(report.results[1].lateOutcome, "item 1 was credited with item 0's late answer").toBeUndefined();
      // Which item stalled, readable on its own — the half that is knowable
      // the moment the item ends, whether or not the host ever answers.
      expect(report.results[0].abandoned).toBe(true);
      expect(report.results[1].abandoned).toBe(false);
    } finally {
      _setBatchTimeoutForTest(45_000);
      stallSyncOn.clear();
    }
  }, 20_000);

  it("keeps a degraded picture re-editable on a host that rejects stale proxies", async () => {
    // The root cause of the shape path's lost tags, from a real 38-item run:
    //
    //   InvalidParam passed to GetItem(id) | code=5010
    //   errorLocation: ShapeCollection.getItem
    //   statement: var shape = shapes.getItem(...); var tags = shape.tags;
    //
    // 28 times. A degraded picture is ONE shape, so it was not "groupable",
    // so the refresh that exists for exactly this trap never ran on it — and
    // its proxy was created in one sync and tagged in another. The trap was
    // never about grouping; it is about using a proxy across a sync boundary,
    // and tagging does that too.
    const deck: FakeSlide[] = [makeSlide("s1")];
    installHost(deck);
    faults.strictTags = true;
    try {
      const report = await insertDemoDeck(
        [
          { scene: buildChart(tinyChart()), title: "One", tagData: `{"i":0}` },
          { scene: buildChart(tinyChart()), title: "Two", tagData: `{"i":1}` },
        ],
        undefined,
        // Budget 0 degrades after the first item, so item 2 arrives as a picture.
        { pictureFor: async () => "iVBORw0KGgo=", shapeBudget: 0 },
      );
      expect(report.degradedAt).toBe(1);
      expect(report.results[1].created).toBe(1); // it IS the picture
      // …and it is re-editable, which is the whole claim the degrade makes.
      expect(report.results[1].tagged).toBe(true);
      const last = deck[deck.length - 1].created.filter((sh) => !sh.deleted);
      expect(last.some((sh) => sh.tagStore.get(CHART_TAG) === `{"i":1}`)).toBe(true);
    } finally {
      faults.strictTags = false;
    }
  }, 20_000);

  it("draws a too-dense chart as a picture once the run has degraded", async () => {
    // The density budget exists to stop a wedge/polygon flood timing the host
    // out. A picture is ONE shape — not a flood — so a run already drawing
    // pictures must not still be skipping its densest charts.
    //
    // It did. A real 38-item web run degraded at item 2 and then skipped and
    // stamped Area (176), Tile map (122), Waffle (103), Sunburst (101),
    // Violin (253) and Smoothed line (101), while the other thirty went on as
    // one-shape pictures in about a second each. Those six are exactly the
    // charts picture mode exists for, and they were the six it refused.
    const deck: FakeSlide[] = [makeSlide("s1")];
    installHost(deck);
    const dense = buildChart({ ...sampleConfig("waffle"), title: "Waffle" });
    expect(estimateOfficeShapes(dense)).toBeGreaterThan(DEMO_SHAPE_BUDGET);
    const report = await insertDemoDeck(
      [
        { scene: buildChart(tinyChart()), title: "One", tagData: `{"i":0}` },
        { scene: dense, title: "Waffle", tagData: `{"i":1}` },
      ],
      undefined,
      // Budget 0 degrades the run after the first item, so the dense chart is
      // reached with a picture available.
      { pictureFor: async () => "iVBORw0KGgo=", shapeBudget: 0 },
    );
    expect(report.degradedAt).toBe(1);
    // Drawn, not skipped — and as exactly one shape.
    expect(report.results[1].status).toBe("rendered");
    expect(report.results[1].created).toBe(1);
    const last = deck[deck.length - 1].created.filter((sh) => !sh.deleted);
    expect(last.some((sh) => sh.imageBase64 === "iVBORw0KGgo=")).toBe(true);
    expect(last.some((sh) => sh.name === "PowerChart:not-complete")).toBe(false);
  }, 20_000);

  it("still skips a too-dense chart when there is no picture to fall back on", async () => {
    // The budget must keep working on a host that offers no rasterizer —
    // otherwise the flood it exists to prevent goes straight through.
    const deck: FakeSlide[] = [makeSlide("s1")];
    installHost(deck);
    const dense = buildChart({ ...sampleConfig("waffle"), title: "Waffle" });
    const report = await insertDemoDeck([{ scene: dense, title: "Waffle", tagData: `{"i":0}` }], undefined, {});
    expect(report.results[0].status).toBe("skipped");
  }, 20_000);

  it("keeps drawing shapes when the host is keeping up", async () => {
    const deck: FakeSlide[] = [makeSlide("s1")];
    installHost(deck);
    const report = await insertDemoDeck(
      [{ scene: buildChart(tinyChart()), title: "One", tagData: `{"i":0}` }],
      undefined,
      { pictureFor: async () => "iVBORw0KGgo=", shapeBudget: 10_000 },
    );
    expect(report.degradedAt).toBeUndefined();
    expect(report.results[0].created).toBeGreaterThan(1);
  }, 20_000);

  it("never degrades when the caller offers no picture to fall back to", async () => {
    const deck: FakeSlide[] = [makeSlide("s1")];
    installHost(deck);
    const report = await insertDemoDeck([{ scene: buildChart(tinyChart()), title: "One" }], undefined, {
      shapeBudget: 0,
    });
    expect(report.degradedAt).toBeUndefined();
  }, 20_000);

  it("traces the run in enough detail to diagnose it afterwards", async () => {
    // Every hard thing in this project was diagnosed after the fact, from a
    // deck and a one-line summary. This is the record that was missing: which
    // item, what landed, what the host refused, and when it was given up on.
    const deck: FakeSlide[] = [makeSlide("s1")];
    installHost(deck);
    setTracing(true);
    try {
      await insertDemoDeck([{ scene: buildChart(tinyChart()), title: "Line", tagData: `{"i":0}` }]);
      const { entries } = traceLog();
      const byScope = (scope: string) => entries.filter((e) => e.scope === scope);
      // The per-item verdict, with the numbers a reader needs to judge it.
      const item = byScope("demo").find((e) => e.message === "item finished");
      expect(item?.data).toMatchObject({ i: 0, title: "Line", status: "rendered" });
      expect(item?.data?.created).toBeGreaterThan(1);
      // And the drawing itself, batch by batch — "died at batch 1 of 4" was
      // the entire diagnosis of the update stall.
      expect(byScope("draw").length).toBeGreaterThan(0);
      expect(byScope("draw")[0].data).toHaveProperty("total");
    } finally {
      setTracing(false);
    }
  }, 20_000);

  it("costs a run nothing while it is off", async () => {
    const deck: FakeSlide[] = [makeSlide("s1")];
    installHost(deck);
    setTracing(false);
    // Switching off deliberately KEEPS the log readable, so measure growth
    // rather than emptiness — the claim is that a disabled trace records
    // nothing new, not that it forgets what it already had.
    const before = traceLog().entries.length;
    await insertDemoDeck([{ scene: buildChart(tinyChart()), title: "Line" }]);
    expect(traceLog().entries).toHaveLength(before);
  }, 20_000);

  it("closes a demo run with the settled truth when asked for it", async () => {
    const deck: FakeSlide[] = [makeSlide("s1")];
    installHost(deck);
    const report = await insertDemoDeck(
      [{ scene: buildChart(tinyChart()), title: "Line", tagData: `{"i":0}` }],
      undefined,
      { reconcile: true },
    );
    expect(report.reconcile).toBeDefined();
    expect(report.reconcile?.plan.verdicts[0]).toMatchObject({ title: "Line", status: "rendered" });
    // A clean run needs no repair, and says so instead of inventing work.
    expect(report.reconcile?.plan.actions).toEqual([]);
    expect(report.blankSlides).toEqual([]);
  });

  it("does not reconcile unless the caller opts in", async () => {
    const deck: FakeSlide[] = [makeSlide("s1")];
    installHost(deck);
    const report = await insertDemoDeck([{ scene: buildChart(tinyChart()), title: "Line" }]);
    expect(report.reconcile).toBeUndefined();
  });
});

describe("stopping work in flight", () => {
  // The stop flag is module state, so a test that set it and threw would arm
  // every test after it. Cleared unconditionally.
  afterEach(() => resetStop());

  const cfgFor = (v: number): ChartConfig => ({
    ...config,
    data: { categories: ["A", "B"], series: [{ name: "S1", values: [v, v + 1] }] },
  });
  const targetsOn = (slide: FakeSlide, n: number) =>
    Array.from({ length: n }, (_, i) => {
      const s = slide.shapes.addGeometricShape("rectangle", { left: 0, top: 0, width: 1, height: 1 });
      return {
        scene: buildChart(cfgFor(i)),
        target: { slideId: slide.id, shapeId: s.id, left: 10, top: 20 },
        opts: { tagData: `{"i":${i}}` },
      };
    });

  it("stops at the next batch and keeps what already committed", async () => {
    // Office.js has no abort: a sync already handed to PowerPoint runs to
    // completion whatever we want. So "stop" means "queue nothing further",
    // and the batches that already landed stay on the slide — which is why
    // this throws rather than returning, so the caller cleans up rather than
    // grouping and tagging a half-drawn chart as a finished one.
    const slide = makeSlide("s1");
    installHost([slide]);
    const scene = buildChart(cfgFor(0));
    let commits = 0;
    let thrown: unknown;
    try {
      await insertSceneIntoSlide(scene, { shapesPerSync: 5 }, (phase) => {
        if (phase === "commit" && ++commits === 1) requestStop();
      });
    } catch (err) {
      thrown = err;
    }
    expect(isStopped(thrown), "did not report a stop").toBe(true);
    // The first batch is on the slide; the rest was never queued.
    const live = slide.created.filter((s) => !s.deleted);
    expect(live.length).toBeGreaterThan(0);
    expect(live.length).toBeLessThan(estimateOfficeShapes(scene));
    // And nothing half-finished was passed off as a chart.
    expect(live.some((s) => s.tagStore.has(CHART_TAG))).toBe(false);
  });

  /**
   * A chart with no config tag is on the slide and is not a PowerChart.
   *
   * `groupAndTagAll` has always answered this honestly — it returns `tagged` —
   * and the answer had nowhere to go: `EditTarget` carried no such field, so
   * the demo path (which has a repair pass) consumed it and the everyday insert
   * and in-place update (which have none) did not. The user gets "Done." in
   * green, clicking the chart says "the selection is not a PowerChart", and
   * reopening the deck loses the config for good. A real host produced this
   * four times in one run: *"a chart's tag could not even be queued"* followed
   * by *"tagging failed — charts are not re-editable until repaired"*.
   *
   * Two undefined `.tags` now, not one. A failure inside the drawing context is
   * no longer the end of the story — `settleAndTagChart` opens a fresh one and
   * writes the tag there, which is what makes that same real-host failure
   * survivable. So one is the case where the chart ends up re-editable after
   * all, and this test is about the case where it does not: the host refuses
   * the settled write too, and the caller has to be told the truth.
   */
  it("says so when the chart landed but its config did not", async () => {
    const slide = makeSlide("s1");
    installHost([slide]);
    faults.strictTags = true;
    // `.tags` undefined for the drawing context, the settle's by-id write, AND
    // the settle's collection-read fallback. Three writers, so refusing "every
    // tag write" now costs three — and the count going UP is the point of the
    // case below it.
    faults.tagsUndefinedOn = 3;
    try {
      const target = await insertSceneIntoSlide(buildChart(config), { tagData: '{"a":1}' });
      // The chart IS there — this is not a failed insert, and reporting it as
      // one would send the user to draw a second copy.
      expect(slide.created.filter((s) => !s.deleted).length, "nothing was drawn").toBeGreaterThan(0);
      expect(target, "lost the target for a chart that is on the slide").toBeTruthy();
      expect(target!.lost, "reported a chart with no config as fully editable").toBe("no-config");
    } finally {
      faults.strictTags = false;
      faults.tagsUndefinedOn = 0;
    }
  });

  it("says nothing of the sort when the config DID land", async () => {
    // The negative control: a flag that is always set is not a signal.
    installHost([makeSlide("s1")]);
    const target = await insertSceneIntoSlide(buildChart(config), { tagData: '{"a":1}' });
    expect(target?.lost, "marked a perfectly good chart as lost").toBeUndefined();
  });

  /**
   * The host from 2026-08-07: it will not name a shape by id, and it will read
   * the collection.
   *
   * Sixty-six errors in that run log, every one `InvalidParam passed to
   * GetItem(id)` at `errorLocation: ShapeCollection.getItem`. Among them was
   * `settleAndTagChart`'s own write — a slide and a shape both resolved fresh
   * inside a first sync of their own, refused anyway. The settle then gave up,
   * on the reasoning that a collection search "would only find a DIFFERENT
   * shape to put this chart's config on", and five charts shipped with no
   * config: `same scale across the deck` reported *"3 of 8 charts carry the
   * shared scale ... the update reported 5×no-config"*.
   *
   * The reasoning is sound with no id and wrong with one. The read loads
   * `items/id`, so the caller's id picks its own shape out of the answer — no
   * guess is involved, and a chart that is not in the answer is simply not
   * tagged. And a collection read is what this host DOES honour: the repair
   * pass landed 23 retags that way in the same run that lost 46 tag writes.
   */
  it("settles the config through a collection read when the host will not name the shape by id", async () => {
    const slide = makeSlide("s1");
    installHost([slide]);
    // Both halves of the real transcript, and both are needed.
    //
    // `refuseTagWrites = 1` is the run log's `tagging failed — charts are not
    // re-editable until repaired`: the drawing context's write goes, and the
    // settle is what has to save the chart. `refuseShapeById` is then the host
    // refusing the settle's by-id write too, which is the line after it.
    //
    // Arming only `refuseShapeById` is NOT enough, and it is worth saying why:
    // the drawing context tags an ungrouped chart through the proxy it created,
    // never by id, so the write lands, the settle never runs, and the case
    // passes against the unfixed file — a guard that proves nothing.
    faults.refuseTagWrites = 1;
    faults.refuseShapeById = true;
    try {
      const target = await insertSceneIntoSlide(buildChart(config), { tagData: '{"a":1}' });
      expect(slide.created.filter((s) => !s.deleted).length, "nothing was drawn").toBeGreaterThan(0);
      // The whole point: re-editable anyway. `lost` set here is the 5×no-config
      // verdict, reproduced.
      expect(target?.lost, "gave up on a chart the collection read could still have tagged").toBeUndefined();
      const tagged = slide.created.filter((s) => !s.deleted && s.tagStore.has(CHART_TAG));
      expect(tagged.length, "no shape on the slide carries the config").toBeGreaterThan(0);
    } finally {
      faults.refuseShapeById = false;
      faults.refuseTagWrites = 0;
    }
  });

  it("gives the two settle writes different names in the source", () => {
    // `settleAndTagChart` writes the tag twice by two different routes — once
    // by shape id, and once through a member of a collection re-read — and both
    // used to pass the SAME label to `boundedSync`. A refusal in a round log
    // then named a write without naming which write.
    //
    // Round 8 was decodable anyway, but only by reasoning across two traces:
    // the absence of `the host refused a settle by id` said the by-id branch
    // had been skipped, so the refusal had to be the collection read's. That is
    // an inference from a missing line, which is the reading this project gets
    // wrong most often. Distinct labels make it direct.
    //
    // Checked against the SOURCE because the labels only reach a log when a
    // write is refused, and arming a fake to refuse both halves in one run
    // exercises the fake's error plumbing rather than this property.
    const src = readFileSync("src/render/powerpoint.ts", "utf8");
    // The whole settle path: the by-id write, the collection re-read, and the
    // write through a member of that read. Sliced by function rather than
    // matched on wording, so renaming a label cannot quietly stop this looking.
    const from = src.indexOf("async function settleAndTagChart");
    const to = src.indexOf("async function settleUntaggedCharts");
    expect(from, "settleAndTagChart is gone — this guard is looking at nothing").toBeGreaterThan(0);
    expect(to, "settleUntaggedCharts is gone — the slice below has no end").toBeGreaterThan(from);
    const labels = [...src.slice(from, to).matchAll(/boundedSync\([^,]+,\s*[`"]([^`"]+)[`"]/g)].map((m) => m[1]);
    // FOUR since the name branch started checking that the chart it matched is
    // not already carrying someone else's config: the by-id write, the
    // collection re-read, that check, and the write through a member of the
    // read. The number is a tripwire for a sync added without thinking — raise
    // it deliberately, as here, or find out why one appeared.
    expect(labels.length, `expected four bounded syncs in the settle path, saw ${JSON.stringify(labels)}`).toBe(4);
    expect(new Set(labels).size, `two calls in the settle path share a label: ${JSON.stringify(labels)}`).toBe(
      labels.length,
    );
  });

  /** The settle pass's summary line, identified by its payload rather than its text. */
  const settleSummary = () =>
    traceLog().entries.find(
      (e) => e.scope === "group" && !!e.data && "charts" in e.data && "settled" in e.data && "lost" in e.data,
    );

  it("does not report a repair the settle did not make", async () => {
    // The 2026-08-08 round printed `settled the config tag the drawing context
    // could not write` five times, each carrying `settled: 0, lost: 1`. The
    // message named an outcome and the numbers underneath contradicted it, so
    // a reader scanning messages — which is how a 190-entry log is read — saw
    // five repairs that never happened.
    //
    // Every tag write refused, including the settle's own, so nothing can be
    // repaired and the line has to say so.
    const slide = makeSlide("s1");
    installHost([slide]);
    faults.refuseTagWrites = 9999;
    faults.refuseShapeById = true;
    setTracing(true);
    try {
      await insertSceneIntoSlide(buildChart(config), { tagData: '{"a":1}' }).catch(() => {});
      // Found by the shape of its DATA, not by its wording. Several lines in
      // this scope mention the settle; only its summary carries these three
      // keys — and picking it by wording would mean the guard could not see
      // the message it exists to judge.
      const line = settleSummary();
      expect(line, "the settle pass left no line at all").toBeTruthy();
      // Non-vacuity: this really is the nothing-was-repaired case. Without it
      // the assertion below would pass on a run where the settle succeeded.
      const data = line?.data as { charts: number; settled: number; lost: number };
      expect(data.charts, "no chart reached the settle").toBeGreaterThan(0);
      expect(data.settled, "the settle was supposed to fail here").toBe(0);
      expect(line?.message, `the log claims a repair it did not make: "${line?.message}"`).toMatch(/could not repair/);
    } finally {
      setTracing(false);
      faults.refuseShapeById = false;
      faults.refuseTagWrites = 0;
    }
  });

  it("says the settle repaired the tag when it did", async () => {
    // The other side of the same mapping, so the message cannot simply be
    // pessimistic and pass the guard above forever.
    const slide = makeSlide("s1");
    installHost([slide]);
    faults.refuseTagWrites = 1;
    faults.refuseShapeById = true;
    setTracing(true);
    try {
      await insertSceneIntoSlide(buildChart(config), { tagData: '{"a":1}' });
      const line = settleSummary();
      const data = line?.data as { settled: number; lost: number };
      expect(data.settled, "the settle was supposed to succeed here").toBeGreaterThan(0);
      expect(data.lost).toBe(0);
      expect(line?.message).toMatch(/repaired every/);
    } finally {
      setTracing(false);
      faults.refuseShapeById = false;
      faults.refuseTagWrites = 0;
    }
  });

  it("leaves every chart untouched when the stop lands before the first one", async () => {
    // updateChartsInSlides checks BEFORE each chart's delete, so a stop taken
    // there costs nothing: no shapes removed, nothing queued. The charts keep
    // their existing targets, which is what the caller needs to stay editable.
    const slide = makeSlide("s1");
    installHost([slide]);
    const items = targetsOn(slide, 3);
    const before = slide.created.filter((s) => !s.deleted).length;
    requestStop();
    const next = await updateChartsInSlides(items);
    // Every target comes back as it went in — none of them redrawn.
    expect(next).toEqual(items.map((it) => it.target));
    expect(slide.created.filter((s) => !s.deleted).length).toBe(before);
    expect(slide.created.some((s) => s.deleted)).toBe(false);
  });

  it("stops a deck run between items, keeping the slides already finished", async () => {
    // The longest thing the add-in does, and the reason a stop is worth having.
    // An item boundary costs nothing to stop at: every slide so far is complete,
    // grouped and tagged, and the next one has not been added.
    const deck: FakeSlide[] = [makeSlide("s1")];
    installHost(deck);
    const report = await insertDemoDeck(
      Array.from({ length: 4 }, (_, i) => ({ scene: buildChart(cfgFor(i)), tagData: `{"i":${i}}` })),
      (done) => {
        if (done === 1) requestStop();
      },
    );
    // One item ran; the run then stopped instead of drawing the other three.
    expect(report.results).toHaveLength(1);
    expect(report.results[0].status).toBe("rendered");
    // And the deck grew by exactly that one slide.
    expect(report.slidesAdded).toBe(1);
  });
});

describe("reading the presentation's slide size", () => {
  // Cached per deck, so a value from one test would answer for the next.
  beforeEach(() => _resetSlideSizeCache());
  afterEach(() => _resetSlideSizeCache());

  it("reads pageSetup directly when the host has 1.10", async () => {
    installHost([makeSlide("s1")]);
    hostSlideSize.cx = 9144000; // 4:3 — 720pt
    const size = await slideSize();
    expect(size).toEqual({ width: 720, height: 540, source: "pageSetup" });
  });

  it("falls back to exporting a slide when the host is below 1.10", async () => {
    // PageSetup arrived in 1.10. A 1.8 host still exports a slide as its own
    // .pptx, and that file declares the SOURCE deck's <p:sldSz> — exact, no
    // guessing, one rung down.
    installHost([makeSlide("s1")], [], undefined, (v) => v !== "1.10");
    hostSlideSize.cx = 9144000;
    const size = await slideSize();
    expect(size).toEqual({ width: 720, height: 540, source: "exportedSlide" });
  });

  it("assumes 16:9 — and says so — when no rung can answer", async () => {
    // The floor. A wrong-but-LABELLED width degrades placement to what it did
    // before; throwing here would take down an insert over a layout hint.
    installHost([makeSlide("s1")], [], undefined, () => false);
    const size = await slideSize();
    expect(size).toEqual({ width: 960, height: 540, source: "assumed" });
  });

  it("caches the answer, and re-reads on request", async () => {
    installHost([makeSlide("s1")]);
    expect((await slideSize()).width).toBe(960);
    // The user changes slide size from PowerPoint's own Design tab.
    hostSlideSize.cx = 9144000;
    expect((await slideSize()).width, "went back to the host for a cached value").toBe(960);
    expect((await slideSize({ refresh: true })).width).toBe(720);
  });

  it("loads pageSetup before reading it", async () => {
    // The bug class this repo keeps finding: a proxy resolved and never loaded
    // answers on the fake and throws PropertyNotLoaded on the host. The fake is
    // honest now, so forgetting the load falls through to the next rung — and
    // the source is how the test can tell that happened.
    installHost([makeSlide("s1")]);
    expect((await slideSize()).source).toBe("pageSetup");
  });

  it("pulls the whole document through the Common API when PowerPointApi cannot answer", async () => {
    // The only rung a 1.4-1.7 host has. `getFileAsync` predates the
    // PowerPointApi requirement sets entirely, so it answers where nothing else
    // does — at the cost of copying the entire deck in 4MB slices to reach two
    // numbers in its first part. Hence last, and hence still worth having.
    installHost([makeSlide("s1")], [], undefined, () => false);
    const { default: JSZip } = await import("jszip");
    const zip = new JSZip();
    zip.file(
      "ppt/presentation.xml",
      `<p:presentation xmlns:p="x"><p:sldSz cx="9144000" cy="6858000"/></p:presentation>`,
    );
    const bytes = await zip.generateAsync({ type: "uint8array" });
    // Two slices, so the reassembly is actually exercised rather than assumed.
    const cut = Math.floor(bytes.length / 2);
    const slices = [bytes.slice(0, cut), bytes.slice(cut)];
    let closed = false;
    vi.stubGlobal("Office", {
      context: {
        host: "PowerPoint",
        requirements: { isSetSupported: () => false },
        document: {
          getFileAsync: (_t: unknown, _o: unknown, cb: (r: unknown) => void) =>
            cb({
              status: "succeeded",
              value: {
                size: bytes.length,
                sliceCount: slices.length,
                getSliceAsync: (i: number, scb: (r: unknown) => void) =>
                  scb({ status: "succeeded", value: { data: slices[i] } }),
                closeAsync: () => {
                  closed = true;
                },
              },
            }),
        },
      },
      FileType: { Compressed: "compressed" },
    });
    const size = await slideSize();
    expect(size).toEqual({ width: 720, height: 540, source: "documentFile" });
    // The handle MUST be released: a leaked one holds the host's copy of the
    // document alive and can block later getFileAsync calls outright.
    expect(closed, "leaked the document file handle").toBe(true);
  });
});

/**
 * The LIVE renderer against a config nobody validated.
 *
 * Every hostile-input sweep this project has ran against the SVG renderer and
 * the pptx one. Neither is the one that runs in a real PowerPoint, and the
 * Office renderer turned out to hold the third independent copy of the same
 * hole: `officeHex` did `color.trim()`, so `style.palette: [1, 2, 3]` threw
 * `color.trim is not a function` and took down a live insert — on the path a
 * user is actually standing on, for a config that came out of the JSON box or
 * a `POWERCHART_CONFIG` tag written in another deck.
 *
 * Three sinks, three separate holes, found one at a time because each sweep
 * only knew about the renderer it was written for.
 */
describe("a hostile config cannot take down a live insert", () => {
  const HOSTILE_CONFIGS: [string, (c: ChartConfig) => unknown][] = [
    ["numeric title", (c) => ({ ...c, title: 2024 })],
    ["numeric categories", (c) => ({ ...c, data: { ...c.data, categories: [1, 2, 3] } })],
    [
      "numeric series name",
      (c) => ({ ...c, data: { ...c.data, series: c.data.series.map((s) => ({ ...s, name: 7 })) } }),
    ],
    ["numeric palette", (c) => ({ ...c, style: { palette: [1, 2, 3] } })],
    // The fourth copy of the same hole, found the same way and in all three
    // sinks at once. `officeHex` hands anything non-alphabetic to `toHex6`, and
    // `toHex6` used to CRASH on a colour whose numbers are not numbers: the
    // regex that finds them matches a bare ".", `parseFloat(".")` is NaN, and
    // the hue sector table has no NaN entry. A malformed colour is exactly what
    // a hand-edited config or a template written in another deck arrives with.
    ["a palette colour whose numbers are not numbers", (c) => ({ ...c, style: { palette: ["hsl(., 50%, 50%)"] } })],
    ["labelContent a bare string", (c) => ({ ...c, decorations: { ...c.decorations, labelContent: "value" } })],
    ["numberFormat null", (c) => ({ ...c, numberFormat: null })],
    ["numeric valueAxisTitle", (c) => ({ ...c, valueAxisTitle: 5 })],
  ];

  for (const [name, mutate] of HOSTILE_CONFIGS) {
    it(
      name,
      async () => {
        const bad: string[] = [];
        for (const { kind } of CHART_KINDS.slice(0, 12)) {
          installHost([makeSlide("s1")]);
          try {
            await insertSceneIntoSlide(buildChart(mutate(sampleConfig(kind)) as ChartConfig), { tagData: "{}" });
          } catch (e) {
            bad.push(`${kind}: ${e instanceof Error ? e.message : String(e)}`);
          }
          vi.unstubAllGlobals();
        }
        expect(bad.slice(0, 4)).toEqual([]);
      },
      30_000,
    );
  }
});

describe("what a deck scan says about itself", () => {
  /**
   * It used to say nothing at all unless it came back short.
   *
   * The battery scans the deck about a dozen times a round — once per scenario,
   * through `probeCharts` — and every one of those was invisible. Round 10 is
   * what that cost: `stop a run part-way` reported 39.4 seconds against 2.6-3.2s
   * in the eight rounds before it, and the log had a 39-second HOLE where the
   * scan was. The stop itself was instant; the verification after it was not,
   * and nothing in the file said so.
   *
   * It is also the one operation the quadratic per-slide cost predicts should
   * grow worst — it reads every slide's shapes, on a deck the battery keeps
   * adding to — and it was the one operation never measured.
   */
  it("reports a CLEAN scan, not only a short one", async () => {
    installHost([makeSlide("s1"), makeSlide("s2")]);
    setTracing(true);
    try {
      const scan = await listChartsInDeck();
      expect(scan.unread, "this fixture was supposed to scan cleanly").toBe(0);
      const line = traceLog().entries.find((e) => e.message === "scanned the deck for charts");
      expect(line, "a clean scan left no trace, so a slow one is a hole in the log").toBeDefined();
      expect(line?.data).toMatchObject({ slides: 2, unread: 0, complete: true });
      expect(typeof line?.data?.ms, "no duration, which is the number the hole was hiding").toBe("number");
    } finally {
      setTracing(false);
    }
  });

  it("still says a short scan was short, in the same line", async () => {
    // The negative control. Moving the trace out of the failure branch must not
    // cost the failure branch its report — a scan that read nothing and one
    // that read everything have to stay distinguishable.
    installHost([makeSlide("s1"), makeSlide("s2")]);
    setTracing(true);
    faults.failSyncOn = 2;
    try {
      await listChartsInDeck();
      const line = traceLog().entries.find((e) => e.message === "scanned the deck for charts");
      expect(line?.data?.complete, "a scan that could not read a page reported itself complete").toBe(false);
      expect(line?.data?.unread, "the unread count went missing with the old message").toBeGreaterThan(0);
    } finally {
      faults.failSyncOn = 0;
      setTracing(false);
    }
  });
});

describe("the pictures a diagnostic round sends back", () => {
  /**
   * Three different reasons a slide comes back without a picture, and for a
   * long time one message for all of them.
   *
   * Every real round said `slides the host would not draw {asked: 22, drew: 12,
   * max: 12}` — which reads as a host refusing ten slides, and the host had
   * refused nothing at all: ten slides were over OUR cap and never asked about.
   * That is the same "never asked looks like answered no" mistake the contract
   * gate used to make, in the line a reader reaches for first when a deck comes
   * back short.
   */
  it("blames the cap for slides it never asked about, not the host", async () => {
    installHost([makeSlide("s1"), makeSlide("s2"), makeSlide("s3")]);
    setTracing(true);
    try {
      const shots = await slideShots(["s1", "s2", "s3"], { max: 1 });
      // Every id still comes back — a capped run that showed the first one only
      // would read as a deck of one.
      expect(shots.map((s) => s.slideId)).toEqual(["s1", "s2", "s3"]);
      expect(shots.filter((s) => s.png)).toHaveLength(1);
      const said = traceLog().entries.filter((e) => e.scope === "host");
      const refusal = said.find((e) => e.message === "slides the host would not draw");
      expect(
        refusal,
        `the host refused nothing and was blamed anyway: ${JSON.stringify(refusal?.data)}`,
      ).toBeUndefined();
      const capped = said.find((e) => e.message === "slides never asked about");
      expect(capped?.data).toMatchObject({ asked: 3, drew: 1, overCap: 2, stopped: 0, max: 1 });
    } finally {
      setTracing(false);
    }
  });

  it("still blames the host when the host is the one that would not draw", async () => {
    // The negative control. A message that never fires is not an improvement on
    // one that fires wrongly — a real refusal has to keep reaching the log.
    installHost([makeSlide("s1"), makeSlide("s2")]);
    setTracing(true);
    faults.emptySlideImage = true;
    try {
      const shots = await slideShots(["s1", "s2"], { max: 12 });
      expect(shots.filter((s) => s.png)).toHaveLength(0);
      const said = traceLog().entries.filter((e) => e.scope === "host");
      expect(said.find((e) => e.message === "slides the host would not draw")?.data).toMatchObject({
        asked: 2,
        drew: 0,
        refused: 2,
      });
      expect(
        said.find((e) => e.message === "slides never asked about"),
        "nothing was capped or stopped",
      ).toBeUndefined();
    } finally {
      faults.emptySlideImage = false;
      setTracing(false);
    }
  });

  it("counts a stop apart from the cap, so a short deck says which", async () => {
    // The third reason, and the one most likely to be misread: a run the user
    // stopped comes back with the same shape of gap as a capped one.
    installHost([makeSlide("s1"), makeSlide("s2")]);
    setTracing(true);
    requestStop();
    try {
      const shots = await slideShots(["s1", "s2"], { max: 12 });
      expect(shots.filter((s) => s.png)).toHaveLength(0);
      const capped = traceLog().entries.find((e) => e.message === "slides never asked about");
      expect(capped?.data).toMatchObject({ asked: 2, drew: 0, overCap: 0, stopped: 2 });
    } finally {
      resetStop();
      setTracing(false);
    }
  });
});

/**
 * `officeHex` handed the host any run of LETTERS verbatim, on the reasoning
 * that Office.js knows the CSS names. It knows the names; it does not know
 * every word. `style.palette: ["banana"]` — or a series colour of
 * `constructor`, which the pane's own saved-template store made reachable — is
 * a run of letters and is not a colour, and `setSolidColor` on a name the host
 * does not know is rejected inside the draw batch. One bad word did not degrade
 * one shape; it took the batch, and with it the chart.
 */
describe("only colour names the host actually knows reach the host", () => {
  const drawWith = async (color: string) => {
    const slide = makeSlide("s1");
    installHost([slide]);
    await insertSceneIntoSlide({
      width: 100,
      height: 60,
      nodes: [{ kind: "rect", x: 0, y: 0, w: 40, h: 20, fill: color, name: "seg-0-0" }],
    });
    return slide.created.map((s) => s.fillColor).filter(Boolean) as string[];
  };

  it("passes a real CSS name through, because Office knows it", async () => {
    expect(await drawWith("steelblue")).toContain("steelblue");
  });

  it("normalises a word that is not a colour instead of handing it over", async () => {
    for (const word of ["banana", "constructor", "notacolour"]) {
      const fills = await drawWith(word);
      expect(fills, `${word} was handed to the host verbatim`).not.toContain(word);
      for (const f of fills) expect(f, `${word} produced ${f}`).toMatch(/^#[0-9a-f]{6}$/i);
    }
  });
});

/**
 * A deck-wide rescale redraws each chart on its own slide, and the trace's
 * per-slide shape counter is the input to this project's only performance
 * claim — that drawing cost grows with what is already on the slide.
 *
 * It pooled them. `updateChartsInSlides` never named the slide it was aimed at,
 * so every chart in the `4feb5be` round keyed on the `(visible)` sentinel and
 * the counter climbed to 260 on a deck whose fullest slide held 24. The number
 * described nothing, and it looked perfectly healthy — a curve, rising, on
 * every line.
 *
 * `onSlideKey` is the only reason it was caught rather than plotted, which is
 * the argument for emitting a key beside any pooled total.
 */
describe("what the per-slide shape counter counts", () => {
  it("keys on the slide each chart is redrawn on, not on one sentinel", async () => {
    // Slide ids nothing else in this file uses. `shapesDrawnOnSlide` is a
    // per-RUN total by design — it answers "how much has this run already put
    // here" — so it is not reset between operations, and a shared `s1` would
    // carry every earlier test's draws into this one's first reading.
    const slides = [makeSlide("counter-a"), makeSlide("counter-b")];
    installHost(slides);
    const cfg = { ...sampleConfig("clustered"), ...DEFAULT_SIZE };
    const scene = buildChart(cfg);
    await insertSceneIntoSlide(scene, { tagData: JSON.stringify(cfg), slideId: "counter-a" });
    await insertSceneIntoSlide(scene, { tagData: JSON.stringify(cfg), slideId: "counter-b" });

    const found = (await listChartsInDeck()).charts;
    expect(found.length, "the two charts did not both land").toBe(2);
    setTracing(true);
    try {
      await updateChartsInSlides(
        found.map((c) => ({ scene, target: c.target, opts: { tagData: JSON.stringify(cfg) } })),
      );
      const batches = traceLog().entries.filter((e) => e.message === "batch issued");
      expect(batches.length, "the redraw issued no batches to check").toBeGreaterThan(1);
      const keys = new Set(batches.map((b) => String(b.data?.onSlideKey)));
      expect(
        keys.has("(visible)"),
        "the redraw pooled its charts under the unnamed-slide sentinel, so the count spans slides",
      ).toBe(false);
      expect(keys.size, `every chart keyed the same: ${[...keys].join(", ")}`).toBe(2);
      // Independent, not merely differently labelled. Both slides took exactly
      // one identical insert before this redraw, so their first batches must
      // start from the same count — if the totals were still pooled, whichever
      // slide was redrawn second would start where the first one finished.
      const firsts = [...keys].map((key) => Number(batches.find((b) => b.data?.onSlideKey === key)!.data?.onSlide));
      expect(
        firsts[0],
        `the two slides' counters started at ${firsts.join(" and ")} after identical work, so one is carrying the other's total`,
      ).toBe(firsts[1]);
    } finally {
      setTracing(false);
    }
  });
});
