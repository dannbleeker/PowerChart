// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from "vitest";
import {
  CHART_TAG,
  getSelectionBounds,
  insertAgendaSlides,
  insertDemoDeck,
  insertSceneIntoSlide,
  isPowerPointHost,
  listChartsInDeck,
  listChartsInSelection,
  loadChartFromSelection,
  onLateSync,
  updateChartInSlide,
  updateChartsInSlides,
} from "../src/render/powerpoint";
import { buildChart, DEFAULT_SIZE } from "../src/core/chart";
import { buildAgendaScene } from "../src/core/agenda";
import type { ChartConfig, MarkerSymbol } from "../src/core/types";

/**
 * Recording doubles for the PowerPoint JS proxy-object API: every shape the
 * renderer creates is captured with the geometry/format calls made on it, so
 * the whole scene→native-shapes mapping is testable without an Office host.
 */

let idSeq = 0;

function makeShape(type: string, geo: string | undefined, box: { left: number; top: number; width: number; height: number }) {
  const tagStore = new Map<string, string>();
  const shape = {
    type,
    geo,
    box,
    fillColor: null as string | null,
    fillCleared: false,
    text: undefined as string | undefined,
    name: undefined as string | undefined,
    rotation: undefined as number | undefined,
    deleted: false,
    id: `shape-${++idSeq}`,
    left: box.left,
    top: box.top,
    width: box.width,
    height: box.height,
    tagStore,
    tags: {
      add: (k: string, v: string) => void tagStore.set(k, v),
      getItemOrNullObject: (k: string) => ({
        isNullObject: !tagStore.has(k),
        value: tagStore.get(k) ?? "",
        load() {},
      }),
    },
    fill: {
      setSolidColor(c: string) {
        shape.fillColor = c;
      },
      clear() {
        shape.fillCleared = true;
      },
    },
    lineFormat: {} as Record<string, unknown>,
    textFrame: {
      textRange: { font: {} as Record<string, unknown>, paragraphFormat: {} as Record<string, unknown> },
    } as Record<string, unknown> & { textRange: { font: Record<string, unknown>; paragraphFormat: Record<string, unknown> } },
    grouped: undefined as unknown[] | undefined,
    delete() {
      shape.deleted = true;
    },
  };
  return shape;
}

type FakeShape = ReturnType<typeof makeShape>;

/** Layout ids passed to slides.add() since the last installHost(). */
const addedWithLayout: (string | undefined)[] = [];

/**
 * The queued-command failure a stalled getItemAt handle produces, surfaced at
 * the NEXT sync — because Office.js reports queued-command errors only at sync,
 * never at the call that queued them. `installHost`'s sync() throws it. See
 * `freshWindowedHandle` below for why a reused fresh-slide handle poisons.
 */
let pendingHostError: Error | null = null;

/**
 * The slide count at the start of the current `PowerPoint.run` — every slide at
 * an index >= this was `add()`ed during this context and so is "fresh". A fresh
 * slide's getItemAt handle is only good within the sync it was acquired in (see
 * `freshWindowedHandle`); a pre-existing slide's handle is durable.
 */
let contextBaseCount = 0;

function makeSlide(id: string) {
  const created: FakeShape[] = [];
  const slide = {
    id,
    created,
    isNullObject: false,
    load() {},
    shapes: {
      items: created,
      load() {},
      addGeometricShape(geo: string, box: FakeShape["box"]) {
        const s = makeShape("geometric", geo, box);
        created.push(s);
        return s;
      },
      addLine(_kind: string, box: FakeShape["box"]) {
        const s = makeShape("line", undefined, box);
        created.push(s);
        return s;
      },
      addTextBox(text: string, box: FakeShape["box"]) {
        const s = makeShape("text", undefined, box);
        s.text = text;
        created.push(s);
        return s;
      },
      addGroup(items: FakeShape[]) {
        const g = makeShape("group", undefined, { left: 0, top: 0, width: 0, height: 0 });
        g.grouped = items;
        created.push(g);
        return g;
      },
      getItemOrNullObject(id: string) {
        return created.find((s) => s.id === id && !s.deleted) ?? { isNullObject: true, delete() {} };
      },
    },
  };
  return slide;
}

type FakeSlide = ReturnType<typeof makeSlide>;

/**
 * The reference `slides.getItemAt(i)` hands back for a FRESHLY-ADDED slide.
 *
 * It draws fine within the sync it was acquired in, but reusing it to draw AFTER
 * a later sync is the object-path rewrite trap: Office.js has by then rewritten
 * its path to `getItem(<web-non-round-trippable id>)`, so the next shape throws
 * "InvalidParam passed to GetItem(id)" (code 5010) at the following sync. That is
 * exactly why the fix re-acquires a brand-new getItemAt proxy every batch — a
 * fresh handle is always inside its own window — and why HOLDING one across
 * batches fails. When poisoned, nothing lands: the queued shapes are detached,
 * not pushed to the real slide's `created`.
 *
 * The old fake could not express this at all: it returned the live slide from
 * getItemAt, so a held handle was as good as a fresh one.
 */
function freshWindowedHandle(real: FakeSlide) {
  const acquiredSync = trips.syncs;
  // Valid only until the next sync moves past the window it was acquired in.
  const ok = () => {
    if (trips.syncs <= acquiredSync) return true;
    pendingHostError = new Error(
      'InvalidParam passed to GetItem(id) | code=5010 | debugInfo={"errorLocation":"SlideCollection.getItem"}',
    );
    return false;
  };
  return {
    id: real.id,
    isNullObject: false,
    load() {},
    shapes: {
      items: real.shapes.items,
      load() {},
      addGeometricShape: (geo: string, box: FakeShape["box"]) =>
        ok() ? real.shapes.addGeometricShape(geo, box) : makeShape("geometric", geo, box),
      addLine: (kind: string, box: FakeShape["box"]) =>
        ok() ? real.shapes.addLine(kind, box) : makeShape("line", undefined, box),
      addTextBox: (text: string, box: FakeShape["box"]) => {
        if (ok()) return real.shapes.addTextBox(text, box);
        const s = makeShape("text", undefined, box);
        s.text = text;
        return s;
      },
      addGroup: (items: FakeShape[]) =>
        ok() ? real.shapes.addGroup(items) : makeShape("group", undefined, { left: 0, top: 0, width: 0, height: 0 }),
      getItemOrNullObject: (id: string) => real.shapes.getItemOrNullObject(id),
    },
  };
}

/**
 * Office round-trips since the last installHost(). Every context.sync() is a
 * trip to PowerPoint and dominates insert latency, so the count is a behaviour
 * worth asserting — see "round-trips do not scale with the chart count".
 */
const trips = { syncs: 0, contexts: 0 };

/**
 * Make the Nth context.sync() of the next run throw. Office.js queues commands
 * and only reports their errors at sync — so this, not a throwing addGroup, is
 * how a host actually refuses something. 0 = never.
 */
let failSyncOn = 0;

/** Install a fake PowerPoint global whose run() drives the mocked context.
 * `supported(version)` models the host's requirement-set support (default: all)
 * — pass a predicate to simulate e.g. PowerPoint on the web lacking grouping. */
function installHost(
  slides: FakeSlide[],
  selectedShapes: FakeShape[] = [],
  selectedSlide = slides[0],
  supported: (version: string) => boolean = () => true,
) {
  // The slide count as of the last COMMITTED sync. getCount() reports THIS, not
  // the live array — so an add() queued in the current batch is invisible to a
  // getCount() in the SAME batch, exactly as PowerPoint web behaves. A getCount
  // result resolves at the next sync to the count from before that sync's adds.
  let committedCount = slides.length;
  const pendingCounts: { value: number }[] = [];
  const context = {
    presentation: {
      slides: {
        items: slides,
        load() {},
        getItem: (id: string) => slides.find((s) => s.id === id)!,
        // Real Office.js hands back a null OBJECT for an unknown id — it does
        // not throw and does not return undefined. A fake that returns
        // undefined would make `slide.isNullObject` a TypeError instead of the
        // false it should be, and hide the very case this models. By id, the
        // reference is always durable.
        getItemOrNullObject: (id: string) => slides.find((s) => s.id === id) ?? { isNullObject: true, load() {} },
        // A pre-existing slide's handle is durable; a freshly-added one's is only
        // good within the sync it was acquired in (see freshWindowedHandle), so
        // HOLDING one across the render's batches is the bug the fix avoids by
        // re-acquiring fresh each batch.
        getItemAt: (i: number) => {
          const s = slides[i];
          if (!s) return s;
          return i >= contextBaseCount ? freshWindowedHandle(s) : s;
        },
        // Resolves at the NEXT sync to the committed count (from before that
        // sync's adds), never to slides.length now — see committedCount.
        getCount: () => {
          const result = { value: committedCount };
          pendingCounts.push(result);
          return result;
        },
        add: (options?: { layoutId?: string }) => {
          addedWithLayout.push(options?.layoutId);
          slides.push(makeSlide(`slide-${slides.length + 1}`));
        },
      },
      // A real deck's master carries several layouts; only one is blank, and
      // its NAME is localised — which is why the renderer matches on type.
      slideMasters: {
        items: [
          {
            id: "master-1",
            layouts: {
              items: [
                { id: "layout-title", name: "Titeldias", type: "titleSlide" },
                { id: "layout-blank", name: "Tom", type: "blank" },
                { id: "layout-content", name: "Titel og indhold", type: "object" },
              ],
            },
          },
        ],
        load() {},
      },
      getSelectedSlides: () => ({ getItemAt: () => selectedSlide }),
      getSelectedShapes: () => ({ items: selectedShapes, load() {} }),
    },
    sync: async () => {
      trips.syncs++;
      if (trips.syncs === failSyncOn) throw new Error("host refused a queued command");
      // A queued-command failure (e.g. drawing on a poisoned getItemAt handle)
      // surfaces here, at the sync, exactly as Office.js reports it.
      if (pendingHostError) {
        const err = pendingHostError;
        pendingHostError = null;
        throw err;
      }
      // Each getCount from this batch resolves to the PRE-batch committed count,
      // then this batch's adds become visible to the NEXT sync's getCount.
      for (const r of pendingCounts) r.value = committedCount;
      pendingCounts.length = 0;
      committedCount = slides.length;
    },
  };
  trips.syncs = 0;
  trips.contexts = 0;
  pendingHostError = null;
  addedWithLayout.length = 0;
  vi.stubGlobal("PowerPoint", {
    run: async <T>(cb: (ctx: typeof context) => Promise<T>) => {
      trips.contexts++;
      // Slides present at the start of this context are "existing"; anything
      // add()ed past here is fresh, and its getItemAt handle is window-limited.
      contextBaseCount = slides.length;
      return cb(context);
    },
    // Real Office.js exposes all 177 presets. A plain object listing only the
    // ones in use today hands back `undefined` for any other name, and the
    // renderer then records a shape with no geometry while this suite still
    // passes green — a test that asserts nothing about the shape it drew.
    // The Proxy makes that impossible: reaching for a preset this stub has not
    // been told about throws instead of returning undefined.
    GeometricShapeType: new Proxy(
      {
        rectangle: "rectangle",
        ellipse: "ellipse",
        triangle: "triangle",
        chevron: "chevron",
        homePlate: "homePlate",
        lineInverse: "lineInverse",
        diamond: "diamond",
        plus: "plus",
      } as Record<string, string>,
      {
        get(target, prop: string) {
          if (!(prop in target)) throw new Error(`office stub: unknown GeometricShapeType "${String(prop)}"`);
          return target[prop];
        },
      },
    ),
    SlideLayoutType: { blank: "blank", titleSlide: "titleSlide", object: "object" },
    ConnectorType: { straight: "straight" },
    ShapeLineDashStyle: { dash: "dash" },
    ShapeAutoSize: { autoSizeNone: "none" },
    TextVerticalAlignment: { top: "top", middle: "middle", bottom: "bottom" },
    ParagraphHorizontalAlignment: { left: "left", center: "center", right: "right" },
  });
  vi.stubGlobal("Office", {
    context: { host: "PowerPoint", requirements: { isSetSupported: (_set: string, version: string) => supported(version) } },
  });
  return context;
}

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
      { kind: "ellipse", cx: 50, cy: 50, rx: 20, ry: 10, fill: "#ff0000", stroke: "#000000", strokeWidth: 2, name: "dot" },
      { kind: "ellipse", cx: 10, cy: 10, rx: 5, ry: 5, fill: "#00ff00" },
    ]);
    const [a, b] = slide.created.filter((s) => s.geo === "ellipse");
    // center − radius, plus the default 60/90pt insert offset
    expect(a.box).toEqual({ left: 90, top: 130, width: 40, height: 20 });
    expect(a.lineFormat).toMatchObject({ color: "#000000", weight: 2 });
    expect(b.lineFormat.visible).toBe(false);
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
      { kind: "line", x1: 10, y1: 50, x2: 200, y2: 50, stroke: "#333333", strokeWidth: 1, dash: [3, 2], name: "connector" },
    ]);
    const line = slide.created.find((s) => s.type === "line")!;
    // Horizontal line: width spans, height is clamped up from 0 so the web host
    // can't blow a zero-thickness box into a giant diagonal.
    expect(line.box.width).toBeGreaterThan(180);
    expect(line.box.height).toBeGreaterThanOrEqual(0.5);
    expect(line.lineFormat.dashStyle).toBe("dash");
  });

  it("draws diagonal lines as thin rotated rectangles (direction-correct on every host)", async () => {
    // Up-right and down-right diagonals a bounding box alone can't distinguish.
    const down = await insert([{ kind: "line", x1: 0, y1: 0, x2: 100, y2: 100, stroke: "#a00000", strokeWidth: 2, name: "d" }]);
    const dr = down.created.find((s) => s.geo === "rectangle")!;
    expect(dr).toBeTruthy();
    expect(dr.fillColor).toBe("#a00000");
    expect(dr.rotation).toBeCloseTo(45, 0); // down-right
    expect(down.created.some((s) => s.type === "line")).toBe(false);

    const up = await insert([{ kind: "line", x1: 0, y1: 100, x2: 100, y2: 0, stroke: "#00a000", strokeWidth: 2, name: "u" }]);
    const ur = up.created.find((s) => s.geo === "rectangle")!;
    expect(ur.rotation).toBeCloseTo(-45, 0); // up-right — the case a box would mirror
  });

  it("draws dashed diagonals as real line shapes, picking the geometry per direction", async () => {
    // A rotated rectangle carries its colour in its fill, which can't be
    // dashed — scatter trend lines and forecast segments came out solid.
    const down = await insert([
      { kind: "line", x1: 0, y1: 0, x2: 100, y2: 60, stroke: "#a00000", strokeWidth: 1.25, dash: [4, 2], name: "trend" },
    ]);
    const dl = down.created.find((s) => s.name === "trend")!;
    expect(dl.type).toBe("line"); // not a filled rectangle
    expect(dl.lineFormat.dashStyle).toBe("dash");
    expect(dl.box).toMatchObject({ width: 100, height: 60 });

    // Up-right: addLine only ever draws the box's top-left→bottom-right
    // diagonal, so this direction needs the lineInverse geometry.
    const up = await insert([
      { kind: "line", x1: 0, y1: 60, x2: 100, y2: 0, stroke: "#a00000", strokeWidth: 1.25, dash: [4, 2], name: "trend" },
    ]);
    const ul = up.created.find((s) => s.name === "trend")!;
    expect(ul.geo).toBe("lineInverse");
    expect(ul.lineFormat.dashStyle).toBe("dash");
  });

  it("draws polygon edges direction-correct, with no zero-thickness boxes", async () => {
    // A violin body: an up-right edge, a horizontal edge and a down-right edge.
    const slide = await insert([
      {
        kind: "polygon",
        points: [{ x: 0, y: 40 }, { x: 50, y: 0 }, { x: 100, y: 40 }, { x: 100, y: 40 }],
        fill: "#eeeeee", stroke: "#3366cc", strokeWidth: 1, name: "violin-0",
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
      { kind: "wedge", cx: 50, cy: 50, r: 30, innerR: 15, startAngle: 0, endAngle: 90, fill: "#333333", stroke: "#ffffff", strokeWidth: 1, name: "ring" },
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
    await insertSceneIntoSlide({ width: 10, height: 10, nodes: [{ kind: "rect", x: 0, y: 0, w: 5, h: 5, fill: "#111111" }] } as never, {});
    expect(slide.created).toHaveLength(1);
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
  it("finds tagged charts across all slides", async () => {
    const s1 = makeSlide("s1");
    const s2 = makeSlide("s2");
    installHost([s1, s2]);
    await insertSceneIntoSlide(buildChart(config), { tagData: '{"a":1}' });
    const g = s2.shapes.addGroup([]);
    g.tagStore.set(CHART_TAG, '{"b":2}');
    s2.shapes.addGeometricShape("rectangle", { left: 0, top: 0, width: 1, height: 1 });

    const found = await listChartsInDeck();
    expect(found).toHaveLength(2);
    expect(found.map((f) => f.target.slideId).sort()).toEqual(["s1", "s2"]);
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
      { scene: buildChart({ ...DEFAULT_SIZE, kind: "pie" as const, data: { categories: ["A", "B"], series: [{ name: "S", values: [3, 1] }] } }), tagData: '{"kind":"pie"}' },
      { scene: buildChart({ ...DEFAULT_SIZE, kind: "clustered" as const, data: { categories: ["A"], series: [{ name: "S", values: [5] }] } }), tagData: '{"kind":"clustered"}' },
      { scene: { width: 100, height: 40, nodes: [{ kind: "rect" as const, x: 0, y: 0, w: 10, h: 10, fill: "#111111" }] } }, // untagged element
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
      const s = slide.created.find(
        (c) => c.geo === "diamond" && Math.abs(c.box.left - (100 + n.cx - n.size)) < 1e-6,
      );
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
      return { scene: buildChart(cfgFor(i)), target: { slideId: slide.id, shapeId: s.id, left: 10, top: 20 }, opts: { tagData: `{"i":${i}}` } };
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
    // 1 resolve slides, 1 resolve old shapes, 1 delete, then each chart's
    // batches, then GROUP.
    failSyncOn = 3 + batches * 3 /* 3 charts */ + 1;
    try {
      const items = targetsOn(slide, 3);
      await expect(updateChartsInSlides(items)).resolves.toBeUndefined();
      // Each chart's OWN config, back on each chart's OWN first shape.
      const tagged = slide.created.filter((s) => s.tagStore.has(CHART_TAG));
      expect(tagged.map((s) => s.tagStore.get(CHART_TAG))).toEqual(['{"i":0}', '{"i":1}', '{"i":2}']);
      expect(tagged.every((s) => s.type !== "group")).toBe(true);
    } finally {
      failSyncOn = 0;
    }
  });

  it("renders one slide per context and reports progress per slide", async () => {
    // One PowerPoint.run per slide isolates a chart the host can't finish and
    // keeps each context light (one chart's shapes, not a chunk's four). Progress
    // is per slide, so a slow host shows slides landing instead of freezing.
    for (const n of [2, 12, 35] as const) {
      installHost([makeSlide("s1")]);
      const seen: string[] = [];
      const failed = await insertDemoDeck(
        Array.from({ length: n }, (_, i) => ({ scene: buildChart(cfgFor(i)), tagData: `{"i":${i}}` })),
        (done, total) => seen.push(`${done}/${total}`),
      );
      expect(failed, `${n} slides`).toEqual([]);
      // One context per SLIDE now.
      expect(trips.contexts, `${n} slides`).toBe(n);
      expect(seen, `${n} slides`).toHaveLength(n);
      expect(seen.at(-1)).toBe(`${n}/${n}`);
      // Monotonic, never over-counting.
      expect(seen.map((x) => Number(x.split("/")[0]))).toEqual([...seen.map((x) => Number(x.split("/")[0]))].sort((a, b) => a - b));
    }
  });

  it("appends every demo slide, each tagged with its own config", async () => {
    const deck: FakeSlide[] = [makeSlide("s1")];
    installHost(deck);
    const n = 35;
    const failed = await insertDemoDeck(Array.from({ length: n }, (_, i) => ({ scene: buildChart(cfgFor(i)), tagData: `{"i":${i}}` })));
    expect(failed).toEqual([]);
    // The fake appends a slide per add(); the original + n new ones.
    expect(deck.length).toBe(1 + n);
    const tags = deck.slice(1).map((s) => s.created.map((c) => c.tagStore.get(CHART_TAG)).find(Boolean));
    expect(tags).toEqual(Array.from({ length: n }, (_, i) => `{"i":${i}}`));
  });

  it("reports the slide a host refuses and finishes the rest", async () => {
    // A chart the host cannot draw (on the real host, a ~200-shape area chart that
    // times out) must not abort the whole deck: it is caught, its index returned,
    // and the remaining slides still render. One sync failure = one lost slide.
    const deck: FakeSlide[] = [makeSlide("s1")];
    installHost(deck);
    const n = 6;
    // Fail a single sync partway in (past the first slide's layout+add+render).
    failSyncOn = 9;
    try {
      const failed = await insertDemoDeck(Array.from({ length: n }, (_, i) => ({ scene: buildChart(cfgFor(i)), tagData: `{"i":${i}}` })));
      expect(failed.length).toBeGreaterThanOrEqual(1);
      expect(failed.length).toBeLessThan(n); // it kept going — not a whole-deck abort
      // At least every non-failed slide drew its chart (a failed slide may retain
      // the partial shapes it had queued before the refused sync).
      const drawn = deck.slice(1).filter((s) => s.created.length > 0).length;
      expect(drawn).toBeGreaterThanOrEqual(n - failed.length);
    } finally {
      failSyncOn = 0;
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
      nodes: Array.from({ length: NODES }, (_, k) => ({ kind: "rect" as const, x: k, y: 0, w: 4, h: 4, fill: "#111111" })),
    };
    const deck: FakeSlide[] = [makeSlide("s1")];
    installHost(deck);
    const n = 6;
    await expect(insertDemoDeck(Array.from({ length: n }, () => ({ scene: bigScene })))).resolves.toEqual([]);
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
    const [done, total] = commits.at(-1)!.match(/(\d+) of (\d+)/)!.slice(1);
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
          cb({ presentation: { slides: { getItemAt: () => slide }, getSelectedSlides: () => ({ getItemAt: () => slide }) }, sync: ctxSync }),
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
      rejectSync({ message: "An internal error has occurred.", code: "GeneralException", debugInfo: { errorLocation: "Shape.name" } });
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
    await insertAgendaSlides([buildAgendaScene(["Intro", "Body"], { highlight: 0 }), buildAgendaScene(["Intro", "Body"], { highlight: 1 })]);
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
    const count = () => slides.reduce((a, s) => a + s.created.length, 0);
    ctx.sync = async () => {
      trips.syncs++;
      worst.n = Math.max(worst.n, count() - last);
      last = count();
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
    // A per-path test would have missed it; this asserts the invariant.
    const scene = () => buildChart(config);
    expect(scene().nodes.length).toBeGreaterThan(10); // must span batches

    const s1 = makeSlide("s1");
    expect(await maxPerSync(() => insertSceneIntoSlide(scene(), { tagData: "{}" }), [s1]), "insert").toBeLessThanOrEqual(10);

    const s2 = makeSlide("s2");
    const old = s2.shapes.addGeometricShape("rectangle", { left: 0, top: 0, width: 1, height: 1 });
    expect(
      await maxPerSync(() => updateChartInSlide(scene(), { slideId: "s2", shapeId: old.id, left: 0, top: 0 }, {}), [s2]),
      "update",
    ).toBeLessThanOrEqual(11); // +1: the pre-existing shape this test planted

    const s3 = makeSlide("s3");
    expect(await maxPerSync(() => insertAgendaSlides([scene(), scene()]), [s3]), "agenda").toBeLessThanOrEqual(10);

    const s4 = makeSlide("s4");
    expect(
      await maxPerSync(() => insertDemoDeck([{ scene: scene() }, { scene: scene() }, { scene: scene() }, { scene: scene() }, { scene: scene() }]), [s4]),
      "demo deck",
    ).toBeLessThanOrEqual(10);
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
        { scene: buildChart(config), target: { slideId: "s-deleted", shapeId: "gone", left: 0, top: 0 }, opts: { tagData: "{}" } },
        { scene: buildChart(config), target: { slideId: "s-live", shapeId: s.id, left: 10, top: 20 }, opts: { tagData: '{"ok":1}' } },
      ]),
    ).resolves.toBeUndefined();
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
      updateChartsInSlides([{ scene: buildChart(config), target: { slideId: "nope", shapeId: "nope", left: 0, top: 0 }, opts: {} }]),
    ).resolves.toBeUndefined();
    expect(slide.created.length).toBe(before);
  });
});
