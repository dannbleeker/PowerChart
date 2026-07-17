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
  updateChartInSlide,
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

function makeSlide(id: string) {
  const created: FakeShape[] = [];
  const slide = {
    id,
    created,
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

/** Install a fake PowerPoint global whose run() drives the mocked context.
 * `supported(version)` models the host's requirement-set support (default: all)
 * — pass a predicate to simulate e.g. PowerPoint on the web lacking grouping. */
function installHost(
  slides: FakeSlide[],
  selectedShapes: FakeShape[] = [],
  selectedSlide = slides[0],
  supported: (version: string) => boolean = () => true,
) {
  const context = {
    presentation: {
      slides: {
        items: slides,
        load() {},
        getItem: (id: string) => slides.find((s) => s.id === id)!,
        getItemAt: (i: number) => slides[i],
        getCount: () => ({ value: slides.length }),
        add: () => void slides.push(makeSlide(`slide-${slides.length + 1}`)),
      },
      getSelectedSlides: () => ({ getItemAt: () => selectedSlide }),
      getSelectedShapes: () => ({ items: selectedShapes, load() {} }),
    },
    sync: async () => {},
  };
  vi.stubGlobal("PowerPoint", {
    run: async <T>(cb: (ctx: typeof context) => Promise<T>) => cb(context),
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
        star5: "star5",
      } as Record<string, string>,
      {
        get(target, prop: string) {
          if (!(prop in target)) throw new Error(`office stub: unknown GeometricShapeType "${String(prop)}"`);
          return target[prop];
        },
      },
    ),
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
            throw new Error("rotation requires PowerPointApi 1.9");
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
    await insertSceneIntoSlide(markerScene(["diamond", "plus", "star5"]), { left: 0, top: 0 });

    // Filled preset geometry is the whole reason a symbol is not a polygon:
    // PowerPoint can only outline a freeform, so a polygon marker would be
    // hollow here while the SVG preview showed it solid.
    for (const preset of ["diamond", "plus", "star5"]) {
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
    await insertSceneIntoSlide(markerScene(["diamond", "triangle", "star5"]), { left: 0, top: 0 });
    const presets = slide.created.filter((s) => ["diamond", "triangle", "star5"].includes(s.geo!));
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
