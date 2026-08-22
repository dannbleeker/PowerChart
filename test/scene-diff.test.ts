import { describe, expect, it } from "vitest";
import { buildChart, DEFAULT_SIZE } from "../src/core/chart";
import { sampleConfig } from "../src/core/samples";
import {
  isUpdatableKind,
  planSceneUpdate,
  sceneFingerprint,
  UPDATE_SHARE_LIMIT,
  worthUpdating,
} from "../src/core/scene-diff";
import type { Scene, SceneNode } from "../src/core/scene";
import type { ChartConfig } from "../src/core/types";

const rect = (name: string, x = 0, fill = "#123456"): SceneNode => ({
  kind: "rect",
  x,
  y: 0,
  w: 10,
  h: 10,
  fill,
  name,
});
const text = (name: string, s = "hello"): SceneNode => ({
  kind: "text",
  x: 0,
  y: 0,
  w: 40,
  h: 10,
  text: s,
  fontSize: 10,
  color: "#000000",
  align: "center",
  valign: "middle",
  name,
});
const scene = (nodes: SceneNode[]): Scene => ({ width: 480, height: 270, nodes });

describe("what a redraw would have been for", () => {
  it("finds the one node a retitle changes", () => {
    const cfg = { ...sampleConfig("clustered"), ...DEFAULT_SIZE } as ChartConfig;
    const plan = planSceneUpdate(buildChart(cfg), buildChart({ ...cfg, title: "Something else" }));
    expect(plan, "a retitle was ruled unsafe to apply").not.toBeNull();
    expect(plan!.changed, "a retitle touched more than the title").toHaveLength(1);
  });

  it("finds the two a single data point changes", () => {
    const cfg = { ...sampleConfig("clustered"), ...DEFAULT_SIZE } as ChartConfig;
    const edited = structuredClone(cfg);
    edited.data!.series[0].values[0] = 7;
    const plan = planSceneUpdate(buildChart(cfg), buildChart(edited));
    expect(plan).not.toBeNull();
    // The bar and its value label. Anything more means the layout moved
    // something this diff is quietly rewriting on every edit.
    expect(plan!.changed).toHaveLength(2);
  });

  it("reports NO changes when the config round-trips without moving anything", () => {
    // A real answer, and not the same as "cannot be done": the caller should
    // write the new config tag and touch no shapes at all.
    const cfg = { ...sampleConfig("clustered"), ...DEFAULT_SIZE } as ChartConfig;
    const plan = planSceneUpdate(buildChart(cfg), buildChart(structuredClone(cfg)));
    expect(plan).toEqual({ changed: [] });
  });
});

describe("when an in-place update must not be attempted", () => {
  it("refuses a chart that changed size", () => {
    // Re-laid-out: every node moves, and the shapes' own box no longer matches
    // the scene's.
    const a = scene([rect("seg-0-0")]);
    expect(planSceneUpdate(a, { ...a, width: 600 })).toBeNull();
    expect(planSceneUpdate(a, { ...a, height: 400 })).toBeNull();
  });

  it("refuses a node count that moved in either direction", () => {
    expect(planSceneUpdate(scene([rect("a")]), scene([rect("a"), rect("b")]))).toBeNull();
    expect(planSceneUpdate(scene([rect("a"), rect("b")]), scene([rect("a")]))).toBeNull();
  });

  it("refuses a reorder even when the same nodes are present", () => {
    // The shape is found POSITIONALLY — anchor is node 0, the parts tag lists
    // the rest in drawing order — so an index meaning a different node in the
    // two scenes writes a bar's geometry onto a label.
    const a = scene([rect("seg-0-0"), text("label-0-0")]);
    const b = scene([text("label-0-0"), rect("seg-0-0")]);
    expect(planSceneUpdate(a, b)).toBeNull();
  });

  it("refuses a renamed node at the same index", () => {
    expect(planSceneUpdate(scene([rect("seg-0-0")]), scene([rect("seg-9-9")]))).toBeNull();
  });

  it("refuses when a CHANGED node has no writable geometry", () => {
    // A wedge is a fan of rotated triangles and an arrowhead is a triangle
    // rotated about a computed box: changing the numbers that produced them
    // means producing them again, and Office.js has no freeform path to edit.
    const wedge = (r: number): SceneNode => ({
      kind: "wedge",
      cx: 50,
      cy: 50,
      r,
      innerR: 0,
      startAngle: 0,
      endAngle: 1,
      fill: "#abcdef",
      name: "slice-0",
    });
    expect(planSceneUpdate(scene([wedge(20)]), scene([wedge(30)]))).toBeNull();
  });

  it("allows an UNCHANGED node of a kind it could never update", () => {
    // It stays exactly where it is, so nothing has to be written to it. Ruling
    // the chart out on its mere presence would disable the fast path for every
    // pie, radar and gauge in the deck.
    const wedge: SceneNode = {
      kind: "wedge",
      cx: 50,
      cy: 50,
      r: 20,
      innerR: 0,
      startAngle: 0,
      endAngle: 1,
      fill: "#abcdef",
      name: "slice-0",
    };
    const plan = planSceneUpdate(scene([wedge, text("title")]), scene([wedge, text("title", "new")]));
    expect(plan).toEqual({ changed: [1] });
  });

  it("refuses a changed node that carries no name", () => {
    const unnamed = (fill: string): SceneNode => ({ kind: "rect", x: 0, y: 0, w: 5, h: 5, fill });
    expect(planSceneUpdate(scene([unnamed("#111111")]), scene([unnamed("#222222")]))).toBeNull();
  });

  it("knows which kinds it can write to", () => {
    expect(isUpdatableKind("rect")).toBe(true);
    expect(isUpdatableKind("text")).toBe(true);
    for (const k of ["wedge", "arrowhead", "polygon", "chevron", "symbol", "ellipse", "line"] as const)
      expect(isUpdatableKind(k), `${k} claims to be writable in place`).toBe(false);
  });
});

describe("the fingerprint that decides whether the old scene is the drawn one", () => {
  it("changes when any node does", () => {
    const a = scene([rect("seg-0-0"), text("title")]);
    expect(sceneFingerprint(a)).toBe(sceneFingerprint(scene([rect("seg-0-0"), text("title")])));
    expect(sceneFingerprint(a)).not.toBe(sceneFingerprint(scene([rect("seg-0-0", 1), text("title")])));
    expect(sceneFingerprint(a)).not.toBe(sceneFingerprint(scene([rect("seg-0-0"), text("title", "x")])));
  });

  it("changes when the frame does, even with identical nodes", () => {
    const a = scene([rect("seg-0-0")]);
    expect(sceneFingerprint(a)).not.toBe(sceneFingerprint({ ...a, width: 600 }));
  });

  it("is a short string on every sample chart", () => {
    // Written to a shape tag on every chart, so it must not be a JSON blob.
    for (const kind of ["clustered", "stacked", "pie", "waterfall"] as const) {
      const fp = sceneFingerprint(buildChart({ ...sampleConfig(kind), ...DEFAULT_SIZE }));
      expect(fp).toMatch(/^[0-9a-z]{1,7}$/);
    }
  });
});

describe("whether a plan is worth taking", () => {
  it("takes a small change and declines a wholesale one", () => {
    // The old comment here read "a redraw is one delete plus one add per node,
    // so an update touching more than half is already doing more host calls per
    // shape saved". That counted an in-place write as ONE call per node; it is
    // about twenty, and the real cost is shapes per BATCH — which is bounded
    // directly now. See `UPDATE_SHARE_LIMIT`.
    expect(worthUpdating({ changed: [0] }, 24)).toBe(true);
    expect(worthUpdating({ changed: [] }, 24)).toBe(true);
    expect(worthUpdating({ changed: Array.from({ length: 22 }, (_, i) => i) }, 24)).toBe(false);
    expect(worthUpdating({ changed: Array.from({ length: 12 }, (_, i) => i) }, 24)).toBe(true);
    expect(UPDATE_SHARE_LIMIT).toBeGreaterThan(0);
  });

  it("admits both edits the battery makes, now that the batch is bounded", () => {
    // BOTH EDGES, because the whole point of this number is which side of it
    // two real edits fall on. `same scale across the deck` makes exactly these:
    // 18-of-24 six times a round and 9-of-16 twice.
    //
    // The history is the argument. 0.8 admitted both with an UNBOUNDED batch
    // and PowerPoint died on all seven attempts of round 150, at
    // 255/284/282/278/257/282s. 0.6 admitted the smaller one only and was safe.
    // `IN_PLACE_WRITES_PER_SYNC` now caps a write at six shapes per sync, and
    // round 152 measured that cap as free — 0.07% on a 164-second scenario —
    // so 0.8 is back to ask whether the crash was ever about the share.
    const upTo = (n: number) => ({ changed: Array.from({ length: n }, (_, i) => i) });
    expect(worthUpdating(upTo(9), 16), "9-of-16 declined").toBe(true);
    expect(worthUpdating(upTo(18), 24), "18-of-24 declined — the round cannot test the bounded batch").toBe(true);
  });

  it("still declines a near-total rewrite", () => {
    // A cap survives, and it is no longer about arithmetic. Where nearly every
    // node changes there is little left to save, and a redraw remains the
    // better-tested way to get a clean chart.
    const upTo = (n: number) => ({ changed: Array.from({ length: n }, (_, i) => i) });
    expect(worthUpdating(upTo(22), 24), "a near-total rewrite took the fast path").toBe(false);
    expect(worthUpdating(upTo(24), 24), "a total rewrite took the fast path").toBe(false);
  });

  it("declines rather than dividing by zero on an empty chart", () => {
    expect(worthUpdating({ changed: [] }, 0)).toBe(false);
  });

  it("puts the measured edits on the right side of the line", () => {
    const cfg = { ...sampleConfig("clustered"), ...DEFAULT_SIZE } as ChartConfig;
    const base = buildChart(cfg);
    const take = (next: Scene) => {
      const p = planSceneUpdate(base, next);
      return p ? worthUpdating(p, next.nodes.length) : false;
    };
    expect(take(buildChart({ ...cfg, title: "Another" })), "a retitle was not worth updating").toBe(true);
    // A rescale moves 18 of 24 — over the line, and the redraw is the honest
    // answer there. Recorded so a future layout change that alters the share
    // shows up as a decision rather than as a silent switch of path.
    //
    // THAT DECISION WAS MADE ONCE AND REVERSED BY THE HOST. `6359d83` raised
    // the limit to 0.8 so this very edit would take the fast path — `max: 105`
    // is what `same scale across the deck` does, eight times a round. Round 150
    // then crashed PowerPoint on all seven attempts, at 255/284/282/278/257/282
    // seconds, and that scenario starts at 280s. See `UPDATE_SHARE_LIMIT`.
    //
    // AND REVERSED AGAIN, deliberately, once the real cost was bounded. Round
    // 150 crashed with an UNBOUNDED batch: eight charts writing eighteen shapes
    // apiece into one sync each. `IN_PLACE_WRITES_PER_SYNC` caps that at six,
    // and round 152 measured the cap as free (0.07% on a 164-second scenario).
    //
    // So this edit takes the fast path again, and the round says whether the
    // crash was per-SYNC size (fixed) or per-CONTEXT accumulation (not).
    expect(take(buildChart({ ...cfg, scale: { max: 105 } })), "a full rescale was still declined").toBe(true);
  });
});
