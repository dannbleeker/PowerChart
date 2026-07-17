/**
 * Office.js renderer: draws a scene as native, individually editable
 * PowerPoint shapes on the current slide, then groups them.
 *
 * This is the same output strategy as think-cell/UpSlide-style tools —
 * charts stay fully editable in PowerPoint (every bar and label is a shape),
 * rather than being pasted as pictures or opaque OLE charts.
 *
 * Requires PowerPointApi 1.4+ (ShapeCollection.addGeometricShape / addLine /
 * addTextBox) — marker symbols are preset geometry, so they need only 1.4 too.
 * Grouping (1.8+) and shape rotation (1.10+) degrade gracefully on older hosts.
 */
import { polar, arrowheadBox, wedgeFanSteps, SYMBOL_PRESET } from "../core/geometry";
import type { Scene, SceneNode, TextNode, WedgeNode } from "../core/scene";

/* global PowerPoint, Office */

export interface InsertOptions {
  /** Top-left of the chart frame on the slide, in points. */
  left?: number;
  top?: number;
  /** Group the shapes after insertion (default true). */
  group?: boolean;
  fontFamily?: string;
  /**
   * Serialized chart model stored as a tag on the inserted group
   * (PowerPointApi 1.3+), so a future version can re-open and re-edit
   * the chart — the think-cell "live chart" pattern.
   */
  tagData?: string;
}

/** Tag key under which the chart's serialized config is persisted. */
export const CHART_TAG = "POWERCHART_CONFIG";

const DEFAULT_FONT = "Segoe UI";

/** Where an existing PowerChart lives on the deck, for in-place update. */
export interface EditTarget {
  slideId: string;
  shapeId: string;
  left: number;
  top: number;
}

export async function insertSceneIntoSlide(scene: Scene, opts: InsertOptions = {}): Promise<void> {
  await PowerPoint.run(async (context) => {
    const slide = getTargetSlide(context);
    const created = renderShapes(slide, scene, opts);
    // Commit the shapes first — so grouping/tagging (which some hosts, notably
    // PowerPoint on the web, don't support) can't roll back the whole insert.
    await context.sync();
    await groupAndTag(context, slide, created, opts);
  });
}

/** Replace an existing PowerChart group with a re-rendered scene, in place. */
export async function updateChartInSlide(scene: Scene, target: EditTarget, opts: InsertOptions = {}): Promise<void> {
  await PowerPoint.run(async (context) => {
    const slide = context.presentation.slides.getItem(target.slideId);
    const old = slide.shapes.getItemOrNullObject(target.shapeId);
    await context.sync();
    if (!old.isNullObject) old.delete();
    const created = renderShapes(slide, scene, { ...opts, left: target.left, top: target.top });
    await context.sync();
    await groupAndTag(context, slide, created, { ...opts, left: target.left, top: target.top });
  });
}

/**
 * Read the PowerChart config back from the current selection (the tag written
 * at insert time). Returns null when the selection is not a PowerChart.
 * Requires PowerPointApi 1.5 (getSelectedShapes).
 */
export async function loadChartFromSelection(): Promise<{ configJson: string; target: EditTarget } | null> {
  return PowerPoint.run(async (context) => {
    const slides = context.presentation.getSelectedSlides();
    const slide = slides.getItemAt(0);
    slide.load("id");
    const shapes = context.presentation.getSelectedShapes();
    shapes.load("items/id,items/left,items/top");
    await context.sync();

    const tags = shapes.items.map((s) => {
      const tag = s.tags.getItemOrNullObject(CHART_TAG);
      tag.load("value");
      return tag;
    });
    await context.sync();

    for (let i = 0; i < shapes.items.length; i++) {
      if (!tags[i].isNullObject && tags[i].value) {
        const s = shapes.items[i];
        return {
          configJson: tags[i].value,
          target: { slideId: slide.id, shapeId: s.id, left: s.left, top: s.top },
        };
      }
    }
    return null;
  });
}

/**
 * Bounds of the currently selected shape when it is NOT a PowerChart —
 * used to insert a new chart into a selected placeholder/frame.
 */
export async function getSelectionBounds(): Promise<{ left: number; top: number; width: number; height: number } | null> {
  try {
    return await PowerPoint.run(async (context) => {
      const shapes = context.presentation.getSelectedShapes();
      shapes.load("items/left,items/top,items/width,items/height");
      await context.sync();
      if (shapes.items.length !== 1) return null;
      const s = shapes.items[0];
      const tag = s.tags.getItemOrNullObject(CHART_TAG);
      tag.load("value");
      await context.sync();
      if (!tag.isNullObject && tag.value) return null; // it's a chart — edit, don't cover
      return { left: s.left, top: s.top, width: s.width, height: s.height };
    });
  } catch {
    return null;
  }
}

/** All PowerCharts in the current selection (for Same Scale on a subset). */
export async function listChartsInSelection(): Promise<{ configJson: string; target: EditTarget }[]> {
  return PowerPoint.run(async (context) => {
    const slide = context.presentation.getSelectedSlides().getItemAt(0);
    slide.load("id");
    const shapes = context.presentation.getSelectedShapes();
    shapes.load("items/id,items/left,items/top");
    await context.sync();
    const tags = shapes.items.map((s) => {
      const tag = s.tags.getItemOrNullObject(CHART_TAG);
      tag.load("value");
      return tag;
    });
    await context.sync();
    return shapes.items
      .map((s, i) => ({ s, tag: tags[i] }))
      .filter(({ tag }) => !tag.isNullObject && tag.value)
      .map(({ s, tag }) => ({
        configJson: tag.value,
        target: { slideId: slide.id, shapeId: s.id, left: s.left, top: s.top },
      }));
  });
}

/**
 * Find every PowerChart in the deck (any shape carrying the config tag),
 * across all slides. Used by "Same scale" to re-render charts together.
 */
export async function listChartsInDeck(): Promise<{ configJson: string; target: EditTarget }[]> {
  return PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items/id");
    await context.sync();

    const perSlide = slides.items.map((slide) => {
      slide.shapes.load("items/id,items/left,items/top");
      return slide;
    });
    await context.sync();

    const lookups: { slideId: string; shape: PowerPoint.Shape; tag: PowerPoint.Tag }[] = [];
    for (const slide of perSlide) {
      for (const shape of slide.shapes.items) {
        const tag = shape.tags.getItemOrNullObject(CHART_TAG);
        tag.load("value");
        lookups.push({ slideId: slide.id, shape, tag });
      }
    }
    await context.sync();

    return lookups
      .filter((l) => !l.tag.isNullObject && l.tag.value)
      .map((l) => ({
        configJson: l.tag.value,
        target: { slideId: l.slideId, shapeId: l.shape.id, left: l.shape.left, top: l.shape.top },
      }));
  });
}

/**
 * Append one agenda slide per chapter, each highlighting its own chapter
 * (think-cell's agenda). Slides are appended at the end of the deck —
 * PowerPointApi's slides.add has no insert-at-position — so move them into
 * place afterwards. Requires PowerPointApi 1.3 (slides.add).
 */
export async function insertAgendaSlides(scenes: Scene[]): Promise<void> {
  await PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    const before = slides.getCount();
    await context.sync();
    for (let i = 0; i < scenes.length; i++) slides.add();
    await context.sync();
    for (let i = 0; i < scenes.length; i++) {
      const slide = slides.getItemAt(before.value + i);
      renderShapes(slide, scenes[i], { left: 0, top: 0, group: false, tagData: undefined });
    }
    await context.sync();
  });
}

/**
 * Testing aid: append one slide per item and render its chart, tagged so each
 * stays re-editable. Shapes are committed for every slide first (one sync),
 * then grouped/tagged per slide — so a host lacking grouping (PowerPoint on the
 * web) never rolls back the already-inserted charts. Requires PowerPointApi 1.3
 * (slides.add).
 */
export async function insertDemoDeck(items: { scene: Scene; tagData?: string }[]): Promise<void> {
  await PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    const before = slides.getCount();
    await context.sync();
    for (let i = 0; i < items.length; i++) slides.add();
    await context.sync();
    const perSlide = items.map((item, i) => {
      const slide = slides.getItemAt(before.value + i);
      const created = renderShapes(slide, item.scene, { left: 60, top: 90 });
      return { slide, created, tagData: item.tagData };
    });
    await context.sync();
    for (const p of perSlide) {
      await groupAndTag(context, p.slide, p.created, { group: true, tagData: p.tagData });
    }
  });
}

/** True when the host advertises the given PowerPointApi requirement set. */
function supports(version: string): boolean {
  try {
    return Office.context.requirements.isSetSupported("PowerPointApi", version);
  } catch {
    return false;
  }
}

/** Add every scene node as a shape (no grouping/tagging). Returns the shapes. */
function renderShapes(slide: PowerPoint.Slide, scene: Scene, opts: InsertOptions): PowerPoint.Shape[] {
  const left = opts.left ?? 60;
  const top = opts.top ?? 90;
  const shapes = slide.shapes;
  const created: PowerPoint.Shape[] = [];
  for (const n of scene.nodes) {
    created.push(...addNode(shapes, n, left, top, opts));
  }
  return created;
}

/**
 * Group the inserted shapes and persist the config tag — each in its OWN sync
 * and gated on host support, so a host that lacks grouping (e.g. PowerPoint on
 * the web) or tags never rolls back the already-committed shapes. The shapes
 * must already be committed (a prior context.sync) before this runs.
 */
async function groupAndTag(
  context: PowerPoint.RequestContext,
  slide: PowerPoint.Slide,
  created: PowerPoint.Shape[],
  opts: InsertOptions,
): Promise<void> {
  let tagTarget: PowerPoint.Shape | undefined = created[0];
  // Grouping is PowerPointApi 1.8+; skip entirely where unsupported.
  if (opts.group !== false && created.length > 1 && supports("1.8")) {
    try {
      const group = (slide.shapes as unknown as { addGroup(items: PowerPoint.Shape[]): PowerPoint.Shape }).addGroup(created);
      group.name = "PowerChart";
      await context.sync();
      tagTarget = group;
    } catch {
      /* grouping failed — shapes stay ungrouped, chart is already on the slide */
    }
  }
  // Tags are PowerPointApi 1.3+; keep the chart re-editable where supported.
  if (opts.tagData && tagTarget && supports("1.3")) {
    try {
      tagTarget.tags.add(CHART_TAG, opts.tagData);
      await context.sync();
    } catch {
      /* tags unavailable — chart is inserted but not re-editable */
    }
  }
}

function getTargetSlide(context: PowerPoint.RequestContext): PowerPoint.Slide {
  try {
    return context.presentation.getSelectedSlides().getItemAt(0);
  } catch {
    return context.presentation.slides.getItemAt(0);
  }
}

/** A straight segment's appearance — the subset of a line node the host needs. */
interface SegmentStyle {
  stroke: string;
  strokeWidth?: number;
  dash?: number[];
  name?: string;
}

/**
 * Draw one straight segment as a native shape, in absolute slide coordinates.
 *
 * PowerPoint's addLine takes only a bounding box, so it can't tell an up-right
 * line from a down-right one — and a zero-thickness box makes the web host
 * substitute a default and draw a giant diagonal. Three cases:
 *
 * - Axis-aligned (baselines, gridlines, connectors, value lines): addLine with
 *   the near-zero dimension clamped. The box is unambiguous.
 * - Dashed diagonal (scatter trend lines, forecast segments, pie breakout
 *   connectors): a real line shape, the only kind that can carry a native dash.
 *   addLine draws the box's top-left→bottom-right diagonal and the lineInverse
 *   geometry draws top-right→bottom-left, so between them the direction is
 *   explicit rather than guessed.
 * - Solid diagonal (line-chart series, the common case): a thin rotated
 *   rectangle, which is direction-correct on every host.
 */
function addSegment(
  shapes: PowerPoint.ShapeCollection,
  x1: number,
  y1: number,
  x2: number,
  y2: number,
  s: SegmentStyle,
): PowerPoint.Shape {
  const w = Math.abs(x2 - x1);
  const h = Math.abs(y2 - y1);
  const box = {
    left: Math.min(x1, x2),
    top: Math.min(y1, y2),
    width: Math.max(w, 0.5),
    height: Math.max(h, 0.5),
  };
  const setDash = (shape: PowerPoint.Shape) => {
    if (!s.dash) return;
    try {
      shape.lineFormat.dashStyle = PowerPoint.ShapeLineDashStyle.dash;
    } catch {
      /* dash style unsupported on this host */
    }
  };

  if (w < 0.5 || h < 0.5 || s.dash) {
    const downRight = (x2 - x1) * (y2 - y1) > 0;
    const line =
      w < 0.5 || h < 0.5 || downRight
        ? shapes.addLine(PowerPoint.ConnectorType.straight, box)
        : shapes.addGeometricShape(PowerPoint.GeometricShapeType.lineInverse, box);
    line.lineFormat.color = s.stroke;
    line.lineFormat.weight = s.strokeWidth ?? 1;
    setDash(line);
    if (s.name) line.name = s.name;
    return line;
  }

  const len = Math.hypot(x2 - x1, y2 - y1);
  const weight = Math.max(0.5, s.strokeWidth ?? 1);
  const rect = shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle, {
    left: (x1 + x2) / 2 - len / 2,
    top: (y1 + y2) / 2 - weight / 2,
    width: len,
    height: weight,
  });
  rect.fill.setSolidColor(s.stroke);
  rect.lineFormat.visible = false;
  try {
    rect.rotation = (Math.atan2(y2 - y1, x2 - x1) * 180) / Math.PI;
  } catch {
    /* rotation unsupported — line renders horizontally */
  }
  if (s.name) rect.name = s.name;
  return rect;
}

function addNode(
  shapes: PowerPoint.ShapeCollection,
  n: SceneNode,
  dx: number,
  dy: number,
  opts: InsertOptions,
): PowerPoint.Shape[] {
  switch (n.kind) {
    case "rect": {
      const shape = shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle, {
        left: dx + n.x,
        top: dy + n.y,
        width: Math.max(0.2, n.w),
        height: Math.max(0.2, n.h),
      });
      shape.fill.setSolidColor(n.fill);
      if (n.stroke && (n.strokeWidth ?? 0) > 0) {
        shape.lineFormat.color = n.stroke;
        shape.lineFormat.weight = n.strokeWidth ?? 1;
      } else {
        shape.lineFormat.visible = false;
      }
      if (n.name) shape.name = n.name;
      return [shape];
    }
    case "line":
      return [addSegment(shapes, dx + n.x1, dy + n.y1, dx + n.x2, dy + n.y2, n)];
    case "ellipse": {
      const shape = shapes.addGeometricShape(PowerPoint.GeometricShapeType.ellipse, {
        left: dx + n.cx - n.rx,
        top: dy + n.cy - n.ry,
        width: Math.max(0.2, n.rx * 2),
        height: Math.max(0.2, n.ry * 2),
      });
      // Stroke-only ellipses (radar circle grid) carry fill "none".
      if (n.fill === "none") shape.fill.clear();
      else shape.fill.setSolidColor(n.fill);
      if (n.stroke && (n.strokeWidth ?? 0) > 0) {
        shape.lineFormat.color = n.stroke;
        shape.lineFormat.weight = n.strokeWidth ?? 1;
      } else {
        shape.lineFormat.visible = false;
      }
      if (n.name) shape.name = n.name;
      return [shape];
    }
    case "chevron": {
      const shape = shapes.addGeometricShape(
        n.flatLeft ? PowerPoint.GeometricShapeType.homePlate : PowerPoint.GeometricShapeType.chevron,
        { left: dx + n.x, top: dy + n.y, width: Math.max(0.2, n.w), height: Math.max(0.2, n.h) },
      );
      shape.fill.setSolidColor(n.fill);
      shape.lineFormat.visible = false;
      if (n.name) shape.name = n.name;
      return [shape];
    }
    case "symbol": {
      // Native preset geometry, so the marker stays FILLED here — the reason a
      // symbol is its own kind rather than a polygon, which PowerPoint can only
      // outline. SYMBOL_PRESET names are GeometricShapeType keys.
      const geo = (PowerPoint.GeometricShapeType as unknown as Record<string, PowerPoint.GeometricShapeType>)[
        SYMBOL_PRESET[n.shape]
      ];
      const shape = shapes.addGeometricShape(geo, {
        left: dx + n.cx - n.size,
        top: dy + n.cy - n.size,
        width: Math.max(0.2, n.size * 2),
        height: Math.max(0.2, n.size * 2),
      });
      shape.fill.setSolidColor(n.fill);
      if (n.stroke && (n.strokeWidth ?? 0) > 0) {
        shape.lineFormat.color = n.stroke;
        shape.lineFormat.weight = n.strokeWidth ?? 1;
      } else {
        shape.lineFormat.visible = false;
      }
      if (n.name) shape.name = n.name;
      return [shape];
    }
    case "wedge":
      return addWedgeFan(shapes, n, dx, dy);
    case "polygon": {
      // No freeform paths in Office.js: draw the outline as connected line
      // segments (translucent fills degrade to outline-only in PowerPoint).
      // These go through addSegment like any other line — passing each edge's
      // bounding box straight to addLine mirrored every up-right edge and gave
      // horizontal ones a zero-height box.
      const created: PowerPoint.Shape[] = [];
      const pts = n.points;
      for (let i = 0; i < pts.length; i++) {
        const a = pts[i];
        const b = pts[(i + 1) % pts.length];
        created.push(
          addSegment(shapes, dx + a.x, dy + a.y, dx + b.x, dy + b.y, {
            stroke: n.stroke ?? n.fill ?? "#000000",
            strokeWidth: n.strokeWidth,
            name: n.name ? `${n.name}-e${i}` : undefined,
          }),
        );
      }
      return created;
    }
    case "text":
      return [addText(shapes, n, dx, dy, opts)];
    case "arrowhead": {
      // No freeform API in Office.js: a rotated geometric triangle whose tip is
      // offset onto (n.x, n.y) about the box centre — see arrowheadBox.
      const box = arrowheadBox(n.x, n.y, n.size, n.angle);
      const shape = shapes.addGeometricShape(PowerPoint.GeometricShapeType.triangle, {
        left: dx + box.left,
        top: dy + box.top,
        width: box.size,
        height: box.size,
      });
      shape.fill.setSolidColor(n.fill);
      shape.lineFormat.visible = false;
      try {
        // Geometric 'triangle' points up (= -90° in scene terms); rotation is
        // exposed from PowerPointApi 1.10 — best effort on older hosts.
        (shape as unknown as { rotation: number }).rotation = box.rotation;
      } catch {
        /* rotation unsupported — arrowhead stays axis-aligned */
      }
      if (n.name) shape.name = n.name;
      return [shape];
    }
  }
}

function addText(
  shapes: PowerPoint.ShapeCollection,
  n: TextNode,
  dx: number,
  dy: number,
  opts: InsertOptions,
): PowerPoint.Shape {
  const shape = shapes.addTextBox(n.text, {
    left: dx + n.x,
    top: dy + n.y,
    width: Math.max(4, n.w),
    height: Math.max(4, n.h),
  });
  shape.fill.clear();
  shape.lineFormat.visible = false;
  const tf = shape.textFrame;
  try {
    tf.wordWrap = false;
    tf.autoSizeSetting = PowerPoint.ShapeAutoSize.autoSizeNone;
    tf.leftMargin = 0;
    tf.rightMargin = 0;
    tf.topMargin = 0;
    tf.bottomMargin = 0;
    tf.verticalAlignment =
      n.valign === "top"
        ? PowerPoint.TextVerticalAlignment.top
        : n.valign === "bottom"
          ? PowerPoint.TextVerticalAlignment.bottom
          : PowerPoint.TextVerticalAlignment.middle;
  } catch {
    /* margin/alignment properties unavailable on this host */
  }
  const font = tf.textRange.font;
  font.size = n.fontSize;
  font.color = n.color;
  font.bold = !!n.bold;
  font.name = n.fontFamily ?? opts.fontFamily ?? DEFAULT_FONT;
  try {
    tf.textRange.paragraphFormat.horizontalAlignment =
      n.align === "left"
        ? PowerPoint.ParagraphHorizontalAlignment.left
        : n.align === "right"
          ? PowerPoint.ParagraphHorizontalAlignment.right
          : PowerPoint.ParagraphHorizontalAlignment.center;
  } catch {
    /* paragraph alignment unavailable */
  }
  if (n.name) shape.name = n.name;
  return shape;
}

/**
 * Approximate a pie wedge with a fan of rotated triangles — Office.js has no
 * adjustable pie geometry or freeform paths. Needs Shape.rotation (1.10);
 * older hosts get no wedge. Doughnut holes are separate ellipse nodes
 * emitted by the layout, so wedges here are always full slices.
 */
function addWedgeFan(
  shapes: PowerPoint.ShapeCollection,
  n: WedgeNode,
  dx: number,
  dy: number,
): PowerPoint.Shape[] {
  const created: PowerPoint.Shape[] = [];
  const cx = dx + n.cx;
  const cy = dy + n.cy;
  const span = n.endAngle - n.startAngle;
  // Annular wedge (sunburst ring / gauge): a triangle can't leave a hole, so
  // the band from innerR→r is drawn as radial rectangles; solid slices keep the
  // triangle fan (which tapers to the centre).
  const annular = n.innerR > 0;
  const midR = annular ? (n.innerR + n.r) / 2 : n.r / 2;
  const bandH = annular ? n.r - n.innerR : n.r;
  // Adaptive fan density (chord sagitta under ~0.5pt), capped — see wedgeFanSteps.
  const { steps, step } = wedgeFanSteps(n.r, span);
  for (let i = 0; i < steps; i++) {
    const mid = n.startAngle + step * (i + 0.5);
    // Slightly overlapping chords hide the seams between fan shapes.
    const chord = 2 * midR * Math.tan(((step / 2) * Math.PI) / 180) + 1;
    const center = polar(cx, cy, midR, mid);
    try {
      const shape = shapes.addGeometricShape(
        annular ? PowerPoint.GeometricShapeType.rectangle : PowerPoint.GeometricShapeType.triangle,
        {
          left: center.x - chord / 2,
          top: center.y - bandH / 2,
          width: chord,
          height: bandH,
        },
      );
      shape.fill.setSolidColor(n.fill);
      shape.lineFormat.visible = false;
      // Unrotated the rectangle's height / the triangle's base points south
      // (180° in wedge terms); rotate so it runs along `mid`.
      (shape as unknown as { rotation: number }).rotation = annular ? mid : mid - 180;
      if (n.name) shape.name = `${n.name}-f${i}`;
      created.push(shape);
    } catch {
      /* rotation unsupported — skip the fan on this host */
      break;
    }
  }
  // Best-effort slice outline: the two radial boundary edges as thin rectangles
  // in the stroke colour (stroking every fan seam would web the slice). This
  // reproduces think-cell's thin separators between adjacent slices. Drawn as
  // rotated rectangles, not addLine, since a line's bounding box can't encode a
  // diagonal's direction.
  if (n.stroke && span < 359.9) {
    const eInner = annular ? n.innerR : 0;
    const eLen = n.r - eInner;
    const eMidR = (eInner + n.r) / 2;
    const sw = n.strokeWidth ?? 1;
    for (const ang of [n.startAngle, n.endAngle]) {
      const c = polar(cx, cy, eMidR, ang);
      try {
        const edge = shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle, {
          left: c.x - sw / 2,
          top: c.y - eLen / 2,
          width: sw,
          height: eLen,
        });
        edge.fill.setSolidColor(n.stroke);
        edge.lineFormat.visible = false;
        (edge as unknown as { rotation: number }).rotation = ang;
        if (n.name) edge.name = `${n.name}-edge`;
        created.push(edge);
      } catch {
        /* rotation unsupported — skip the separator */
      }
    }
  }
  return created;
}

/**
 * Read the presentation's theme accent colors (Accent1-6) from the current
 * slide's color scheme — the deck's actual corporate palette. Requires
 * PowerPointApi 1.10 (ThemeColorScheme); returns null on older hosts.
 */
export async function loadThemePalette(): Promise<string[] | null> {
  try {
    return await PowerPoint.run(async (context) => {
      const slide = context.presentation.getSelectedSlides().getItemAt(0);
      const scheme = (slide as unknown as { themeColorScheme: { getThemeColor(c: string): { value: string } } })
        .themeColorScheme;
      const accents = ["Accent1", "Accent2", "Accent3", "Accent4", "Accent5", "Accent6"].map((a) =>
        scheme.getThemeColor(a),
      );
      await context.sync();
      const palette = accents
        .map((r) => r.value)
        .filter(Boolean)
        .map((c) => (c.startsWith("#") ? c : `#${c}`).toLowerCase());
      return palette.length >= 3 ? palette : null;
    });
  } catch {
    return null; // no selection, or host below PowerPointApi 1.10
  }
}

/** True when running inside an Office host with the PowerPoint JS API. */
export function isPowerPointHost(): boolean {
  return (
    typeof Office !== "undefined" &&
    typeof PowerPoint !== "undefined" &&
    !!Office.context?.host
  );
}
