/**
 * Office.js renderer: draws a scene as native, individually editable
 * PowerPoint shapes on the current slide, then groups them.
 *
 * This is the same output strategy as think-cell/UpSlide-style tools —
 * charts stay fully editable in PowerPoint (every bar and label is a shape),
 * rather than being pasted as pictures or opaque OLE charts.
 *
 * Requires PowerPointApi 1.4+ (ShapeCollection.addGeometricShape / addLine /
 * addTextBox). Grouping and arrowhead rotation degrade gracefully on older hosts.
 */
import type { Scene, SceneNode, TextNode } from "../core/scene";

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
    renderScene(getTargetSlide(context), scene, opts);
    await context.sync();
  });
}

/** Replace an existing PowerChart group with a re-rendered scene, in place. */
export async function updateChartInSlide(scene: Scene, target: EditTarget, opts: InsertOptions = {}): Promise<void> {
  await PowerPoint.run(async (context) => {
    const slide = context.presentation.slides.getItem(target.slideId);
    const old = slide.shapes.getItemOrNullObject(target.shapeId);
    await context.sync();
    if (!old.isNullObject) old.delete();
    renderScene(slide, scene, { ...opts, left: target.left, top: target.top });
    await context.sync();
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
      renderScene(slide, scenes[i], { left: 0, top: 0, group: false, tagData: undefined });
    }
    await context.sync();
  });
}

function renderScene(slide: PowerPoint.Slide, scene: Scene, opts: InsertOptions): void {
  const left = opts.left ?? 60;
  const top = opts.top ?? 90;
  const shapes = slide.shapes;
  const created: PowerPoint.Shape[] = [];

  for (const n of scene.nodes) {
    const shape = addNode(shapes, n, left, top, opts);
    if (shape) created.push(shape);
  }

  let tagTarget: PowerPoint.Shape | undefined = created[0];
  if (opts.group !== false && created.length > 1) {
    try {
      // PowerPointApi 1.8+. On older hosts the shapes are simply left ungrouped.
      tagTarget = (shapes as unknown as { addGroup(items: PowerPoint.Shape[]): PowerPoint.Shape }).addGroup(created);
      tagTarget.name = "PowerChart";
    } catch {
      /* grouping unavailable — leave shapes ungrouped */
    }
  }
  if (opts.tagData && tagTarget) {
    try {
      // PowerPointApi 1.3+ — persists the data model in the document.
      tagTarget.tags.add(CHART_TAG, opts.tagData);
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

function addNode(
  shapes: PowerPoint.ShapeCollection,
  n: SceneNode,
  dx: number,
  dy: number,
  opts: InsertOptions,
): PowerPoint.Shape | null {
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
      return shape;
    }
    case "line": {
      const line = shapes.addLine(PowerPoint.ConnectorType.straight, {
        left: dx + Math.min(n.x1, n.x2),
        top: dy + Math.min(n.y1, n.y2),
        width: Math.abs(n.x2 - n.x1),
        height: Math.abs(n.y2 - n.y1),
      });
      line.lineFormat.color = n.stroke;
      line.lineFormat.weight = n.strokeWidth ?? 1;
      if (n.dash) {
        try {
          line.lineFormat.dashStyle = PowerPoint.ShapeLineDashStyle.dash;
        } catch {
          /* dash style unsupported on this host */
        }
      }
      if (n.name) line.name = n.name;
      return line;
    }
    case "ellipse": {
      const shape = shapes.addGeometricShape(PowerPoint.GeometricShapeType.ellipse, {
        left: dx + n.cx - n.rx,
        top: dy + n.cy - n.ry,
        width: Math.max(0.2, n.rx * 2),
        height: Math.max(0.2, n.ry * 2),
      });
      shape.fill.setSolidColor(n.fill);
      if (n.stroke && (n.strokeWidth ?? 0) > 0) {
        shape.lineFormat.color = n.stroke;
        shape.lineFormat.weight = n.strokeWidth ?? 1;
      } else {
        shape.lineFormat.visible = false;
      }
      if (n.name) shape.name = n.name;
      return shape;
    }
    case "text":
      return addText(shapes, n, dx, dy, opts);
    case "arrowhead": {
      // No freeform API in Office.js: approximate with a rotated triangle.
      const s = n.size * 2;
      const shape = shapes.addGeometricShape(PowerPoint.GeometricShapeType.triangle, {
        left: dx + n.x - s / 2,
        top: dy + n.y - s / 2,
        width: s,
        height: s,
      });
      shape.fill.setSolidColor(n.fill);
      shape.lineFormat.visible = false;
      try {
        // Geometric 'triangle' points up (= -90° in scene terms); rotation is
        // exposed from PowerPointApi 1.9 — best effort on older hosts.
        (shape as unknown as { rotation: number }).rotation = n.angle + 90;
      } catch {
        /* rotation unsupported — arrowhead stays axis-aligned */
      }
      if (n.name) shape.name = n.name;
      return shape;
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

/** True when running inside an Office host with the PowerPoint JS API. */
export function isPowerPointHost(): boolean {
  return (
    typeof Office !== "undefined" &&
    typeof PowerPoint !== "undefined" &&
    !!Office.context?.host
  );
}
