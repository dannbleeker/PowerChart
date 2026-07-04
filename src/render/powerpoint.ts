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
}

const DEFAULT_FONT = "Segoe UI";

export async function insertSceneIntoSlide(scene: Scene, opts: InsertOptions = {}): Promise<void> {
  const left = opts.left ?? 60;
  const top = opts.top ?? 90;

  await PowerPoint.run(async (context) => {
    const slide = getTargetSlide(context);
    const shapes = slide.shapes;
    const created: PowerPoint.Shape[] = [];

    for (const n of scene.nodes) {
      const shape = addNode(shapes, n, left, top, opts);
      if (shape) created.push(shape);
    }

    if (opts.group !== false && created.length > 1) {
      try {
        // PowerPointApi 1.8+. On older hosts the shapes are simply left ungrouped.
        (shapes as unknown as { addGroup(items: PowerPoint.Shape[]): PowerPoint.Shape }).addGroup(created);
      } catch {
        /* grouping unavailable — leave shapes ungrouped */
      }
    }
    await context.sync();
  });
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
