/**
 * Build a real .pptx in the browser, so a whole deck can be handed to
 * PowerPoint in ONE call.
 *
 * Every other path in this renderer draws shape by shape through Office.js,
 * which is the only thing Office.js offers and the one thing PowerPoint on the
 * web handles badly: a 12-item demo run issues 19 slide adds and hundreds of
 * queued commands, and the host is free to drop an add, stall a sync past its
 * timeout, or refuse a group at any of them. The 2026-07-30 run took 680
 * seconds and shipped duplicate slides.
 *
 * A file has none of those surfaces. Grouping is in the file; the config tag is
 * in the file; there is exactly one call to lose.
 *
 * The painting is NOT reimplemented here — `skill/scripts/pptx-paint.mjs` is
 * the same pure node-to-pptxgenjs mapping the headless skill renderer uses, so
 * the add-in's file output and the skill's file output stay one implementation
 * rather than two that drift. What this module adds is the two things
 * pptxgenjs cannot express and SSF Charts cannot do without:
 *
 *   - a real `<p:grpSp>` per chart, so the chart is one draggable object;
 *   - a `POWERCHART_CONFIG` tag on that group, so it stays re-editable.
 *
 * Both are injected into the generated zip afterwards (see `ooxml.ts`).
 * pptxgenjs has no API for either — it can only write flat shapes onto a
 * slide.
 */
import { arrowheadBox, annularSectorPoints, dashKind, symbolPreset } from "../core/geometry";
import type { Scene } from "../core/scene";
import { sceneFingerprint } from "../core/scene-diff";
import { IN, hexOr, makeAddNode } from "../../skill/scripts/pptx-paint.mjs";
import { injectGroupsAndTags, EMU_PER_POINT, type SlideDressing } from "./ooxml";

/** 16:9 at PowerPoint's default size, in inches — the skill renderer's deck. */
export const DECK_SIZE = { w: 13.333, h: 7.5 };

/** One slide to generate. */
export interface DeckItem {
  scene: Scene;
  /** Slide title, carried into the slot tag so a repair pass can pair it up. */
  title?: string;
  /** Serialized ChartConfig — written as the group's `POWERCHART_CONFIG` tag. */
  configJson?: string;
  /** Item index, written as the slide's `POWERCHART_DEMO_SLOT` tag. */
  slot?: number;
  /** Run token, written alongside the slot — see `SlideDressing.run`. */
  run?: string;
  /** Slide background; defaults to white, matching the skill renderer. */
  background?: string;
}

/**
 * Render items to a base64 .pptx.
 *
 * pptxgenjs is loaded with a dynamic `import()` so it never enters the pane's
 * first-load bundle — it is ~1 MB of library that only a deck insert needs,
 * and the pane is fetched over the network every time PowerPoint opens it.
 */
export async function buildDeckBase64(
  items: DeckItem[],
  /**
   * The slide size of the deck this file is going INTO, in points.
   *
   * Omitted, the generated deck is 16:9 — which is what it always was, and
   * which is wrong on a 4:3 presentation in two compounding ways: the file
   * declares a size the destination does not share (so the host rescales every
   * slide on insert, moving every chart), and the centring below divides by a
   * width the destination does not have (so the chart is off-centre before the
   * rescale even starts). Passing the real size makes both go away, because a
   * generated deck that agrees with its destination is one nothing rescales.
   */
  slideSizePt?: { width: number; height: number },
): Promise<{ base64: string; shapesPerSlide: number[] }> {
  const { default: PptxGen } = await import("pptxgenjs");
  const pres = new PptxGen();
  const size = slideSizePt ? { w: slideSizePt.width / 72, h: slideSizePt.height / 72 } : DECK_SIZE;
  pres.defineLayout({ name: "WIDE", width: size.w, height: size.h });
  pres.layout = "WIDE";
  const addNode = makeAddNode({ dashKind, annularSectorPoints, symbolPreset, arrowheadBox });

  const dressing: SlideDressing[] = [];
  for (const item of items) {
    const slide = pres.addSlide();
    // `hexOr`, not `hex` — see the note at its definition. An unrecognised
    // background must fall back to white, because `hex`'s black is right for ink
    // and catastrophic for the paint everything else is read against.
    slide.background = { color: hexOr(item.background, "#ffffff") };
    // Centre the scene on the slide — the same arithmetic the skill renderer
    // uses, so a chart lands in the same place whichever path produced it.
    const dx = (size.w - item.scene.width * IN) / 2;
    const dy = (size.h - item.scene.height * IN) / 2;
    for (const node of item.scene.nodes) addNode(slide, node, dx, dy);
    dressing.push({
      configJson: item.configJson,
      // THE FINGERPRINT OF WHAT THIS BUILDER ACTUALLY DREW. Computed from the
      // same scene whose nodes are written above, so it describes the shapes on
      // the slide rather than whatever the stored config re-renders to later.
      // Only alongside a config: a chart with no config is not re-editable and
      // has nothing to diff against. See `SlideDressing.sceneTag`.
      sceneTag: item.configJson ? sceneFingerprint(item.scene) : undefined,
      slot: item.slot,
      title: item.title,
      run: item.run,
      // Where the live renderer would have put the chart's frame, so an edit
      // of a generated chart anchors the same way one drawn shape-by-shape
      // does. Points, matching CHART_ORIGIN_TAG's units.
      origin: item.configJson ? [dx / IN, dy / IN, dx / IN, dy / IN] : undefined,
    });
  }

  const base64 = (await pres.write({ outputType: "base64" })) as string;
  // Declare the same size the layout above was built for. pptxgenjs writes its
  // own rounded EMU (13.333in is 305 EMU short of the exact 13⅓), so the
  // declaration is rewritten from the size we actually used rather than left to
  // whatever the library rounded to.
  return injectGroupsAndTags(
    base64,
    dressing,
    slideSizePt
      ? { cx: Math.round(slideSizePt.width * EMU_PER_POINT), cy: Math.round(slideSizePt.height * EMU_PER_POINT) }
      : undefined,
  );
}
