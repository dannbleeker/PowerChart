import type { Scene, SceneNode } from "./scene";
import { finiteNodes, textWidth } from "./scene";
import { clipToWidth } from "./elements";
import { clampDim } from "./chart";
import { DEFAULT_STYLE } from "./style";

/** Standard 16:9 slide in points. */
export const SLIDE = { width: 960, height: 540 };

export interface AgendaOptions {
  title?: string;
  /** Chapter to highlight (the slide's position in the deck); -1 = none (overview slide). */
  highlight?: number;
  width?: number;
  height?: number;
}

/**
 * think-cell-style agenda (chapter) slide: the agenda list with the current
 * chapter highlighted and the others dimmed. One scene per chapter; insert
 * each before its section.
 */
export function buildAgendaScene(chapters: string[], opts: AgendaOptions = {}): Scene {
  // A `string[]` in the types is not an array in the object someone passed.
  // Exported from `src/index.ts`, so the caller may be anyone — and an agenda
  // with no chapters is an ordinary thing to ask for, not a crash. Same rule
  // and same reason as `buildProcessFlow`.
  const list: string[] = Array.isArray(chapters) ? chapters.map((c) => String(c ?? "")) : [];
  // The same clamp `buildChart` puts on a chart's dimensions, for the same
  // reason and from the same place — that function's own comment says one rule
  // in one place, and this was a third place without it. `??` catches only null
  // and undefined, so `{kind: "agenda", width: "wide"}` out of the skill's
  // caller JSON put a non-number into every arithmetic below: all seven nodes
  // came out with NaN coordinates, and they were emitted, because this is the
  // one scene builder in the repo that did not end with `finiteNodes`.
  //
  // In the SVG that is a visibly broken slide. In the pptx path it is what
  // `finiteNodes` exists to prevent — `x="Infinity"` is not an Int64 and
  // Microsoft's validator rejects the whole deck.
  const width = clampDim(opts.width, SLIDE.width);
  const height = clampDim(opts.height, SLIDE.height);
  const highlight = opts.highlight ?? -1;
  const s = DEFAULT_STYLE;
  const accent = s.palette[0];

  const marginX = width * 0.09;
  const titleY = height * 0.09;
  const listY = height * 0.28;
  const rowH = Math.min(46, (height * 0.62) / Math.max(1, list.length));
  const fsTitle = 28;
  // The chapter text starts after the number column (fs * 2.6) and its box ends
  // at width - marginX. The row font came from the chapter COUNT alone, so a
  // long title ran off the right edge of the slide — and neither PowerPoint
  // renderer wraps it. Shrink until the widest title fits, then ellipsize.
  const itemW = (f: number) => width - marginX * 2 - f * 2.6;
  let fs = Math.min(18, rowH * 0.42);
  const overflows = (f: number) => list.some((c, i) => textWidth(c, f, i === highlight) > itemW(f));
  while (fs > 9 && overflows(fs)) fs -= 0.5;

  const nodes: SceneNode[] = [
    {
      kind: "text",
      x: marginX,
      y: titleY,
      w: width - marginX * 2,
      h: fsTitle * 1.5,
      text: opts.title ?? "Agenda",
      fontSize: fsTitle,
      bold: true,
      color: s.text,
      align: "left",
      valign: "top",
      name: "agenda-title",
    },
    {
      kind: "line",
      x1: marginX,
      y1: titleY + fsTitle * 1.8,
      x2: width - marginX,
      y2: titleY + fsTitle * 1.8,
      stroke: s.gridline,
      strokeWidth: 1,
      name: "agenda-rule",
    },
  ];

  list.forEach((chapter, i) => {
    const y = listY + i * rowH;
    const active = i === highlight;
    if (active) {
      nodes.push({
        kind: "rect",
        x: marginX - 10,
        y: y - rowH * 0.12,
        w: width - marginX * 2 + 20,
        h: rowH * 0.9,
        fill: "#eaf2fc",
        name: `agenda-hl-${i}`,
      });
    }
    nodes.push(
      {
        kind: "text",
        x: marginX,
        y,
        w: fs * 2.2,
        h: rowH * 0.7,
        text: String(i + 1),
        fontSize: fs,
        bold: true,
        color: active ? accent : s.mutedText,
        align: "left",
        valign: "middle",
        name: `agenda-num-${i}`,
      },
      {
        kind: "text",
        x: marginX + fs * 2.6,
        y,
        w: itemW(fs),
        h: rowH * 0.7,
        text: clipToWidth(chapter, fs, itemW(fs), active),
        fontSize: fs,
        bold: active,
        color: active ? s.text : s.mutedText,
        align: "left",
        valign: "middle",
        name: `agenda-item-${i}`,
      },
    );
  });

  // Unreachable today, and deliberately here anyway: with the clamp above there
  // is no route to a non-finite number in this function, so nothing is being
  // filtered. It is the contract EVERY scene builder follows, and the value of
  // that is a reader not having to re-derive per builder whether an
  // un-openable file can escape — which is the derivation this file got wrong.
  // If the clamp is ever weakened, the failure is a blank slide rather than a
  // deck PowerPoint refuses to open.
  return { width, height, nodes: finiteNodes(nodes) };
}
