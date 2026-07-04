import type { Scene, SceneNode } from "./scene";
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
  const width = opts.width ?? SLIDE.width;
  const height = opts.height ?? SLIDE.height;
  const highlight = opts.highlight ?? -1;
  const s = DEFAULT_STYLE;
  const accent = s.palette[0];

  const marginX = width * 0.09;
  const titleY = height * 0.09;
  const listY = height * 0.28;
  const rowH = Math.min(46, (height * 0.62) / Math.max(1, chapters.length));
  const fsTitle = 28;
  const fs = Math.min(18, rowH * 0.42);

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

  chapters.forEach((chapter, i) => {
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
        w: width - marginX * 2 - fs * 2.6,
        h: rowH * 0.7,
        text: chapter,
        fontSize: fs,
        bold: active,
        color: active ? s.text : s.mutedText,
        align: "left",
        valign: "middle",
        name: `agenda-item-${i}`,
      },
    );
  });

  return { width, height, nodes };
}
