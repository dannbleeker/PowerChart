// @vitest-environment jsdom
import { describe, it, expect } from "vitest";
import JSZip from "jszip";
import { buildDeckBase64 } from "../src/render/pptx-deck";
import { buildChart } from "../src/core/chart";
import { sampleConfig } from "../src/core/samples";
import type { ChartConfig } from "../src/core/types";
import fs from "node:fs";
const OUT: string[] = [];
const log = (s: string) => { OUT.push(s); };
const P = "/tmp/claude-0/-home-user-PowerChart/206f6d00-2ecf-5f5e-bcd9-49746102260f/scratchpad/t7x.txt";

async function report(label: string, cfg: ChartConfig) {
  const scene = buildChart(cfg);
  const built = await buildDeckBase64([{ scene, title: label, configJson: JSON.stringify(cfg), slot: 0 }]);
  const zip = await JSZip.loadAsync(built.base64, { base64: true });
  const slide = await zip.file("ppt/slides/slide1.xml")!.async("string");
  let origin: number[] | null = null;
  for (const f of Object.keys(zip.files).filter((f) => /^ppt\/tags\/tag\d+\.xml$/.test(f))) {
    const t = await zip.file(f)!.async("string");
    const m = /name="POWERCHART_ORIGIN" val="([^"]*)"/.exec(t);
    if (m) origin = JSON.parse(m[1].replace(/&quot;/g, String.fromCharCode(34)));
  }
  const g = /<p:grpSpPr><a:xfrm><a:off x="(-?\d+)" y="(-?\d+)"\/>/.exec(slide.slice(slide.indexOf("<p:grpSp>")));
  if (!g) { log(label + ": NO GROUP (single-shape slide)"); return; }
  const gl = Number(g[1]) / 12700, gt = Number(g[2]) / 12700;
  const [fl, ft, al, at] = origin as number[];
  // powerpoint.ts:1297-1298 with the chart UNTOUCHED (live pos == group frame)
  const redrawLeft = fl + (gl - al);
  const redrawTop = ft + (gt - at);
  log(
    label +
      "\\n  POWERCHART_ORIGIN = [" + [fl, ft, al, at].map((n) => n.toFixed(3)).join(", ") + "]  (frameLeft, frameTop, groupLeft, groupTop)" +
      "\\n  group <a:off>     = left " + gl.toFixed(3) + "pt, top " + gt.toFixed(3) + "pt" +
      "\\n  anchor error      = left " + (gl - al).toFixed(3) + "pt, top " + (gt - at).toFixed(3) + "pt" +
      "\\n  redraw frame (untouched chart) = " + redrawLeft.toFixed(3) + ", " + redrawTop.toFixed(3) +
      "  vs drawn at " + fl.toFixed(3) + ", " + ft.toFixed(3) +
      "  => CHART MOVES " + (redrawLeft - fl).toFixed(2) + "pt right, " + (redrawTop - ft).toFixed(2) + "pt down",
  );
}

describe("origin anchor", () => {
  it("reports", async () => {
    const wf = sampleConfig("waterfall");
    await report("waterfall WITH title (the case the tests use)", wf);
    await report("waterfall, title cleared to empty string", { ...wf, title: "" });
    await report("waterfall, no title key at all", { ...wf, title: undefined } as ChartConfig);
    const pie = sampleConfig("pie");
    await report("pie, no title", { ...pie, title: undefined } as ChartConfig);
    await report("mekko WITH its sample title", sampleConfig("mekko"));
    fs.writeFileSync(P, OUT.join("\n\n"));
    expect(true).toBe(true);
  }, 120000);
});
