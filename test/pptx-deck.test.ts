import { describe, expect, it } from "vitest";
import JSZip from "jszip";
import { buildDeckBase64 } from "../src/render/pptx-deck";
import {
  canonicalSlideSize,
  injectGroupsAndTags,
  groupSlideShapes,
  tagFirstShape,
  tagPart,
  topLevelElements,
  xmlAttr,
} from "../src/render/ooxml";
import { buildChart } from "../src/core/chart";
import { sampleConfig } from "../src/core/samples";
import type { ChartConfig } from "../src/core/types";

/** A minimal slide tree of the shape pptxgenjs writes. */
function slideXml(shapes: string): string {
  return (
    `<?xml version="1.0" encoding="UTF-8"?><p:sld xmlns:p="p" xmlns:a="a" xmlns:r="r"><p:cSld><p:spTree>` +
    `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
    `<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm></p:grpSpPr>` +
    `${shapes}</p:spTree></p:cSld></p:sld>`
  );
}

/** One shape with a frame, at EMU coordinates. */
function sp(id: number, x: number, y: number, cx: number, cy: number, inner = ""): string {
  return (
    `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="part-${id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
    `<p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm></p:spPr>${inner}</p:sp>`
  );
}

describe("ooxml: splitting a shape tree", () => {
  it("returns each top-level shape, not the first closing tag it meets", () => {
    // A shape that CONTAINS a p:sp is exactly what breaks a textual
    // <p:sp>…</p:sp> match: it pairs the outer open with the inner close.
    const nested = `<p:sp><p:x>${sp(9, 0, 0, 1, 1)}</p:x></p:sp>`;
    const parts = topLevelElements(`${sp(2, 0, 0, 10, 10)}${nested}${sp(3, 0, 0, 10, 10)}`);
    expect(parts).toHaveLength(3);
    expect(parts[1]).toBe(nested);
  });

  it("keeps a self-closing element as its own top-level entry", () => {
    expect(topLevelElements(`<p:a/><p:b><p:c/></p:b>`)).toEqual(["<p:a/>", "<p:b><p:c/></p:b>"]);
  });
});

describe("ooxml: grouping a slide's shapes", () => {
  it("frames the group on its children's bounding box", () => {
    const xml = groupSlideShapes(slideXml(sp(2, 100, 200, 300, 400) + sp(3, 50, 900, 100, 100)));
    const off = /<a:off x="(\d+)" y="(\d+)"\/><a:ext cx="(\d+)" cy="(\d+)"\/><a:chOff/.exec(xml)!;
    expect(off.slice(1)).toEqual(["50", "200", "350", "800"]); // x 50..400, y 200..1000
  });

  it("repeats the frame as chOff/chExt, so the group transform is an identity", () => {
    // PowerPoint's own grouping does this, and it is why children keep their
    // absolute coordinates. Any other child extent silently rescales the chart.
    const xml = groupSlideShapes(slideXml(sp(2, 100, 200, 300, 400) + sp(3, 50, 900, 100, 100)));
    const m =
      /<a:off x="(\d+)" y="(\d+)"\/><a:ext cx="(\d+)" cy="(\d+)"\/><a:chOff x="(\d+)" y="(\d+)"\/><a:chExt cx="(\d+)" cy="(\d+)"\/>/.exec(
        xml,
      )!;
    expect([m[5], m[6], m[7], m[8]]).toEqual([m[1], m[2], m[3], m[4]]);
  });

  it("leaves every child's own coordinates untouched", () => {
    const shapes = sp(2, 100, 200, 300, 400) + sp(3, 50, 900, 100, 100);
    const xml = groupSlideShapes(slideXml(shapes));
    expect(xml).toContain(shapes);
  });

  it("gives the group an id no shape on the slide is using", () => {
    const xml = groupSlideShapes(slideXml(sp(2, 0, 0, 1, 1) + sp(7, 0, 0, 1, 1)));
    expect(xml).toContain(`<p:cNvPr id="8" name="PowerChart"/>`);
  });

  it("hangs the config tag off the group, where the live renderer puts it", () => {
    const xml = groupSlideShapes(slideXml(sp(2, 0, 0, 1, 1) + sp(3, 0, 0, 1, 1)), "rId4");
    expect(xml).toContain(`<p:nvPr><p:custDataLst><p:tags r:id="rId4"/></p:custDataLst></p:nvPr>`);
  });

  it("does not group a single shape — there is nothing to hold together", () => {
    const one = slideXml(sp(2, 0, 0, 1, 1));
    expect(groupSlideShapes(one)).toBe(one);
  });

  it("refuses a tree it does not recognise instead of writing a broken slide", () => {
    expect(() => groupSlideShapes("<p:sld/>")).toThrow(/no <p:spTree>/);
  });
});

describe("ooxml: tags and deck size", () => {
  it("escapes a config JSON payload into an attribute", () => {
    const part = tagPart([["POWERCHART_CONFIG", `{"title":"A & <B>"}`]]);
    expect(part).toContain(`val="{&quot;title&quot;:&quot;A &amp; &lt;B&gt;&quot;}"`);
    expect(xmlAttr(`'`)).toBe("&apos;");
  });

  it("carries several tags in one part, since a shape may reference only one", () => {
    const part = tagPart([
      ["POWERCHART_CONFIG", "{}"],
      ["POWERCHART_ORIGIN", "[0,0,0,0]"],
    ]);
    expect(part.match(/<p:tag /g)).toHaveLength(2);
  });

  it("rewrites the slide size to PowerPoint's exact 16:9", () => {
    // pptxgenjs writes 12191695 for the same nominal 13.333in — 305 EMU short.
    // A deck whose size differs from the destination's invites a rescale on
    // insert, and a rescale moves every chart.
    const out = canonicalSlideSize(`<p:presentation><p:sldSz cx="12191695" cy="6858000"/></p:presentation>`);
    expect(out).toContain(`<p:sldSz cx="12192000" cy="6858000"/>`);
  });
});

describe("ooxml: a slide with nothing to group", () => {
  it("tags the shape itself, so a one-picture chart is still re-editable", () => {
    // Image mode is one picture. There is no group to hang the config on, and
    // without this such a chart carries no POWERCHART_CONFIG at all — the
    // difference between a picture of a chart and a chart.
    const one = slideXml(`<p:pic><p:nvPicPr><p:cNvPr id="2" name="img"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr></p:pic>`);
    expect(groupSlideShapes(one)).toBe(one); // nothing to group
    expect(tagFirstShape(one, "rId7")).toContain(
      `<p:nvPr><p:custDataLst><p:tags r:id="rId7"/></p:custDataLst></p:nvPr>`,
    );
  });

  it("extends an nvPr that already has content rather than replacing it", () => {
    const withPh = slideXml(
      `<p:sp><p:nvSpPr><p:cNvPr id="2" name="s"/><p:cNvSpPr/><p:nvPr><p:ph type="body"/></p:nvPr></p:nvSpPr></p:sp>`,
    );
    const out = tagFirstShape(withPh, "rId3");
    // The placeholder survives, and CT_ApplicationNonVisualDrawingProps wants
    // ph before custDataLst — which is the order this produces.
    expect(out).toContain(`<p:nvPr><p:ph type="body"/><p:custDataLst><p:tags r:id="rId3"/></p:custDataLst></p:nvPr>`);
  });

  it("leaves a slide it cannot read alone", () => {
    expect(tagFirstShape("<p:sld/>", "rId1")).toBe("<p:sld/>");
  });
});

describe("ooxml: dressing and slides must stay in step", () => {
  it("refuses a deck with more slides than dressing entries", async () => {
    // `dressing[i]` IS `slide{i+1}.xml` — the pairing is strictly positional.
    // A caller whose two counts drift apart puts every later chart's config on
    // the wrong slide, and the file still opens cleanly, so nothing surfaces
    // it until someone edits a chart and overwrites a different one. Too MANY
    // dressing entries already failed loudly on the missing slide part; too
    // few used to leave the trailing slides silently ungrouped and untagged.
    const built = await buildDeckBase64([
      { scene: buildChart(sampleConfig("line")), title: "A" },
      { scene: buildChart(sampleConfig("line")), title: "B" },
    ]);
    await expect(injectGroupsAndTags(built.base64, [{ title: "A" }])).rejects.toThrow(/2 slide\(s\) but 1 dressing/);
  }, 30_000);
});

describe("ooxml: parsing bytes another library wrote", () => {
  it("does not end a tag at a raw > inside an attribute value", () => {
    // pptxgenjs writes `typeface="…"` WITHOUT escaping it — the one unescaped
    // attribute in the library — and font names are free-form user strings.
    // Read naively, the `>` ends the tag early and the depth counter desyncs
    // for the rest of the slide: this returns nothing, `groupSlideShapes` sees
    // fewer than two children, and the chart is never grouped at all.
    const nasty = `<p:sp><p:nvSpPr><p:cNvPr id="2" name="a"/></p:nvSpPr><a:latin typeface="Ba>d"/></p:sp>`;
    const parts = topLevelElements(`${nasty}${sp(3, 0, 0, 10, 10)}`);
    expect(parts).toHaveLength(2);
    expect(parts[0]).toBe(nasty);
  });

  it("numbers new tag parts above the ones already in the deck", async () => {
    // `zip.file()` overwrites. Starting at tag1 unconditionally destroyed any
    // pre-existing tag part and every slide pointing at it — safe today only
    // because pptxgenjs happens to emit none.
    const built = await buildDeckBase64([
      { scene: buildChart(sampleConfig("line")), title: "A", configJson: `{"kind":"line"}` },
    ]);
    const zip = await JSZip.loadAsync(built.base64, { base64: true });
    zip.file("ppt/tags/tag1.xml", "<keep-me/>");
    const out = await injectGroupsAndTags(await zip.generateAsync({ type: "base64" }), [
      { configJson: `{"kind":"line"}`, slot: 0, title: "A" },
    ]);
    const after = await JSZip.loadAsync(out.base64, { base64: true });
    expect(await after.file("ppt/tags/tag1.xml")!.async("string")).toBe("<keep-me/>");
    expect(Object.keys(after.files).filter((f) => /^ppt\/tags\/tag\d+\.xml$/.test(f)).length).toBeGreaterThan(1);
  }, 30_000);
});

describe("building a deck in-process", () => {
  const cfg = (title: string): ChartConfig => ({ ...sampleConfig("line"), title });

  it("produces a .pptx whose charts are grouped, tagged and slot-stamped", async () => {
    const configJson = JSON.stringify(cfg("Line"));
    const built = await buildDeckBase64([
      { scene: buildChart(cfg("Line")), title: "Line", configJson, slot: 3 },
      { scene: buildChart(cfg("Second")), title: "Second", slot: 4 },
    ]);
    const zip = await JSZip.loadAsync(built.base64, { base64: true });
    // The count the caller verifies against has to be what THIS renderer drew,
    // not what the Office.js one would have.
    expect(built.shapesPerSlide).toHaveLength(2);
    expect(built.shapesPerSlide[0]).toBeGreaterThan(1);

    const slides = Object.keys(zip.files).filter((f) => /^ppt\/slides\/slide\d+\.xml$/.test(f));
    expect(slides).toHaveLength(2);

    const slide1 = await zip.file("ppt/slides/slide1.xml")!.async("string");
    // One group per slide, holding the chart.
    expect(slide1.match(/<p:grpSp>/g)).toHaveLength(1);
    expect(slide1).toContain(`name="PowerChart"`);
    // The slot tag is a slide-level custDataLst, right after the shape tree.
    expect(slide1).toContain(`</p:spTree><p:custDataLst><p:tags r:id=`);

    // The config round-trips through the tag part unharmed.
    const tags = Object.keys(zip.files).filter((f) => /^ppt\/tags\/tag\d+\.xml$/.test(f));
    expect(tags.length).toBeGreaterThanOrEqual(3); // config+origin, slot, slot
    const partTexts = await Promise.all(tags.map((t) => zip.file(t)!.async("string")));
    const configTag = partTexts.find((t) => t.includes("POWERCHART_CONFIG"))!;
    const val = /name="POWERCHART_CONFIG" val="([^"]*)"/.exec(configTag)![1];
    const decoded = val
      .replace(/&quot;/g, '"')
      .replace(/&apos;/g, "'")
      .replace(/&lt;/g, "<")
      .replace(/&gt;/g, ">")
      .replace(/&amp;/g, "&");
    expect(JSON.parse(decoded)).toEqual(JSON.parse(configJson));

    // Every tag part is declared, or PowerPoint rejects the file.
    const types = await zip.file("[Content_Types].xml")!.async("string");
    for (const t of tags) {
      expect(types).toContain(`PartName="/${t}"`);
    }
    // …and reachable from its slide.
    const rels = await zip.file("ppt/slides/_rels/slide1.xml.rels")!.async("string");
    expect(rels.match(/relationships\/tags/g)).toHaveLength(2); // group tag + slot tag
    expect(rels).toContain("../tags/");

    const presentation = await zip.file("ppt/presentation.xml")!.async("string");
    expect(presentation).toContain(`<p:sldSz cx="12192000" cy="6858000"/>`);
  }, 30_000);

  it("leaves a chart without config ungrouped-but-slotted rather than inventing a tag", async () => {
    const built = await buildDeckBase64([{ scene: buildChart(cfg("Plain")), title: "Plain", slot: 0 }]);
    const zip = await JSZip.loadAsync(built.base64, { base64: true });
    const slide = await zip.file("ppt/slides/slide1.xml")!.async("string");
    expect(slide).toContain("<p:grpSp>");
    // The group carries no tag reference — there was no config to write.
    expect(slide).toContain(`<p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>`);
    const tags = Object.keys(zip.files).filter((f) => /^ppt\/tags\/tag\d+\.xml$/.test(f));
    expect(tags).toHaveLength(1);
    expect(await zip.file(tags[0])!.async("string")).toContain("POWERCHART_DEMO_SLOT");
  }, 30_000);
});
