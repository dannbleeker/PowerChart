import { describe, expect, it } from "vitest";
import JSZip from "jszip";
// @ts-expect-error — plain .mjs tool, no types; the audit is deliberately
// independent of src/ so it cannot inherit a bug from the code it audits.
import { readDeckBytes, faultsIn } from "../scripts/verify-deck.mjs";
import { tagPart, tagFirstShape } from "../src/render/ooxml";
import { buildDeckBase64 } from "../src/render/pptx-deck";
import { buildChart } from "../src/core/chart";
import { sampleConfig } from "../src/core/samples";

/**
 * The deck auditor.
 *
 * Every hard diagnosis in this project was settled by reading the .pptx rather
 * than the add-in's own report — including the one where the report was what
 * was wrong. This tool is that reading, committed. Its value depends entirely
 * on it disagreeing with a broken file, so the negative controls below matter
 * more than the happy path: a checker that passes everything is decoration.
 */

/** A small real deck: two tagged charts and one untagged. */
async function deck(): Promise<Uint8Array> {
  const cfg = (title: string) => ({ ...sampleConfig("line"), title });
  const built = await buildDeckBase64([
    { scene: buildChart(cfg("A")), title: "A", configJson: JSON.stringify(cfg("A")), slot: 0, run: "run-x" },
    { scene: buildChart(cfg("B")), title: "B", configJson: JSON.stringify(cfg("B")), slot: 1, run: "run-x" },
    { scene: buildChart(cfg("C")), title: "C", slot: 2, run: "run-x" },
  ]);
  return Uint8Array.from(atob(built.base64), (c) => c.charCodeAt(0));
}

/**
 * Mutate whichever part matches `find` — by CONTENT, not by filename.
 *
 * Tag parts are numbered in write order, so `tag2.xml` is not reliably slide
 * 2's slot tag. Targeting by name silently edited the wrong part and left one
 * of these corruptions a no-op, which made its test pass against an unmodified
 * deck: the exact vacuous-guard failure this repo has a rule about.
 */
async function mutateWhere(bytes: Uint8Array, find: RegExp, edit: (s: string) => string): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(bytes);
  const names = Object.keys(zip.files).filter((f) => /^ppt\/tags\/tag\d+\.xml$/.test(f));
  for (const n of names) {
    const text = await zip.file(n)!.async("string");
    if (!find.test(text)) continue;
    const next = edit(text);
    expect(next, `mutation of ${n} changed nothing`).not.toBe(text);
    zip.file(n, next);
    return zip.generateAsync({ type: "uint8array" });
  }
  throw new Error(`no tag part matched ${find}`);
}

/** Rebuild a deck's bytes after mutating one part — the corruption harness. */
async function mutate(bytes: Uint8Array, path: string, edit: (s: string) => string | null): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(bytes);
  const next = edit(await zip.file(path)!.async("string"));
  if (next === null) zip.remove(path);
  else zip.file(path, next);
  return zip.generateAsync({ type: "uint8array" });
}

describe("verify-deck: reading what a deck actually holds", () => {
  it("reports each slide's slot, run, chart object and config", async () => {
    const audit = await readDeckBytes(await deck());
    expect(audit.rows).toHaveLength(3);
    expect(audit.rows.map((r: { title: string }) => r.title)).toEqual(["A", "B", "C"]);
    expect(audit.rows.map((r: { slot: number }) => r.slot)).toEqual([0, 1, 2]);
    expect(audit.rows.every((r: { run: string }) => r.run === "run-x")).toBe(true);
    // Two carry a config, one does not — the distinction the whole tool exists
    // to make, and the one a real run's log got wrong for eleven charts.
    expect(audit.rows.map((r: { config: boolean }) => r.config)).toEqual([true, true, false]);
    expect(audit.rows.every((r: { chartObject: boolean }) => r.chartObject)).toBe(true);
    expect(faultsIn(audit)).toEqual([]);
  }, 30_000);

  it("does not confuse a picture-chart with a grouped one", async () => {
    // A degraded run leaves ONE shape named PowerChart with no group. Both are
    // "one object", and the difference is what says whether shapes survived.
    const bytes = await mutate(await deck(), "ppt/slides/slide1.xml", (s) =>
      s.replace(
        /<p:grpSp>[\s\S]*<\/p:grpSp>/,
        `<p:sp><p:nvSpPr><p:cNvPr id="9" name="PowerChart"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr></p:sp>`,
      ),
    );
    const audit = await readDeckBytes(bytes);
    expect(audit.rows[0]).toMatchObject({ chartObject: true, picture: true, groups: 0 });
    expect(audit.rows[1]).toMatchObject({ chartObject: true, picture: false, groups: 1 });
  }, 30_000);
});

describe("verify-deck: the faults it must catch", () => {
  // Each case corrupts a real deck one way and asserts the checker notices.
  // Without these the tool could return "no structural faults" unconditionally
  // and every test above would still pass.

  it("catches a slot tag whose payload will not parse", async () => {
    // The fault existed and could not fire. `slot = { malformed }` has no `i`,
    // so the row's `slot: slot?.i ?? null` collapsed it to null, and the check
    // read `r.slot !== null && typeof r.slot !== "number"` — satisfied only by a
    // slot tag that parses FINE and carries a non-numeric `i`, which is the one
    // case it was not written for. A deck this repo emitted with a corrupt slot
    // tag was certified "no structural faults", exit 0; triage then counted the
    // slide as untagged and reported its item lost, blaming the host for
    // dropping a slide sitting right there.
    const broken = await mutateWhere(await deck(), /POWERCHART_DEMO_SLOT/, (s) =>
      s.replace(/(name="POWERCHART_DEMO_SLOT" val=")[^"]*"/, '$1{&quot;i&quot;:0,&quot;title&quot;:&quot;X"'),
    );
    expect(faultsIn(await readDeckBytes(broken)).join(" ")).toMatch(/slot tag is malformed/i);
    // And the case the old check DID cover still fires: parses, `i` is a string.
    const stringy = await mutateWhere(await deck(), /POWERCHART_DEMO_SLOT/, (s) =>
      s.replace(/(name="POWERCHART_DEMO_SLOT" val=")[^"]*"/, '$1{&quot;i&quot;:&quot;0&quot;}"'),
    );
    expect(faultsIn(await readDeckBytes(stringy)).join(" ")).toMatch(/slot tag is malformed/i);
  }, 30_000);

  it("catches a slide pointing at a tag part that is not there", async () => {
    const bytes = await mutate(await deck(), "ppt/tags/tag1.xml", () => null);
    const faults = faultsIn(await readDeckBytes(bytes));
    expect(faults.join(" ")).toMatch(/missing tag part/i);
  }, 30_000);

  it("catches two shapes on one slide sharing a cNvPr id", async () => {
    // The classic trigger for PowerPoint's "repair" dialog, and the one defect
    // nothing else here can see: the schema does not express uniqueness, so
    // the OOXML validator passes a slide with a duplicated id (measured), and
    // a shape-tree walk that counts elements never compares them.
    //
    // Ids must be unique per slide INCLUDING inside groups, so the check reads
    // every cNvPr in the part rather than only the top level.
    const bytes = await mutate(await deck(), "ppt/slides/slide1.xml", (s) =>
      s.replace(/(<p:cNvPr[^>]*\bid=")(\d+)(")/g, (m, a, _id, c) => `${a}7${c}`),
    );
    const faults = faultsIn(await readDeckBytes(bytes));
    expect(faults.join(" ")).toMatch(/used more than once/i);
    expect(faults.join(" ")).toMatch(/repair/i);
  }, 30_000);

  it("does not cry duplicate over a healthy deck", async () => {
    // The other half: a check that fires on everything is the same as no check.
    const faults = faultsIn(await readDeckBytes(await deck()));
    expect(faults.join(" ")).not.toMatch(/used more than once/i);
  }, 30_000);

  it("catches a tag part no slide references", async () => {
    const zip = await JSZip.loadAsync(await deck());
    zip.file("ppt/tags/tag99.xml", "<p:tagLst/>");
    const faults = faultsIn(await readDeckBytes(await zip.generateAsync({ type: "uint8array" })));
    expect(faults.join(" ")).toMatch(/not referenced by any slide/i);
  }, 30_000);

  it("catches a tag part with no [Content_Types] override — PowerPoint rejects the file", async () => {
    const bytes = await mutate(await deck(), "[Content_Types].xml", (s) =>
      s.replace(/<Override PartName="\/ppt\/tags\/tag1\.xml"[^>]*\/>/, ""),
    );
    const faults = faultsIn(await readDeckBytes(bytes));
    expect(faults.join(" ")).toMatch(/Content_Types/i);
  }, 30_000);

  it("catches a config that is not valid JSON", async () => {
    const bytes = await mutate(await deck(), "ppt/tags/tag1.xml", (s) => s.replace(/val="[^"]*"/, `val="{not json"`));
    const faults = faultsIn(await readDeckBytes(bytes));
    expect(faults.join(" ")).toMatch(/not valid JSON/i);
  }, 30_000);

  it("catches a config with no object to load it from", async () => {
    // A tag hanging off nothing is unreachable from the pane: the chart looks
    // editable in the file and cannot be opened.
    const bytes = await mutate(await deck(), "ppt/slides/slide1.xml", (s) =>
      s.replace(/name="PowerChart"/, `name="something-else"`),
    );
    const faults = faultsIn(await readDeckBytes(bytes));
    expect(faults.join(" ")).toMatch(/no "PowerChart" object/i);
  }, 30_000);

  /**
   * The check that could not see itself failing.
   *
   * `config` was read from the slide's relationship part, and shape-level and
   * slide-level tags are declared in the same file — so "the tag is on the
   * chart's `p:grpSp`" (re-editable) and "the tag part exists and nothing in the
   * shape tree points at it" (dead) were the same measurement. The only guard,
   * `config && !chartObject`, is satisfied by an element named PowerChart
   * existing anywhere at top level, which it still does here.
   *
   * So a regression that put the `r:id` on the slide's own `custDataLst`
   * instead of the group's would have reported "re-editable", exited 0, kept CI
   * green, and left every generated chart dead to the pane. `ooxml.ts` has a
   * branch that mints the rid, writes the part, adds the relationship and then
   * consumes it only `if (item.group !== false)`, so the case is reachable.
   *
   * This matters beyond the tool: the last real-host deck was cleared on this
   * tool's "31 re-editable" line.
   */
  it("catches a config tag the chart object does not point at", async () => {
    // Move the tag reference off the group and onto the slide itself — the same
    // rels entry, the same tag part, a different anchor.
    const bytes = await mutate(await deck(), "ppt/slides/slide1.xml", (s) => {
      const ref = /<p:custDataLst><p:tags r:id="([^"]+)"\/><\/p:custDataLst>/.exec(s);
      expect(ref, "the deck under test carries no shape-level tag to move").toBeTruthy();
      return s.replace(ref![0], "").replace("</p:spTree>", `</p:spTree>${ref![0]}`);
    });
    const audit = await readDeckBytes(bytes);
    // The old reading still says yes, which is the point: nothing about the
    // relationship part changed.
    expect(audit.rows[0].config, "the fixture did not keep the tag part reachable").toBe(true);
    expect(audit.rows[0].chartObject, "the fixture removed the chart object too").toBe(true);
    expect(audit.rows[0].configOnChart, "counted an orphaned config as re-editable").toBe(false);
    expect(faultsIn(audit).join(" ")).toMatch(/not referenced by the "PowerChart" object/i);
  }, 30_000);

  it("still calls a healthy deck's configs re-editable", async () => {
    // The negative control. A check that answers "orphaned" for everything
    // would pass the case above and be worthless.
    const audit = await readDeckBytes(await deck());
    expect(audit.rows.map((r: { configOnChart: boolean }) => r.configOnChart)).toEqual([true, true, false]);
    expect(faultsIn(audit)).toEqual([]);
  }, 30_000);

  it("catches two slides claiming the same run and slot", async () => {
    const bytes = await mutateWhere(await deck(), /&quot;i&quot;:1/, (s) =>
      s.replace(/&quot;i&quot;:1/, "&quot;i&quot;:0"),
    );
    const faults = faultsIn(await readDeckBytes(bytes));
    expect(faults.join(" ")).toMatch(/same run and slot/i);
  }, 30_000);

  it("does NOT call two different runs in one deck a fault", async () => {
    // Inserting the demo deck twice is the case the run token exists to
    // survive. Flagging it would train the reader to ignore the tool.
    const bytes = await mutateWhere(await deck(), /&quot;i&quot;:1/, (s) => s.replace(/run-x/, "run-y"));
    // Non-vacuity: the deck really does now hold two runs.
    const audit = await readDeckBytes(bytes);
    expect(new Set(audit.rows.map((r: { run: string }) => r.run)).size).toBe(2);
    expect(faultsIn(audit)).toEqual([]);
  }, 30_000);

  it("does NOT call a missing chart a fault — that is a bad run, not a bad file", async () => {
    const zip = await JSZip.loadAsync(await deck());
    zip.remove("ppt/slides/slide3.xml");
    zip.remove("ppt/slides/_rels/slide3.xml.rels");
    const audit = await readDeckBytes(await zip.generateAsync({ type: "uint8array" }));
    expect(audit.rows).toHaveLength(2);
    // The orphaned slot tag part is a genuine fault; the absent chart is not.
    expect(faultsIn(audit).every((f: string) => /not referenced/.test(f))).toBe(true);
  }, 30_000);
});

describe("counting what is inside the chart object", () => {
  it("reports the group's children, not just the slide's top-level shapes", async () => {
    // `shapes` counts a slide's direct children, where a 40-shape chart and a
    // 1-shape degraded picture both read as 1 — which is the number a reader
    // most often wants and the one this could not previously give. The fake
    // host builds its slides from it, so a wrong count there makes every
    // generated chart read back as wreckage.
    const audited = await readDeckBytes(await deck());
    const chart = audited.rows.find((r: { slot: number | null }) => r.slot === 0)!;
    expect(chart.shapes).toBe(1); // one group at the top level…
    expect(chart.chartShapes).toBeGreaterThan(1); // …holding the whole chart
  });

  it("counts a degraded picture as the one shape it is", async () => {
    // No container to discount: the picture IS the chart object.
    const built = await buildDeckBase64([
      { scene: { width: 100, height: 100, nodes: [] }, title: "flat", slot: 0, run: "run-x" },
    ]);
    const audited = await readDeckBytes(Uint8Array.from(atob(built.base64), (c) => c.charCodeAt(0)));
    expect(audited.rows[0].chartShapes).toBeLessThanOrEqual(1);
  });
});

/**
 * Files this repo writes that PowerPoint would refuse, or that its own repair
 * paths could not find.
 */
describe("what the writer must never emit", () => {
  it("strips XML-illegal characters from a tag part, not just from the slide", () => {
    // `xmlAttr` escaped `& < > " '` and nothing else, and it is the only writer
    // of `ppt/tags/tagN.xml` — where the whole serialized ChartConfig lands as
    // an attribute value. `JSON.stringify` escapes C0 controls and lone
    // surrogates but NOT U+FFFE/U+FFFF, so those reached the part raw: the slide
    // XML came out clean, the tags part was not well-formed, and PowerPoint
    // offered to repair a file the renderer had called a success. The two
    // sibling sinks have run `xmlText` first for exactly this reason.
    const hostile = `ok${String.fromCharCode(0xffff)}${String.fromCharCode(0xfffe)}`;
    const part = tagPart([["POWERCHART_CONFIG", JSON.stringify({ title: hostile })]]);
    for (const code of [0xffff, 0xfffe]) {
      expect(part.includes(String.fromCharCode(code)), `U+${code.toString(16)} survived into a tag part`).toBe(false);
    }
    // And the ordinary escaping still happens.
    expect(tagPart([["k", `a"b&c<d`]])).toContain("&quot;");
  });

  it("names the shape it tags when a chart is too small to group", () => {
    // "PowerChart" is the single string every downstream consumer keys on —
    // `retagSlideChart`, `rescueGroupAndTag`, `slideHoldsOnlyChart`,
    // `snapshotAddedSlides` and this file's own chart-object test. The
    // fewer-than-two-children fallback hung the config on pptxgenjs's default
    // `Text 0`, so an ordinary "no data yet" chart shipped re-editable in the
    // file and invisible to every path that looks for one.
    const slide =
      `<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
      `<p:grpSpPr/></p:grpSpPr>` +
      `<p:sp><p:nvSpPr><p:cNvPr id="2" name="Text 0"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr></p:sp></p:spTree>`;
    const out = tagFirstShape(slide, "rId9");
    expect(out).toContain('name="PowerChart"');
    expect(out, "the shape TREE was renamed instead of the shape").toContain('<p:cNvPr id="1" name=""/>');
    expect(out).toContain('<p:tags r:id="rId9"/>');
  });
});
