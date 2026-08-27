import { describe, expect, it } from "vitest";
import { readFileSync, writeFileSync, mkdtempSync } from "fs";
import { tmpdir } from "os";
import { join } from "path";
import JSZip from "jszip";
// @ts-expect-error — plain .mjs tool, no types; the gate is deliberately
// independent of src/ so it cannot inherit a bug from the code it validates.
import { unexpectedFindings } from "../scripts/validate-ooxml.mjs";

/** One schema finding, as the validator reports it. */
interface Finding {
  id: string;
  path: string;
  xPath?: string;
  description: string;
  errorType: string;
}

/**
 * Schema validation of the generated deck, and the baseline it needs.
 *
 * `verify-deck.mjs` asks "did SSF Charts write what it meant to" by scanning
 * for elements it knows about. This asks "is this a legal .pptx at all",
 * against the grammar. The two catch nearly disjoint sets, which is the whole
 * reason for having both — and the reason neither can be dropped in favour of
 * the other.
 *
 * The interesting case is the baseline. One finding on the committed deck is
 * deliberate and upstream, so the gate must accept exactly that one and
 * nothing else. A baseline that swallowed a second finding by accident would
 * be indistinguishable, from CI's point of view, from no gate at all.
 */
const DECK = "examples/showcase.pptx";

const tmp = () => mkdtempSync(join(tmpdir(), "pc-ooxml-"));

async function findings(file: string): Promise<{ all: Finding[]; unexpected: Finding[] }> {
  return unexpectedFindings(file) as Promise<{ all: Finding[]; unexpected: Finding[] }>;
}

async function corrupt(edit: (zip: JSZip) => void | Promise<void>): Promise<string> {
  const zip = await JSZip.loadAsync(readFileSync(DECK));
  await edit(zip);
  const out = join(tmp(), "bad.pptx");
  writeFileSync(out, await zip.generateAsync({ type: "nodebuffer" }));
  return out;
}

describe("the OOXML schema gate", () => {
  it("passes the committed deck, with exactly one baselined finding", async () => {
    const { all, unexpected } = await findings(DECK);
    expect(unexpected, `unbaselined: ${JSON.stringify(unexpected)}`).toHaveLength(0);
    // Pinned at one. If the deck ever validates clean the baseline entry is
    // dead and should go; if it grows a second finding, that is a real defect
    // wearing the first one's clothes.
    expect(all).toHaveLength(1);
    expect(all[0].id).toBe("Sch_UnexpectedElementContentExpectingComplex");
    expect(all[0].path).toBe("/ppt/presentation.xml");
  }, 60_000);

  it("catches a dangling relationship id", async () => {
    // The class `verify-deck` structurally cannot see: it scans for elements
    // it knows about, and a broken r:embed is a reference, not an element.
    const bad = await corrupt(async (zip) => {
      const name = "ppt/slides/slide122.xml";
      const xml = await zip.file(name)!.async("string");
      zip.file(name, xml.replace(/r:embed="[^"]*"/, 'r:embed="rId999"'));
    });
    const { unexpected } = await findings(bad);
    expect(unexpected.length).toBeGreaterThan(0);
    expect(unexpected.map((e) => e.id).join(" ")).toMatch(/RelationshipId/i);
  }, 60_000);

  it("catches malformed XML in a slide part", async () => {
    const bad = await corrupt(async (zip) => {
      const name = "ppt/slides/slide1.xml";
      const xml = await zip.file(name)!.async("string");
      zip.file(name, xml.replace("</p:sld>", "<p:oops></p:sld>"));
    });
    const { unexpected } = await findings(bad);
    expect(unexpected.length).toBeGreaterThan(0);
  }, 60_000);

  it("does not baseline away a SECOND finding in the same part", async () => {
    // The baseline matches on path + id + xPath, never on the description —
    // so it must not act as a blanket pardon for /ppt/presentation.xml. This
    // corrupts that exact part a different way and insists the gate still
    // fires. Without it, the one accepted finding would quietly cover every
    // future defect in the file it lives in.
    const bad = await corrupt(async (zip) => {
      const name = "ppt/presentation.xml";
      const xml = await zip.file(name)!.async("string");
      zip.file(name, xml.replace(/<p:sldSz[^>]*\/>/, '<p:sldSz cx="-1" cy="-1"/>'));
    });
    const { all, unexpected } = await findings(bad);
    expect(all.length).toBeGreaterThan(1);
    expect(unexpected.length, "a second finding in the baselined part was swallowed").toBeGreaterThan(0);
  }, 60_000);
});
