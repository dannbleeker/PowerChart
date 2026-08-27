#!/usr/bin/env node
/**
 * Validate a generated .pptx against the Open XML schema, and fail on anything
 * that is not already known and accepted.
 *
 *   node scripts/validate-ooxml.mjs <deck.pptx> [more.pptx …] [--json]
 *
 * This is a complement to `verify-deck.mjs`, not an overlap with it. That tool
 * asks "did SSF Charts write what it meant to" — slot tags, groups, config
 * parts — by scanning for the elements it knows about. This one asks "is this
 * a legal .pptx at all", against the grammar, and the two catch almost
 * disjoint sets. Measured on `examples/showcase.pptx` by injecting defects:
 * the validator catches dangling relationship ids, malformed XML, duplicate
 * attributes, negative extents and wrong child ordering, none of which a
 * regex scan can see; `verify-deck` catches unreferenced tag parts and missing
 * content-type overrides, which the schema does not describe.
 *
 * Neither catches a duplicate `cNvPr` id inside one `spTree` — the classic
 * trigger for PowerPoint's repair dialog. `verify-deck.mjs` now does; see
 * `duplicateShapeIds` there.
 *
 * ## Why there is a baseline at all
 *
 * Schema-valid and PowerPoint-valid are different sets, and for one element
 * they are in direct conflict. `CT_Presentation` requires `notesMasterIdLst`
 * before `sldIdLst`; pptxgenjs deliberately emits it after, with three
 * comments in its source explaining why:
 *
 *     // IMPORTANT: In this order (matches PPT2019) PPT will give corruption
 *     //   message on open!
 *     // IMPORTANT: Placing this before `<p:sldIdLst>` causes warning in
 *     //   modern powerpoint!
 *
 * So the conformant ordering is the one PowerPoint complains about. That
 * finding is accepted here rather than fixed, because fixing it would break
 * the decks in the host we ship to. Everything else is a failure.
 *
 * A baseline entry matches on `path` + `id` + `xPath` — never on the
 * description, which is long, and whose wording is the validator's to change.
 */
import { validateFile } from "@xarsh/ooxml-validator";
import { isMain } from "./is-main.mjs";

/**
 * Findings that are known, understood, and deliberately not fixed.
 *
 * Each needs a reason, and the reason has to be about the FILE rather than
 * about the inconvenience of fixing it. An entry with no explanation is a
 * silenced test.
 */
const BASELINE = [
  {
    path: "/ppt/presentation.xml",
    xPath: "/p:presentation[1]",
    id: "Sch_UnexpectedElementContentExpectingComplex",
    why: "pptxgenjs emits <p:notesMasterIdLst> after <p:sldIdLst> on purpose — the schema-conformant order makes PowerPoint report the file as corrupt (see its source, three IMPORTANT comments above the emit). Upstream, deliberate, and not ours to fix.",
  },
];

const known = (e) => BASELINE.find((b) => b.path === e.path && b.xPath === e.xPath && b.id === e.id);

/** Validate one file. Returns the findings that are NOT baselined. */
export async function unexpectedFindings(file) {
  const result = await validateFile(file);
  const errors = result?.errors ?? [];
  return { all: errors, unexpected: errors.filter((e) => !known(e)) };
}

async function main(argv) {
  const json = argv.includes("--json");
  const files = argv.filter((a) => !a.startsWith("--"));
  if (!files.length) {
    console.error("usage: node scripts/validate-ooxml.mjs <deck.pptx> [more.pptx …] [--json]");
    return 2;
  }
  const report = [];
  let bad = 0;
  for (const file of files) {
    let found;
    try {
      found = await unexpectedFindings(file);
    } catch (err) {
      // A file the validator cannot even open is the worst outcome, not an
      // excuse to pass: that is what PowerPoint's repair dialog looks like
      // from the outside.
      console.error(`  ${file}: could not be validated — ${err.message}`);
      bad++;
      continue;
    }
    report.push({ file, ...found });
    const baselined = found.all.length - found.unexpected.length;
    if (found.unexpected.length) {
      bad++;
      console.error(`\n  ${file}: ${found.unexpected.length} schema finding(s) that are not baselined`);
      for (const e of found.unexpected) {
        console.error(`    ${e.errorType}  ${e.id}`);
        console.error(`      at ${e.path} ${e.xPath ?? ""}`);
        console.error(`      ${e.description}`);
      }
    } else {
      console.log(`  ${file}: schema-valid${baselined ? ` (${baselined} known finding(s) baselined)` : ""}`);
    }
  }
  if (json) console.log(JSON.stringify(report, null, 2));
  return bad ? 1 : 0;
}

// The FOURTH spelling of this guard, and the fourth one wrong on Windows:
// `file://${process.argv[1]}` builds `file://C:\repo\scripts\validate-ooxml.mjs`
// against an `import.meta.url` of `file:///C:/repo/scripts/validate-ooxml.mjs`.
// Never equal — so the OOXML grammar gate printed nothing and exited 0 on the
// owner's box, a clean pass from a check that never opened the file. Its three
// siblings were fixed hours earlier and a grep for THEIR wording did not reach
// this one, which is why `test/is-main.test.ts` now checks the shape of every
// guard in the repo rather than any single spelling of it.
if (isMain(import.meta.url, process.argv[1])) {
  main(process.argv.slice(2)).then((code) => {
    process.exitCode = code;
  });
}
