#!/usr/bin/env node
/**
 * Diff a real PowerPoint's answer sheet against the fake host's.
 *
 * This is the payoff for `src/render/host-probe.ts`. Every Office.js assertion
 * in this repo runs against a fake, and the fake has never been checked against
 * the thing it stands for. Its faults are behaviours a real host taught us; its
 * happy path is a set of things we assumed. Where an assumption is wrong, every
 * test resting on it is confidently wrong with it, and a green suite says
 * nothing about that.
 *
 *   node scripts/host-diff.mjs <real-host-answers.json> [fake-answers.json]
 *
 * The fake's sheet is the one CI freezes in `test/host-probe.test.ts`; pass it
 * explicitly, or let this read the baseline committed beside it.
 *
 * Only the `answer` field is compared. Timings and error text differ between
 * any two runs of the same host and would bury the signal.
 *
 * Exit 0 when the two agree, 1 when they do not, 2 when a file cannot be read —
 * so it can gate as well as inform.
 */
import { readFileSync } from "fs";

/** The fake's frozen answers, as asserted in `test/host-probe.test.ts`. */
export const FAKE_BASELINE = {
  "load-isnullobject-populates": "unreadable",
  "load-id-populates-isnullobject": "yes",
  "getitemornullobject-missing": "null-object",
  "shape-proxy-survives-one-sync": "yes",
  "shapes-items-count-honest": "at-least-5",
  "getcount-populates-same-sync": "yes",
  "tags-add-same-key-twice": "overwrites",
  "tags-on-fresh-shape": "yes",
  "delete-then-lookup": "reports-gone",
  "addgroup-returns-usable": "yes",
  "group-reports-its-children": "two",
  "tag-on-group-survives": "yes",
  "getitemat-past-end": "threw",
  "untrack-available": "no",
};

/**
 * What each divergence would MEAN, so the report says something actionable
 * rather than just "these differ".
 *
 * Keyed by probe id. Absent is fine — the diff still reports the divergence,
 * it just cannot say what rests on it.
 */
const WHAT_IT_MEANS = {
  "load-isnullobject-populates":
    "`queueNullCheck` loads 'id' instead of 'isNullObject' precisely because the flag cannot be loaded by name. If a real host populates it, that whole comment is wrong for this host — and the workaround is merely harmless rather than necessary. ANSWERED: PowerPoint on the web (2026-08-04) said yes, and read the flag back as false. The negative is host-specific; the workaround stays because the host it was written for is real too.",
  "load-id-populates-isnullobject":
    "If a real host does NOT populate the flag from a real property load, `queueNullCheck` does not work and every `isLive` check is answering 'not live' for live objects. `isLive` treats unreadable as NOT live, so the failure mode is refusing to act on slides that are fine.",
  "shape-proxy-survives-one-sync":
    "office-js#2903. The fake keeps proxies alive by default, which is the kindness that hid a whole class of stale-proxy bug until a human found it in a real host. If a real host refuses a one-sync-old proxy, `applyWebProfile` should be the default rather than a named profile.",
  "shapes-items-count-honest":
    "`faults.hollowReads` models a host answering SHORT without throwing — a readback asked about 19 shapes and was told 3. If a real host is honest, the readback paging and the re-read are more caution than the platform needs.",
  "tags-add-same-key-twice":
    "Re-editing a chart rewrites POWERCHART_CONFIG on the same shape every time. If a host appends rather than overwrites, a chart edited ten times carries ten configs and the reader picks one arbitrarily — silently editing the wrong data.",
  "tags-on-fresh-shape":
    "`faults.tagsUndefinedOn` models `.tags` coming back undefined, where reading `.add` throws SYNCHRONOUSLY and escapes the tagging loop — losing the config for every chart after it in the batch, not just the one.",
  "delete-then-lookup":
    "`deleteSlideById` re-checks from a FRESH context because the same-context answer was not trusted. If a host answers honestly here, that second round trip is removable.",
  "group-reports-its-children":
    "The single most load-bearing answer here. A chart IS a group, and the readback measures whether a chart survived by counting what is inside it — so a host that groups successfully and then reports no children makes every chart read back as wreckage, and the repair pass 'fixes' charts that were never broken.",
  "tag-on-group-survives":
    "Where a chart's config actually lives. Tags on a plain shape are a separate question; if a GROUP behaves differently, every chart in every deck is un-re-editable and nothing else here would say so.",
  "getitemat-past-end":
    "Nothing in this repo currently depends on the answer — it is here to find out before something does.",
  "untrack-available":
    "The fake does not implement `untrack`, so every `untrack()` call in `powerpoint.ts` is a no-op under test and the proxy-release path is entirely unexercised. A real host saying 'yes' does not fix that; it means the path is real and still untested.",
};

/** Read a sheet, whichever shape it arrived in. */
export function answersOf(sheet) {
  if (sheet?.kind === "powerchart-host-answers" && Array.isArray(sheet.answers)) {
    return Object.fromEntries(sheet.answers.map((a) => [a.id, a.answer]));
  }
  // A bare map, e.g. the committed baseline.
  if (sheet && typeof sheet === "object" && !Array.isArray(sheet)) return sheet;
  return null;
}

/**
 * Compare two answer sheets.
 *
 * `onlyReal` and `onlyFake` matter as much as the disagreements: a question one
 * side has never been asked is a gap in the comparison, not a match, and
 * silently treating it as agreement is how a diff comes to mean nothing.
 */
export function diffAnswers(real, fake) {
  const ids = [...new Set([...Object.keys(real), ...Object.keys(fake)])].sort();
  const agree = [];
  const differ = [];
  const onlyReal = [];
  const onlyFake = [];
  for (const id of ids) {
    const r = real[id];
    const f = fake[id];
    if (r === undefined) onlyFake.push(id);
    else if (f === undefined) onlyReal.push(id);
    else if (r === f) agree.push(id);
    else differ.push({ id, real: r, fake: f, means: WHAT_IT_MEANS[id] });
  }
  return { agree, differ, onlyReal, onlyFake };
}

const invokedDirectly = process.argv[1] && process.argv[1].endsWith("host-diff.mjs");
if (invokedDirectly) {
  const [realPath, fakePath] = process.argv.slice(2).filter((a) => !a.startsWith("--"));
  if (!realPath) {
    console.error("usage: node scripts/host-diff.mjs <real-host-answers.json> [fake-answers.json]");
    process.exit(2);
  }
  let realSheet, fakeSheet;
  try {
    realSheet = JSON.parse(readFileSync(realPath, "utf8"));
    fakeSheet = fakePath ? JSON.parse(readFileSync(fakePath, "utf8")) : FAKE_BASELINE;
  } catch (err) {
    console.error(`could not read the answer sheets: ${err.message}`);
    process.exit(2);
  }
  const real = answersOf(realSheet);
  const fake = answersOf(fakeSheet);
  if (!real || !fake) {
    console.error("that file is not an answer sheet (expected kind: powerchart-host-answers)");
    process.exit(2);
  }

  const { agree, differ, onlyReal, onlyFake } = diffAnswers(real, fake);
  const sets = Array.isArray(realSheet.requirementSets) ? realSheet.requirementSets.join(", ") : "unknown";
  console.log(
    `\n  REAL HOST ${realSheet.source ?? "?"} · build ${realSheet.build ?? "?"}` +
      `\n  requirement sets: ${sets}` +
      `\n  ${agree.length} of ${agree.length + differ.length} answers match the fake\n`,
  );

  if (differ.length) {
    console.log(`  ${differ.length} DIVERGENCE(S) — the fake and the real host disagree:\n`);
    for (const d of differ) {
      console.log(`    ${d.id}`);
      console.log(`      real host: ${d.real}`);
      console.log(`      fake:      ${d.fake}`);
      if (d.means) console.log(`      what rests on it: ${d.means}`);
      console.log("");
    }
    console.log("  Each of these is one of two things: the fake lies and the tests built on it");
    console.log("  are worth less than they look, or the host does something we did not know.\n");
  }
  // A question one side was never asked is a hole in the comparison, and
  // reporting it as agreement is how a diff stops meaning anything.
  if (onlyReal.length)
    console.log(`  ${onlyReal.length} question(s) the fake has no answer for: ${onlyReal.join(", ")}`);
  if (onlyFake.length)
    console.log(`  ${onlyFake.length} question(s) the real sheet is missing (older build?): ${onlyFake.join(", ")}`);
  if (!differ.length && !onlyReal.length && !onlyFake.length) {
    console.log("  The fake agrees with this host on every question it was asked.\n");
  }
  process.exit(differ.length ? 1 : 0);
}
