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
 * Exit 0 when the two agree, 1 when they do not — or when the real host never
 * got as far as answering — and 2 when a file cannot be read, so it can gate as
 * well as inform.
 */
import { readFileSync } from "fs";

/** The fake's frozen answers, as asserted in `test/host-probe.test.ts`. */
export const FAKE_BASELINE = {
  "load-isnullobject-populates": "unreadable",
  "load-id-populates-isnullobject": "yes",
  "getitemornullobject-missing": "null-object",
  "shape-add-fresh-slide-proxy": "yes",
  "shape-add-held-slide-proxy": "yes",
  "shape-add-positional-slide-proxy": "yes",
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
  "shape-add-fresh-slide-proxy":
    "Whether this host will take a shape at all on a slide added moments ago, asked through a slide proxy resolved in the same sync as the add. A 'no' here is the end of the scratch-slide probe design, and would say something much larger about drawing onto new slides.",
  "shape-add-held-slide-proxy":
    "The same add through a slide proxy resolved one sync earlier — what Office.js has by then rewritten to `slides.getItem(id)`, and what a freshly-added slide's id does not round-trip through on the web. Expected to be the one that fails. If it does, `getItemOrNullObject` handles on new slides are single-sync objects everywhere in this file, exactly as `SlideThunk` already says of `getItemAt`.",
  "shape-add-positional-slide-proxy":
    "The third way to name the same slide. If by-index works where by-id fails, the ID is what this host will not take for a new slide, and a write path that holds one is fixed by counting rather than by re-resolving.",
  "shape-proxy-survives-one-sync":
    "office-js#2903. The fake keeps proxies alive by default, which is the kindness that hid a whole class of stale-proxy bug until a human found it in a real host. If a real host refuses a one-sync-old proxy, `applyWebProfile` should be the default rather than a named profile. ANSWERED, sideways: the 2026-08-04 self-test run threw `InvalidParam passed to GetItem(id)` at `ShapeCollection.getItem` while grouping a chart's shapes, five charts in a row — so on that host the answer is no. The probe's own attempt never reached the question.",
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
 * Answers that mean the question was never put, so they are never divergences.
 *
 * The probe's own vocabulary, kept in step with `NOT_ASKED` in
 * `src/render/host-probe.ts`. No probe can produce either word as an answer,
 * which is what makes them safe to read this way.
 */
const NEVER_ASKED = new Set(["no-scratch-slide", "no-scratch-shape"]);

/**
 * Compare two answer sheets.
 *
 * Three ways a question can fail to be a match, and all three used to be one:
 *
 * - `onlyReal` / `onlyFake` — one side was never asked. A gap, not agreement.
 * - `notAsked` — the real host never got far enough to answer. The probe run
 *   says so in a word no probe can produce, because it learned the hard way
 *   what happens otherwise: PowerPoint on the web refused eight setups on
 *   2026-08-04 and this tool reported eight host divergences from questions
 *   nobody had asked.
 * - `differ` — the only one that is actually a finding.
 */
export function diffAnswers(real, fake) {
  const ids = [...new Set([...Object.keys(real), ...Object.keys(fake)])].sort();
  const agree = [];
  const differ = [];
  const notAsked = [];
  const onlyReal = [];
  const onlyFake = [];
  for (const id of ids) {
    const r = real[id];
    const f = fake[id];
    if (r === undefined) onlyFake.push(id);
    else if (f === undefined) onlyReal.push(id);
    else if (NEVER_ASKED.has(r)) notAsked.push({ id, why: r });
    else if (r === f) agree.push(id);
    else differ.push({ id, real: r, fake: f, means: WHAT_IT_MEANS[id] });
  }
  return { agree, differ, notAsked, onlyReal, onlyFake };
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

  const { agree, differ, notAsked, onlyReal, onlyFake } = diffAnswers(real, fake);
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
  if (notAsked.length) {
    console.log(`  ${notAsked.length} QUESTION(S) NEVER PUT — the host would not set the probe up:\n`);
    for (const n of notAsked) console.log(`    ${n.id}  (${n.why})`);
    console.log("\n  Not divergences. Nothing is known about what this host does here —");
    console.log("  the run could not get as far as asking. Fix the setup, run it again.\n");
  }
  // A question one side was never asked is a hole in the comparison, and
  // reporting it as agreement is how a diff stops meaning anything.
  if (onlyReal.length)
    console.log(`  ${onlyReal.length} question(s) the fake has no answer for: ${onlyReal.join(", ")}`);
  if (onlyFake.length)
    console.log(`  ${onlyFake.length} question(s) the real sheet is missing (older build?): ${onlyFake.join(", ")}`);
  if (!differ.length && !notAsked.length && !onlyReal.length && !onlyFake.length) {
    console.log("  The fake agrees with this host on every question it was asked.\n");
  }
  // A question that was never put is not agreement, so it fails the gate too —
  // an incomplete sheet exiting 0 is exactly the false all-clear this tool
  // spent a round learning not to give.
  process.exit(differ.length || notAsked.length ? 1 : 0);
}
