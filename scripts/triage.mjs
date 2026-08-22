#!/usr/bin/env node
/**
 * Read a real run's .pptx and its run log together, and say where they differ.
 *
 * These two files are the whole evidence base for every host bug this project
 * has fixed, and joining them has always been done by hand: unzip the deck,
 * pretty-print 160 KB of JSON, line the slots up, squint. That is a slow job
 * and — twice now — a wrong one. One analysis read a truncated log and blamed
 * the wrong subsystem; another read a trace buffer that spanned three runs and
 * found a contradiction that was not there. Both cost a full round trip: a
 * deploy, a run in a real PowerPoint, an upload, and a re-read.
 *
 * The join itself is mechanical, so it belongs in a script. Slot by slot:
 * what the run believed it did, what the file actually holds, and the verdict
 * when those disagree.
 *
 *   node scripts/triage.mjs <deck.pptx> <run-log.json> [--all] [--json]
 *
 * Arguments may be given in either order — the extension says which is which.
 *
 * Two files, two authorities, and they are NOT equal: the deck is what the
 * user has, the log is what the add-in thought. Where they conflict the deck
 * wins, and the conflict is the finding. (A 39-slide run reported 20 tagged
 * charts; the file carried 31. The log agreed with itself and was false.)
 *
 * Exit code 0 when the two agree, 1 when they do not, 2 when a file cannot be
 * read — so it can gate as well as inform.
 */
import { readFileSync, readdirSync } from "fs";
import { join } from "path";
import { readDeck, faultsIn } from "./verify-deck.mjs";

/**
 * Verdicts, worst first — the order they are printed and counted in.
 *
 * `repaired` and `foreign` are not faults: the first is the end-of-run repair
 * pass doing its job, the second is an earlier run's slide sitting in the same
 * deck, which is the case the run token exists to survive.
 */
const BAD = new Set(["lost", "duplicated", "tag-lost", "not-editable", "no-config", "blank", "orphan"]);

/**
 * What the deck holds for one slot, against what the run believed.
 *
 * `item.chart` — was this item MEANT to carry a config — is load-bearing here,
 * and its absence is why the first version of this file called seven perfectly
 * healthy slides broken. Not every demo item is a chart: the title page, the
 * contents pages and several elements are drawn as `PowerChart` objects with
 * no config by design, and a rule of "a chart object with no config is not
 * re-editable" flags every one of them. What the run *intended* cannot be
 * recovered from the file, and on the shape path it cannot be recovered from
 * `tagged` either — false there means both "never had one" and "the write was
 * lost". So the log states it, and a log too old to state it gets the weaker
 * `no-config` verdict rather than a confident wrong one.
 */
function verdictFor(item, rows) {
  if (rows.length === 0) return item.status === "skipped" ? "skipped" : "lost";
  if (rows.length > 1) return "duplicated";
  const row = rows[0];
  if (row.shapes === 0) return "blank";
  // `configOnChart`, NOT `config`. `verify-deck.mjs` draws that distinction on
  // purpose and says why on the field itself: `config` only reports that the
  // tag part EXISTS, which is equally true of a tag nothing in the shape tree
  // points at. The pane loads a chart by walking from the PowerChart object to
  // its tag rid, so a config anchored on the SLIDE is a config the user cannot
  // reach — the chart is not re-editable, which is the whole thing this report
  // is about.
  //
  // triage was never migrated when that field was added, so for the exact case
  // `verify-deck` was extended to catch it printed `ok` and counted zero
  // disagreements, while verify-deck's own report on the same deck said
  // `orphan`. `--json` consumers saw `verdict: "ok"`.
  //
  // Falls back to `config` for a deck read by an older `verify-deck` that did
  // not report the anchor: unknown must not become a fault, and the two agree
  // wherever the anchor is known.
  const reachable = row.configOnChart ?? row.config;
  if (reachable) {
    // The repair pass got there after the run gave up on the tag — a success,
    // and one worth seeing, because it is the difference between a fix working
    // and a fix never having been needed.
    return item.tagged ? "ok" : "repaired";
  }
  // The run says it wrote a config tag and the file has none. On the file path
  // this is impossible by construction (the tag is written into the .pptx
  // before the host ever sees it), so it means the deck lost it afterwards; on
  // the shape path it means the write was reported as landing and did not.
  if (item.tagged) return "tag-lost";
  if (!row.chartObject) return "ok";
  if (item.chart === true) return "not-editable";
  if (item.chart === false) return "ok"; // never had a config to lose
  return "no-config"; // a log from before `chart` existed — intent unknown
}

/** A one-line description of what the file holds for a slot. */
function describeRow(row) {
  if (!row) return "—";
  const kind = row.picture ? "picture" : row.groups ? "group" : row.chartObject ? "shape" : "loose";
  // Names the orphan case rather than calling it `config`: a tag the pane
  // cannot reach is not the same fact as a tag it can, and the deck column is
  // where a reader looks to see which.
  const cfg = row.configOrphaned ? "config-ORPHANED" : (row.configOnChart ?? row.config) ? "config" : "no-config";
  return `${kind} ${row.shapes}sh ${cfg}${row.stamped ? " NOT-COMPLETE" : ""}`;
}

export function triage(deck, log) {
  // Logs written before the run token existed have no join key of their own.
  // The deck's commonest token is the best available stand-in and is very
  // probably right, but it IS a guess, so it gets labelled as one rather than
  // quietly presented as fact.
  const tally = new Map();
  for (const r of deck.rows) if (r.run) tally.set(r.run, (tally.get(r.run) ?? 0) + 1);
  const commonest = [...tally.entries()].sort((a, b) => b[1] - a[1])[0]?.[0] ?? null;
  const run = log.run ?? commonest;
  const inferredRun = !log.run;

  const mine = new Map();
  const foreign = [];
  const unowned = [];
  for (const row of deck.rows) {
    if (row.run && row.run !== run) foreign.push(row);
    else if (row.slot === null) unowned.push(row);
    else {
      if (!mine.has(row.slot)) mine.set(row.slot, []);
      mine.get(row.slot).push(row);
    }
  }

  const items = log.items ?? [];
  const slots = items.map((item, i) => {
    const rows = mine.get(i) ?? [];
    return {
      slot: i,
      title: item.title,
      status: item.status,
      wroteTag: !!item.tagged,
      deck: describeRow(rows[0]),
      rows: rows.length,
      verdict: verdictFor(item, rows),
    };
  });
  // A slide carrying a slot this run never issued. Only possible when a tag
  // survived from a longer run into a shorter one, or when a slot tag was
  // written wrong — either way nothing owns it and nothing will clean it.
  // Outside the run's range in EITHER direction. Only `>= items.length` was
  // checked, so a slide tagged with a negative slot — a tag written wrong, or
  // carried over from something that was not this — matched no item, fell
  // outside the orphan test, and was reported by nothing at all.
  const orphans = [...mine.keys()]
    .filter((s) => !Number.isInteger(s) || s < 0 || s >= items.length)
    .map((s) => ({ slot: s, verdict: "orphan", indexes: mine.get(s).map((r) => r.index + 1) }));

  const counts = {};
  for (const s of [...slots, ...orphans]) counts[s.verdict] = (counts[s.verdict] ?? 0) + 1;
  const disagreements = slots.filter((s) => BAD.has(s.verdict)).length + orphans.length;
  return { run, inferredRun, slots, orphans, foreign, unowned, counts, disagreements };
}

/**
 * The runs inside a log file, however that file is shaped.
 *
 * One click can now take both insert paths, so a log holds a list. Logs from
 * before that — every artifact captured so far — are a single run at the top
 * level, and are read as a list of one rather than being refused: the whole
 * point of this tool is reading the runs that already exist.
 */
export function runsIn(log) {
  return Array.isArray(log?.runs) ? log.runs : [log];
}

/**
 * The host self-test's verdicts, if this log is one.
 *
 * A self-test log carries no runs at all — the scenarios are not inserts and
 * own no slots — so `runsIn` gives an empty list and the whole report used to
 * come out blank, exit 0, and look like a clean deck. It is a file someone
 * hands this tool expecting an answer.
 */
export function selfTestIn(log) {
  return Array.isArray(log?.selftest) ? log.selftest : [];
}

/**
 * A crash log — the steps a run wrote before it stopped existing.
 *
 * Not a run log and not a self-test log: it holds no slots, no verdicts and no
 * structured trace, because it is written a line at a time to browser storage
 * by a run that may be about to die. It carries the one thing the other two
 * cannot: what a run that never ENDED was doing. Both files this tool was
 * built for are produced at the end of a run, and the runs worth reading are
 * the ones that do not get there.
 */
/**
 * One banked finding in a line: a scenario verdict as its verdict, an answer
 * sheet as its count, anything else as JSON.
 *
 * Kept shallow on purpose. This runs on the file from a round that died, and the
 * reader wants to know which scenarios got a verdict and whether the probe half
 * survived — not to read a sheet pretty-printed down the terminal.
 */
export function describeFinding(value) {
  if (value && typeof value === "object" && typeof value.name === "string" && "ok" in value) {
    const state = value.skipped ? "SKIPPED" : value.ok ? "passed" : "FAILED";
    return `${state}${value.detail ? ` — ${value.detail}` : ""}${value.ms ? ` (${(value.ms / 1000).toFixed(1)}s)` : ""}`;
  }
  if (value && typeof value === "object" && Array.isArray(value.answers)) {
    const asked = value.answers.filter((a) => a && typeof a.answer === "string").length;
    return `answer sheet, ${asked} of ${value.answers.length} question(s) answered`;
  }
  return typeof value === "string" ? value : JSON.stringify(value);
}

export function crashLogIn(log) {
  return log?.kind === "powerchart-crash-log" && Array.isArray(log.steps) ? log : null;
}

/**
 * Failure signatures that belong to PowerPoint, not to PowerChart.
 *
 * The first hour of the last diagnosis went into establishing that
 * `InvalidParam passed to GetItem(id)` was a host bug rather than ours. It is,
 * it is reported, and most of these are still open — so the trace can just say
 * so. A reader who sees an issue number stops looking for the mistake in this
 * repo.
 *
 * Matched on a substring of the reason string the trace recorded, FIRST MATCH
 * WINS, so the order below is part of the mapping rather than decoration.
 * Checked 2026-08-01; re-check before trusting the status (see RESEARCH.md 4b).
 */
const SELECTION_WEDGE =
  "office-js#3083 / #3698 family (open) — a programmatic setSelectedShapes wedges the web host's " +
  "selection subsystem; every selection call after it goes silent";

const KNOWN_HOST_BUGS = [
  // MOST SPECIFIC FIRST, and this is load-bearing rather than tidy. A wedged
  // selection subsystem does not throw — it goes quiet, so what the trace
  // records is a timeout, whose text is "did not respond while <phase>". The
  // generic sync-hang note below matches that too, and on a first-match-wins
  // table it would claim every one of them and send the reader to #5022: a
  // different bug, a different cause, and a fix that does not apply. The
  // phase name is the only thing that separates them, which is one more
  // reason every bounded wait carries one.
  ["stopped answering selection calls", SELECTION_WEDGE],
  ["did not respond while reading the selected chart", SELECTION_WEDGE],
  ["did not respond while selecting a shape", SELECTION_WEDGE],
  ["did not respond while clearing the shape selection", SELECTION_WEDGE],
  ["InvalidParam passed to GetItem", "office-js#2903 (closed: not planned) — stale shape proxy on web"],
  ["not available", "office-js#6363 (open: regression) — loaded property missing after sync"],
  [
    "did not respond",
    "office-js#5022 (closed: completed 2024-11-18, unverified on this host) — sync hangs after add/delete/re-read",
  ],
  [
    "Timed out",
    "office-js#5022 (closed: completed 2024-11-18, unverified on this host) — sync hangs after add/delete/re-read",
  ],
];

/** The known-bug note for a problem string, if there is one. */
export function knownBug(text) {
  const hit = KNOWN_HOST_BUGS.find(([sig]) => text.includes(sig));
  return hit ? hit[1] : null;
}

function pad(s, n) {
  return String(s ?? "—")
    .padEnd(n)
    .slice(0, n);
}

function report(deck, log, run, t, showAll) {
  const secs = ((run.totalMs ?? 0) / 1000).toFixed(1);
  console.log(
    `\n  run ${t.run ?? "unknown"}${t.inferredRun ? " (INFERRED from the deck — the log carries no run id)" : ""}` +
      `\n  build ${log.build ?? run.build ?? "?"} · ${log.host ?? run.host ?? "?"}` +
      `\n  path ${run.path ?? "?"} · ${secs}s\n`,
  );

  console.log(
    `  DECK ${deck.rows.length} slide(s): ${deck.rows.length - t.foreign.length - t.unowned.length} from this run, ` +
      `${t.foreign.length} from other run(s), ${t.unowned.length} carrying no slot tag`,
  );

  const shown = showAll ? t.slots : t.slots.filter((s) => BAD.has(s.verdict) || s.verdict === "repaired");
  console.log(
    `\n  SLOTS ${t.slots.length} expected · ${t.slots.length - (t.counts.lost ?? 0)} present · ` +
      Object.entries(t.counts)
        .sort((a, b) => b[1] - a[1])
        .map(([k, n]) => `${n} ${k}`)
        .join(" · "),
  );
  if (shown.length) {
    console.log(`\n  ${pad("#", 5)}${pad("title", 26)}${pad("log", 11)}${pad("tag?", 6)}${pad("deck", 30)}verdict`);
    for (const s of shown)
      console.log(
        `  ${pad(s.slot, 5)}${pad(s.title, 26)}${pad(s.status, 11)}${pad(s.wroteTag ? "wrote" : "no", 6)}` +
          `${pad(s.rows > 1 ? `${s.rows} slides claim it` : s.deck, 30)}${s.verdict}`,
      );
    if (!showAll) console.log(`  (agreeing slots hidden — pass --all to see all ${t.slots.length})`);
  }
  for (const o of t.orphans) console.log(`  orphan slot ${o.slot} on slide(s) ${o.indexes.join(", ")}`);

  reportTrace(run.trace);

  console.log(
    t.disagreements
      ? `\n  ${t.disagreements} slot(s) where the deck and the log disagree\n`
      : `\n  deck and log agree on every slot\n`,
  );
}

/**
 * The trace tallies, from wherever the trace hangs.
 *
 * A run carries its own; a self-test round is not a run and hangs one at the
 * top level, which is why this is a function rather than four lines inside
 * `report`. Both shapes have the same tallies and deserve the same reading.
 */
function reportTrace(trace) {
  if (trace?.summary) {
    console.log(`\n  TRACE ${trace.entries?.length ?? 0} entries${trace.dropped ? `, ${trace.dropped} dropped` : ""}`);
    for (const s of trace.summary.steps.slice(0, 8)) console.log(`    ${pad(s.n, 6)}${s.scope}  ${s.message}`);
    // Every phase an error escaped, whether or not it made the top 8 above.
    //
    // These are the most locating lines in the whole log and the least likely
    // to be common enough to rank: a run with one fatal read has one of them
    // against hundreds of ordinary steps. The `problems` tally below carries
    // what the HOST said, truncated to keep one debugInfo blob from swamping
    // it; this carries what the add-in was DOING, which is the half that used
    // to be missing entirely and took a session to reconstruct.
    const phases = trace.summary.steps.filter(
      (s) => s.scope === "error" && !trace.summary.steps.slice(0, 8).includes(s),
    );
    if (phases.length) {
      console.log(`  phases an error escaped:`);
      for (const s of phases) console.log(`    ${pad(s.n, 6)}${s.message}`);
    }
    if (trace.summary.problems.length) {
      console.log(`  problems:`);
      for (const p of trace.summary.problems.slice(0, 8)) {
        console.log(`    ${pad(p.n, 6)}${p.text}`);
        const known = knownBug(p.text);
        if (known) console.log(`    ${pad("", 6)}  ^ known host bug: ${known}`);
      }
    }
  } else if (trace) {
    console.log(`\n  TRACE ${trace.entries?.length ?? 0} entries (written before summaries existed — no tallies)`);
  }
}

/** The self-test verdicts, with the host bugs their own words give away. */
/**
 * Pool the rasterise/cheap-read arms across MANY rounds.
 *
 * `does a rasterise poison the next draw` is a counterbalanced control that
 * makes four draws a round — two after a rasterise, two after a cheap read,
 * interleaved so position cannot account for the difference. Four is enough to
 * spot a call that fails EVERY time and far too few for anything smaller, so
 * for eleven rounds it has correctly reported "no pattern" and correctly been
 * unable to say more.
 *
 * Pooled over eight rounds the same arms read 5 stalls in 16 draws after a
 * rasterise against 2 in 16 after a cheap read — 31% against 12.5%. That is a
 * direction, and it is still not an answer: Fisher's exact on those counts is
 * p≈0.39. Separating rates that close needs somewhere near 60-100 draws an arm,
 * which is 30-50 more rounds at two an arm each.
 *
 * So the verdict a single round can reach is not the interesting one, and the
 * evidence was accumulating in files nobody was adding up. This adds them up.
 *
 * A draw's outcome is NOT its `batch issued` line: that is logged before the
 * sync, deliberately, because the sync is where a bad host goes quiet — so
 * every stall has a `batch issued` of its own immediately before it. Pairing
 * the two is how a first pass at this produced 0 stalls in 32 draws, and then
 * an unrelated cut of the same data produced a 6x effect that was not there.
 * The outcome is whether a `gave up waiting` lands between this draw's step and
 * the next one.
 *
 * Both mistakes were made while that line was NAMED `batch committed`, which is
 * why it is not called that any more — but round files saved before 2026-08-11
 * still carry the old name, and this function reads neither, by design.
 */
/**
 * The two populations a draw batch falls into, computed rather than eyeballed.
 *
 * "FAST IS THE BROKEN MODE" rests on the batch times being BIMODAL with an
 * empty band between them, and every number supporting it has been read off a
 * round by hand. That is how the band's edge came to be quoted as 29.2s and
 * stayed quoted after a later round put three surviving batches at 29.3, 30.8
 * and 31.1 — a measurement drifting in the prose because nothing recomputed it.
 *
 * The split is taken at the LARGEST gap in the sorted times, which is what a
 * reader does by eye and cannot then mis-remember. A round with no gap worth
 * the name says so rather than inventing a boundary: `bimodal` is false unless
 * the gap is bigger than the whole spread of the fast group, which is the
 * property "an empty band" actually means.
 *
 * The other number this file re-derives by hand — the per-slide cost slope —
 * is deliberately NOT computed here. `onSlide` is a NET counter that resets as
 * each chart replaces itself, so one slide's points are a repeated ramp rather
 * than a series: slide 257 of round 17 carries onSlide 58 and 68 four times
 * each, non-monotonic. Fitting a line through that produced 0.29s per shape
 * against the 0.44-0.49 the rasterise arms give, which is a number belonging to
 * neither population. A slope needs an add-only path and this trace does not
 * mark one; read the arms by hand until it does.
 */
export function batchPopulations(log) {
  const entries = log?.trace?.entries;
  if (!Array.isArray(entries)) return null;
  const times = entries
    .filter((e) => e.message === "batch issued" && typeof e.data?.prevBatchMs === "number")
    .map((e) => e.data.prevBatchMs)
    .sort((a, b) => a - b);
  if (times.length < 4) return null;
  let at = 0;
  let gap = 0;
  for (let i = 0; i < times.length - 1; i++)
    if (times[i + 1] - times[i] > gap) {
      gap = times[i + 1] - times[i];
      at = i;
    }
  const fast = times.slice(0, at + 1);
  const slow = times.slice(at + 1);
  const spread = fast[fast.length - 1] - fast[0];
  // Both groups need at least two members. On evenly spread times every gap is
  // equal, the argmax picks the FIRST, and a one-member fast group has a spread
  // of zero — which any gap beats, so the rule would answer "bimodal" for the
  // most uniform data there is. Found by the negative test, not by reasoning.
  const bimodal = fast.length >= 2 && slow.length >= 2 && gap > spread;
  return { times, fast, slow, gap, bimodal, max: times[times.length - 1] };
}

export function poolRasteriseArms(logs) {
  const arms = { rasterise: { ok: 0, stall: 0 }, "cheap read": { ok: 0, stall: 0 } };
  let rounds = 0;
  for (const log of logs) {
    const entries = log?.trace?.entries;
    if (!Array.isArray(entries)) continue;
    rounds++;
    const es = [...entries].sort((a, b) => (a.ms ?? 0) - (b.ms ?? 0));
    const draws = [];
    es.forEach((e, i) => {
      const m = /^after a (rasterise|cheap read) #\d+: drawing$/.exec(String(e.data?.what ?? ""));
      if (m) draws.push([i, m[1]]);
    });
    draws.forEach(([i, arm], k) => {
      const end = k + 1 < draws.length ? draws[k + 1][0] : es.length;
      let stalled = false;
      for (let z = i + 1; z < end; z++) {
        const e = es[z];
        if (e.scope === "host" && e.message === "gave up waiting" && /drawing shapes/.test(String(e.data?.what ?? "")))
          stalled = true;
        if (/scenario (passed|FAILED|skipped)/.test(String(e.message))) break;
      }
      arms[arm][stalled ? "stall" : "ok"]++;
    });
  }
  return { rounds, arms };
}

/**
 * Every draw in the round, not just the scenario's four.
 *
 * Round 28 is why this exists. `does a rasterise poison the next draw` PASSED,
 * and the same round skipped `the chart is actually visible` on "PowerPoint did
 * not respond while drawing shapes 1-9 of 9 (45s)" with the trace adding "the
 * last thing the host answered was 'rasterising a slide', 0s earlier". A draw
 * stalling straight after a rasterise, in the round whose rasterise scenario had
 * just reported no effect — because the scenario counts only the four draws it
 * makes itself and that one was not one of them. The evidence was being thrown
 * away by the thing built to collect it.
 *
 * A round issues about forty draws. Counting all of them is a tenfold bigger
 * sample per round, which turns "30-50 more rounds" into a handful:
 *
 *   rounds 23, 26, 27, 28 — 157 draws, against 16 the arms would have counted
 *   after a rasterise      1 stalled /  36
 *   after anything else    0 stalled / 121
 *
 * REPORTED SEPARATELY FROM THE ARMS, AND WEAKER THAN THEM, WHICH IS THE WHOLE
 * POINT OF KEEPING BOTH. The arms are counterbalanced — interleaved so position
 * cannot account for a difference — and that is what makes four draws worth
 * anything. This population is observational: draws that follow a rasterise
 * follow it because of which scenario they belong to, and those scenarios differ
 * in shape count and in what they ask of the host. So it can raise a suspicion
 * and it cannot settle one. Two populations, honestly labelled, beat one
 * population quietly mixing the two kinds of evidence.
 */
/**
 * Rasterise labels from rounds archived BEFORE `op` existed.
 *
 * NOT a guess at wording — an ENUMERATION of the archive. A rasterise names
 * itself unambiguously in two places that do not depend on the label at all:
 * the success line `rasterised a slide` carries `label`, and the visibility
 * scenario's `visibility step` carries `what`. Reading every round through
 * those gives the complete set, and it is closed: old archives do not change.
 *
 * WHAT THIS RECOVERS TODAY: NOTHING, AND THAT IS MEASURED, NOT ASSUMED.
 *
 * The first version of this comment claimed these four labels were 35 of 43
 * labelled rasterises and that 81% of the population was missing from the
 * pooled answer. THAT WAS WRONG, and it was wrong in the way this repo keeps
 * being wrong: a number counted against the wrong denominator. These labels do
 * lack the string "rasteris" — but `isRasterise` tests the label AND THE
 * MESSAGE, and the message on a successful rasterise is `rasterised a slide`,
 * which matches. They were classified correctly all along.
 *
 * Pooled over all 90 archived rounds, with and without this set:
 *
 *     rasterise      ok 449, stall 1      (identical both ways)
 *     anything else  ok 3302, stall 1     (identical both ways)
 *
 * So this set is belt-and-braces, not a repair. It is kept because it makes the
 * classifier independent of a message string that nobody has promised to keep,
 * and because the equality above is now a fact on the record rather than an
 * assumption. If it ever starts changing a number, something upstream renamed a
 * trace message and that is worth knowing.
 *
 * The REAL breakage this pair found was in `chartIsVisible`, which matches
 * `lastStall.what` alone — no message to fall back on — and therefore genuinely
 * did stop firing. See round 113.
 */
const RASTERISE_LABELS_BEFORE_OP = new Set([
  "an end-of-round slide shot",
  "the visibility BEFORE render",
  "the visibility AFTER render",
  "the visibility CONTROL render (same slide, back to back)",
]);

export function poolEveryDraw(logs) {
  const isDraw = (e) =>
    (e.scope === "draw" && e.message === "batch issued") ||
    /^after a (?:rasterise|cheap read) #\d+: drawing$/.test(String(e.data?.what ?? ""));
  // A rasterise EVENT is never itself a draw. The scenario's own arm markers say
  // "after a rasterise #0: drawing" — that is a draw which FOLLOWS a rasterise,
  // and reading it as one would tar the next draw with a rasterise that had
  // already been accounted for. Caught by the test below rather than by reading:
  // the untagged-draw case came out one short and the miscount was this.
  // ON `op` FIRST, then the enumerated legacy labels, then the prose.
  //
  // THIS ONE WAS NOT BROKEN — checked, and the check is the point. The sibling
  // in `chartIsVisible` was broken by call sites being given individual names,
  // so this classifier was the obvious next casualty: it also identifies a
  // rasterise by matching prose. It survives only because it happens to test
  // the MESSAGE as well as the label, and the message `rasterised a slide`
  // still contains "rasteris".
  //
  // That is luck, not design. `op` makes it design. The pooled numbers are
  // identical before and after (see `RASTERISE_LABELS_BEFORE_OP`), which is
  // exactly what a belt-and-braces change should look like and is recorded so
  // nobody later mistakes this for a fix that moved something.
  const isRasterise = (e) =>
    !isDraw(e) &&
    (e.data?.op === "rasterise" ||
      RASTERISE_LABELS_BEFORE_OP.has(String(e.data?.what ?? "")) ||
      RASTERISE_LABELS_BEFORE_OP.has(String(e.data?.label ?? "")) ||
      /rasteris/i.test(`${String(e.data?.what ?? "")} ${String(e.message ?? "")}`));
  const isStall = (e) =>
    e.scope === "host" && e.message === "gave up waiting" && /drawing shapes/.test(String(e.data?.what ?? ""));

  const after = { rasterise: { ok: 0, stall: 0 }, "anything else": { ok: 0, stall: 0 } };
  let rounds = 0;
  for (const log of logs) {
    const entries = log?.trace?.entries;
    if (!Array.isArray(entries)) continue;
    rounds++;
    const es = [...entries].sort((a, b) => (a.ms ?? 0) - (b.ms ?? 0));
    const at = [];
    es.forEach((e, i) => {
      if (isDraw(e)) at.push(i);
    });
    let prev = 0;
    at.forEach((i, k) => {
      // A rasterise anywhere since the PREVIOUS draw, so the classification is
      // about what the host did immediately before this draw and nothing older.
      let rasterised = false;
      for (let z = prev; z < i; z++) if (isRasterise(es[z])) rasterised = true;
      const end = k + 1 < at.length ? at[k + 1] : es.length;
      let stalled = false;
      for (let z = i + 1; z < end; z++) if (isStall(es[z])) stalled = true;
      after[rasterised ? "rasterise" : "anything else"][stalled ? "stall" : "ok"]++;
      prev = i;
    });
  }
  return { rounds, after };
}

/**
 * Judge a prediction against a round.
 *
 * A change made because of a round should say what the NEXT round will show, and
 * then be judged by it in as many words. That discipline is worth having and it
 * is worth nothing if nobody checks: #468 staked itself on three questions
 * answering for the first time in eleven rounds, round 27 said no, and the only
 * reason the failure was recorded rather than quietly forgotten is that someone
 * remembered making the claim.
 *
 * Deliberately few claim kinds. A prediction that cannot be evaluated by a
 * machine is a paragraph, and paragraphs are what this replaces — but the prose
 * still travels with it, because "it failed" is the least interesting half and
 * `because` is where the thinking is.
 */
export function judgePrediction(prediction, log) {
  const NOT_ASKED = new Set(["no-scratch-slide", "no-scratch-shape", "not-asked"]);
  // A THIRD CATEGORY, and keeping it out of NOT_ASKED is the point. A
  // CONDITIONAL probe's question only exists when the host misbehaves in a
  // particular way, and on a round where it behaved there is nothing to
  // measure — which is neither "the harness could not set it up" (ours to fix)
  // nor "the host would not answer" (a fact about the host).
  //
  // `does-a-failed-group-poison-the-tag` is why this exists. It answers
  // `no-refusal` — "the host grouped through a slide handle two syncs old, so
  // the refusal was never provoked", its own words — in 94 of the 114 rounds
  // that carry it, and `tags-gone` in the other 20. `no-refusal` fell through
  // to the pattern test, matched nothing, and was reported FAILED: the house
  // defect exactly, an unmeasured thing reported as a negative measurement.
  const UNPROVOKED = new Set(["no-refusal"]);
  const answers = log?.hostAnswers?.answers;
  const byId = new Map((Array.isArray(answers) ? answers : []).map((a) => [a.id, a]));
  const c = prediction.claim ?? {};
  const seen = (id) => byId.get(id);

  if (c.kind === "probe-answers" || c.kind === "probe-starves") {
    const want = c.kind === "probe-answers";
    const missing = (c.ids ?? []).filter((id) => !seen(id));
    if (missing.length) return { verdict: "undetermined", why: `not on this sheet: ${missing.join(", ")}` };
    const wrong = (c.ids ?? []).filter((id) => NOT_ASKED.has(seen(id).answer) === want);
    return wrong.length
      ? { verdict: "FAILED", why: `${wrong.map((id) => `${id}=${seen(id).answer}`).join(", ")}` }
      : { verdict: "held", why: (c.ids ?? []).map((id) => `${id}=${seen(id).answer}`).join(", ") };
  }
  if (c.kind === "trace-line-present") {
    // A CLAIM ABOUT WHAT THE TRACE SAYS, which the ledger could not express.
    // Its four kinds were all about scenarios and probes, so a question whose
    // whole answer is "does this line appear" had to live in prose — and prose
    // is what this ledger exists to replace.
    //
    // `insteadOf` is what makes it judgeable rather than merely true-or-silent:
    // a round where NEITHER line appears did not ask the question, and that is
    // `undetermined`, not a refutation. Without it, a round that simply never
    // reached the code would read as evidence against the claim.
    const entries = log?.trace?.entries;
    if (!Array.isArray(entries) || !entries.length)
      return { verdict: "undetermined", why: "the round carries no trace to look in" };
    // BY SCOPE AS WELL AS BY MESSAGE, because a claim naming one message is only
    // as good as the number of messages the thing can produce. Round 102: the
    // deck-style probe has TWO failure lines, `spending-the-bad-first-call-cures-
    // the-read` named one, the round produced the other — and the ledger read
    // that as a cure. A false HELD from a malformed claim, which is the same
    // defect this file keeps finding elsewhere.
    //
    // `scope` claims the whole family: no entry from that scope at all.
    const saw = (m) => entries.some((e) => String(e.message) === m);
    // What actually matched, so the verdict names the right reason. Reporting
    // `still carries <message>` for a SCOPE match would be a report describing
    // evidence it did not use.
    const sawScope = () => (c.scope ? entries.find((e) => String(e.scope) === c.scope) : undefined);
    // A CLAIM THAT A SYMPTOM IS GONE, which is what a fix predicts and what this
    // kind could not say. `absent: true` inverts the reading: the named line
    // must NOT appear, and `insteadOf` — a line every round carries — is what
    // proves the round ran at all, so silence from a round that never started
    // stays `undetermined` rather than counting as a cure.
    if (c.absent) {
      if (saw(c.message)) return { verdict: "FAILED", why: `the trace still carries \`${c.message}\`` };
      const other = sawScope();
      if (other)
        return {
          verdict: "FAILED",
          why: `the \`${c.scope}\` scope is still there, as \`${String(other.message)}\``,
        };
      if (c.insteadOf && saw(c.insteadOf)) return { verdict: "held", why: `no \`${c.message}\` in a round that ran` };
      return { verdict: "undetermined", why: "the round did not run far enough to show the line is gone" };
    }
    if (saw(c.message)) return { verdict: "held", why: `the trace carries \`${c.message}\`` };
    if (c.insteadOf && saw(c.insteadOf))
      return { verdict: "FAILED", why: `the trace carries \`${c.insteadOf}\` instead` };
    return {
      verdict: "undetermined",
      why: c.insteadOf
        ? `neither line appeared — the round never reached the code that writes them`
        : `\`${c.message}\` did not appear, and no alternative was named to tell absence from silence`,
    };
  }
  if (c.kind === "scenario-passes") {
    const st = log?.selftest ?? {};
    const all = Object.keys(st).map((k) => st[k]);
    const named = (c.names ?? []).map((n) => [n, all.find((x) => x?.name === n)]);
    // A QUESTION THE HOST NEVER ANSWERED IS NOT A REFUTATION, and this branch
    // called both of those FAILED until 2026-08-19 — while `probe-answers` and
    // `probe-detail-matches` directly above already answered `undetermined` for
    // exactly the same situation, and this file's own doctrine says "a skip is
    // not a flip" and "a miss is not a failure" in as many words.
    //
    // Round 088 is what makes it live rather than theoretical: `insert onto a
    // slide that already has content` came back `skipped: true, ok: false`
    // because PowerPoint stopped answering mid-draw. Any prediction naming that
    // scenario would have been recorded FAILED — the strongest verdict the
    // ledger has — on a round where the host declined to be asked.
    const missing = named.filter(([, s]) => !s).map(([n]) => n);
    if (missing.length) return { verdict: "undetermined", why: `not on this sheet: ${missing.join(", ")}` };
    const skipped = named.filter(([, s]) => s.skipped).map(([n]) => n);
    if (skipped.length) return { verdict: "undetermined", why: `the host never ran: ${skipped.join(", ")}` };
    const wrong = named.filter(([, s]) => !s.ok).map(([n]) => n);
    return wrong.length ? { verdict: "FAILED", why: wrong.join(", ") } : { verdict: "held", why: "" };
  }
  if (c.kind === "probe-detail-matches") {
    const a = seen(c.id);
    if (!a) return { verdict: "undetermined", why: `${c.id} not on this sheet` };
    // A never-put question has no detail worth matching, and calling that a
    // failure would blame the prediction for the host refusing the question.
    if (NOT_ASKED.has(a.answer)) return { verdict: "undetermined", why: `${c.id} was never put (${a.answer})` };
    if (UNPROVOKED.has(a.answer))
      return { verdict: "undetermined", why: `${c.id}'s precondition did not occur this round (${a.answer})` };
    // THE ANSWER KEY AS WELL AS THE PROSE. An `answer` is a stable enum the
    // probe picks; a `detail` is a sentence someone rewords. #520 was staked on
    // `InvalidParam|5010|GeneralException` — error codes this probe has never
    // once emitted in 114 rounds — so on the 20 rounds where it DID fire and
    // report `.tags was undefined after the refused group`, the exact outcome
    // the prediction claimed, the ledger would still have printed FAILED.
    //
    // Round 102 taught the same lesson from the trace side: a claim is only as
    // good as the set of strings the thing it watches can produce. Matching the
    // key gives a claim something to anchor on that a rewrite cannot move.
    const hit = new RegExp(c.pattern, "i").test(`${a.answer ?? ""} ${a.detail ?? ""}`);
    return { verdict: hit ? "held" : "FAILED", why: `${a.answer} — ${String(a.detail ?? "").slice(0, 90)}` };
  }
  return { verdict: "undetermined", why: `unknown claim kind ${String(c.kind)}` };
}

/**
 * The round a prediction may be judged on: the newest one taken AFTER the build
 * it was staked on, or nothing.
 *
 * The "nothing" is the half that was wrong, and it produced a false HELD the
 * first time a prediction was staked on a build with no round yet. The old rule
 * read "find the prompting build in the archive and take everything after it",
 * and when that build was NOT in the archive — which is the normal state for a
 * prediction written the moment a change lands — it fell back to judging against
 * every round, i.e. the newest one, which is OLDER than the change. `same scale
 * across the deck` passes in the recent archive, so a prediction about a change
 * that has never been run came out `held` on evidence recorded before it
 * existed.
 *
 * That is the same defect `roundsToJudgeOn` already describes ("round 27
 * standing in for a round 29 that does not exist yet"), reached through the
 * not-found branch instead of through inequality. A prediction whose build has
 * not been rounded has NO round to judge on, and saying so is the whole answer.
 */
export function roundToJudgeOn(logs, afterBuild, buildOf, madeOn) {
  // THE NEWEST OF THE ELIGIBLE ROUNDS, and everything about WHICH rounds are
  // eligible now lives in `roundsToJudgeOn` with the code that decides it.
  // Prefer that one: taking the last of a population is exactly the n=1 reading
  // that let #520 sit open for 130 rounds, and it survives here only because
  // the single-round claim kinds have no population to pool.
  return roundsToJudgeOn(logs, afterBuild, buildOf, madeOn).slice(-1)[0];
}

/**
 * EVERY round a prediction may be judged on, newest last — the population
 * `roundToJudgeOn` takes the last of.
 *
 * A single round is n=1, and this file spends a page arguing that a single
 * round's verdict is a fact about the code AND the host's mood that afternoon.
 * The prediction ledger was the one reader that ignored it: `roundToJudgeOn`
 * hands back the newest round and the verdict is whatever that afternoon said.
 *
 * For a CONDITIONAL probe that is not a rounding error, it is the whole answer.
 * `does-a-failed-group-poison-the-tag` puts its question in 20 of 114 rounds,
 * so the newest round is 82% likely to carry no measurement at all — and the
 * one time in five it does, that reading stands alone against nineteen others
 * nobody looks at. #520 sat OPEN for 130 rounds with its answer recorded
 * twenty times.
 */
export function roundsToJudgeOn(logs, afterBuild, buildOf, madeOn) {
  // THE LAST ROUND ON THAT BUILD, NOT THE FIRST. `findIndex` stopped at the
  // earliest one, and the cycle archives two rounds at 16:9 plus one at 4:3 on
  // a single build — so `slice(madeAt + 1)` still held the prompting build's own
  // siblings, and the newest of them could be handed back as the round that
  // judges a change made after all three were taken. The discipline this loop is
  // built on is "run the same build twice"; the judge has to know that a build
  // can appear more than once or it quietly rules on its own control.
  let madeAt = -1;
  for (let i = 0; i < logs.length; i++) if (buildOf(logs[i]) === afterBuild) madeAt = i;
  // Rounds are ordered oldest-first, so anything at or before the prompting
  // round cannot test the change — and a build with no round at all is past the
  // end of the archive, not the start of it.
  //
  // WHICH LEAVES A PREDICTION STAKED ON A BUILD NOBODY EVER ROUNDS, and that is
  // the normal case rather than the exotic one: a claim is written the moment a
  // change lands, and the next merge supersedes that commit before any round
  // runs. Matching the build exactly would then answer `no round yet` forever —
  // an entry that can never be judged, which is the same as no entry at all.
  //
  // A round's build stamp carries its DATE (`95170cf · 2026-08-17 09:03Z`), so
  // an entry that says when it was made can be judged on any round taken after
  // that day. The build match stays first: it is exact, and dates are only as
  // good as the stamp.
  return madeAt !== -1 ? logs.slice(madeAt + 1) : madeOn ? logs.filter((l) => stampDate(l, buildOf) > madeOn) : [];
}

/**
 * The verdict across every eligible round, and the split that produced it.
 *
 * THE SPLIT IS THE POINT, not the headline. A probe that answers sometimes
 * gives a population, and a population can disagree with itself — which is a
 * finding about the host, not a tie to be broken quietly. So when the decided
 * rounds say both things, this says so rather than picking one.
 */
export function judgeAcross(prediction, rounds, buildOf) {
  const held = [];
  const failed = [];
  let undecided = 0;
  for (const r of rounds) {
    const { verdict, why } = judgePrediction(prediction, r);
    if (verdict === "held") held.push({ build: buildOf(r), why });
    else if (verdict === "FAILED") failed.push({ build: buildOf(r), why });
    else undecided++;
  }
  const decided = held.length + failed.length;
  const verdict = !decided ? "undetermined" : held.length && failed.length ? "BOTH" : held.length ? "held" : "FAILED";
  const last = (failed.length ? failed : held).slice(-1)[0];
  return { verdict, held: held.length, failed: failed.length, undecided, decided, rounds: rounds.length, last };
}

/** When a round was taken, out of its build stamp, or "" when it does not carry one. */
function stampDate(log, buildOf) {
  // `buildOf` yields the hash alone, so read the raw stamp — the date is the
  // half this needs and the half the hash throws away.
  void buildOf;
  // THE TIME AS WELL AS THE DAY, because the comparison above is strict `>` and
  // a prediction is nearly always judged by a round taken the SAME day it was
  // staked. Date-only, `"2026-08-19" > "2026-08-19"` is false, so the round that
  // exists is refused and the report says `no round yet` — which is precisely
  // what round 088 did on 2026-08-19 to the #586 entry staked that morning.
  //
  // Lexicographic order does the rest with no extra branch: a stamp is
  // `hash · YYYY-MM-DD HH:MMZ`, so `"2026-08-19 13:58Z" > "2026-08-19"` is true
  // (same prefix, longer string), while `"2026-08-18 07:11Z" > "2026-08-19"` is
  // still false. A `madeOn` carrying its own time compares by time; one carrying
  // only a day means start of that day, which is what staking on a day means.
  const m = /·\s*(\d{4}-\d{2}-\d{2}(?:[ T]\d{2}:\d{2})?)/.exec(String(log?.build ?? ""));
  return m ? m[1] : "";
}

/**
 * A scenario whose verdict counts "N of M" — and what M was in the rounds before.
 *
 * ROUND 088 IS WHY THIS EXISTS. `same scale across the deck` passed with
 * `6 of 6 charts carry the shared scale` where all 63 rounds before it had run
 * EIGHT. Its verdict is `scaled === charts.length`, measured against a
 * population it discovers rather than one it is given — `probeCharts(prefix)`
 * returns whatever the deck scan finds, and the only guard is `< 2`. So a run
 * that finds six charts and scales six reads exactly like one that found eight
 * and scaled eight, and `scenarioRegressions` compares PASS to PASS and says
 * "no scenario regressed".
 *
 * That is the same defect as the suite-size high-water mark: a guard that cannot
 * see its own population shrink. A scenario could fall to `2 of 2` and every
 * reading in this file would still be green.
 *
 * The denominator is not a fault on its own — round 088's six is downstream of a
 * host stall that skipped the scenario which seeds the probe charts. It is a
 * REASON TO READ, and quoting the pass without it is quoting a ratio whose
 * bottom half moved.
 */
export function poolScenarioPopulations(logs) {
  const seen = new Map();
  for (const log of logs) {
    const st = log?.selftest ?? [];
    for (const s of Object.keys(st).map((k) => st[k])) {
      if (!s?.name) continue;
      // The verdict's own "N of M" — the shape every counting scenario here uses.
      const m = /\b(\d+) of (\d+)\b/.exec(String(s.detail ?? ""));
      if (!m) continue;
      if (!seen.has(s.name)) seen.set(s.name, []);
      seen.get(s.name).push({ build: String(log.build ?? ""), of: Number(m[2]), ok: Boolean(s.ok) });
    }
  }
  const shrunk = [];
  for (const [name, hist] of seen) {
    // THREE PRIORS MINIMUM, because "usually" needs more than one observation.
    // Round 112 fired this on `insert onto a slide that already has content —
    // 2 this round, usually 16 over 1 prior round(s)`: a scenario whose verdict
    // only recently started carrying an "N of M" count, so its entire history
    // was a single round. One number is not a norm, and this project's own noise
    // floor — one build run twice scoring 1 and 5 — is the reason to say so.
    if (hist.length < MIN_PRIORS_FOR_A_BASELINE + 1) continue;
    const now = hist[hist.length - 1];
    const priors = hist.slice(0, -1).map((h) => h.of);
    // The population it has USUALLY had. Not the mean: a single small round
    // would drag the bar down and hide the next one.
    const usual = priors.sort((a, b) => b - a)[Math.floor(priors.length / 2)];
    if (now.of < usual) shrunk.push({ name, now: now.of, usual, ok: now.ok, rounds: priors.length });
  }
  return shrunk;
}

/**
 * Whether the charts this round drew ended up GROUPED, which no verdict reports.
 *
 * ROUNDS 092 AND 093 ARE WHY. One build, run twice, nothing changed between them:
 *
 *     092   20 charts grouped, 0 refusals, deck 0,4,2,5,1,1,1
 *     093   15 charts grouped, 4 refusals, deck 0,4,2,17,24,24,24
 *
 * Three slides ended holding TWENTY-FOUR shapes each instead of one — three
 * charts that did not group, seventy-two shapes loose on the deck — and both
 * rounds reported `13/13` and the byte-identical verdict line `8 of 8 charts
 * carry the shared scale ... 8 still re-editable`.
 *
 * The scenario is not lying. It asks whether the config survived, and it did:
 * the ungrouped fallback keeps the tag. It simply cannot see grouping, which is
 * what the last several changes here have been about — so `scenarioRegressions`
 * compares PASS to PASS across a round that left seventy-two loose shapes and a
 * round that left none.
 *
 * Counted from the TRACE, which is exact — `charts` on each `grouped the chart's
 * shapes` line, and every `not grouping:` line — with the deck printed beside it
 * as corroboration rather than as the measure. A slide's shape count needs a
 * threshold to interpret, and a guard sized by guesswork is how this instrument
 * has been wrong before.
 */
/**
 * How many earlier rounds a "usually" needs before it is allowed to be printed.
 *
 * Two emitters have now shipped a baseline computed from ONE prior round, and
 * the gate printed `usually 16 over 1 prior round(s)` in a real run before
 * anyone noticed. This project's own noise floor — one build run twice, scoring
 * 1 and 5 with nothing changed — is the argument: a single prior cannot
 * distinguish a trend from the host's mood.
 */
const MIN_PRIORS_FOR_A_BASELINE = 3;

export function poolGroupingOutcome(logs) {
  const per = [];
  for (const log of logs) {
    const entries = log?.trace?.entries;
    if (!Array.isArray(entries)) continue;
    let grouped = 0;
    let refused = 0;
    for (const e of entries) {
      const m = String(e.message ?? "");
      if (m === "grouped the chart's shapes") grouped += Number(e.data?.charts) || 0;
      else if (/^not grouping/.test(m)) refused += 1;
    }
    if (!grouped && !refused) continue;
    const deck = (log?.deck?.inventory ?? []).map((sl) => sl.count ?? sl.shapes?.length ?? 0);
    per.push({ build: String(log.build ?? "").split(" ")[0], grouped, refused, deck });
  }
  if (!per.length) return null;
  const now = per[per.length - 1];
  const priors = per.slice(0, -1);
  // THREE PRIORS MINIMUM — the same rule `poolScenarioPopulations` needed, and
  // the same defect it had. A median of one observation is not a "usually", and
  // a median of ZERO observations used to be reported here as `0`, which is this
  // project's house defect exactly: UNREADABLE PRINTED AS A NEGATIVE. A reader
  // seeing "usually 0 refused" cannot tell a clean history from no history.
  // `null` is the honest value and the printer must say so out loud.
  const refusedMedian =
    priors.length >= MIN_PRIORS_FOR_A_BASELINE
      ? priors.map((p) => p.refused).sort((a, b) => a - b)[Math.floor(priors.length / 2)]
      : null;
  // THE DENOMINATOR, AND THE ONLY READING THAT SURVIVES A CHANGE OF POPULATION.
  //
  // `grouped` per round ran 15-20 for the whole archive and then halved to 9 at
  // round 153, and nothing here could say why — or that it had happened. The
  // cause is benign and complete: the in-place update started working. Six more
  // charts per round are updated in place, an in-place update never redraws, and
  // a chart that is not redrawn is never regrouped. 15 - 9 = 6 and 11 - 5 = 6,
  // exactly, in the round it changed.
  //
  // Benign, and it silently rebased every grouping figure in this file. "0
  // refused (usually 2)" reads as an improvement when half the attempts stopped
  // happening — fewer refusals out of fewer tries. The RATE is the reading that
  // holds across the change: round 160 grouped 9 of 9 attempts, round 141
  // grouped 10 of 19. Same instrument, and only one of those two numbers can be
  // compared to the other.
  //
  // This is the `same scale across the deck` trap again — "6 of 6" and "8 of 8"
  // are both a pass — and this file argues it at length one screen up while
  // reporting a bare count here.
  const attempts = now.grouped + now.refused;
  const recent = per.slice(-RECENT_IN_A_ROW).map((p) => p.grouped + p.refused);
  return { now, refusedMedian, rounds: priors.length, attempts, recent };
}

/**
 * The host-cost meter every scenario verdict has carried since round 023, which
 * no script has ever read.
 *
 * Four counters — `errors`, `idRefusals`, `generalExceptions`, `emptyReReads` —
 * are recorded per scenario, per round, as a delta across that scenario. 86 of
 * 86 archived rounds carry them. `grep friction scripts/` returned nothing.
 *
 * It answers per scenario what `poolScenarioPopulations` can only ask per round:
 * whether a scenario passed on an easier host than it used to.
 *
 * TWO COUNTERS ARE NOT SIGNALS AND THE REPORT MUST SAY SO, or this becomes one
 * more number read as meaning something:
 *
 *   - `generalExceptions` has never been non-zero. Not once, in any scenario, in
 *     any round.
 *   - `stop a run part-way` reports exactly one `error` every round — the
 *     deliberate abort, counted as an error. A constant is not a measurement.
 *
 * Both are DERIVED here rather than written down, because a hardcoded conclusion
 * keeps printing after it stops being true — which is exactly how the deck-style
 * probe lied for three rounds.
 */
export function poolScenarioFriction(logs) {
  const KEYS = ["errors", "idRefusals", "generalExceptions", "emptyReReads"];
  const per = new Map();
  let rounds = 0;
  for (const log of logs) {
    const entries = log?.trace?.entries;
    if (!Array.isArray(entries)) continue;
    let any = false;
    for (const e of entries) {
      const f = e?.data?.friction;
      const name = e?.data?.name;
      if (!f || !name) continue;
      any = true;
      if (!per.has(name)) per.set(name, { name, n: 0, seen: new Map(), sum: {} });
      const row = per.get(name);
      row.n += 1;
      for (const k of KEYS) {
        const v = Number(f[k]) || 0;
        row.sum[k] = (row.sum[k] ?? 0) + v;
        // Every distinct per-round value, so a CONSTANT can be told from a
        // number that merely happens to be large.
        if (!row.seen.has(k)) row.seen.set(k, new Set());
        row.seen.get(k).add(v);
      }
    }
    if (any) rounds += 1;
  }
  const rows = [...per.values()].sort((a, b) => b.sum.errors + b.sum.idRefusals - (a.sum.errors + a.sum.idRefusals));
  // A counter that has never once been non-zero anywhere carries no information.
  const dead = KEYS.filter((k) => rows.every((r) => r.sum[k] === 0));
  // A scenario/counter pair with exactly one distinct value across every round
  // is a constant — true of the deliberate abort, and worth naming as such.
  const constant = [];
  for (const r of rows)
    for (const k of KEYS) {
      const vals = r.seen.get(k);
      if (r.n > 2 && vals && vals.size === 1 && [...vals][0] !== 0)
        constant.push({ name: r.name, key: k, value: [...vals][0] });
    }
  return { rounds, rows, dead, constant };
}

/** Open predictions, judged against the newest round given. */
/**
 * Is this entry stamped with a day and no time, on a claim that has no build in
 * the archive to anchor it?
 *
 * `stampDate` compares lexicographically and a bare `2026-08-22` means the
 * START of that day, so a claim staked at 22:55 is judged against every round
 * taken earlier the same day — including the rounds that motivated it. The
 * first rule this ledger is built on is that a prediction is never judged
 * against the round that prompted it, and a date-only stamp walks around it.
 *
 * ONLY WHEN THE BUILD IS ABSENT. A `afterBuild` that IS in the archive pins the
 * position exactly and the date is never consulted, so warning there would be
 * noise on every correctly-staked entry.
 */
export function stakedWithoutATime(prediction, logs, buildOf) {
  const madeOn = String(prediction?.madeOn ?? "");
  // A DAY WITH NO TIME, or a time with no ZONE. Both make the comparison mean
  // something other than what the writer meant, and both fail quietly.
  //
  // The zone half was found immediately after the day half was fixed, by
  // stepping straight into it. #684 was re-stamped `2026-08-22 23:00` — local
  // time — while every round stamp is UTC (`c7e2876 · 2026-08-22 21:14Z`).
  // Lexicographically "21:14Z" < "23:00", so the two rounds that were taken to
  // judge it read as older than the claim and the entry said `no round yet`.
  // That is the worst direction to fail in: `no round yet` reads as patience.
  const bare = /^\d{4}-\d{2}-\d{2}$/.test(madeOn);
  const zoneless = /^\d{4}-\d{2}-\d{2}[ T]\d{2}:\d{2}$/.test(madeOn);
  if (!bare && !zoneless) return false;
  // Only when the build is absent: a build that IS in the archive pins the
  // position exactly and the date is never consulted.
  return !(logs ?? []).some((l) => buildOf(l) === prediction.afterBuild);
}

function reportPredictions(logs, load = readFileSync) {
  let ledger;
  try {
    ledger = JSON.parse(load("rounds/predictions.json", "utf8"));
  } catch {
    return; // no ledger, nothing to say
  }
  const open = ledger.filter((p) => p.outcome === "open");
  if (!open.length || !logs.length) return;
  const buildOf = (log) => String(log?.build ?? "").split(" ")[0];
  console.log(`\n  OPEN PREDICTIONS`);
  for (const p of open) {
    // NEVER judged against the round that prompted it. A prediction is made
    // BECAUSE of a round, and the change it predicts about is not in that
    // round's build — so judging there fails every prediction the moment it is
    // written. #472's came out "FAILED / value=undefined" against round 28,
    // which is the round that raised the question and could not carry the
    // instrument that answers it.
    // By POSITION in the archive, not by "any build that is not this one".
    // Filtering on inequality picked the newest round that merely differed,
    // which for a prediction made on the newest round meant judging it against
    // an OLDER one — round 27 standing in for a round 29 that does not exist
    // yet. Rounds are ordered oldest-first, so anything at or before the
    // prompting round cannot test the change.
    // A DATE WITHOUT A TIME MEANS THE START OF THAT DAY, and that is a foot-gun
    // rather than a convenience. `stampDate` compares lexicographically, so a
    // claim staked at 22:55 with a bare `2026-08-22` is judged against every
    // round taken earlier the same day — INCLUDING the rounds that motivated it.
    // The first rule this ledger is built on is that a prediction is never
    // judged against the round that prompted it, and a date-only stamp walks
    // straight around it. #684 did exactly that on the evening it was written
    // and came out `BOTH`, on evidence recorded before the change existed.
    //
    // Warned rather than refused: the entries from 2026-08-16 onward carry
    // date-only stamps and are correctly judged, because the rounds that would
    // confuse them were taken on other days.
    if (stakedWithoutATime(p, logs, buildOf))
      console.log(
        `    ^ NOTE  ${p.id} is stamped \`${p.madeOn}\`, which the judge cannot compare safely.\n` +
          `            A day with no time is read from the START of that day, so rounds taken earlier —\n` +
          `            including the ones that prompted the claim — count as evidence. A time with no\n` +
          `            zone is compared against UTC round stamps, so a local-time stamp silently\n` +
          `            back-dates or forward-dates itself. Stake as \`YYYY-MM-DD HH:MMZ\`, in UTC.`,
      );
    const eligible = roundsToJudgeOn(logs, p.afterBuild, buildOf, p.madeOn);
    if (!eligible.length) {
      console.log(`    no round yet   ${p.id}  (${p.madeIn}, made on ${p.afterBuild})`);
      continue;
    }
    // ACROSS THE POPULATION, not on whichever round happened to be newest.
    const t = judgeAcross(p, eligible, buildOf);
    const where = t.decided
      ? `judged on ${t.rounds} round(s), ${t.decided} of which measured it`
      : `no round measured it`;
    console.log(`    ${t.verdict.padEnd(13)} ${p.id}  (${p.madeIn}) — ${where}`);
    if (t.decided) console.log(`                  held ${t.held} · FAILED ${t.failed} · never put ${t.undecided}`);
    if (t.last) console.log(`                  latest reading, ${t.last.build}: ${t.last.why}`);
    if (t.verdict === "BOTH")
      console.log(
        `                  the archive says BOTH about code that did not change. Read the split, not a headline.`,
      );
  }
  console.log(
    `    A prediction that came out is only half of it — record what happened in\n` +
      `    rounds/predictions.json and say so in the change that acts on it.`,
  );
}

/**
 * What each self-test scenario has said, round by round.
 *
 * A single round's verdict is not a fact about the code — it is a fact about the
 * code AND the host's mood that afternoon, and the two are not separable from
 * one round. `explode a degraded picture` FAILED in round 23, PASSED in 26,
 * FAILED in 27 and 28, with no change to it in between. Read one of those as a
 * regression and you go looking for a bug that is not there; read the next as a
 * fix and you close a question that is still open. Both happened here before
 * this existed, and the only reason it was caught is that someone laid four
 * rounds side by side by hand.
 *
 * A SKIP IS NOT A FLIP, and keeping the distinction is the point. `the chart is
 * actually visible` reads `pass pass pass skip` — it has never once disagreed
 * with itself; the host simply stopped answering during the fourth. A scenario
 * that has only ever passed-or-skipped is sometimes-unmeasured, which is a
 * completely different thing from one that has genuinely said both pass and fail
 * about the same code.
 */
export function scenarioHistory(logs) {
  const hist = new Map();
  for (const log of logs) {
    const st = log?.selftest;
    if (!st) continue;
    for (const key of Object.keys(st)) {
      const s = st[key];
      if (!s?.name) continue;
      if (!hist.has(s.name)) hist.set(s.name, []);
      hist.get(s.name).push(s.skipped ? "skip" : s.ok ? "pass" : "FAIL");
    }
  }
  return [...hist.entries()].map(([name, verdicts]) => {
    const measured = verdicts.filter((v) => v !== "skip");
    return {
      name,
      verdicts,
      skips: verdicts.length - measured.length,
      // Disagreement about the same code, which is the thing worth flagging.
      // Computed over MEASURED rounds only, so a skip can never manufacture one.
      flips: new Set(measured).size > 1,
    };
  });
}

/** Scenario verdicts across rounds, so one round is not mistaken for a trend. */
function reportStability(logs) {
  if (logs.length < 2) return;
  const hist = scenarioHistory(logs);
  if (!hist.length) return;
  const flipping = hist.filter((h) => h.flips);
  const skipped = hist.filter((h) => !h.flips && h.skips);
  console.log(`\n  SCENARIO VERDICTS ACROSS ${logs.length} ROUNDS`);
  for (const h of hist) {
    const mark = h.flips ? "FLIPS" : h.skips ? "skips" : "     ";
    console.log(`    ${mark} ${h.name.padEnd(46)} ${h.verdicts.join(" ")}`);
  }
  if (flipping.length)
    console.log(
      `\n    ${flipping.length} scenario(s) have said BOTH pass and fail about code that did not\n` +
        `    change between them. Do not read a single round's verdict on those as a\n` +
        `    regression or as a fix — they are reporting the host's mood.`,
    );
  if (skipped.length)
    console.log(
      `    ${skipped.length} more were SKIPPED in some rounds: never a disagreement, just\n` +
        `    rounds where the host stopped answering before the question was put.`,
    );
}

/**
 * The tag-failure faults, per BUILD, across every round.
 *
 * WHY THIS EXISTS, and it is the only reason: on 2026-08-15 a fix to the tag
 * anchor was nearly reported as working because round 043 (with it) looked like
 * round 042 (without it). It did — and so did rounds 041 and 042, which have NO
 * renderer change between them at all:
 *
 *     041  ca866e3   tags-undefined 1   group-5010 1   no-queue 1   tagging-failed 4
 *     042  a54401c   tags-undefined 5   group-5010 5   no-queue 5   tagging-failed 8
 *
 * A five-fold swing across a merge of a probe and some documentation. These
 * counts track the host's REGIME, not the code, so comparing two rounds by eye
 * is measuring mood — the same trap the rasterise arms exist to avoid, in a
 * place nobody had noticed it.
 *
 * Grouped by build rather than by round so repeat rounds on one build pool
 * instead of competing, and the spread WITHIN a build is printed, because that
 * spread is the noise floor any claim about a fix has to clear.
 */
export function poolTagFaults(logs) {
  const KINDS = [
    [/Cannot read properties of undefined \(reading 'add'\)/, "tags-undefined"],
    [/at=writing the chart's config tag/, "cfg-tag-5010"],
    [/at=grouping the chart's shapes/, "group-5010"],
    [/no chart's tag could be queued/, "no-queue"],
  ];
  const byBuild = new Map();
  for (const log of logs) {
    const entries = log?.trace?.entries;
    if (!Array.isArray(entries)) continue;
    const build = String(log.build ?? "?").split(" ")[0];
    const counts = Object.fromEntries(KINDS.map(([, name]) => [name, 0]));
    counts["tagging-failed"] = 0;
    for (const e of entries) {
      if (/^tagging failed/.test(String(e.message ?? ""))) counts["tagging-failed"]++;
      const blob = JSON.stringify(e.data ?? {});
      for (const [re, name] of KINDS) if (re.test(blob)) counts[name]++;
    }
    if (!byBuild.has(build)) byBuild.set(build, []);
    byBuild.get(build).push(counts);
  }
  return byBuild;
}

/**
 * Does a chart that GETS GROUPED keep its config?
 *
 * THE QUESTION NOBODY ASKED FOR ELEVEN ROUNDS, and the archive had the answer
 * the whole time. Pooled over every round on 2026-08-15:
 *
 *     grouped      64 charts, 1 lost its tag    (1.6%)
 *     NOT grouped  62 charts, 41 lost their tag (66%)
 *
 * Per round it is almost mechanical — three grouped and none lost, two or three
 * ungrouped and two lost, round after round.
 *
 * It reframes the whole tag effort. When grouping succeeds the tag goes onto the
 * GROUP, a handle made in that batch and never resolved, and it lands. When
 * grouping is skipped the tag falls back to a `created` handle and is refused
 * two times in three. Which handle the fallback uses — the question four rounds
 * and a renderer change went into — is a question about the losing path.
 * `not grouping: no member handle this host will accept` carries `refreshed: 0`,
 * so what decides a chart's config is whether the pre-grouping RE-READ returned
 * anything.
 *
 * Reported every round from now on, because the cost of not reporting it was
 * eleven rounds of looking one level too low.
 */
/**
 * How much of the round's grouping this pool can actually SEE.
 *
 * `poolGroupVsTag` joins a chart's messages by `data.chart`, and that field is
 * written in exactly one place: the `traceAbout({ chart })` inside the deck-wide
 * rescale in `src/taskpane/selftest.ts`. Every other grouping in a round — the
 * ordinary inserts, the probe, anything a single-chart batch does — carries no
 * such key and is invisible here.
 *
 * MEASURED ARCHIVE-WIDE, the blindness is uneven and that is what makes it
 * dangerous rather than merely partial: across 57 rounds the pool sees 229 of
 * 638 `grouped the chart` events (36%) but 109 of 127 `not grouping` events
 * (86%). The two columns of a comparison are sampled at wildly different rates,
 * so the RATIO between them is biased, not just small.
 *
 * This is the shape of the 333/333 incident, in which a pooled figure was
 * reported as 100% because the window collecting it silently dropped every
 * declining case. The remedy is the same: print the denominator the reader would
 * otherwise assume.
 *
 * AND WIDENING IT IS NOT AVAILABLE, which is worth writing down so the next
 * reader does not spend the evening finding out again. Closing the gap would
 * mean joining each `grouped the chart` with the `tagging failed` for the SAME
 * chart, and that second line is emitted per BATCH on purpose — its own comment
 * says "the batch covers several charts, so one id would be a guess about
 * which", because a tag write fails for a whole sync rather than for one chart
 * in it. Attributing a batch failure to individual charts would manufacture
 * precision the host never gave, which is a worse fault than a narrow number
 * honestly labelled.
 *
 * Checked, not assumed: `data.charts` is 1 on every one of these entries across
 * the archive, so entries and charts do coincide here and the fractions below
 * compare like with like.
 */
export function poolGroupVsTagCoverage(logs) {
  const out = { groupedSeen: 0, groupedTotal: 0, ungroupedSeen: 0, ungroupedTotal: 0 };
  for (const log of logs) {
    for (const e of log?.trace?.entries ?? []) {
      const m = String(e.message ?? "");
      const keyed = (e.data ?? {}).chart !== undefined;
      if (m.startsWith("grouped the chart")) {
        out.groupedTotal++;
        if (keyed) out.groupedSeen++;
      } else if (m.startsWith("not grouping")) {
        out.ungroupedTotal++;
        if (keyed) out.ungroupedSeen++;
      }
    }
  }
  return out;
}

export function poolGroupVsTag(logs) {
  const out = { grouped: 0, groupedLost: 0, ungrouped: 0, ungroupedLost: 0 };
  for (const log of logs) {
    const entries = log?.trace?.entries;
    if (!Array.isArray(entries)) continue;
    const per = new Map();
    for (const e of entries) {
      const chart = (e.data ?? {}).chart;
      if (!chart) continue;
      if (!per.has(chart)) per.set(chart, []);
      per.get(chart).push(String(e.message ?? ""));
    }
    for (const msgs of per.values()) {
      const lost = msgs.some((m) => m.startsWith("tagging failed"));
      // A chart can only be counted once, and `grouped` wins: the two messages
      // are mutually exclusive per chart by construction, but a future retry
      // that produced both would otherwise be counted in both columns and quietly
      // flatten the very difference this exists to show.
      if (msgs.some((m) => m.startsWith("grouped the chart"))) {
        out.grouped++;
        if (lost) out.groupedLost++;
      } else if (msgs.some((m) => m.startsWith("not grouping"))) {
        out.ungrouped++;
        if (lost) out.ungroupedLost++;
      }
    }
  }
  return out;
}

/**
 * Did the chart land on a slide that already had shapes, or on a fresh one?
 *
 * THE ROOT, found 2026-08-15, and the cleanest separation this project has:
 *
 *     slide already had shapes  82 chart(s), 81 grouped = 99%
 *     freshly added, empty      74 chart(s),  1 grouped =  1%
 *
 * Not a tendency — a switch. And it completes the chain: a chart on a freshly
 * added slide gets a short or empty pre-grouping re-read, so it is not grouped,
 * so its tag falls back to a `created` handle, which this host refuses about
 * seven times in ten. Charts on an established slide group and keep their config.
 *
 * It is not a new problem either. It is THE problem, one level below everything
 * the tag work was aimed at, and this repo already knew a freshly-added slide is
 * special: `shape-add-held-slide-proxy` answers `threw`, its id does not
 * round-trip, and the #108-#111 saga was four attempts at drawing on one.
 *
 * `onSlide` is the shape count the DRAW recorded for the slide before it began,
 * which is why this pools over rounds that were archived long before anyone
 * thought to ask.
 */
export function poolFreshVsEstablished(logs) {
  const out = { established: 0, establishedGrouped: 0, fresh: 0, freshGrouped: 0 };
  for (const log of logs) {
    const entries = log?.trace?.entries;
    if (!Array.isArray(entries)) continue;
    const onSlide = new Map();
    const grouped = new Set();
    const decided = new Set();
    for (const e of entries) {
      const d = e.data ?? {};
      if (!d.chart) continue;
      // The FIRST batch's reading: what was on the slide before this chart.
      if (typeof d.onSlide === "number" && !onSlide.has(d.chart)) onSlide.set(d.chart, d.onSlide);
      if (/^grouped the chart/.test(String(e.message))) {
        grouped.add(d.chart);
        decided.add(d.chart);
      }
      if (/^not grouping/.test(String(e.message))) decided.add(d.chart);
    }
    for (const [chart, n] of onSlide) {
      // Only charts whose grouping was DECIDED. A chart the round never reached
      // has no outcome to attribute to its slide.
      if (!decided.has(chart)) continue;
      if (n > 0) {
        out.established++;
        if (grouped.has(chart)) out.establishedGrouped++;
      } else {
        out.fresh++;
        if (grouped.has(chart)) out.freshGrouped++;
      }
    }
  }
  return out;
}

/**
 * How long a window counts as "now".
 *
 * A GUESS ABOUT WHERE THE LAST REGIME BOUNDARY IS, and it should be read as
 * one. 20 was chosen to exclude everything before the settled retry (rounds
 * 064/065) — and it does, while sitting straight across a second boundary
 * nobody had found: rounds 137-142 failed to group 1,4,4,5,9,2 charts and
 * rounds 143-160 failed 3 in eighteen. So the "current rate" this produced was
 * itself smeared by a dead era, which is the exact defect it was written to
 * fix, one level down and in my own hand.
 *
 * Hence `recentSequence` below. A window is a guess; a sequence is evidence.
 */
export const RECENT_ROUNDS = 20;

/**
 * The fresh-versus-established split over recent rounds only.
 *
 * WHY IT EXISTS. The all-time number is known to be unreadable and the report
 * said so in prose and printed it anyway: "pooled over 39 rounds that predate
 * the retry, so it will climb slowly and should not be read as the current
 * rate". A figure a reader is told to mentally discount is a figure nobody can
 * use, and the answer to a stale window is not a caveat — it is a second
 * window.
 *
 * READ IT WITH `recentFreshSequence`, NEVER ALONE. Over successive windows this
 * same population reads 80% at 30 rounds, 91% at 20, 100% at 12 — and at 8
 * there are no fresh-slide charts in it at all, because the in-place update now
 * handles those charts and a chart that is not redrawn never lands on a fresh
 * slide. A percentage over an emptying population is the thing to watch for
 * here, and only the sequence shows it emptying.
 *
 * `null` below the window rather than a short-window figure, because a rate
 * from five rounds is the thing this exists to stop.
 */
export function recentFreshVsEstablished(logs, n = RECENT_ROUNDS) {
  return (logs?.length ?? 0) > n ? poolFreshVsEstablished(logs.slice(-n)) : null;
}

/**
 * Fresh-slide charts per round, and how many of them grouped — in sequence.
 *
 * The reading no window can give. A rate needs a population, and this one is
 * shrinking for a reason that has nothing to do with grouping: the in-place
 * update took the charts. `0/0` rounds are the signal, not a gap.
 */
export function recentFreshSequence(logs, n = RECENT_IN_A_ROW) {
  return (logs ?? []).slice(-n).map((log) => {
    const f = poolFreshVsEstablished([log]);
    return `${f.freshGrouped}/${f.fresh}`;
  });
}

/**
 * Charts that cannot follow a drag — the failure a passing scenario was hiding.
 *
 * WHY IT NEEDED ITS OWN NUMBER. `an update follows a moved chart` passes, and it
 * tests ONE chart. Rounds 073 and 074 lost the origin tag on **9 of 19 and 8 of
 * 17** charts in the same rounds — roughly half the population, every one of
 * which would fail to follow a user's drag, while the scenario reported the
 * round trip holding.
 *
 * That is the shape of thing this project keeps finding: a green verdict over a
 * sample, with the population telling a different story. `does a rasterise
 * poison the next draw` counted only its own four draws; the fresh-slide split
 * sat unqueried for eleven rounds. A scenario samples; a pooled count does not.
 *
 * NOT A RATE, deliberately. Only the FAILURES are traced — a successful origin
 * write says nothing — so there is no honest denominator here, and inventing one
 * by guessing at the chart count would be the kind of number this file has
 * already had to correct once. A count that climbed from 0 to 8-9 a round is the
 * signal; anyone wanting the ratio should read `grouped the chart's shapes`
 * beside it and say so out loud.
 */
export function poolOriginTagLosses(logs) {
  const out = { rounds: 0, charts: 0, worst: 0 };
  for (const log of logs) {
    const entries = log?.trace?.entries;
    if (!Array.isArray(entries)) continue;
    let here = 0;
    for (const e of entries) {
      if (!/^origin tag lost/.test(String(e.message))) continue;
      here += Number(e.data?.charts) || 0;
    }
    if (!here) continue;
    out.rounds++;
    out.charts += here;
    out.worst = Math.max(out.worst, here);
  }
  return out;
}

/** What each scenario cost the host — the meter nothing read for 63 rounds. */
function reportScenarioFriction(logs) {
  const o = poolScenarioFriction(logs);
  if (!o.rounds) return;
  console.log(`
  WHAT EACH SCENARIO COST THE HOST — pooled over ${o.rounds} round(s)`);
  console.log("    scenario                                    n    errors  idRefusals  emptyReReads");
  for (const r of o.rows.slice(0, 8))
    console.log(
      "    " +
        r.name.slice(0, 40).padEnd(42) +
        String(r.n).padStart(3) +
        String(r.sum.errors).padStart(9) +
        String(r.sum.idRefusals).padStart(12) +
        String(r.sum.emptyReReads).padStart(14),
    );
  // DERIVED, NOT DECLARED. A counter that has never moved and a counter that
  // never varies are both worthless, and printing them beside real numbers is
  // how a reader comes to trust one of them.
  if (o.dead.length)
    console.log(`    ${o.dead.join(", ")} has NEVER been non-zero in any scenario in any round — it measures nothing.`);
  for (const c of o.constant)
    console.log(
      `    \`${c.name}\` reports ${c.key}=${c.value} EVERY round — a constant, not a signal (its deliberate abort).`,
    );
  console.log('    This is the per-scenario answer to "did it pass on an easier host", and it has');
  console.log("    been recorded since round 023 without anything reading it.");
}

/** Charts that would not follow a drag, which no scenario counts. */
function reportUpdateShortfalls(logs) {
  const o = poolUpdateShortfalls(logs);
  if (!o.updates) return;
  console.log(`
  WHAT AN UPDATE LEFT ON THE SLIDE — pooled over ${logs.length} round(s)`);
  console.log(
    `    ${o.updates} in-place update(s) measured across ${o.rounds} round(s); ${o.blind} touched a chart
` +
      `    AT RISK — ungrouped and with no parts list, which is what a refused \`reading back
` +
      `    an ungrouped chart's shape ids\` leaves behind.`,
  );
  // THE POPULATION `atRisk` CANNOT SEE, printed beside it so a zero above is
  // never read alone. #586 groups the majority the host names and leaves the
  // remainder out of the parts tag on purpose; that chart reports `group` at
  // update time, counts as safe, and still has loose shapes in its own box.
  if (o.subsetGroups)
    console.log(
      `    PLUS ${o.strandedByDesign} shape(s) left loose ON PURPOSE across ${o.subsetGroups} subset group(s) — #586's
` +
        `    trade, not a fault, and invisible to \`atRisk\` because the host calls a subset
` +
        `    group a group. Count these before calling any zero above an all-clear.`,
    );
  else
    console.log(
      `    #586's subset branch has not run in this pool: 0 partially-grouped chart(s). It is
` +
        `    reached only by a re-read still SHORT after the settled retry, which no round
` +
        `    has recorded. Read that as a FLOOR, not a count: until 2026-08-20 the COLD
` +
        `    read's outcome was never traced at all, so "none since the retry" compared a
` +
        `    post-settle read against 42 archived COLD ones. Different units. The cold read
` +
        `    is traced now, and the next rounds can say whether the fault went or is hidden.`,
    );
  if (o.unsettledKept)
    console.log(
      `    ${o.unsettledKept} reading(s) the two host reads DISAGREED on, kept and flagged rather than
` +
        `    dropped (${o.unsettledGrowth} shape(s) of growth between them, not pooled above). The second
` +
        `    read is taken: across the archive the deck adjudicates every one of these and backs the
` +
        `    second 48 times against the first's zero. A fifth of this instrument used to vanish here.`,
    );
  // THE INSTRUMENT'S OWN ERROR RATE, which is a number worth printing: this
  // reading has produced four false positives in two days, every one of them a
  // host lag the deck later contradicted.
  if (o.deckContradicted)
    console.log(
      `    ${o.deckContradicted} reading(s) DISCARDED because the deck disagreed — they claimed more shapes
` +
        `    on a slide than it finished the round holding, which a round that only adds
` +
        `    cannot do. Host lag outlasting the settle delay; two agreeing reads inside one
` +
        `    lag window agree on the stale number.`,
    );
  const usable = o.updates - o.unitMismatch - o.deckContradicted;
  if (usable > 0) {
    console.log(`    the slide GREW by ${String(o.blindGrowth).padStart(5)} shape(s), worst single update ${o.worst}`);
    console.log(`    on charts that HAD a list ${String(o.sightedGrowth).padStart(4)} — an ordinary change of size`);
    // ONLY WITH SOMETHING TO CONCLUDE FROM. This fired on a pool whose every
    // reading had just been excluded as unit-mismatched, announcing that the
    // growth "is not happening" on the strength of no measurements at all —
    // which is a report lying in a new direction rather than the old one.
    // GATED ON THE POPULATION THE QUESTION IS ABOUT, which is charts that were
    // ungrouped AND had no parts list — the only ones that can strand anything,
    // counted per update from the host's own `type`. The round-wide
    // `not grouping` count is a poor stand-in: it says a round contained such a
    // chart, not that an UPDATE ever touched one.
    if (o.blindGrowth === 0) {
      if (o.atRisk === 0)
        console.log(
          `    NOT AN ALL-CLEAR: no update here touched a chart that could strand anything
` +
            `    (0 at risk — every one was grouped, or had its parts list). Zero growth over
` +
            `    zero at-risk charts is the question never being put.
` +
            `    The round held ${o.ungroupedCharts} ungrouped chart(s); wait for one an update actually edits.`,
        );
      else
        console.log(
          `    ZERO, over ${o.atRisk} chart(s) that COULD have stranded: "the chart grows by a whole
` + `    chart on every edit" is a danger the code describes, not something happening.`,
        );
    }
  } else {
    console.log(`    no usable reading yet — every measurement here predates the instrument's fix.`);
  }
  if (o.unitMismatch)
    console.log(
      `    ${o.unitMismatch} update(s) are NOT counted above, from the two builds whose readings
` +
        `    cannot be believed. Round 082 subtracted mismatched units — "23 stranded" beside
` +
        `    its own before:3 after:3. Round 084 measured in one unit but read the host ONCE,
` +
        `    and every non-zero number it produced was the host lagging an addGroup it had
` +
        `    already committed: four slides "grew" by 23 while the deck ended with one grouped
` +
        `    chart on each. A reading now carries \`settled\` only when two reads agreed.`,
    );
}

function reportOriginTagLosses(logs) {
  const o = poolOriginTagLosses(logs);
  if (!o.charts) return;
  console.log(`
  CHARTS THAT CANNOT FOLLOW A DRAG — pooled over ${logs.length} round(s)`);
  console.log(
    `    origin tag lost  ${String(o.charts).padStart(4)} chart(s) across ${o.rounds} round(s), worst round ${o.worst}`,
  );
  console.log(
    `    The chart is re-editable and its config is intact — what it loses is the ability to
` +
      `    redraw where the USER left it, so an update snaps it back to where it was inserted.
` +
      `    \`an update follows a moved chart\` passes while this climbs, because a scenario tests
` +
      `    ONE chart and this counts them all. Rounds 073/074 lost 9 of 19 and 8 of 17.
` +
      `    No rate is printed: only failures are traced, so there is no honest denominator.`,
  );
}

/**
 * Which slide size a round ran at — its PROFILE, as a comparable string.
 *
 * `16:9` for everything filed before 2026-08-16, and that default is a fact
 * rather than a guess: all 53 rounds archived before the field existed were run
 * on a widescreen deck, which `docs/ROUNDS.md` states once so no reader has to
 * infer it per file.
 *
 * The profile exists because a 4:3 round and a 16:9 round are DIFFERENT
 * EXPERIMENTS, and pooling them is the rounds 24-and-25 mistake — "differed only
 * in this, and were compared as though they did not". Round 077 scored 10 of 13
 * where 16:9 scored 13 of 13 twice, so the difference is large enough to swamp
 * anything a pooled number would say.
 */
/**
 * Rounds whose two independent slide-size readings disagree.
 *
 * THE ROUND THAT MADE THIS NECESSARY: 115 and 116 recorded `720x540` from the
 * pane while the driver had read `960x540` off live `PageSetup`, twice, and
 * printed it before each round started. Every profile comparison in this file
 * groups by `slideSize`, so both rounds were filed into the wrong arm — the
 * exact failure `PW_EXPECT_SIZE` exists to prevent, sailing past it because the
 * guard and the archive read different sources and nobody compared them.
 *
 * A ROUND WITH NO DRIVER READING IS NOT A ROUND THAT AGREES. It is a round with
 * one opinion, and it is skipped rather than counted as a pass — every round
 * archived before `driverSlideSize` existed is in that state, and reporting 116
 * of them as "consistent" would be a lie told by a denominator.
 */
/**
 * Is the SECOND round of a pair systematically worse than the first?
 *
 * The pair exists to separate a real fault from the host's mood, and both halves
 * have been read as two samples of ONE condition. They are not. Pooled over
 * every build this archive has run twice — 31 pairs on 2026-08-20 — the second
 * round had more post-retry failures in 15, fewer in 2, and tied in 14.
 *
 * TIES ARE NOT COIN FLIPS. Most are older rounds whose counters sat at zero
 * both times, so counting them as evidence of symmetry is how this nearly got
 * waved away as noise: 15 against 2 among the pairs that moved is not a mood.
 *
 * Only the first two rounds of a build count, so a build run three times cannot
 * vote twice — and if it is ever run three times deliberately, that is the
 * experiment this finding asks for.
 */
/**
 * How long a round's own trace says it took, in seconds — or null.
 *
 * ALREADY IN EVERY ROUND FILE AND SURFACED BY NOTHING. Every entry carries an
 * `ms` offset, so the last one is the round's span; no reader has ever printed
 * it, and it turns out to be the strongest single predictor of a bad round in
 * this archive. The slower half of the rounds whose instruments existed average
 * 5.1 post-retry failures against 2.8, and 80 deck shapes against 48.
 *
 * It matters because the second round of a pair is 2.0-2.4x slower than the
 * first in all four pairs measured, and the likeliest reason is that the first
 * round gets mined WHILE THE SECOND RUNS. A reader who cannot see the span
 * cannot tell a degraded round from a clean one, and will read the difference
 * as evidence about the code.
 *
 * Null when the trace carries no usable offsets — an unreadable span must not
 * become a fast one.
 */
export function roundSpanSeconds(log) {
  const es = log?.trace?.entries;
  if (!Array.isArray(es) || !es.length) return null;
  const ms = es.map((e) => Number(e.ms) || 0);
  // LAST MINUS FIRST, NOT THE LAST. `ms` counts from the PANE'S load, not the
  // round's start, and the pane is not reloaded between rounds — so a round that
  // inherits its pane starts its clock wherever the previous round left off.
  //
  // The first version of this took `Math.max(...)` and therefore reported every
  // reused-pane round's duration as ITS OWN plus the previous round's. That
  // produced "the second round of a pair is 2.0-2.4x slower", which was
  // published twice before the arithmetic was checked. The real figure is about
  // 1.35x. See `paneAgeAtStartSeconds` for the fact the inflated number was
  // accidentally encoding.
  const span = Math.max(...ms) - Math.min(...ms);
  return span > 0 ? Math.round(span / 1000) : null;
}

/**
 * How long the pane had already been alive when this round started, in seconds.
 *
 * THE VARIABLE THAT ACTUALLY PREDICTS A BAD ROUND, and it was hiding inside a
 * metric that had been reported as duration. `ms` counts from the pane's load,
 * so the FIRST entry's offset is the pane's age at the moment the round began.
 *
 * Rounds 110-123, split on it:
 *
 *     fresh pane (< 200s)   post-retry 0, 2, 0, 0, 0, 1, 0   deck 14-45, mostly 16
 *     reused pane           post-retry 0, 5, 7, 3, 8, 7, 2   deck 16-97, mostly 60+
 *
 * Mean post-retry 0.43 against 4.57. This is what "the second round of a pair is
 * worse" always was: the second round is the one that inherits a pane. Position,
 * profile and observer load were all stand-ins for it.
 *
 * Null when unreadable — an unknown pane age must not read as a fresh one.
 */
export function paneAgeAtStartSeconds(log) {
  const es = log?.trace?.entries;
  if (!Array.isArray(es) || !es.length) return null;
  const ms = es.map((e) => Number(e.ms)).filter((n) => Number.isFinite(n));
  return ms.length ? Math.round(Math.min(...ms) / 1000) : null;
}

/** Under this many seconds old at a round's start, a pane counts as fresh. */
export const FRESH_PANE_SECONDS = 200;

/**
 * The fallback and repair signals the trace records and nothing reads.
 *
 * 71 of the 87 distinct messages in this archive are never named by triage, the
 * gate or the cycle. Most are narrative and that is fine. These four are not:
 * each records a path the code took because its FIRST choice failed, thousands
 * of times, with nobody watching the rate. Banded by round, mean per round:
 *
 *     rounds      tagging-failed  redraw-instead  scratch-retry  scratch-wrecked
 *       1- 40          5.2             9.3            9.3            10.7
 *      41- 80          6.6             9.1           14.1            11.1
 *      81-110          0.4            12.9           14.8            11.2
 *     111-141          0.2            13.0           14.6            10.7
 *
 * TWO THINGS WERE HAPPENING AND NOBODY COULD SEE EITHER. Tagging failures
 * collapsed thirtyfold around round 81 — a real win, never verified, and if it
 * ever regresses nothing will say so. Meanwhile in-place updates began falling
 * back to a full redraw 40% more often, which is slower and touches more of the
 * deck, and that drift went unremarked across sixty rounds.
 *
 * A count that no one reads is not observability. This makes the gate say them.
 */
export const FALLBACK_SIGNALS = {
  "tagging failed — charts are not re-editable until repaired": "charts left un-tagged",
  "not updating in place — redrawing instead": "in-place update fell back to a redraw",
  // ITS OWN SIGNAL, not folded into the line above. The in-place update either
  // declines by rule or THROWS, and this tracked only the first — the same
  // two-of-three gap `poolInPlaceUpdates` had, missed when that one was fixed.
  // Kept separate rather than merged because a write the HOST threw out is a
  // different event from the differ declining work a redraw does better, and
  // drift in one must not be absorbed by the other: rounds 144 and 145 carried
  // 2 and 3 of these against a flat zero everywhere else, and nothing said so.
  "in-place update refused — redrawing instead": "in-place update refused BY THE HOST",
  "took another scratch slide after giving up on the last": "scratch slide retried",
  "giving up the scratch slide this question wrecked": "scratch slide wrecked",
};

/**
 * Each fallback's count this round against the median of its priors.
 *
 * THREE PRIORS MINIMUM, the same rule `poolScenarioPopulations` and
 * `poolGroupingOutcome` needed: a "usually" from one observation is not a
 * baseline, and this project's own noise floor is the argument.
 */
/**
 * Has the in-place chart update EVER worked?
 *
 * It was added in #405 ("Change one thing, write one shape") and #406 was titled
 * "The in-place update fired zero times and would not say why" — that PR added
 * the fallback trace to diagnose it. THE DIAGNOSTIC HAS BEEN ANSWERING EVER
 * SINCE AND NOBODY HAS READ IT: across 117 archived rounds the success line
 * `updated only the shapes that changed` appears ZERO times, while the fallback
 * fires 12-13 times a round.
 *
 * The reason is one reason, 12 of those 13: "the chart has no parts list, so its
 * nodes cannot be mapped". The remaining one is the picture path, which is
 * legitimate — a picture is not in the scene, so the scene cannot decide it.
 *
 * A feature that has never once run in production is not a feature; it is a
 * branch that costs a fallback every time. This makes the gate say so.
 */
/**
 * The in-place update has THREE outcomes and this counted two.
 *
 * `tryInPlaceUpdate` declines by rule (`not updating in place`, carrying a
 * `why`) or THROWS (`in-place update refused`, carrying an `error` and no
 * `why`). Only the first was counted, so every host-side refusal was invisible
 * here — including the three `InvalidArgument | errorLocation=Shape.textFrame`
 * in round 145 that turned out to be the last thing standing between this
 * feature and working. They sat in round 144's trace too, uncounted.
 *
 * The reason breakdown exists for the same reason. A total says a feature fell
 * back; only the reasons say whether that was the differ declining work a
 * redraw does better (correct) or the host refusing a write (a defect).
 */
function inPlaceErrorKey(err) {
  const code = err.split(" | ")[0]?.trim() || "threw";
  const where = /"errorLocation":"([^"]+)"/.exec(err)?.[1];
  return where ? `${code} at ${where}` : code;
}

export function poolInPlaceUpdates(logs) {
  let ok = 0,
    fell = 0,
    threw = 0,
    rounds = 0;
  const reasons = new Map();
  // ONE EXAMPLE, not just a tally. A residual bucket reported as a number is
  // read as noise; reported as a line it gets opened. That distinction cost two
  // rounds of a host defect sitting in plain sight.
  const unexplained = [];
  const note = (key) => reasons.set(key, (reasons.get(key) ?? 0) + 1);
  for (const log of logs ?? []) {
    const entries = log?.trace?.entries;
    if (!Array.isArray(entries)) continue;
    rounds++;
    for (const e of entries) {
      const m = String(e.message ?? "");
      if (m === "updated only the shapes that changed") ok++;
      else if (m === "not updating in place — redrawing instead") {
        fell++;
        const why = e.data?.why;
        if (why) note(String(why));
        else unexplained.push(m);
      } else if (m === "in-place update refused — redrawing instead") {
        threw++;
        const err = String(e.data?.error ?? "");
        if (err) note(inPlaceErrorKey(err));
        else unexplained.push(m);
      }
    }
  }
  return {
    ok,
    fell,
    threw,
    rounds,
    unexplained,
    reasons: [...reasons].sort((a, b) => b[1] - a[1]).map(([why, n]) => ({ why, n })),
  };
}

/**
 * What the driver had to do before each round would start.
 *
 * A successful recovery erases its own evidence — the round that follows looks
 * like one that never needed rescuing — so `driverRun` records the attempts and
 * the stops behind them. This pools that across the archive.
 *
 * REPORTED AS COUNTS, NOT A RATE, and deliberately. `driverRun` is only as old
 * as round 150, so most of the archive carries nothing, and a percentage over a
 * denominator of four would be a number with a decimal point and no meaning.
 * `rounds` is the rounds that CARRY the field; anything older is not a clean
 * round, it is an unmeasured one.
 */
/**
 * Trace signals a round RECORDS that no tool READS.
 *
 * This repo has found the same thing by hand twice. `poolFallbackRates` exists
 * because the fallback lines had been "recorded thousands of times and read by
 * nothing"; `poolInPlaceUpdates` because the in-place diagnostic "has been
 * answering ever since and nobody has read it". Both were discovered by someone
 * scrolling a round file, months after the data started arriving.
 *
 * Round 153 records 49 distinct messages and 35 of them are read by nothing.
 * Most of those are noise and should stay unread — the point is not to pool
 * them all, it is to stop the next `poolFallbackRates` waiting months for
 * someone to notice it by hand.
 *
 * A FLOOR, NOT A COUNT, and the gate says so. "Read" here means the tool source
 * mentions the message verbatim, so a matcher built by concatenation reads as
 * unread. That errs toward offering too much rather than hiding something, and
 * a detector that reports a floor must say which it is — see
 * `poolFallbackRates` and the dead-detector guard in the tests.
 */
export function unreadSignals(log, toolSource, top = 6) {
  const entries = log?.trace?.entries;
  if (!Array.isArray(entries)) return [];
  const counts = new Map();
  for (const e of entries) {
    const m = String(e?.message ?? "");
    if (m) counts.set(m, (counts.get(m) ?? 0) + 1);
  }
  const src = String(toolSource ?? "");
  return [...counts]
    .filter(([m]) => !src.includes(m))
    .sort((a, b) => b[1] - a[1])
    .slice(0, top)
    .map(([message, n]) => ({ message, n }));
}

export function poolDriverRuns(logs) {
  let rounds = 0,
    clean = 0;
  const causes = new Map();
  const attempts = new Map();
  // THE ARM OF THE EXPERIMENT — SPLIT ON THE PANE, NOT ON THE FLAG.
  //
  // The first version of this split on `driverRun.fresh`, and round 167 showed
  // within one pair why that is wrong: 166 ran WITHOUT `--fresh` and started on
  // a 69-second pane anyway, because a merge preceded it and recovery reloads a
  // stale pane. The flag and the pane are different variables, and the whole
  // PAIR POSITION argument is about the pane — `--fresh` is merely one way to
  // get a fresh one. Splitting on the flag files a fresh-pane round under
  // "aged" and makes both arms mean nothing.
  //
  // The house defect, in an instrument built four hours earlier to fix the same
  // defect: measuring the PROXY instead of the thing.
  //
  // `flagDisagreed` counts rounds where the two answers differ, so the proxy's
  // unreliability is on the page rather than assumed away. `unlabelled` counts
  // rounds from before either was recorded, so a rate over a handful is never
  // quoted as if it covered the archive.
  const arms = { fresh: 0, freshRefusedNone: 0, aged: 0, agedRefusedNone: 0, unlabelled: 0, flagDisagreed: 0 };
  for (const log of logs ?? []) {
    const dr = log?.driverRun;
    if (!dr || typeof dr.attempts !== "number") continue;
    rounds++;
    attempts.set(dr.attempts, (attempts.get(dr.attempts) ?? 0) + 1);
    if (dr.attempts <= 1) clean++;
    for (const c of dr.recovered ?? []) causes.set(String(c), (causes.get(String(c)) ?? 0) + 1);
    const age = paneAgeAtStartSeconds(log);
    if (typeof age !== "number") {
      arms.unlabelled++;
      continue;
    }
    const paneWasFresh = age < FRESH_PANE_SECONDS;
    if (typeof dr.fresh === "boolean" && dr.fresh !== paneWasFresh) arms.flagDisagreed++;
    const refusedNone = !(log?.trace?.entries ?? []).some((e) => /^not grouping/.test(String(e.message ?? "")));
    if (paneWasFresh) {
      arms.fresh++;
      if (refusedNone) arms.freshRefusedNone++;
    } else {
      arms.aged++;
      if (refusedNone) arms.agedRefusedNone++;
    }
  }
  return {
    rounds,
    clean,
    arms,
    attempts: [...attempts].sort((a, b) => a[0] - b[0]).map(([n, of]) => ({ attempts: n, rounds: of })),
    causes: [...causes].sort((a, b) => b[1] - a[1]).map(([cause, n]) => ({ cause, n })),
  };
}

/** How many rounds of a signal to print in sequence, so a step is visible. */
export const RECENT_IN_A_ROW = 8;

export function poolFallbackRates(logs) {
  const per = [];
  for (const log of logs ?? []) {
    const entries = log?.trace?.entries;
    if (!Array.isArray(entries)) continue;
    const c = {};
    for (const key of Object.keys(FALLBACK_SIGNALS)) c[key] = 0;
    for (const e of entries) {
      const m = String(e.message ?? "");
      if (m in c) c[m]++;
    }
    per.push(c);
  }
  if (per.length < MIN_PRIORS_FOR_A_BASELINE + 1) return [];
  const now = per[per.length - 1];
  const priors = per.slice(0, -1);
  // DRIFT, NOT JUST TODAY. Comparing this round against the median of ALL
  // priors cannot see a slow climb: the median absorbs it. The signal that
  // motivated this instrument rose from 9.3 to 13.0 per round over sixty
  // rounds, and by the time it was noticed "now" and "usually" were both 13 —
  // so a now-against-median check would have reported it as normal for as long
  // as it kept getting worse.
  //
  // The oldest third against the newest third catches exactly that shape, and
  // is why this returns both readings rather than one.
  const third = Math.max(1, Math.floor(per.length / 3));
  const med = (xs) => [...xs].sort((a, b) => a - b)[Math.floor(xs.length / 2)];
  const out = [];
  for (const [key, label] of Object.entries(FALLBACK_SIGNALS)) {
    const median = med(priors.map((p) => p[key]));
    const oldest = med(per.slice(0, third).map((p) => p[key]));
    const newest = med(per.slice(-third).map((p) => p[key]));
    // THE LAST FEW VALUES IN SEQUENCE, because every summary above is a median
    // and A MEDIAN CANNOT SEE A STEP. `in-place update fell back to a redraw`
    // reads 13,13,11,10,10,8,8,2,2,2,2,2 over the last twelve rounds — the
    // feature started working and the fallback collapsed — and all three
    // readings miss it: all-time median 12, last-20 median 10, and the thirds
    // say "RISING, 8 to 13". A reader is told to chase a regression that ended
    // five rounds ago, next to a `now` of 2 that contradicts it.
    //
    // A wider window is not the fix; a median over ANY fixed window smears a
    // step by construction. The sequence is the evidence, and it costs eight
    // numbers.
    const recent = per.slice(-RECENT_IN_A_ROW).map((p) => p[key]);
    out.push({ key, label, now: now[key], median, oldest, newest, recent, rounds: priors.length, span: third });
  }
  return out;
}

export function poolPairPosition(logs) {
  const KIND = (m, d) => {
    const k = String(d.kind ?? "");
    if (k) return k;
    if (/came back empty/.test(m)) return "empty";
    if (/named none/.test(m)) return "zero-match";
    if (/matched only some/.test(m)) return "short";
    return null;
  };
  const byBuild = new Map();
  for (const log of logs ?? []) {
    const build = String(log?.build ?? "").split(" ")[0];
    if (!build) continue;
    let post = 0;
    for (const e of log?.trace?.entries ?? []) {
      const d = e.data ?? {};
      if (d.afterRetry === true && KIND(String(e.message ?? ""), d)) post++;
    }
    const deck = (log?.deck?.inventory ?? []).map((s) => s.count ?? s.shapes?.length ?? 0);
    if (!byBuild.has(build)) byBuild.set(build, []);
    byBuild.get(build).push({ post, age: paneAgeAtStartSeconds(log), deck: deck.reduce((a, b) => a + b, 0) });
  }
  let pairs = 0,
    worse = 0,
    better = 0,
    tied = 0,
    secondFresh = 0;
  for (const rounds of byBuild.values()) {
    if (rounds.length < 2) continue;
    const [a, b] = rounds;
    pairs++;
    if (b.post > a.post) worse++;
    else if (b.post < a.post) better++;
    else tied++;
    // WHETHER THE SECOND ROUND STARTED ON A FRESH PANE, which is the whole
    // reason this asymmetry exists. Counted so the report can distinguish
    // pairs that PREDATE the between-rounds reload from ones that do not,
    // instead of announcing a fixed problem forever from historical data.
    if (typeof b.age === "number" && b.age < FRESH_PANE_SECONDS) secondFresh++;
  }
  return { pairs, worse, better, tied, secondFresh };
}

export function poolProfileDisagreements(logs) {
  const out = [];
  for (const log of logs ?? []) {
    const pane = log?.slideSize;
    const driver = log?.driverSlideSize;
    if (!driver || !pane || typeof pane.width !== "number") continue;
    // The driver records a PROFILE STRING ("16:9"), the pane records points.
    // Compare them as profiles, which is the unit every consumer groups by.
    const paneProfile = roundProfile(log);
    if (String(driver) !== paneProfile)
      out.push({
        build: String(log.build ?? "").split(" ")[0],
        pane: paneProfile,
        driver: String(driver),
        source: pane.source ?? "?",
      });
  }
  return out;
}

export function roundProfile(log) {
  const s = log?.slideSize;
  if (!s || typeof s.width !== "number" || typeof s.height !== "number") return "16:9";
  // Named ratios for the two PowerPoint offers, and the raw size for anything
  // else — a custom deck should be visibly its own profile, not silently folded
  // into whichever named one it is nearest.
  const r = s.width / s.height;
  if (Math.abs(r - 16 / 9) < 0.01) return "16:9";
  if (Math.abs(r - 4 / 3) < 0.01) return "4:3";
  return `${Math.round(s.width)}x${Math.round(s.height)}`;
}

/**
 * Scenarios that PASS at one slide size and FAIL at another, in the same cycle.
 *
 * A different question from the regression gate, and it needs its own answer.
 * The gate asks "did this fall against its own history"; this asks "did this
 * profile fail what another profile passed tonight". Round 077 was exactly the
 * second — 10 of 13 at 4:3 against 13 of 13 at 16:9 — and nothing would have
 * said so automatically.
 *
 * BUILD-SCOPED, because that is the only way the comparison is fair. Two rounds
 * on different builds differ for reasons that have nothing to do with slide
 * size, and this must never report those.
 *
 * A scenario failing in BOTH profiles is not divergence — it is an ordinary bug,
 * already visible everywhere else, and reporting it here would bury the one
 * signal this exists for.
 */
export function profileDivergence(logs) {
  const byBuild = new Map();
  for (const log of logs) {
    const build = String(log?.build ?? "").split(" ")[0];
    if (!build) continue;
    const prof = roundProfile(log);
    if (!byBuild.has(build)) byBuild.set(build, new Map());
    const profs = byBuild.get(build);
    if (!profs.has(prof)) profs.set(prof, new Map());
    const seen = profs.get(prof);
    for (const sc of log?.selftest ?? []) {
      if (!sc?.name) continue;
      // A scenario that DID NOT MEASURE says nothing about its profile — the
      // same rule the regression gate follows, for the same reason.
      if (sc.skipped) continue;
      // BOTH outcomes kept, not collapsed to the worst. A profile that
      // disagrees with ITSELF is flaky, not different from another profile, and
      // the two want opposite responses — see below.
      const at = seen.get(sc.name) ?? { pass: 0, fail: 0 };
      if (sc.ok) at.pass++;
      else at.fail++;
      seen.set(sc.name, at);
    }
  }
  const out = [];
  for (const [build, profs] of byBuild) {
    if (profs.size < 2) continue;
    const names = new Set([...profs.values()].flatMap((m) => [...m.keys()]));
    for (const name of names) {
      const passedIn = [];
      const failedIn = [];
      const unstableIn = [];
      for (const [prof, m] of profs) {
        const at = m.get(name);
        if (!at) continue;
        // A profile that both passed and failed this scenario on one build has
        // not told us anything about its slide size — it has told us the
        // scenario is flaky. Reporting that as divergence sends someone to
        // investigate an aspect ratio for a difference the profile produces on
        // its own, which is worse than saying nothing.
        if (at.pass && at.fail) unstableIn.push(prof);
        else if (at.pass) passedIn.push(prof);
        else failedIn.push(prof);
      }
      if (passedIn.length && failedIn.length) out.push({ build, name, passedIn, failedIn, unstableIn });
      else if (unstableIn.length && (passedIn.length || failedIn.length))
        // Worth saying, and NOT as divergence. This is the shape the check's
        // first live outing produced: `explode a degraded picture` passed at
        // 4:3 and then passed once and failed once at 16:9, on build 17a8204.
        // "Diverged between slide sizes" was true of the worst reading and
        // misleading about the cause.
        out.push({ build, name, passedIn, failedIn, unstableIn, flaky: true });
    }
  }
  return out.sort((a, b) => a.build.localeCompare(b.build) || a.name.localeCompare(b.name));
}

/**
 * Scenarios that were passing and have stopped — the only automatic check this
 * project has on a round's own result.
 *
 * WHY IT EXISTS. Rounds 070-072 took `same scale across the deck` from 35
 * consecutive failures to three consecutive passes. Nothing protected that: a
 * later build could take it back to 3 of 8 and no gate would fail, because every
 * round result in this repo is read by a person and then filed. Three rounds of
 * host time bought a result with no guard on it.
 *
 * ESTABLISHED MEANS PASSED IN ALL OF THE LAST `window` ROUNDS, and the threshold
 * is the project's own: three is what `docs/ROUNDS.md` asks for "where a claim
 * depends on it". This is what keeps the gate off the host's mood. A scenario
 * that fails half the time was never established, so its next failure is not a
 * regression and is not reported; one that has passed three times running and
 * then fails is exactly the thing a person would want stopped at.
 *
 * A scenario absent from the older rounds is NEW, not regressed — it cannot have
 * been established, and reporting it would make every added scenario look like a
 * fault on its first bad round.
 */
export function scenarioRegressions(rounds, window = 3) {
  if (!Array.isArray(rounds) || rounds.length < window + 1) return [];
  // WITHIN ONE PROFILE ONLY. A nightly cycle runs 16:9 twice and 4:3 once, so a
  // 4:3 round judged against three 16:9 rounds would be flagged for scoring
  // differently — which it does, by design. The gate would fire every night,
  // and a gate that cries wolf gets switched off; this file has already watched
  // that happen twice. See `roundProfile`.
  const profile = roundProfile(rounds[rounds.length - 1]);
  const sameProfile = rounds.filter((r) => roundProfile(r) === profile);
  if (sameProfile.length < window + 1) return [];
  rounds = sameProfile;
  // `true` passed, `false` failed, `null` DID NOT MEASURE. The third value is
  // what the gate got wrong on its first live outing: round 073 flagged `explode
  // a degraded picture` as having stopped passing, for a result whose own words
  // were "proves nothing either way". A scenario declining to conclude is an
  // absence of evidence, not a fall, and a gate that fires on one is a gate that
  // gets switched off — which this repo has already watched happen once.
  //
  // Read from the `skipped` FLAG, never from the detail text. The prose is
  // edited whenever a message is improved, and a gate keyed to it would go quiet
  // the first time someone reworded a sentence.
  const scenariosOf = (r) => {
    const out = new Map();
    for (const sc of r?.selftest ?? []) if (sc?.name) out.set(sc.name, sc.skipped ? null : !!sc.ok);
    return out;
  };
  const all = rounds.map(scenariosOf);
  const newest = all[all.length - 1];
  // The `window` rounds BEFORE the newest — the newest is what is being judged.
  const before = all.slice(-1 - window, -1);
  const out = [];
  for (const [name, ok] of newest) {
    // Passed, or did not measure — neither is a regression.
    if (ok !== false) continue;
    // Established: present AND passing in every one of the previous rounds.
    const established = before.every((r) => r.get(name) === true);
    if (!established) continue;
    // LIFETIME, not the window. `passedIn` is the window SIZE — a constant that
    // is true of every regression this function can return, so printing it as
    // "had passed the previous 3 rounds running" said nothing at all. Whether a
    // scenario has failed once in 124 rounds or twelve times is the difference
    // between a blip and a habit, and it is what decides whether to chase it.
    let ran = 0,
      failed = 0;
    for (const r of all) {
      const v = r.get(name);
      // `undefined` never ran, `null` declined to conclude. Neither is evidence.
      if (v === undefined || v === null) continue;
      ran++;
      if (!v) failed++;
    }
    out.push({ name, passedIn: window, ran, failed });
  }
  return out.sort((a, b) => a.name.localeCompare(b.name));
}

/**
 * Did the chart span sync batches, and did it group?
 *
 * THE SHARPEST SEPARATION IN THE ARCHIVE, found 2026-08-16 while chasing what
 * looked like a rasterise effect:
 *
 *     spanned batches   452 draw(s), 353 grouped = 78%
 *     one batch only    214 draw(s),  49 grouped = 23%
 *
 * And it is OURS, not the host's. `refreshShapes` is set from `spansBatches()`,
 * so only a multi-batch chart gets the pre-grouping re-read that resolves its
 * shapes by id. A single-batch chart hands `addGroup` the raw `created` proxies
 * — which this host refuses with InvalidParam 5010 — and the failed group then
 * takes the tag with it: `target.tags` comes back undefined, 155 times out of
 * 155 across the archive, every one immediately after a 5010 group.
 *
 * THE RASTERISE WAS A RED HERRING, recorded so nobody re-finds it. Draws after a
 * rasterise group 22% of the time against 27% for every other draw — nothing. It
 * looked like 22% against 93% until the arms were split by batch count, because
 * the scenarios that rasterise happen to draw small charts. That is the exact
 * confound `poolEveryDraw` already warns about, met from a different direction.
 */
export function poolBatchSpanVsGroup(logs) {
  const out = { multi: 0, multiGrouped: 0, single: 0, singleGrouped: 0 };
  for (const log of logs) {
    const entries = log?.trace?.entries;
    if (!Array.isArray(entries)) continue;
    entries.forEach((e, i) => {
      if (!/^batch issued/.test(String(e.message))) return;
      const d = e.data ?? {};
      const total = Number(d.total) || 0;
      // The LAST batch of a draw only, or a multi-batch chart is counted once
      // per batch and swamps the single-batch arm with copies of itself.
      if (Number(d.upTo) !== total || !total) return;
      const multi = total > (Number(d.perSync) || 10);
      // UNTIL THE NEXT DRAW, not a fixed number of entries away — and that is a
      // correction, not a refinement. This read `i + 4` until 2026-08-16, which
      // was true of the traces it was written against and stopped being true the
      // moment every groupable chart started re-reading the slide first: the
      // extra entries pushed the verdict out of the window and the report showed
      // ZERO single-batch draws, in the same change that altered the single-batch
      // path. An instrument that goes blind exactly where it is being used is
      // worse than no instrument.
      //
      // The outcome set was short too. `not grouping` — the honest decline — was
      // not counted at all, so a chart that declined looked like a chart that had
      // never been decided.
      for (let k = i + 1; k < entries.length; k++) {
        const m = String(entries[k].message);
        // The next draw begins: this one was never resolved either way.
        if (/^batch issued/.test(m)) break;
        const ok = /^grouped the chart/.test(m);
        const bad =
          /^not grouping/.test(m) ||
          (/grouping the chart/.test(m) && /5010/.test(JSON.stringify(entries[k].data ?? {})));
        if (!ok && !bad) continue;
        if (multi) {
          out.multi++;
          if (ok) out.multiGrouped++;
        } else {
          out.single++;
          if (ok) out.singleGrouped++;
        }
        break;
      }
    });
  }
  return out;
}

/** The batch-span split, which is where grouping is really decided. */
function reportBatchSpanVsGroup(logs) {
  const b = poolBatchSpanVsGroup(logs);
  if (!b.multi && !b.single) return;
  const pct = (n, d) => (d ? `${Math.round((100 * n) / d)}%` : "—");
  console.log(`\n  DID THE CHART SPAN BATCHES — pooled over ${logs.length} round(s)`);
  console.log(
    `    spanned batches  ${String(b.multi).padStart(4)} draw(s), ${String(b.multiGrouped).padStart(4)} grouped = ${pct(b.multiGrouped, b.multi)}`,
  );
  console.log(
    `    one batch only   ${String(b.single).padStart(4)} draw(s), ${String(b.singleGrouped).padStart(4)} grouped = ${pct(b.singleGrouped, b.single)}`,
  );
  console.log(
    `    OURS, not the host's. refreshShapes USED to be set from spansBatches(), so only a\n` +
      `    multi-batch chart got the pre-grouping re-read that resolves its shapes by id; a\n` +
      `    single-batch chart handed addGroup the raw created proxies, which this host refuses\n` +
      `    with 5010, and the failed group took the tag with it (target.tags undefined, 155 of\n` +
      `    155). Every groupable chart asks for the refresh since 2026-08-16. That removed the\n` +
      `    doomed attempt but did NOT make those charts group — rounds 066 and 067 both show\n` +
      `    5 of 5 declining with 'not grouping'. Their re-read comes back non-empty but\n` +
      `    matching none of our ids, and the settle retry only fires on empty or partial.\n` +
      `    A rasterise beforehand looked like the cause and is not: 22% vs 27% once the arms\n` +
      `    are split by batch count. The scenarios that rasterise just draw small charts.`,
  );
}

/**
 * Questions that produced NOTHING, round after round.
 *
 * Every round pays for the whole probe battery in host time and scratch slides,
 * and a question that never answers costs exactly as much as one that does. The
 * per-round report names the ones "never put" in that round; nothing said which
 * ones have never been put in ANY round, so a permanently dead question read as
 * bad luck twelve times running.
 *
 * It was found by hand on 2026-08-16 and the answer was worth the query: SIX
 * questions were 0-for-12, and four of them are the group cluster —
 * `addgroup-returns-usable`, `group-children-via-getcount`,
 * `grouped-child-by-id-from-slide`, `tag-on-group-survives`. The probe has been
 * blind on groups for twelve rounds, and rounds 064/065 then answered the most
 * important of them FROM PRODUCTION, twice, in one evening.
 *
 * SPLIT BY KIND, because the two want opposite fixes and pooling them hides
 * that:
 *
 *   never asked   `no-scratch-slide` / `no-scratch-shape` — the harness could
 *                 not set the question up. A HARNESS problem, ours to fix.
 *   unanswerable  `unreadable` — the question was put and the host would not
 *                 answer. A HOST fact, and a real (if annoying) finding.
 *
 * A question in the first group for many rounds should be moved into production
 * instrumentation or retired; the second is telling you something about the host
 * and should be left alone.
 */
export function poolStarvedQuestions(logs) {
  const NEVER = /^no-scratch/;
  const UNANSWERABLE = /^unreadable/;
  // WHICH QUESTIONS THE BUILD STILL ASKS. This report tells the reader to fix
  // or retire something, and for eight rounds it named two questions that had
  // ALREADY BEEN RETIRED — `grouped-child-by-id-from-slide` and
  // `tag-on-group-survives`, dropped on 2026-08-21 and last seen in round 149,
  // still listed at "125 round(s)" and still filed under OURS TO FIX. A pooled
  // count over the whole archive cannot tell a live starving probe from a dead
  // one, and a report that demands action on work already done is the same
  // defect as a conclusion hardcoded into an instrument: it keeps printing
  // after it stops being true.
  //
  // A WINDOW, not the newest round alone. A single sheet can be short because
  // the host died mid-probe, which would read a live question as retired; three
  // non-empty sheets is enough that a genuinely live probe appears in one, and
  // few enough that a retirement shows up within a round or two of landing.
  const recent = logs.filter((l) => (l?.hostAnswers?.answers ?? []).length).slice(-3);
  const live = new Set(recent.flatMap((l) => (l.hostAnswers.answers ?? []).map((a) => a?.id)));
  const seen = new Map();
  for (const log of logs) {
    for (const a of log?.hostAnswers?.answers ?? []) {
      if (!a?.id) continue;
      const v = a.answer == null ? "(none)" : String(a.answer);
      const t = seen.get(a.id) ?? { rounds: 0, never: 0, unanswerable: 0, answered: 0, last: "" };
      t.rounds++;
      if (NEVER.test(v)) {
        t.never++;
        t.last = v;
      } else if (UNANSWERABLE.test(v)) {
        t.unanswerable++;
        t.last = v;
      } else t.answered++;
      seen.set(a.id, t);
    }
  }
  // Only the ones that have NEVER produced an answer. A question that answers
  // sometimes is doing its job and does not belong in a report about waste.
  return (
    [...seen.entries()]
      .filter(([, t]) => t.answered === 0 && t.rounds > 1)
      // `retired` rather than dropped: the archive still holds the rounds these
      // starved in, and a reader comparing an old report to this one deserves to
      // see WHY a row moved rather than find it simply gone.
      .map(([id, t]) => ({ id, ...t, retired: !live.has(id) }))
      .sort((a, b) => b.rounds - a.rounds || a.id.localeCompare(b.id))
  );
}

/** The dead questions, named, so a starved probe stops reading as bad luck. */
/**
 * What PRODUCTION has seen of the question `shape-resolve-held-slide-proxy`
 * cannot ask.
 *
 * That probe has answered `no-scratch-shape` in all 133 archived rounds and
 * structurally cannot do better: it needs an id for a freshly added shape and
 * this host refuses to give one. The one production site that still resolves a
 * shape by id through a slide handle a sync old is `deleteShapesById`, so this
 * pools what that sweep saw.
 *
 * BOTH DIRECTIONS, and that is the point. The sweep used to trace only its
 * failures, so the archive held 133 rounds of silence that could mean either
 * "the host resolved everything" or "the sweep never ran" — and in fact it
 * means the second, which nothing could say until the positive line existed.
 */
export function poolAgedHandleResolves(logs) {
  let resolved = 0;
  let refused = 0;
  let rounds = 0;
  for (const log of logs ?? []) {
    const entries = log?.trace?.entries;
    if (!Array.isArray(entries)) continue;
    let any = false;
    for (const e of entries) {
      const m = String(e.message ?? "");
      if (m === "resolved a shape by id through a slide handle a sync old") {
        resolved += Number(e.data?.resolved ?? 0) || 0;
        any = true;
      } else if (m === "wreckage the host would not resolve") {
        refused += Number(e.data?.unresolved ?? 0) || 0;
        any = true;
      }
    }
    if (any) rounds++;
  }
  return { resolved, refused, rounds, of: (logs ?? []).length };
}

function reportStarvedQuestions(logs) {
  const dead = poolStarvedQuestions(logs);
  if (!dead.length) return;
  console.log(`\n  QUESTIONS THAT NEVER ANSWERED — pooled over ${logs.length} round(s)`);
  const show = (title, rows, note) => {
    if (!rows.length) return;
    console.log(`    ${title}`);
    for (const r of rows) console.log(`      ${r.id.padEnd(44)} ${String(r.rounds).padStart(3)} round(s)  ${r.last}`);
    console.log(`      ${note}`);
  };
  const live = dead.filter((d) => !d.retired);
  show(
    "never asked — the harness could not set the question up (OURS to fix)",
    live.filter((d) => d.never >= d.unanswerable),
    "Move it into production instrumentation or retire it; it costs a scratch slide either way.",
  );
  // THE PRODUCTION WITNESS for the one question above that will never answer
  // itself. Printed under the row that raises it, because a reader told to
  // "move it into production instrumentation" deserves to see whether that has
  // already happened and what it found.
  if (live.some((d) => d.id === "shape-resolve-held-slide-proxy")) {
    const w = poolAgedHandleResolves(logs);
    console.log(
      `      ^ production witness (deleteShapesById): ${w.resolved} resolved, ${w.refused} refused, ` +
        `across ${w.rounds} of ${w.of} round(s)`,
    );
    if (!w.rounds)
      console.log(`        The sweep only runs when a round leaves wreckage, and none has. Still unanswered.`);
  }
  show(
    "asked, and the host would not answer (a fact ABOUT the host)",
    live.filter((d) => d.unanswerable > d.never),
    "Leave it. An unanswerable question is a finding, not a failure.",
  );
  // ALREADY DONE, and said so. These rows sat in the OURS-TO-FIX bucket for
  // eight rounds after the work landed.
  show(
    "already retired — the archive remembers them, the build no longer asks",
    dead.filter((d) => d.retired),
    "Nothing to do. Here so a row that vanished reads as finished rather than as lost.",
  );
}

/** Fresh slide versus established — where a config is really decided. */
function reportFreshVsEstablished(logs) {
  const f = poolFreshVsEstablished(logs);
  if (!f.established && !f.fresh) return;
  const pct = (n, d) => (d ? `${Math.round((100 * n) / d)}%` : "—");
  console.log(`\n  WHICH SLIDE THE CHART LANDED ON — pooled over ${logs.length} round(s)`);
  console.log(
    `    slide already had shapes  ${String(f.established).padStart(3)} chart(s), ${String(f.establishedGrouped).padStart(3)} grouped = ${pct(f.establishedGrouped, f.established)}`,
  );
  console.log(
    `    freshly added, empty      ${String(f.fresh).padStart(3)} chart(s), ${String(f.freshGrouped).padStart(3)} grouped = ${pct(f.freshGrouped, f.fresh)}`,
  );
  // THE CURRENT RATE, because the pooled one is known to be unreadable and this
  // report said so in prose and printed it anyway: "pooled over 39 rounds that
  // predate the retry, so it will climb slowly and should not be read as the
  // current rate". A number a reader is told to mentally discount is a number
  // nobody can use — and the answer is not a caveat, it is a second window.
  //
  // 20 rounds: long enough to be a rate rather than an afternoon, short enough
  // that nothing before the settled retry (rounds 064/065) is in it.
  const r = recentFreshVsEstablished(logs);
  if (r && (r.established || r.fresh)) {
    console.log(`    last ${RECENT_ROUNDS} round(s) — the rate that is actually current:`);
    console.log(
      `      slide already had shapes  ${String(r.established).padStart(3)} chart(s), ${String(r.establishedGrouped).padStart(3)} grouped = ${pct(r.establishedGrouped, r.established)}`,
    );
    console.log(
      `      freshly added, empty      ${String(r.fresh).padStart(3)} chart(s), ${String(r.freshGrouped).padStart(3)} grouped = ${pct(r.freshGrouped, r.fresh)}`,
    );
    // A WINDOW IS A GUESS, A SEQUENCE IS EVIDENCE — and this population is
    // emptying, which no percentage can show. Over successive windows it reads
    // 80% at 30 rounds, 91% at 20, 100% at 12, and at 8 there are no fresh-slide
    // charts left to measure: the in-place update took them, and a chart that is
    // not redrawn never lands on a fresh slide. `0/0` rounds are the signal.
    const seq = recentFreshSequence(logs);
    if (seq.length) console.log(`      grouped/landed on a fresh slide, per round: [${seq.join(" ")}]`);
  }
  if (f.established && f.fresh)
    console.log(
      `    A chart on a freshly added slide USED NOT TO GROUP: its pre-grouping re-read came\n` +
        `    back short or empty, so it fell through ungrouped and lost its config. Since the\n` +
        `    settled retry those charts DO group — rounds 064 and 065, both, on the same two\n` +
        `    charts. The all-time percentage above is pooled over 39 rounds that predate the\n` +
        `    retry and cannot be read as the current rate; the recent window is what to quote.\n` +
        `    What such a chart still loses is the TAG, refused through the GROUP handle: the\n` +
        `    group hangs off a slide handle Office has rewritten to slides.getItem(id), and a\n` +
        `    freshly added slide's id does not round-trip on this host.\n` +
        `    See the #108-#111 saga and shape-add-held-slide-proxy.`,
    );
}

/** Grouped versus ungrouped, and what it costs a chart's config. */
function reportGroupVsTag(logs) {
  const g = poolGroupVsTag(logs);
  if (!g.grouped && !g.ungrouped) return;
  const pct = (n, d) => (d ? `${((100 * n) / d).toFixed(0)}%` : "—");
  const cov = poolGroupVsTagCoverage(logs);
  console.log(`\n  DOES GROUPING SAVE THE CONFIG — over the charts of the DECK-WIDE RESCALE, ${logs.length} round(s)`);
  // THE DENOMINATOR THE READER WOULD OTHERWISE ASSUME. These two columns are
  // joined by `data.chart`, which only the rescale scenario writes, so this is
  // one scenario's charts and not the round's. Printed because the blindness is
  // UNEVEN — it drops far more of the grouped column than the ungrouped one —
  // and an uneven filter biases the ratio rather than merely shrinking it.
  console.log(
    `    scope: ${cov.groupedSeen}/${cov.groupedTotal} grouped and ${cov.ungroupedSeen}/${cov.ungroupedTotal} ungrouped` +
      ` events carry a chart key; the rest of the round is not in these numbers.`,
  );
  console.log(
    `    grouped      ${String(g.grouped).padStart(3)} chart(s), ${String(g.groupedLost).padStart(3)} lost the tag = ${pct(g.groupedLost, g.grouped)}`,
  );
  console.log(
    `    NOT grouped  ${String(g.ungrouped).padStart(3)} chart(s), ${String(g.ungroupedLost).padStart(3)} lost the tag = ${pct(g.ungroupedLost, g.ungrouped)}`,
  );
  // Said out loud, because the number's whole value is what it implies about
  // where to work, and that was missed for eleven rounds while it sat here.
  if (g.grouped && g.ungrouped && g.ungroupedLost / g.ungrouped > 2 * (g.groupedLost / g.grouped || 0.001))
    console.log(
      `    A chart that groups usually keeps its config; one that cannot usually loses it.\n` +
        `    READ THE SPLIT ABOVE BEFORE QUOTING THIS. Until round 064 a freshly-added slide's\n` +
        `    chart COULD NOT GROUP, so it could never appear in the grouped column — this ratio\n` +
        `    was measured on a population that excluded the hard case by construction. Round 064\n` +
        `    made those charts group and they lost their tag anyway (from: group, 5010, on the\n` +
        `    slide the run had just added). The rule is closer to "a slide that was already\n` +
        `    there saves the config", with grouping standing in for it.`,
    );
}

/** The tag-fault table, and the noise floor it implies. */
function reportTagFaults(logs) {
  const byBuild = poolTagFaults(logs);
  if (byBuild.size < 2) return;
  const kinds = ["tags-undefined", "cfg-tag-5010", "group-5010", "no-queue", "tagging-failed"];
  console.log(`\n  TAG FAULTS PER BUILD — pooled over ${logs.length} round(s)`);
  console.log(`    ${"build".padEnd(9)}${"n".padEnd(4)}${kinds.map((k) => k.padStart(16)).join("")}`);
  const spread = {};
  for (const [build, rounds] of byBuild) {
    const cell = (k) => {
      const vs = rounds.map((r) => r[k]);
      const lo = Math.min(...vs);
      const hi = Math.max(...vs);
      (spread[k] ??= []).push(lo, hi);
      return (lo === hi ? String(lo) : `${lo}-${hi}`).padStart(16);
    };
    console.log(`    ${build.padEnd(9)}${String(rounds.length).padEnd(4)}${kinds.map(cell).join("")}`);
  }
  // THE NOISE FLOOR, measured WITHIN one build wherever a build has been run
  // twice — which is the only measurement that owes nothing to an argument. A
  // spread across builds can always be answered with "but the code changed";
  // the same build twice cannot.
  const within = [];
  for (const [build, rounds] of byBuild) {
    if (rounds.length < 2) continue;
    for (const k of kinds) {
      const vs = rounds.map((r) => r[k]);
      const lo = Math.min(...vs);
      const hi = Math.max(...vs);
      if (hi > lo) within.push({ build, k, lo, hi });
    }
  }
  const worst = within.sort((a, b) => b.hi - b.lo - (a.hi - a.lo))[0];
  if (worst)
    console.log(
      `    NOISE FLOOR, from one build run ${byBuild.get(worst.build).length}× : ${worst.build} scored ` +
        `${worst.lo} and ${worst.hi} for ${worst.k}\n` +
        `    with NOTHING changed between them. A difference smaller than that is not evidence of a change.\n` +
        `    Judge a change on a count that did NOT move, or on a trace line that appears where none did.\n` +
        // CORRECTED 2026-08-20. This line said the spread WAS "the host's mood",
        // which is a claim about its CAUSE and it is wrong. The spread has a
        // direction: pooled over all 31 pairs this archive holds, the SECOND
        // round of a build is worse than the first in 15 of the 17 pairs that
        // moved at all. Mood is symmetric; this is not. Calling a directional
        // effect noise is how it stayed invisible through nineteen pairs.
        `    NOT symmetric: see PAIR POSITION below — the second run is usually the worse one,\n` +
        `    so a floor measured this way includes an effect as well as noise.`,
    );
  else
    console.log(
      `    No build has been run twice, so there is no measured noise floor and no difference\n` +
        `    in this table can be attributed to a change. Run one build twice before believing it.`,
    );
}

/** The pooled arms, printed — or nothing when no round carried the scenario. */
function reportPool(logs) {
  const { rounds, arms } = poolRasteriseArms(logs);
  const total = Object.values(arms).reduce((n, a) => n + a.ok + a.stall, 0);
  if (!total) return;
  console.log(`\n  DOES A RASTERISE POISON THE NEXT DRAW — pooled over ${rounds} round(s)`);
  for (const [name, a] of Object.entries(arms)) {
    const n = a.ok + a.stall;
    const rate = n ? `${((100 * a.stall) / n).toFixed(1)}%` : "—";
    console.log(
      `    after a ${name.padEnd(11)} ${String(a.stall).padStart(3)} stalled / ${String(n).padStart(3)} drawn = ${rate}`,
    );
  }
  const n = Math.min(...Object.values(arms).map((a) => a.ok + a.stall));
  if (n < 60)
    console.log(
      `    NOT an answer yet: ${n} draws in the smaller arm. Telling rates this close apart\n` +
        `    needs nearer 60-100 an arm, which is dozens more rounds at two an arm each.`,
    );

  // The same question asked of every draw the round made, which is ~10x the
  // sample and weaker evidence. See `poolEveryDraw`.
  const every = poolEveryDraw(logs);
  const seen = Object.values(every.after).reduce((t, a) => t + a.ok + a.stall, 0);
  if (!seen) return;
  console.log(`\n  EVERY DRAW IN THE ROUND — observational, not counterbalanced`);
  for (const [name, a] of Object.entries(every.after)) {
    const m = a.ok + a.stall;
    const rate = m ? `${((100 * a.stall) / m).toFixed(1)}%` : "—";
    const label = name === "rasterise" ? "after a rasterise" : "after anything else";
    console.log(
      `    ${label.padEnd(20)} ${String(a.stall).padStart(3)} stalled / ${String(m).padStart(3)} drawn = ${rate}`,
    );
  }
  console.log(
    `    Draws follow a rasterise because of WHICH SCENARIO they belong to, and those\n` +
      `    differ in shape count and in what they ask of the host. Suspicion, not verdict —\n` +
      `    the arms above are the controlled measurement, this is the one with the numbers.`,
  );
}

/** The batch-time split and the per-slide slope, printed when a round has them. */
export function reportBatchCost(log) {
  const pop = batchPopulations(log);
  if (pop) {
    console.log(`\n  DRAW BATCHES ${pop.times.length} timed, slowest ${(pop.max / 1000).toFixed(1)}s`);
    if (pop.bimodal)
      console.log(
        `    two populations: ${pop.fast.length} fast (${(pop.fast[0] / 1000).toFixed(1)}-${(pop.fast[pop.fast.length - 1] / 1000).toFixed(1)}s) ` +
          `and ${pop.slow.length} slow (${(pop.slow[0] / 1000).toFixed(1)}-${(pop.slow[pop.slow.length - 1] / 1000).toFixed(1)}s), ` +
          `nothing in the ${(pop.gap / 1000).toFixed(1)}s between`,
      );
    else console.log(`    one population — no empty band worth the name (largest gap ${(pop.gap / 1000).toFixed(1)}s)`);
  }
}

export function reportSelfTest(selftest) {
  if (!selftest.length) return 0;
  const failed = selftest.filter((s) => !s.ok && !s.skipped).length;
  console.log(`\n  SELF-TEST ${selftest.filter((s) => s.ok).length} of ${selftest.length} scenarios passed\n`);
  for (const s of selftest) {
    const mark = s.skipped ? "skip" : s.ok ? "ok" : "FAIL";
    console.log(`  ${pad(mark, 6)}${pad(s.name, 46)}${s.detail}`);
    // A scenario's own words are a problem string like any other, and the
    // one place a host bug is stated in plain language rather than in a
    // host error code. Annotating only the `problems` tally missed it.
    const known = knownBug(s.detail ?? "");
    if (known) console.log(`  ${pad("", 6)}${pad("", 46)}  ^ known host bug: ${known}`);
  }
  console.log("");
  return failed;
}

/**
 * What the round left in the deck, from the round's OWN evidence.
 *
 * The deck-evidence change made the file self-contained on purpose — "there is
 * no deck to save and no screenshot to take" — and then nothing read the part
 * that replaced them. Working the 2026-08-09 round out by hand took a dozen
 * ad-hoc queries to reach three numbers that are right here: how many added
 * slides carry shapes, how many read back empty, and how many of the empties
 * the host's own rasteriser AGREES are empty.
 *
 * That last column is the one worth having. An empty readback and a blank
 * picture are two witnesses; an empty readback alone is one, and this host is
 * known to answer a shape collection short without throwing
 * (`shapes-items-count-honest`). A slide with no shot is not evidence of
 * anything — the rasterise pass is capped — so it is counted apart rather than
 * folded in with the confirmed blanks.
 */
export function deckEvidence(deck) {
  if (!deck || !Array.isArray(deck.inventory) || !deck.inventory.length) return null;
  const shots = new Map((deck.shots ?? []).filter((s) => s.png).map((s) => [s.slideId, s.png]));
  /**
   * How big a slide's PNG can be and still be nothing.
   *
   * This used to be `Math.min` over the very shots being classified, and that
   * cannot work: the smallest picture satisfies `<= itself` unconditionally, so
   * the tool could never report a round with NO blank in it. At least one added
   * slide was always printed under "read back empty AND rasterise blank" — the
   * two-witness line a maintainer reads as proven data loss — however much
   * chart was on it. It got the right answer on all ten real rounds only
   * because each contains genuine blanks; it breaks in exactly the regime the
   * column exists for, the round after the drawing bug is fixed, where every
   * slide has content and one answers its shape collection short (documented
   * host behaviour). The honest verdict there is "the readback is lying", and
   * the tool said "confirmed blank".
   *
   * A relative test cannot replace it either. Rounds 12 and 13 ran with the
   * picture cap at 12 and EVERY shot they took is a blank, so there is no
   * contrast in the round to measure against — judging by bimodality calls
   * those 23 genuinely blank slides content.
   *
   * So the assumption is made explicit instead of hidden in a `min`. Every
   * round this tool reads is a 960×540 deck from this add-in, and the measured
   * spread is not close: a blank is 1146 bytes across all ten rounds, and the
   * smallest picture with a chart on it is 5142. Two kilobytes sits 4.5× above
   * the one and 2.5× below the other.
   */
  const BLANK_PNG_CEILING = 2048;
  const bytes = (png) => Math.round((png.length * 3) / 4);
  const added = new Set(deck.newSlides ?? []);
  /**
   * An id whose second half is `0`, e.g. `256#0` against `288#3603562595`.
   *
   * Round 041 added seven slides and two came back shaped like that, in the same
   * round that reported `delete-by-id left slides behind and this host does not
   * list the ids they were added under` and left one added slide blank. Those
   * may be one fact or three; what is certain is that nobody saw it, because
   * both halves were in the log and nothing put them side by side. It took a
   * hand query to notice, which is the definition of a thing this tool should
   * have said.
   *
   * Reported as a SHAPE, not a diagnosis. What `#0` means to this host is not
   * known — only that it is not what a slide it has finished adding looks like.
   */
  const unfinishedId = (id) => /#0$/.test(String(id ?? ""));
  const out = {
    scanned: deck.inventory.length,
    added: added.size,
    withShapes: 0,
    confirmed: 0,
    unseen: 0,
    lying: 0,
    oddIds: [...added].filter(unfinishedId),
    oddAndBlank: 0,
  };
  for (const s of deck.inventory) {
    if (!added.has(s.slideId)) continue;
    // `count` is the host's own number and `shapes.length` is what the scan
    // managed to list; the SMALLER would call a partial listing empty, so the
    // larger is what decides whether anything is there.
    const n = Math.max(s.count ?? 0, s.shapes?.length ?? 0);
    if (n > 0) {
      out.withShapes++;
      continue;
    }
    const png = shots.get(s.slideId);
    if (!png) out.unseen++;
    else if (bytes(png) <= BLANK_PNG_CEILING) {
      out.confirmed++;
      // The join that had to be done by hand. If the blank slides ARE the
      // oddly-named ones, that is one finding rather than two — and if they are
      // not, the coincidence is dead and nobody spends a round on it.
      if (unfinishedId(s.slideId)) out.oddAndBlank++;
    }
    // A picture that is not blank and not obviously a chart cannot be called
    // either way — and calling it blank is the expensive mistake, because that
    // is the reading that alleges data loss.
    else out.lying++;
  }
  return out;
}

function reportDeckEvidence(deck) {
  const e = deckEvidence(deck);
  if (!e) return;
  console.log(
    `\n  DECK EVIDENCE ${e.scanned} slide(s) scanned, ${e.added} added by this round:` +
      `\n    ${e.withShapes} carry shapes` +
      `\n    ${e.confirmed} read back empty AND rasterise blank` +
      `\n    ${e.unseen} read back empty, no picture taken (the rasterise pass is capped — not evidence)` +
      (e.lying ? `\n    ${e.lying} read back empty but rasterise with CONTENT — the readback is lying` : "") +
      (e.oddIds.length
        ? `\n    ${e.oddIds.length} added with an id ending #0 (${e.oddIds.slice(0, 4).join(", ")}` +
          `${e.oddIds.length > 4 ? ", …" : ""}) — not the shape this host gives a slide it has finished adding` +
          `\n      of which ${e.oddAndBlank} also read back empty AND rasterised blank`
        : ""),
  );
  if (deck.gap) console.log(`    scan gap: ${deck.gap}`);
}

/**
 * Did an in-place update leave the rest of an ungrouped chart on the slide?
 *
 * The pooled reading for `shapes left on the slide after an in-place update`.
 * The question it settles has been open for 56 rounds: `reading back an
 * ungrouped chart's shape ids` is the last `GetItem(id)` refusal still firing,
 * each failure costs a chart its parts list, and the comment beside that read
 * says such a chart "grows by a whole chart on every edit" — which nothing has
 * ever observed, because no round recorded whether one of those charts was the
 * one an update touched.
 *
 * SPLIT BY WHETHER THE CHART COULD HAVE STRANDED ANYTHING, because that is the
 * whole point. Growth on a chart that had its parts list is an ordinary change
 * of size; growth on one that was UNGROUPED AND UNLISTED is the stranding.
 * Pooling the two would destroy the only distinction the instrument draws.
 *
 * `atRisk` is that population, counted from the host's own shape type, and the
 * report refuses to call zero growth an all-clear when it is zero — a grouped
 * chart is deleted whole and can strand nothing, so a round that grouped
 * everything never put the question. Round 082 was exactly that: 20 of 20
 * grouped, and it would have read as an all-clear.
 *
 * Readings from before 2026-08-16 carry `shortfall`/`unexplained` instead of
 * `growth` — a subtraction across three different units that summed to zero on
 * every line. They are counted as `unitMismatch` and never mixed in.
 */
export function poolUpdateShortfalls(logs) {
  const out = {
    rounds: 0,
    updates: 0,
    blind: 0,
    blindGrowth: 0,
    sightedGrowth: 0,
    worst: 0,
    unitMismatch: 0,
    ungroupedCharts: 0,
    atRisk: 0,
    /** Readings whose two host reads disagreed — kept, flagged, never pooled. */
    unsettledKept: 0,
    unsettledGrowth: 0,
    /** Shapes #586 left loose ON PURPOSE, and the charts that left them. */
    strandedByDesign: 0,
    subsetGroups: 0,
    deckContradicted: 0,
  };
  for (const log of logs) {
    const entries = log?.trace?.entries;
    if (!Array.isArray(entries)) continue;
    // THE DECK IS THE ONE SOURCE THAT HAS NEVER BEEN WRONG HERE. It is taken at
    // the end of the round, long after any host lag has settled, and it has
    // caught three phantom readings running — the 084 investigation and both of
    // round 086's.
    //
    // A round only ADDS shapes to the slides it keeps, so a reading claiming
    // MORE shapes than the slide finished with is claiming shapes that never
    // existed. Round 086 read `after: 24` on a slide the inventory then showed
    // holding 1, and both of its host reads agreed on the 24 — the lag outlasted
    // the settle delay, so `settled` was true and wrong.
    const finalCount = new Map();
    for (const s of log?.deck?.inventory ?? []) {
      const id = s.slideId ?? s.id;
      const n = s.count ?? s.shapes?.length;
      if (id && typeof n === "number") finalCount.set(id, n);
    }
    // WHETHER THE QUESTION COULD BE PUT AT ALL. A round that grouped everything
    // cannot answer this either way, and round 082 was exactly that: 20 grouped,
    // 0 not. Without this the report would read "no growth" from such a round
    // and sound like an all-clear.
    out.ungroupedCharts += entries.filter((e) => /^not grouping/.test(String(e.message))).length;
    // AND "A GROUP IS DELETED WHOLE" STOPPED BEING TRUE ON 2026-08-19. This
    // comment used to say stranding was possible only for an ungrouped chart.
    // #586 groups the majority the host will name and, in its own words, leaves
    // "the stranded remainder deliberately not written into the parts tag" —
    // so a chart can now be GROUPED and still leave shapes loose inside its own
    // box, and the next update deletes the group and walks past them.
    //
    // `atRisk` cannot see it. It is read from the host's own shape type at
    // update time (`powerpoint.ts`), where a subset group and a whole one are
    // the same word, and telling them apart there would cost a load per chart.
    // The ROUND file has both halves though — the draw pass records
    // `partial=N left=i:k` — so the join belongs here, and a zero from `atRisk`
    // can no longer be read as an all-clear on its own.
    for (const e of entries) {
      if (String(e.message) !== "grouped the chart's shapes") continue;
      // `left` is `index:count` per partially-grouped chart, comma-joined.
      for (const pair of String(e.data?.left ?? "").split(",")) {
        const k = Number(pair.split(":")[1]);
        if (Number.isFinite(k) && k > 0) {
          out.strandedByDesign += k;
          out.subsetGroups += 1;
        }
      }
    }
    let seen = false;
    for (const e of entries) {
      if (!/^shapes left on the slide/.test(String(e.message))) continue;
      const d = e.data ?? {};
      // `charts` and `withParts` are no longer read here. They said how many
      // charts an update touched and how many carried a parts list, and this
      // pool used the second as a stand-in for "could strand something" — which
      // round 086 disproved, because a grouped chart has no parts list either.
      // `atRisk` names the population directly. Both fields stay in the trace:
      // they are still worth having when reading a line by hand.
      seen = true;
      out.updates++;
      // ROUND 082 AND EARLIER SPOKE A DIFFERENT LANGUAGE. Those entries carry
      // `shortfall`/`unexplained` — a subtraction across three different units
      // that summed to zero on every line and measured nothing. Counted, never
      // mixed in: pooling them with a real growth would resurrect the artifact
      // this pool was rewritten to stop reporting.
      // Two shapes of unusable reading, counted together because the
      // response is the same: do not pool it. Pre-2026-08-16 entries speak the
      // mismatched-unit language; round 084 speaks `growth` but from a single
      // host read, and every non-zero number it produced was the host lagging a
      // group it had already committed.
      // ABSENT and FALSE are different now, and lumping them lost the
      // distinction. Absent means a pre-2026-08-16 build whose numbers are in
      // mismatched units and cannot be read at all. FALSE means a reading the
      // instrument deliberately kept: the two host reads disagreed, and the
      // SECOND one was taken because across the archive the deck adjudicates all
      // 76 such readings and backs the second 48 times against the first's zero.
      //
      // Still not pooled with settled readings — an unsettled number is weaker
      // evidence and mixing them would hide that. Counted and reported, because
      // a fifth of this instrument's output used to vanish in silence.
      if (d.growth === undefined || d.settled === undefined) {
        out.unitMismatch++;
        continue;
      }
      if (d.settled === false) {
        out.unsettledKept++;
        if (Number(d.growth) > 0) out.unsettledGrowth += Number(d.growth);
        continue;
      }
      // Checked BEFORE the reading is pooled, and counted rather than dropped
      // silently: an instrument's own error rate is a number worth reporting,
      // and this one has had four false readings in two days.
      const ended = finalCount.get(d.slideId);
      if (typeof ended === "number" && Number(d.after) > ended) {
        out.deckContradicted++;
        continue;
      }
      const atRisk = Number(d.atRisk) || 0;
      out.atRisk += atRisk;
      const growth = Number(d.growth) || 0;
      // AT RISK, not merely list-less. This bucketed on `withParts === 0`, and a
      // GROUPED chart has no parts list either — so every grouped chart landed
      // in the stranding column, where by construction it cannot belong: a group
      // is deleted whole and leaves nothing behind. Round 086 put a growth of 23
      // there from a chart whose own line said `atRisk: 0`.
      //
      // `atRisk` is the population the question is about: ungrouped AND with no
      // parts list, read from the host's own shape type. Growth anywhere else is
      // an ordinary change of size or an instrument artifact, and either way it
      // is not stranding.
      if (atRisk > 0) {
        out.blind++;
        out.blindGrowth += growth;
        out.worst = Math.max(out.worst, growth);
      } else out.sightedGrowth += growth;
    }
    if (seen) out.rounds++;
  }
  return out;
}

/**
 * One trace entry reduced to the SHAPE of the thing it says.
 *
 * Numbers and ids are the noise here: `repaired 3 tags on slide 7f2a91c4` and
 * `repaired 5 tags on slide 0b41ee02` are one event happening twice, and a
 * signature that keeps their digits counts them as two things that each
 * happened once. So digits collapse to `N` and hex runs to `#`.
 *
 * THE ERROR CLASS IS PART OF THE SHAPE, and leaving it out cost a real reading.
 * A first cut hashed `scope` and `message` alone, and round 077's 52
 * `UnexpectedError`s — the single loudest fact in that round — did not appear in
 * its report at all, because they live in `data.error` and every one of them
 * shared a message with its healthy counterpart. Folding the class in costs
 * almost nothing: the archive's whole vocabulary grew from 74 signatures to 81
 * across 58 rounds.
 *
 * The class only, never the full text — error strings carry ids and offsets, and
 * hashing those would give every failure its own signature and report a
 * vocabulary that grows forever.
 */
export function traceSignature(entry) {
  const norm = (s) =>
    String(s ?? "")
      .replace(/[0-9a-f]{8,}/gi, "#")
      .replace(/\d+/g, "N");
  const err = entry?.data?.error;
  const cls = err ? "!" + (norm(err).match(/[A-Za-z][A-Za-z0-9_]+/)?.[0] ?? "err") : "";
  return `${entry?.scope ?? "?"}|${norm(entry?.message).slice(0, 60)}${cls}`;
}

/**
 * What did the NEWEST round say that its 57 predecessors did not?
 *
 * The instrument this exists to replace is a person reading 95K characters of
 * trace. Measured on round 081: 512 entries, 44 distinct signatures, NONE of
 * them new against every prior round. That is the ordinary case, and paying
 * ~50k tokens to rediscover it is the waste — two rounds a night is ~100k tokens
 * of reading to learn nothing.
 *
 * THIS DOES NOT REPLACE READING THE TRACE. It routes it. `docs/ROUNDS.md` is
 * explicit that the headline is never all of a round says, and a summary that
 * stood in for the trace would be that mistake with a script behind it. What
 * this does is count, which a person reading 512 entries cannot do reliably —
 * rounds 077 and 078 are the same two signatures at the same two counts, and
 * establishing that by eye took an entire extra round.
 *
 * THREE BUCKETS, AND THE SPLIT IS THE POINT. A count that leaves a baseline it
 * had is a different event from a shape the archive has barely produced, and
 * that one from a shape it has never produced at all.
 *
 * The archive shows all three. At round 077 the `UnexpectedError` signatures
 * were NOVEL — a first appearance, 18 of them — while `settle pass: repaired
 * every config tag` beside them had been seen rarely and jumped to 14. At rounds
 * 079 and 081 the only report is `re-reading the slide's shapes again after a
 * settle delay`, 11 times against a median of 0: the re-read EXISTED before
 * build 17a8204 and was rare, and `needsPreGroupRefresh` widened which charts
 * reach it. That is a fix working, reported as such. Filing it under the same
 * heading as round 077's errors would mean every fix this project lands
 * announces itself as a fault on the night it starts working, and
 * `docs/BACKLOG.md` records what happens to a report like that: it gets ignored,
 * then switched off.
 *
 * It is the same mistake `profileDivergence` made on its first live outing —
 * collapsing two causes into one worst-case reading — caught there by reading
 * the output against the rounds it described. Split here from the start.
 */
export function traceNovelty(rounds, { minCount = 10, factor = 3, window = 5 } = {}) {
  const entriesOf = (r) => (Array.isArray(r?.trace?.entries) ? r.trace.entries : []);
  const countsOf = (r) => {
    const m = new Map();
    for (const e of entriesOf(r)) {
      const s = traceSignature(e);
      m.set(s, (m.get(s) ?? 0) + 1);
    }
    return m;
  };
  if (!rounds?.length) return { build: null, novel: [], sinceBuild: [], spikes: [], priors: 0, vocabulary: 0 };
  const newest = rounds[rounds.length - 1];
  const priors = rounds.slice(0, -1).map(countsOf);
  const vocabulary = new Set(priors.flatMap((m) => [...m.keys()]));
  const seen = new Set(vocabulary);
  const novel = [];
  const sinceBuild = [];
  const spikes = [];
  // A BASELINE NEEDS ROUNDS TO BE A BASELINE. With three prior rounds a median
  // is whatever the middle one happened to do, and every count looks like a
  // spike against it. The archive has 57, so this only guards the fresh-clone
  // case — but it guards it silently rather than reporting an alarm built on
  // nothing.
  const enough = priors.length >= 5;
  for (const [sig, n] of countsOf(newest)) {
    if (!seen.has(sig)) {
      novel.push({ sig, n });
      continue;
    }
    if (!enough || n < minCount) continue;
    // A RECENT WINDOW, not the whole archive. Taking the median over every
    // prior round meant a signature stayed "new behaviour" until it had
    // appeared in more than HALF the archive — so one that shipped twenty
    // rounds ago was still being announced as new, and the denominator kept
    // growing underneath it.
    //
    // Measured: `re-reading the slide's shapes again after a settle delay` first
    // appeared in round 064 and has sat at 10-11 ever since. It was reported as
    // NEW BEHAVIOUR in fifteen separate rounds and blamed on nine different
    // builds, the last of them a commit that only changed a slide counter. That
    // is the "cries wolf, gets switched off" failure this file's own header
    // warns about, and it drowned the one signature round 086 had actually
    // changed.
    //
    // Five rounds is the same order as this project's noise floor — 1-versus-5
    // for the same fault with nothing changed — so a signature absent from all
    // five and present now is genuinely new to recent history.
    const recent = priors.slice(-window);
    const vals = recent.map((m) => m.get(sig) ?? 0).sort((a, b) => a - b);
    const median = vals[Math.floor(vals.length / 2)];
    // The median is the honest centre for a host whose mood swings 4-of-5 to
    // 1-of-5 with nothing changed; a mean would let one bad night set the
    // baseline for every round after it.
    if (median === 0) {
      // THE BUILD IT STARTED IN, not the build being judged. This reported the
      // newest round's build unconditionally, which is how nine innocent
      // commits got named for one 064-era signature.
      const at = priors.findIndex((m) => (m.get(sig) ?? 0) > 0);
      const startedIn = at === -1 ? null : String(rounds[at]?.build ?? "").split(" ")[0] || null;
      sinceBuild.push({ sig, n, median, startedIn });
    } else if (n > factor * median) spikes.push({ sig, n, median });
  }
  const bySize = (a, b) => b.n - a.n;
  return {
    build: String(newest?.build ?? "").split(" ")[0] || null,
    novel: novel.sort(bySize),
    sinceBuild: sinceBuild.sort(bySize),
    spikes: spikes.sort(bySize),
    priors: priors.length,
    vocabulary: vocabulary.size,
  };
}

const invokedDirectly = process.argv[1] && process.argv[1].endsWith("triage.mjs");
if (invokedDirectly) {
  const args = process.argv.slice(2);
  const flags = args.filter((a) => a.startsWith("--"));
  const paths = args.filter((a) => !a.startsWith("--"));
  const deckPath = paths.find((p) => p.endsWith(".pptx"));
  // A DIRECTORY expands to the rounds inside it, which is how `npm run rounds`
  // stays correct as the archive grows. The obvious alternative — listing the
  // files in package.json, or `rounds/*.json` — fails twice over: the list has
  // to be edited for every round anyone adds, and npm scripts run through
  // cmd.exe on Windows, which does not expand a glob at all. Both failures are
  // silent, and both end with the pooled view quietly reading fewer rounds than
  // the archive holds, which is the exact state this archive exists to end.
  const logPaths = paths.flatMap((p) => {
    if (p.endsWith(".json")) return [p];
    let entries;
    try {
      entries = readdirSync(p);
    } catch {
      return [];
    }
    // ROUNDS ONLY — `NNN-<build>.json`, the naming `test/rounds.test.ts`
    // enforces. Taking every `.json` in the directory looked obviously right and
    // was wrong within the hour: `rounds/predictions.json` lives beside them, so
    // the ledger was counted as a round. The stability header said 5 rounds
    // while the pooling underneath said 4, and the open prediction was judged
    // against the ledger rather than against the newest round — reporting
    // "not on this sheet" for a question that was answered.
    return entries
      .filter((f) => /^\d{3}-.*\.json$/.test(f))
      .sort()
      .map((f) => join(p, f));
  });
  /**
   * The round REPORTED ON is the newest; every round is still pooled.
   *
   * A directory expands sorted, and this read `[0]` — the OLDEST round in the
   * archive. So `npm run rounds`, the one command a loop glances at between
   * rounds, printed two-day-old deck evidence and a two-day-old self-test above
   * a current grid, with nothing saying they came from different days. Nobody
   * noticed for nineteen rounds because the grid underneath was right.
   *
   * An explicit list of files keeps its order — someone naming two rounds means
   * the first one — so only the directory case flips.
   */
  const namedDirectly = paths.some((p) => p.endsWith(".json"));
  const logPath = namedDirectly ? logPaths[0] : logPaths[logPaths.length - 1];
  if (!logPath) {
    console.error("usage: node scripts/triage.mjs <deck.pptx> <run-log.json> [--all] [--json]");
    console.error("       node scripts/triage.mjs <crashed-run.json>            (no deck needed)");
    console.error("       node scripts/triage.mjs <round-*.json>                (pools the counterbalanced arms)");
    process.exit(2);
  }
  let log;
  try {
    log = JSON.parse(readFileSync(logPath, "utf8"));
  } catch (err) {
    console.error(`could not read the run: ${err.message}`);
    process.exit(2);
  }
  // Every round given, for the pooled view. Unreadable ones are skipped rather
  // than fatal: pooling is a bonus on top of reading the FIRST file, and one
  // bad archive should not stop the run in front of you being triaged.
  const pooled = logPaths
    .map((p) => {
      try {
        return JSON.parse(readFileSync(p, "utf8"));
      } catch {
        return null;
      }
    })
    .filter(Boolean);
  // A crashed run answers on its own. There is no deck to join it to — the run
  // never reached the point of producing one, which is the whole reason this
  // file exists — so requiring a .pptx would refuse the one artifact a crash
  // leaves behind.
  const crash = crashLogIn(log);
  if (crash && !deckPath) {
    const ran = crash.finishedAt ? "finished" : "NEVER REPORTED FINISHING";
    console.log(
      `\n  CRASHED RUN "${crash.label}" — ${ran}` +
        `\n  build ${crash.build} · ${crash.host}` +
        `\n  started ${crash.startedAt}${crash.dropped ? ` · ${crash.dropped} earlier step(s) dropped` : ""}` +
        `\n  ${crash.steps.length} step(s), OLDEST FIRST — the last line is where it stopped\n`,
    );
    for (const line of crash.steps) console.log(`    ${line}`);
    // What the run CONCLUDED, as opposed to what it narrated.
    //
    // A crashed round's steps say a scenario started; they never said what it
    // decided, because verdicts only existed in memory until the run ended and
    // these runs do not end. They are recorded as they land now, so a crash log
    // carries them — and this is the reader that has to show them, or they are
    // back to being invisible in a different place.
    if (crash.findings?.length) console.log(`\n  ${crash.findings.length} finding(s) banked before it stopped:\n`);
    for (const { key, value } of crash.findings ?? []) {
      console.log(`    ${key}: ${describeFinding(value)}`);
    }
    const known = crash.steps.map((l) => knownBug(l)).filter(Boolean);
    if (known.length) console.log(`\n  known host bug: ${[...new Set(known)].join("\n  known host bug: ")}`);
    console.log("");
    process.exit(crash.finishedAt ? 0 : 1);
  }
  // A round with no .pptx is the NORMAL case now, and this used to refuse it.
  //
  // The starred runbook step tells the owner to send one file and nothing else,
  // because the round carries every shape on every slide and the host's own
  // picture of each — "there is no deck to save and no screenshot to take". A
  // file produced by following that instruction landed here on 2026-08-09 with
  // 12 self-test verdicts and 246 trace entries, and got `usage:` and exit 2.
  // It was then read by hand, which is the job this script exists to remove.
  //
  // Only the slot join needs the .pptx. The self-test, the trace and the round's
  // own deck evidence are all readable without one, so they are reported and the
  // join is skipped — and the exit code still means what it meant.
  if (!deckPath) {
    const selftest = selfTestIn(log);
    if (!selftest.length && !log?.trace && !log?.deck) {
      console.error("usage: node scripts/triage.mjs <deck.pptx> <run-log.json> [--all] [--json]");
      console.error("       node scripts/triage.mjs <crashed-run.json>            (no deck needed)");
      console.error("       node scripts/triage.mjs <round.json>                  (self-test / trace half)");
      process.exit(2);
    }
    const runs = runsIn(log).filter((r) => r?.run);
    console.log(
      `\n  ROUND — no deck supplied, so slots are not joined` +
        `\n  build ${log.build ?? "?"} · ${log.host ?? "?"}` +
        (runs.length ? `\n  ${runs.length} insert run(s) in this file — pass the .pptx to check their slots` : ""),
    );
    reportDeckEvidence(log.deck);
    reportTrace(log.trace);
    reportBatchCost(log);
    const failed = reportSelfTest(selftest);
    reportStability(pooled);
    reportPredictions(pooled);
    reportFreshVsEstablished(pooled);
    reportGroupVsTag(pooled);
    reportBatchSpanVsGroup(pooled);
    reportOriginTagLosses(pooled);
    reportUpdateShortfalls(pooled);
    reportScenarioFriction(pooled);
    reportStarvedQuestions(pooled);
    reportTagFaults(pooled);
    reportPool(pooled);
    process.exit(failed ? 1 : 0);
  }
  let deck;
  try {
    deck = await readDeck(deckPath);
  } catch (err) {
    console.error(`could not read the run: ${err.message}`);
    process.exit(2);
  }
  const runs = runsIn(log);
  const results = runs.map((run) => ({ run, t: triage(deck, run) }));
  const selftest = selfTestIn(log);
  const faults = faultsIn(deck);
  if (flags.includes("--json")) {
    console.log(
      JSON.stringify({ faults, selftest, runs: results.map(({ run, t }) => ({ path: run.path, ...t })) }, null, 2),
    );
  } else {
    // Faults belong to the FILE, not to a run, so they are reported once
    // above the runs rather than repeated under each of them.
    if (faults.length) {
      console.log(`\n  ${faults.length} STRUCTURAL FAULT(S) — this repo wrote the file wrong:`);
      for (const f of faults) console.log(`    - ${f}`);
    }
    for (const { run, t } of results) report(deck, log, run, t, flags.includes("--all"));
    reportDeckEvidence(log.deck);
    // The trace belongs to the FILE, like the faults above, and it was reachable
    // only from the no-deck branch. So the tool's own documented invocation —
    // `triage.mjs <deck.pptx> <run-log.json>`, the one in its usage line and in
    // CLAUDE.md — printed 28 FEWER lines than the degraded one-file form and
    // said nothing about it. What went missing: the entry histogram, "phases an
    // error escaped", the problems tally, and every `known host bug: office-js#…`
    // annotation — the most locating lines in the whole log. A round carrying a
    // trace and no self-test went further and reported "this log holds no runs
    // and no self-test" over 186 entries, then exited 0.
    reportTrace(log.trace);
    reportBatchCost(log);
    reportSelfTest(selftest);
    reportStability(pooled);
    reportPredictions(pooled);
    reportFreshVsEstablished(pooled);
    reportGroupVsTag(pooled);
    reportBatchSpanVsGroup(pooled);
    reportOriginTagLosses(pooled);
    reportStarvedQuestions(pooled);
    reportTagFaults(pooled);
    reportPool(pooled);
    if (!results.length && !selftest.length && !log?.trace)
      console.log("\n  this log holds no runs and no self-test\n");
  }
  const disagreements =
    results.reduce((n, { t }) => n + t.disagreements, 0) + selftest.filter((s) => !s.ok && !s.skipped).length;
  process.exit(disagreements || faults.length ? 1 : 0);
}
