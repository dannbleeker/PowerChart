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
export function poolEveryDraw(logs) {
  const isDraw = (e) =>
    (e.scope === "draw" && e.message === "batch issued") ||
    /^after a (?:rasterise|cheap read) #\d+: drawing$/.test(String(e.data?.what ?? ""));
  // A rasterise EVENT is never itself a draw. The scenario's own arm markers say
  // "after a rasterise #0: drawing" — that is a draw which FOLLOWS a rasterise,
  // and reading it as one would tar the next draw with a rasterise that had
  // already been accounted for. Caught by the test below rather than by reading:
  // the untagged-draw case came out one short and the miscount was this.
  const isRasterise = (e) => !isDraw(e) && /rasteris/i.test(`${String(e.data?.what ?? "")} ${String(e.message ?? "")}`);
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
  if (c.kind === "scenario-passes") {
    const st = log?.selftest ?? {};
    const all = Object.keys(st).map((k) => st[k]);
    const wrong = (c.names ?? []).filter((n) => {
      const s = all.find((x) => x?.name === n);
      return !s || !s.ok;
    });
    return wrong.length ? { verdict: "FAILED", why: wrong.join(", ") } : { verdict: "held", why: "" };
  }
  if (c.kind === "probe-detail-matches") {
    const a = seen(c.id);
    if (!a) return { verdict: "undetermined", why: `${c.id} not on this sheet` };
    // A never-put question has no detail worth matching, and calling that a
    // failure would blame the prediction for the host refusing the question.
    if (NOT_ASKED.has(a.answer)) return { verdict: "undetermined", why: `${c.id} was never put (${a.answer})` };
    const hit = new RegExp(c.pattern, "i").test(String(a.detail ?? ""));
    return { verdict: hit ? "held" : "FAILED", why: String(a.detail ?? "").slice(0, 90) };
  }
  return { verdict: "undetermined", why: `unknown claim kind ${String(c.kind)}` };
}

/** Open predictions, judged against the newest round given. */
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
    const madeAt = logs.findIndex((l) => buildOf(l) === p.afterBuild);
    const since = madeAt === -1 ? logs : logs.slice(madeAt + 1);
    const judged = since[since.length - 1];
    if (!judged) {
      console.log(`    no round yet   ${p.id}  (${p.madeIn}, made on ${p.afterBuild})`);
      continue;
    }
    const { verdict, why } = judgePrediction(p, judged);
    console.log(`    ${verdict.padEnd(13)} ${p.id}  (${p.madeIn}) — judged on ${buildOf(judged)}`);
    if (why) console.log(`                  ${why}`);
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

/** Charts that would not follow a drag, which no scenario counts. */
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
  const newest = scenariosOf(rounds[rounds.length - 1]);
  // The `window` rounds BEFORE the newest — the newest is what is being judged.
  const before = rounds.slice(-1 - window, -1).map(scenariosOf);
  const out = [];
  for (const [name, ok] of newest) {
    // Passed, or did not measure — neither is a regression.
    if (ok !== false) continue;
    // Established: present AND passing in every one of the previous rounds.
    const established = before.every((r) => r.get(name) === true);
    if (established) out.push({ name, passedIn: window });
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
  return [...seen.entries()]
    .filter(([, t]) => t.answered === 0 && t.rounds > 1)
    .map(([id, t]) => ({ id, ...t }))
    .sort((a, b) => b.rounds - a.rounds || a.id.localeCompare(b.id));
}

/** The dead questions, named, so a starved probe stops reading as bad luck. */
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
  show(
    "never asked — the harness could not set the question up (OURS to fix)",
    dead.filter((d) => d.never >= d.unanswerable),
    "Move it into production instrumentation or retire it; it costs a scratch slide either way.",
  );
  show(
    "asked, and the host would not answer (a fact ABOUT the host)",
    dead.filter((d) => d.unanswerable > d.never),
    "Leave it. An unanswerable question is a finding, not a failure.",
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
  if (f.established && f.fresh)
    console.log(
      `    A chart on a freshly added slide USED NOT TO GROUP: its pre-grouping re-read came\n` +
        `    back short or empty, so it fell through ungrouped and lost its config. Since the\n` +
        `    settled retry those charts DO group — rounds 064 and 065, both, on the same two\n` +
        `    charts. The percentage above is pooled over 39 rounds that predate the retry, so\n` +
        `    it will climb slowly and should not be read as the current rate.\n` +
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
  console.log(`\n  DOES GROUPING SAVE THE CONFIG — pooled over ${logs.length} round(s)`);
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
        `    with NOTHING changed between them. A difference smaller than that is the host's mood.\n` +
        `    Judge a change on a count that did NOT move, or on a trace line that appears where none did.`,
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
export function traceNovelty(rounds, { minCount = 10, factor = 3 } = {}) {
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
    const vals = priors.map((m) => m.get(sig) ?? 0).sort((a, b) => a - b);
    const median = vals[Math.floor(vals.length / 2)];
    // The median is the honest centre for a host whose mood swings 4-of-5 to
    // 1-of-5 with nothing changed; a mean would let one bad night set the
    // baseline for every round after it.
    if (median === 0) sinceBuild.push({ sig, n, median });
    else if (n > factor * median) spikes.push({ sig, n, median });
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
