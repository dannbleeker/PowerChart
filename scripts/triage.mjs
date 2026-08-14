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
import { readFileSync } from "fs";
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
  const out = { scanned: deck.inventory.length, added: added.size, withShapes: 0, confirmed: 0, unseen: 0, lying: 0 };
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
    else if (bytes(png) <= BLANK_PNG_CEILING) out.confirmed++;
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
      (e.lying ? `\n    ${e.lying} read back empty but rasterise with CONTENT — the readback is lying` : ""),
  );
  if (deck.gap) console.log(`    scan gap: ${deck.gap}`);
}

const invokedDirectly = process.argv[1] && process.argv[1].endsWith("triage.mjs");
if (invokedDirectly) {
  const args = process.argv.slice(2);
  const flags = args.filter((a) => a.startsWith("--"));
  const paths = args.filter((a) => !a.startsWith("--"));
  const deckPath = paths.find((p) => p.endsWith(".pptx"));
  const logPaths = paths.filter((p) => p.endsWith(".json"));
  const logPath = logPaths[0];
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
    reportPool(pooled);
    if (!results.length && !selftest.length && !log?.trace)
      console.log("\n  this log holds no runs and no self-test\n");
  }
  const disagreements =
    results.reduce((n, { t }) => n + t.disagreements, 0) + selftest.filter((s) => !s.ok && !s.skipped).length;
  process.exit(disagreements || faults.length ? 1 : 0);
}
