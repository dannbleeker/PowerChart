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
  if (row.config) {
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
  return `${kind} ${row.shapes}sh ${row.config ? "config" : "no-config"}${row.stamped ? " NOT-COMPLETE" : ""}`;
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

  const trace = run.trace;
  if (trace?.summary) {
    console.log(`\n  TRACE ${trace.entries?.length ?? 0} entries${trace.dropped ? `, ${trace.dropped} dropped` : ""}`);
    for (const s of trace.summary.steps.slice(0, 8)) console.log(`    ${pad(s.n, 6)}${s.scope}  ${s.message}`);
    if (trace.summary.problems.length) {
      console.log(`  problems:`);
      for (const p of trace.summary.problems.slice(0, 8)) console.log(`    ${pad(p.n, 6)}${p.text}`);
    }
  } else if (trace) {
    console.log(`\n  TRACE ${trace.entries?.length ?? 0} entries (written before summaries existed — no tallies)`);
  }

  console.log(
    t.disagreements
      ? `\n  ${t.disagreements} slot(s) where the deck and the log disagree\n`
      : `\n  deck and log agree on every slot\n`,
  );
}

const invokedDirectly = process.argv[1] && process.argv[1].endsWith("triage.mjs");
if (invokedDirectly) {
  const args = process.argv.slice(2);
  const flags = args.filter((a) => a.startsWith("--"));
  const paths = args.filter((a) => !a.startsWith("--"));
  const deckPath = paths.find((p) => p.endsWith(".pptx"));
  const logPath = paths.find((p) => p.endsWith(".json"));
  if (!deckPath || !logPath) {
    console.error("usage: node scripts/triage.mjs <deck.pptx> <run-log.json> [--all] [--json]");
    process.exit(2);
  }
  let deck, log;
  try {
    deck = await readDeck(deckPath);
    log = JSON.parse(readFileSync(logPath, "utf8"));
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
    if (selftest.length) {
      console.log(`\n  SELF-TEST ${selftest.filter((s) => s.ok).length} of ${selftest.length} scenarios passed\n`);
      for (const s of selftest) {
        const mark = s.skipped ? "skip" : s.ok ? "ok" : "FAIL";
        console.log(`  ${pad(mark, 6)}${pad(s.name, 36)}${s.detail}`);
      }
      console.log("");
    }
    if (!results.length && !selftest.length) console.log("\n  this log holds no runs and no self-test\n");
  }
  const disagreements =
    results.reduce((n, { t }) => n + t.disagreements, 0) + selftest.filter((s) => !s.ok && !s.skipped).length;
  process.exit(disagreements || faults.length ? 1 : 0);
}
