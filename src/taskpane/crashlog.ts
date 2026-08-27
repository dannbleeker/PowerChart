/**
 * The record that survives the runs which do not end.
 *
 * "Download run log" can only hand out a run that finished — the file is built
 * from a value assigned after the last `await` returns. That is precisely the
 * wrong contract for this project. Two rounds of real-host testing have now
 * been lost to runs that never finished: one wedged at 1819 seconds and had to
 * be killed by closing the tab, one taken out at 108 seconds by PowerPoint's
 * own *"Sorry, we ran into a problem"*. Neither produced a log. Both were the
 * most interesting runs of their session.
 *
 * The Live steps list was the first answer and it half-works: it is on screen
 * the whole time, so a screenshot survives a crash. But a screenshot holds
 * twenty lines of a three-hundred-line run, it cannot be searched, and it
 * cannot be handed to `npm run triage`.
 *
 * So the steps are also written to `localStorage` as they happen. It is the
 * only store in a task pane that outlives the pane's own JavaScript context:
 * the tab dies, the add-in reloads, and the bytes are still there. On the next
 * open the pane notices a run that was never marked finished and offers it for
 * download — the evidence from the crash, retrieved after the crash.
 *
 * Three things this deliberately does NOT do:
 *
 * - **Structured TRACE entries.** It keeps the formatted one-line steps, not
 *   the trace's objects. A real run's entries pretty-print to 160 KB and a
 *   stalled one emits thousands; lines are ~100 bytes and are what a person
 *   actually reads.
 *
 *   It does now keep the run's FINDINGS — the probe's answer sheet, each
 *   scenario's verdict (see `findings`). That is not a softening of the rule
 *   above, it is the rule read properly: a dozen small objects is not thousands
 *   of large ones, and these are the part of a round everyone wants. The
 *   sentence this replaces said the structured log "still rides in the normal
 *   run log, which exists whenever the run got far enough to have one" — true,
 *   and it is the whole problem. These runs do not get that far.
 * - **Flush per step.** A synchronous `localStorage.setItem` per trace entry
 *   would put a serialize-and-write in the hot path of a run that is already
 *   struggling. Debounced instead, so a crash costs at most the last window.
 * - **Throw.** Storage can be absent, disabled, or full, and none of that is a
 *   reason for a run to fail. Every path here swallows.
 */

/**
 * The run being written now. Overwritten at the start of each run.
 */
const LIVE_KEY = "ssf-charts.crashlog.live.v1";
/**
 * The most recent run that ended without being marked finished — i.e. crashed.
 *
 * Kept separately from the live slot so that starting a new run cannot destroy
 * the evidence from the one that died. Without this, the natural reaction to a
 * crash (reload the pane, run it again) would erase exactly what was worth
 * keeping, and it would do so before anyone had a chance to look.
 */
const KEPT_KEY = "ssf-charts.crashlog.kept.v1";

/**
 * Steps kept. Mirrors the trace's own ring so the two cannot disagree about
 * how much history exists, and bounds the write at roughly 200 KB against a
 * store that gives about 5 MB.
 */
const MAX_STEPS = 2000;

/** How long writes are batched for. A crash costs at most this much record. */
const FLUSH_MS = 400;

/**
 * Findings kept, and how big one may be.
 *
 * A round has one answer sheet and about a dozen verdicts, so the count is
 * headroom rather than a limit anything real approaches. The BYTE cap is the
 * one doing work: a finding is caller-supplied, and the store this shares with
 * the 2000-step log is the same ~5 MB that a long crashed run already fills.
 * Over the cap the finding is replaced by a note saying so — losing one
 * oversized value, never the record.
 */
const MAX_FINDINGS = 40;
const MAX_FINDING_BYTES = 128 * 1024;

/** Does this name a crash log, under either name it has had? */
export function isCrashLogKind(kind: unknown): boolean {
  return kind === "ssf-charts-crash-log" || kind === "powerchart-crash-log";
}

/** One captured run, as written to storage and handed back for download. */
export interface CrashLog {
  /**
   * Names the file for whatever reads it — the shape is not a run log.
   *
   * BOTH SPELLINGS ARE READ, one is written. Every crash already archived under
   * `crashes/` carries `powerchart-crash-log`, and the reader below selects on
   * it; narrowing to the new name would make this build refuse to open its own
   * crash history. New logs are stamped `ssf-charts-crash-log`.
   */
  kind: "ssf-charts-crash-log" | "powerchart-crash-log";
  build: string;
  host: string;
  /** What the run was: "self-test", "demo deck (file)", … */
  label: string;
  /**
   * Which run this is, counting up across pane reloads.
   *
   * Two unfinished records can exist at once, and the one worth offering is the
   * later one — so they have to be orderable. A timestamp is the obvious key
   * and the wrong one: two runs can start inside a millisecond, and a clock
   * that a user or an NTP sync moves backwards would silently invert the
   * answer. A counter carried in the record cannot do either.
   */
  seq: number;
  startedAt: string;
  /** Set only when the run ended normally. */
  finishedAt?: string;
  /**
   * Set when this run's file was actually handed to the browser.
   *
   * Separate from `finishedAt`, because the two are different facts and using
   * one for the other cost a real round. A round that completes prints "Saved
   * as one file" and used to mark itself finished BEFORE the download was even
   * attempted — so a finished run stopped being recoverable on the strength of
   * a save nobody had tried yet, let alone watched. PowerPoint then died, the
   * owner reopened the pane, and there was nothing to offer: the record was
   * there and the recovery would not look at it, because it had finished.
   *
   * The absence of THIS is the signal now. `finishedAt` still says whether the
   * run crashed, which is what the offer's wording needs.
   */
  savedAt?: string;
  /** How many steps were recorded before the cap dropped the oldest. */
  dropped: number;
  /** Formatted one-line steps, OLDEST FIRST — the reading order for a file. */
  steps: string[];
  /**
   * The round's FINDINGS, as opposed to its narration.
   *
   * The "no structured entries" rule above is about the trace: thousands of
   * objects, 160 KB, and lines are what a person reads. It was never about
   * these. A round produces one probe answer sheet and about a dozen scenario
   * verdicts — a few KB, the part everyone actually wants, and until now the
   * part a crash destroyed. Both were held in a module variable and written to
   * a file after the last `await`, so a tab that died mid-battery took the
   * probe's answers with it even though the probe had finished minutes earlier
   * and the code's own comment calls that half "complete, cheap, and the half
   * most likely to be worth reading".
   *
   * Steps alone cannot replace them: a step says a scenario started, not what
   * it concluded, and `npm run triage` reads verdicts.
   *
   * Bounded on both counts (see `MAX_FINDINGS`), because the argument against
   * structure here is size and it stays answered rather than merely asserted.
   */
  findings?: { key: string; value: unknown }[];
}

/**
 * The store, or null where there is not one. Never throws.
 *
 * Touched rather than merely looked up: a `localStorage` that exists and throws
 * on access is a real configuration — third-party cookies blocked, and a task
 * pane IS a third-party frame — and it throws at USE, not at lookup.
 *
 * Touched with a READ. The probe used to be a write, which conflated two
 * different questions: "is there a store" and "is there room in it". A store
 * with no room left answers no to the second and yes to the first, and gating
 * everything on the write meant a crashed run's record — sitting right there,
 * needing no quota to read — could not be handed back. The way such a store
 * gets full, on this origin, is our own 2000-step log: so that is not a corner
 * case, it is what a long crashed run does to the next one.
 *
 * Whether a write will fit is `flush`'s question, and it is the one place that
 * can do anything useful about the answer.
 */
function store(): Storage | null {
  try {
    const s = window.localStorage;
    s.getItem("ssf-charts.crashlog.probe");
    return s;
  } catch {
    return null;
  }
}

function read(key: string): CrashLog | null {
  const s = store();
  if (!s) return null;
  try {
    const raw = s.getItem(key);
    if (!raw) return null;
    const parsed = JSON.parse(raw) as CrashLog;
    // A record written by an older build, or a half-written one from a crash
    // mid-write, must not reach the UI as if it were sound.
    return parsed && isCrashLogKind(parsed.kind) && Array.isArray(parsed.steps) ? parsed : null;
  } catch {
    return null;
  }
}

/** Live state. Held in memory and mirrored to storage on the flush timer. */
let live: CrashLog | null = null;
let timer: ReturnType<typeof setTimeout> | undefined;

/**
 * Write the live record out.
 *
 * On a full store the record is halved and retried once rather than abandoned:
 * the newest half of a long run is worth more than nothing, and "the log could
 * not be saved" is the one message this module must never be the cause of.
 */
function flush(): void {
  const s = store();
  if (!s || !live) return;
  try {
    s.setItem(LIVE_KEY, JSON.stringify(live));
  } catch {
    try {
      const half = Math.floor(live.steps.length / 2);
      live.dropped += half;
      live.steps = live.steps.slice(half);
      s.setItem(LIVE_KEY, JSON.stringify(live));
    } catch {
      /* nothing to be done, and nothing worth failing a run over */
    }
  }
}

function schedule(): void {
  if (timer !== undefined) return;
  timer = setTimeout(() => {
    timer = undefined;
    flush();
  }, FLUSH_MS);
}

/**
 * Write out now, without waiting for the timer.
 *
 * The pane calls this when the page is going away. `pagehide` fires on a tab
 * close and on a navigation, and it is the last synchronous moment an add-in
 * gets — so the difference between having it and not is the final 400 ms of a
 * run, which on a run that ends by dying is the only part anyone wanted.
 *
 * It does NOT fire on the crash that matters most: PowerPoint's own "Sorry, we
 * ran into a problem" takes the frame without warning. That case is covered by
 * the timer alone, which is why the timer is short.
 */
export function flushCrashLog(): void {
  if (timer !== undefined) {
    clearTimeout(timer);
    timer = undefined;
  }
  flush();
}

/**
 * Start recording a run, preserving whatever the previous one left behind.
 *
 * The promotion is the point: if the last live record was never finished, it
 * belonged to a run that died, and it is moved somewhere the new run cannot
 * overwrite it. The obvious reaction to a crash is to reload and try again,
 * and without this that reaction is what destroys the evidence.
 */
export function beginCrashLog(meta: { build: string; host: string; label: string }): void {
  const s = store();
  if (!s) return;
  const previous = read(LIVE_KEY);
  // Promoted on UNSAVED, not on unfinished — the same correction `savedAt`
  // makes to `recoverCrashLog`. A finished run whose file never reached the
  // disk is exactly as worth keeping as one that crashed, and starting the
  // next run was the moment that record used to be overwritten.
  if (previous && !previous.savedAt && (previous.steps.length || previous.findings?.length)) {
    try {
      s.setItem(KEPT_KEY, JSON.stringify(previous));
    } catch {
      /* the store is full; the live record below matters more */
    }
  }
  live = {
    kind: "ssf-charts-crash-log",
    build: meta.build,
    host: meta.host,
    label: meta.label,
    // One past the highest run this browser has recorded. Read from both slots
    // because either can hold the most recent one, depending on whether the
    // last run was promoted.
    seq: Math.max(previous?.seq ?? 0, read(KEPT_KEY)?.seq ?? 0) + 1,
    startedAt: new Date().toISOString(),
    dropped: 0,
    steps: [],
  };
  flush();
}

/**
 * Record one step, oldest first.
 *
 * Takes the line the Live steps list has already formatted rather than a trace
 * entry: one formatter, so the file and the screen cannot describe the same run
 * differently, and a caller that has not switched tracing on writes nothing.
 */
export function recordCrashStep(line: string): void {
  if (!live) return;
  live.steps.push(line);
  if (live.steps.length > MAX_STEPS) {
    live.steps.shift();
    live.dropped++;
  }
  schedule();
}

/**
 * Record one FINDING — a probe answer sheet, a scenario verdict — as soon as it
 * exists, rather than when the run ends.
 *
 * Called under the same rule as `recordCrashStep`: never throws, never blocks,
 * and does nothing at all when no run is being recorded. `key` is what the
 * reader will look for (`"hostAnswers"`, `"selftest:<scenario>"`), and the same
 * key may be recorded twice — the later value wins, so a caller can record a
 * growing result without the record growing with it.
 */
export function recordCrashFinding(key: string, value: unknown): void {
  if (!live) return;
  let stored: unknown = value;
  try {
    const size = JSON.stringify(value)?.length ?? 0;
    if (size > MAX_FINDING_BYTES) stored = `[dropped: ${size} bytes, over the ${MAX_FINDING_BYTES}-byte cap]`;
  } catch {
    // A value that will not serialise cannot be written to the store at all,
    // and it must not take the record down on its way past.
    stored = "[dropped: not serialisable]";
  }
  live.findings ??= [];
  const at = live.findings.findIndex((f) => f.key === key);
  if (at >= 0) live.findings[at] = { key, value: stored };
  else live.findings.push({ key, value: stored });
  // Oldest out. A round records its answer sheet first, so if anything is ever
  // going to be dropped here it should be the fortieth verdict rather than the
  // sheet — but nothing real gets near this, and the cap exists so the claim
  // "bounded" is true rather than intended.
  while (live.findings.length > MAX_FINDINGS) live.findings.shift();
  schedule();
}

/**
 * Mark the run finished, and say whether its file was actually handed over.
 *
 * `saved` defaults to false and callers must pass it deliberately, because the
 * default that cost a round was the other one. Marking a run finished used to
 * be enough to make it unrecoverable, and the round path called it BEFORE
 * attempting the download — so the record was written off on the strength of a
 * save that had not happened yet.
 */
export function endCrashLog(saved = false): void {
  if (!live) return;
  live.finishedAt = new Date().toISOString();
  if (saved) live.savedAt = live.finishedAt;
  if (timer !== undefined) {
    clearTimeout(timer);
    timer = undefined;
  }
  flush();
  live = null;
}

/**
 * The run's file has now reached the user — stop offering it back.
 *
 * The explicit save, from a button the user pressed, which is the strongest
 * evidence this pane can have that they have the thing. Reaches the LIVE slot
 * as well as the KEPT one: an auto-download that was blocked leaves a finished
 * run in LIVE, and pressing *Download run log* afterwards is precisely the
 * recovery that must clear it.
 */
export function markCrashLogSaved(): void {
  const s = store();
  if (!s) return;
  const stamp = new Date().toISOString();
  if (live) {
    live.savedAt = stamp;
    flush();
  }
  for (const key of [LIVE_KEY, KEPT_KEY]) {
    const rec = read(key);
    if (!rec || rec.savedAt) continue;
    rec.savedAt = stamp;
    try {
      s.setItem(key, JSON.stringify(rec));
    } catch {
      /* full store — the record simply stays on offer, which is the safe way round */
    }
  }
}

/**
 * The most recent run whose file the user has not got — crashed, or finished
 * and never saved.
 *
 * It used to be "ended without finishing", and that is the narrower fact. A
 * round can complete, print "Saved as one file", and leave the user with
 * nothing: the browser can refuse a download in a task pane (a nested
 * cross-origin frame — the same refusal the Copy button already handles), and
 * a host that dies moments later takes an in-flight one with it. Both happened
 * on 2026-08-06, and the pane had already written the record off as finished.
 *
 * Both slots can hold one. A pane that reloads straight after a crash has the
 * dead run still in the LIVE slot, never promoted, because promotion happens
 * when the next run starts and no next run has. Start one anyway and it is
 * promoted to KEPT — so after two crashes in a row, both slots are full.
 *
 * The later one wins, and that ordering is the whole point of this function
 * rather than an incidental. The sequence that produces two is the one people
 * actually perform: a run dies, you reopen, you do the obvious thing and try
 * again rather than downloading first, and the second run dies too. The record
 * worth having then is the SECOND — the run you were watching, on the build you
 * were testing. Checking the kept slot first handed back the first.
 */
export function recoverCrashLog(): CrashLog | null {
  // Worth offering when it holds ANYTHING a reader wants — a step or a finding.
  //
  // The gate used to be `steps.length > 0`, i.e. "did this run narrate itself",
  // and narration is optional: `recordCrashStep` writes nothing unless Verbose
  // trace is ticked. So a round that crashed with the box unticked was
  // unrecoverable even when the probe half had finished and its answer sheet
  // was sitting right there. The runbook does say to tick it; a recovery path
  // that depends on the owner having done so is one that fails on the round
  // where they did not.
  const worthKeeping = (r: CrashLog | null): r is CrashLog =>
    !!r && !r.savedAt && (r.steps.length > 0 || (r.findings?.length ?? 0) > 0);
  const crashed = [read(KEPT_KEY), read(LIVE_KEY)].filter(worthKeeping);
  // `seq` counts up across reloads; a record written before it existed sorts
  // as 0, which puts it behind any newer one — the right way round.
  return crashed.sort((a, b) => (b.seq ?? 0) - (a.seq ?? 0))[0] ?? null;
}

/** Forget the recovered run, once it has been saved or dismissed. */
export function clearCrashLog(): void {
  const s = store();
  if (!s) return;
  try {
    s.removeItem(KEPT_KEY);
    // Only the LIVE slot of a run that is not the one recording now. Clearing
    // it while a run is in flight would delete that run's evidence — the exact
    // failure this module exists to prevent, committed by its own cleanup.
    if (!live) s.removeItem(LIVE_KEY);
  } catch {
    /* nothing to clear it with */
  }
}

/**
 * Test-only: drop in-memory state so one test's run cannot leak into the next.
 * Storage is the fixture's own to clear; this is the module's half.
 */
export function _resetCrashLogForTest(): void {
  if (timer !== undefined) clearTimeout(timer);
  timer = undefined;
  live = null;
}
