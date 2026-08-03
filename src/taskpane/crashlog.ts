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
 * - **Structured entries.** It keeps the formatted one-line steps, not the
 *   trace's objects. A real run's entries pretty-print to 160 KB and a stalled
 *   one emits thousands; lines are ~100 bytes and are what a person actually
 *   reads. The full structured log still rides in the normal run log, which
 *   exists whenever the run got far enough to have one.
 * - **Flush per step.** A synchronous `localStorage.setItem` per trace entry
 *   would put a serialize-and-write in the hot path of a run that is already
 *   struggling. Debounced instead, so a crash costs at most the last window.
 * - **Throw.** Storage can be absent, disabled, or full, and none of that is a
 *   reason for a run to fail. Every path here swallows.
 */

/**
 * The run being written now. Overwritten at the start of each run.
 */
const LIVE_KEY = "powerchart.crashlog.live.v1";
/**
 * The most recent run that ended without being marked finished — i.e. crashed.
 *
 * Kept separately from the live slot so that starting a new run cannot destroy
 * the evidence from the one that died. Without this, the natural reaction to a
 * crash (reload the pane, run it again) would erase exactly what was worth
 * keeping, and it would do so before anyone had a chance to look.
 */
const KEPT_KEY = "powerchart.crashlog.kept.v1";

/**
 * Steps kept. Mirrors the trace's own ring so the two cannot disagree about
 * how much history exists, and bounds the write at roughly 200 KB against a
 * store that gives about 5 MB.
 */
const MAX_STEPS = 2000;

/** How long writes are batched for. A crash costs at most this much record. */
const FLUSH_MS = 400;

/** One captured run, as written to storage and handed back for download. */
export interface CrashLog {
  /** Names the file for whatever reads it — the shape is not a run log. */
  kind: "powerchart-crash-log";
  build: string;
  host: string;
  /** What the run was: "self-test", "demo deck (file)", … */
  label: string;
  startedAt: string;
  /** Set only when the run ended normally. Its ABSENCE is the whole signal. */
  finishedAt?: string;
  /** How many steps were recorded before the cap dropped the oldest. */
  dropped: number;
  /** Formatted one-line steps, OLDEST FIRST — the reading order for a file. */
  steps: string[];
}

/** The store, or null where there is not one. Never throws. */
function store(): Storage | null {
  try {
    // Touched, not just read: a `localStorage` that exists but throws on access
    // is a real configuration (third-party cookies blocked, and a task pane IS
    // a third-party frame), and it throws at USE rather than at lookup.
    const s = window.localStorage;
    const probe = "powerchart.crashlog.probe";
    s.setItem(probe, "1");
    s.removeItem(probe);
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
    return parsed && parsed.kind === "powerchart-crash-log" && Array.isArray(parsed.steps) ? parsed : null;
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
  if (previous && !previous.finishedAt && previous.steps.length) {
    try {
      s.setItem(KEPT_KEY, JSON.stringify(previous));
    } catch {
      /* the store is full; the live record below matters more */
    }
  }
  live = {
    kind: "powerchart-crash-log",
    build: meta.build,
    host: meta.host,
    label: meta.label,
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

/** Mark the run finished. Its absence is what identifies a crash. */
export function endCrashLog(): void {
  if (!live) return;
  live.finishedAt = new Date().toISOString();
  if (timer !== undefined) {
    clearTimeout(timer);
    timer = undefined;
  }
  flush();
  live = null;
}

/**
 * A run that ended without finishing, if there is one — the crashed run.
 *
 * Checks the kept slot first and the live slot second, because a pane that
 * reloads after a crash has the dead run still sitting in the live slot: it was
 * never promoted, since promotion happens when the NEXT run starts and no next
 * run has. Both are the same question asked at different moments.
 */
export function recoverCrashLog(): CrashLog | null {
  const kept = read(KEPT_KEY);
  if (kept && !kept.finishedAt && kept.steps.length) return kept;
  const stale = read(LIVE_KEY);
  if (stale && !stale.finishedAt && stale.steps.length) return stale;
  return null;
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
