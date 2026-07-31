/**
 * An opt-in activity log, for the runs nobody can watch.
 *
 * Everything hard about this project has been diagnosed after the fact, from
 * a deck and a one-line summary: a chart reported failed that had landed
 * twice, a banner contradicting its own slide, a verification pass that
 * returned nothing and said nothing about why. Each time the missing thing
 * was the same — an ordered record of what the add-in actually did, while it
 * was doing it. `console.log` does not survive a crashed tab, and PowerPoint
 * on the web has crashed the tab.
 *
 * So: an in-memory ring the pane can hand back as JSON alongside the run
 * report. Off by default — the cost of a disabled `trace()` is one boolean
 * test, and nothing is retained — and switched on from the pane when someone
 * is actually chasing something.
 *
 * Deliberately independent of the pane and of Office.js, so the renderer can
 * write to it without importing either. If the UI toggle is dropped before
 * shipping, this module and its call sites stay: `setTracing(true)` from a
 * console is all a future investigation needs.
 */

export interface TraceEntry {
  /** Milliseconds since tracing was switched on — relative, so a log reads as
   *  a timeline rather than a column of epoch numbers. */
  ms: number;
  /** Which part of the add-in spoke: "demo", "insert", "repair", "host". */
  scope: string;
  message: string;
  /** Structured extras. Kept small — this is a breadcrumb, not a heap dump. */
  data?: Record<string, unknown>;
}

/**
 * Entries kept before the oldest is dropped. A stalled run on the web can
 * spend fifteen minutes emitting steps; the last few hundred are what
 * anyone reads, and an unbounded array in a struggling browser tab is the
 * last thing this should add to.
 */
const MAX_ENTRIES = 2000;

let enabled = false;
let startedAt = 0;
let entries: TraceEntry[] = [];
let dropped = 0;

/** Turn the log on or off. Switching on clears whatever came before. */
export function setTracing(on: boolean): void {
  enabled = on;
  if (on) {
    entries = [];
    dropped = 0;
    startedAt = Date.now();
  }
}

/** Whether tracing is currently recording. */
export function tracing(): boolean {
  return enabled;
}

/**
 * Record one step. A no-op — and specifically NOT a stringification — when
 * tracing is off, so call sites can be liberal without costing a live run.
 */
export function trace(scope: string, message: string, data?: Record<string, unknown>): void {
  if (!enabled) return;
  // Copy the payload rather than holding the caller's object. A log is a
  // record of what was true AT THE MOMENT it was written; keeping the live
  // reference meant a caller who reused or mutated its object afterwards
  // rewrote history — silently, and in a file someone may already have
  // downloaded and be reading as fact.
  entries.push({ ms: Date.now() - startedAt, scope, message, ...(data ? { data: { ...data } } : {}) });
  if (entries.length > MAX_ENTRIES) {
    entries.shift();
    dropped++;
  }
}

/** The log so far, oldest first, plus how many entries fell off the front. */
export function traceLog(): { entries: TraceEntry[]; dropped: number } {
  // Entries copied too, for the same reason in the other direction: a reader
  // that mutated what it was handed corrupted the next reader's copy.
  return { entries: entries.map((e) => ({ ...e, ...(e.data ? { data: { ...e.data } } : {}) })), dropped };
}
