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

/**
 * How much of one payload value may appear on a step line.
 *
 * A cap rather than a ban. `formatTraceValue` used to drop anything whose
 * `typeof` was "object", which is every array — and the payloads that matter
 * most are arrays: `degradation curves` carries the two timing series that
 * whole experiment exists to produce, and `tag pass over a page` carries
 * `withoutTag`, the list of charts that lost their config. Both vanished from
 * the on-screen step list AND from the crash log, which share this formatter,
 * so an experiment that killed the tab left behind its verdict and none of its
 * numbers. A breadcrumb should be short; it should not be silent.
 */
const STEP_VALUE_CHARS = 140;

/**
 * One payload value as it appears on a step line, or undefined to leave it off.
 *
 * Left off: `undefined`, and the empty string. Both mean "nothing to say", and
 * `key=` with nothing after it is noise on a line read at a glance. `0`,
 * `false` and `null` are all kept — they are answers.
 */
export function formatTraceValue(v: unknown): string | undefined {
  if (v === undefined || v === "") return undefined;
  if (v === null || typeof v !== "object") return String(v);
  // Objects and arrays, compactly. A circular payload is a caller's mistake
  // rather than a reason to lose the line, so it degrades to a shape.
  let text: string;
  try {
    text = JSON.stringify(v) ?? String(v);
  } catch {
    text = Array.isArray(v) ? `[${v.length} items]` : "{…}";
  }
  return text.length > STEP_VALUE_CHARS ? `${text.slice(0, STEP_VALUE_CHARS - 1)}…` : text;
}

/**
 * One trace entry as one line — the single rendering the pane's step list and
 * the crash log both use.
 *
 * Shared on purpose: the file and the screen must never describe the same run
 * differently, and they did for as long as the formatter lived inside a DOM
 * closure where nothing could test it.
 */
export function formatTraceLine(e: TraceEntry): string {
  const bits = Object.entries(e.data ?? {})
    .map(([k, v]) => {
      const text = formatTraceValue(v);
      return text === undefined ? undefined : `${k}=${text}`;
    })
    .filter((b): b is string => b !== undefined)
    .join(" ");
  const secs = String(Math.round(e.ms / 100) / 10).padStart(6);
  return `${secs}s  ${e.scope}  ${e.message}${bits ? `  ${bits}` : ""}`;
}

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
 * The trace's own clock — ms since tracing was switched on — or `null` when it
 * is off and there is therefore no timeline to be on.
 *
 * Exported so that anything ELSE stamping times into the same round file can
 * use one origin. It could not, and that cost a real misreading: the host
 * probe stamped its samples from its own `runStarted`, which begins when
 * `runHostProbes` is called — 7.9 seconds after `setTracing(true)` in the
 * 2026-08-11 round. Two time series in one file, on two origins, with nothing
 * anywhere saying so. Every cross-reference between a sample and a trace line
 * is the whole analysis method for these rounds, and every one of them was off
 * by that constant: pass 2 began at 41.6s by the trace and at 34.4s by the
 * samples, which reads as a sample arriving before the pass that produced it.
 *
 * The same family as `idleMs` and `afterAnswering` — a number that cannot be
 * compared against the numbers beside it is not yet a measurement.
 */
export function traceElapsed(): number | null {
  return enabled ? Date.now() - startedAt : null;
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
  const entry: TraceEntry = { ms: Date.now() - startedAt, scope, message, ...(data ? { data: { ...data } } : {}) };
  entries.push(entry);
  // Never let a watcher's failure cost the log the entry it was recording.
  try {
    watcher?.(entry);
  } catch {
    /* the window is broken; the record is not */
  }
  if (entries.length > MAX_ENTRIES) {
    entries.shift();
    dropped++;
  }
}

/**
 * A marker for "everything from here on", so a caller can later ask for only
 * the entries its own operation produced.
 *
 * The buffer spans every operation since tracing was switched on, separated by
 * nothing but a `pane: action started` line. A run log that carried the whole
 * buffer therefore carried other runs too — and pairing the wrong run's trace
 * with a report's per-item numbers is exactly the wrong turn that costs an
 * hour. It cost one.
 */
/**
 * Something that wants to SEE each step as it happens, not afterwards.
 *
 * The log is the record; this is the window. They are not the same need: a run
 * log is downloadable only once the run ends, and the runs that most need
 * explaining are exactly the ones that never end — a host that wedges, or a
 * PowerPoint that puts up "Sorry, we ran into a problem" and takes the pane's
 * memory with it. What is already on screen survives that, and can be
 * screenshotted or copied. Twice now a real-host failure has cost a whole round
 * because the only evidence died with the tab.
 *
 * One subscriber, replaced rather than accumulated: there is one pane.
 */
let watcher: ((e: TraceEntry) => void) | undefined;

/** Watch every step as it is recorded. Pass undefined to stop watching. */
export function onTrace(cb: ((e: TraceEntry) => void) | undefined): void {
  watcher = cb;
}

export function traceMark(): number {
  return entries.length + dropped;
}

/** Tallies that make a long trace legible without reading all of it. */
export interface TraceSummary {
  /** Every distinct scope+message, commonest first. */
  steps: { scope: string; message: string; n: number }[];
  /** Every distinct failure reason, commonest first — see `PROBLEM_KEYS`. */
  problems: { text: string; n: number }[];
}

/**
 * The `data` keys that carry a reason something went wrong.
 *
 * Call sites reach for whichever word fits the sentence, so a histogram that
 * only knew `error` missed two thirds of them. A run's problems are worth more
 * than its consistency about naming them.
 */
const PROBLEM_KEYS = ["error", "why", "reason"];

/** Longest kept per distinct problem string. Office.js error text carries a
 *  `debugInfo` blob that repeats the message and swamps the tally. */
const PROBLEM_MAX = 120;

function summarise(list: TraceEntry[]): TraceSummary {
  const steps = new Map<string, { scope: string; message: string; n: number }>();
  const problems = new Map<string, number>();
  for (const e of list) {
    // The pair, as JSON rather than as a delimited string.
    //
    // This used to join with a literal NUL — a separator that cannot appear in
    // either half, which is the right property and the wrong character. One NUL
    // byte makes grep and ripgrep classify the whole FILE as binary, so every
    // codebase search silently skipped `trace.ts`: a sweep for the step
    // formatter's drop rule, which lived here and was wrong for months, would
    // never have matched it. A file nothing can search is a file nothing will
    // fix.
    const key = JSON.stringify([e.scope, e.message]);
    const hit = steps.get(key);
    if (hit) hit.n++;
    else steps.set(key, { scope: e.scope, message: e.message, n: 1 });
    for (const k of PROBLEM_KEYS) {
      const v = e.data?.[k];
      if (typeof v !== "string" || !v) continue;
      const text = v.slice(0, PROBLEM_MAX);
      problems.set(text, (problems.get(text) ?? 0) + 1);
    }
  }
  const byCount = <T extends { n: number }>(a: T, b: T) => b.n - a.n;
  return {
    steps: [...steps.values()].sort(byCount),
    problems: [...problems.entries()].map(([text, n]) => ({ text, n })).sort(byCount),
  };
}

/**
 * The log, oldest first, plus how many entries fell off the front.
 *
 * Pass a mark from `traceMark()` to get only what happened after it.
 *
 * `summary` comes FIRST because it is what gets read first. A real run's log
 * is 160 KB of pretty-printed JSON and its 276 entries do not fit on a screen
 * or in one read; the two questions actually asked of it — what did this run
 * do, and what went wrong — are both answered by a tally at the top. The
 * entries stay for when the tally raises a question.
 */
export function traceLog(since = 0): { summary: TraceSummary; entries: TraceEntry[]; dropped: number } {
  // `dropped` shifts the buffer's start, so a mark taken before entries fell
  // off has to be measured against the same absolute scale.
  const from = Math.max(0, since - dropped);
  const slice = entries.slice(from).map((e) => ({ ...e, ...(e.data ? { data: { ...e.data } } : {}) }));
  return { summary: summarise(slice), entries: slice, dropped };
}
