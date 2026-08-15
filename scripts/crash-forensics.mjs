/**
 * Pull PowerPoint's own account of a crash out of the browser tab.
 *
 * THREE TIMES this was done by hand, and each pass was the same ten commands in
 * the same order: list the requests, find where the document's data channel
 * stopped, walk the telemetry batches backwards until one carries an error
 * name, print the thirty lines before it. Each pass took a quarter of an hour
 * and produced the same three lines:
 *
 *     OnServerFindSucceeded could not find target slide, time elapsed: 449 ms
 *     GlobalErrorHandler:DisplayErrorDialog: 5341289
 *     ErrorDialog::ShowErrorDialog BSQMErrorCode: 5341289;
 *       ErrorName: errorLocalChangeLostSingleUser
 *
 * That is the only evidence in existence about why this host dies — it is not
 * documented anywhere public — and it lives in a request log that the browser
 * discards when the tab reloads. Which is exactly what recovery does next. So
 * every wedge until now has been diagnosed from memory of the last one, or not
 * at all.
 *
 * The add-in cannot collect this. Office ships it from the document's own frame
 * to `RemoteUls.ashx`, and the pane is a different origin with no access to
 * either. The driver, sitting outside in the browser, is the only thing that
 * can — and it already knows the moment a crash happens, because it watches for
 * the dialog.
 *
 * Reads only. Nothing here clicks, navigates or reloads: it runs BEFORE
 * recovery, on a tab that is about to be thrown away.
 */

/** Office's telemetry lines that mean the host gave up, in the host's words. */
const FATAL = /ErrorName:\s*(\w+)|could not find target slide|DisplayErrorDialog:\s*(\d+)/;

/**
 * The request indices worth opening, newest first.
 *
 * Bounded on purpose. A wedged tab carries tens of thousands of requests and
 * each body is a separate CLI round trip, so an unbounded scan would take
 * longer than the round did. The crash is always near the end — the document's
 * data channel stops and only telemetry keeps flowing — so the last few dozen
 * batches hold it or nothing does.
 */
export function telemetryIndices(requestsOutput, limit = 40) {
  return String(requestsOutput ?? "")
    .split("\n")
    .filter((l) => /RemoteUls|RemoteTelemetry/.test(l))
    .map((l) => Number(/^(\d+)\./.exec(l.trim())?.[1]))
    .filter((n) => Number.isFinite(n))
    .slice(-limit)
    .reverse();
}

/** The lines around the first fatal entry in one telemetry batch, or null. */
export function fatalWindow(body, before = 30, after = 3) {
  let batch;
  try {
    batch = JSON.parse(String(body ?? ""));
  } catch {
    return null;
  }
  const entries = Array.isArray(batch?.L) ? batch.L : [];
  const at = entries.findIndex((e) => FATAL.test(String(e?.M ?? "")));
  if (at < 0) return null;
  return entries
    .slice(Math.max(0, at - before), at + after)
    .map((e) => `${e?.T ?? "?"} | ${String(e?.M ?? "").slice(0, 200)}`)
    .join("\n");
}

/**
 * Everything worth keeping about a crash, as one report.
 *
 * `sh` is the driver's CLI caller. Every failure here is swallowed and named in
 * the report rather than thrown: this runs on a host that has just died, so
 * half the reads are expected to come back empty, and a forensics pass that
 * takes the driver down with it would be worse than none.
 */
export async function collectCrashEvidence(sh, { at = "unknown", limit = 40 } = {}) {
  const parts = [`# PowerPoint crash — ${at}`, ""];
  const safe = (label, fn) => {
    try {
      const out = fn();
      parts.push(`## ${label}`, "", String(out ?? "").trim() || "(nothing)", "");
    } catch (err) {
      parts.push(`## ${label}`, "", `(could not read: ${err?.message ?? err})`, "");
    }
  };

  safe("Console errors", () =>
    String(sh("console", "error") ?? "")
      .split("\n")
      .filter((l) => /ERROR|net::/.test(l))
      .slice(-12)
      .join("\n"),
  );

  // The document's data channel, and where it stopped. `GetPopWacUpdates` going
  // quiet is the tell that separates a crashed session from a slow one.
  safe("Document channel, last 12", () =>
    String(sh("requests") ?? "")
      .split("\n")
      .filter((l) => /edit\.svc|podedit|RemoteSessionTermination/.test(l))
      .slice(-12)
      .map((l) => l.replace(/\?[^\s]*/, "").trim())
      .join("\n"),
  );

  // Guarded like everything else, and this one caught itself: the scan sat
  // outside the `safe` wrapper at first and threw the whole pass away on a host
  // where the browser had gone, which is one of the two states it exists for.
  let found = null;
  // No initialiser: the only path that reads it is the one that assigned it,
  // because the catch below returns early.
  let scanned;
  try {
    const indices = telemetryIndices(sh("requests"), limit);
    scanned = indices.length;
    for (const i of indices) {
      const window = fatalWindow(sh("request-body", String(i)));
      if (window) {
        found = { i, window };
        break;
      }
    }
  } catch (err) {
    parts.push("## PowerPoint's own error log (ULS)", "", `(could not read: ${err?.message ?? err})`, "");
    return parts.join("\n");
  }
  parts.push(
    "## PowerPoint's own error log (ULS)",
    "",
    found
      ? `From request ${found.i}:\n\n${found.window}`
      : `(no fatal entry in the last ${scanned} telemetry batches — the quiet form of the wedge looks like this)`,
    "",
  );

  return parts.join("\n");
}
