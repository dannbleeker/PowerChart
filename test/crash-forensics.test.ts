import { describe, expect, it } from "vitest";
// @ts-expect-error — plain .mjs tool, no types.
import * as forensics from "../scripts/crash-forensics.mjs";
const { telemetryIndices, fatalWindow, collectCrashEvidence, scrub } = forensics;

/**
 * The evidence this collects exists for a few minutes and then is gone: it
 * lives in the browser's request log, and recovery reloads the tab. Three
 * crashes were diagnosed by hand before this existed and each pass cost a
 * quarter of an hour to reach the same three lines.
 */
describe("pulling PowerPoint's own account of a crash", () => {
  const uls = (messages: string[]) => JSON.stringify({ T: 1, L: messages.map((M, i) => ({ T: i, M })) });

  it("keeps live session tokens out of the report", () => {
    // THIS REPO IS PUBLIC. A wedged PowerPoint logs from URLs carrying usid,
    // hid, postmessagetoken and the user's filename — live session tokens and a
    // personal document name. The report is scrubbed on the way IN as well as
    // gitignored, because it also gets pasted into issues and chat by hand, and
    // only one of those two defences travels with the text.
    const line =
      "[ERROR] Failed to load resource @ https://powerpoint.officeapps.live.com/pods/ppt.aspx?usid=abc-123&hid=def-456&postmessagetoken=ghi&filename=Presentation63.pptx";
    const clean = scrub(line);
    expect(clean, "the log message itself is the evidence and must survive").toContain("Failed to load resource");
    for (const secret of ["usid", "hid=", "postmessagetoken", "Presentation63"]) {
      expect(clean, `leaked ${secret}`).not.toContain(secret);
    }
    // A bare URL inside a message keeps its origin — which is diagnostic — and
    // loses the query, which is where the tokens live.
    expect(scrub("POST https://x.live.com/edit.svc/Foo?usid=abc-123")).toBe("POST https://x.live.com/edit.svc/Foo?…");
  });

  it("reads the fatal window out of a telemetry batch", () => {
    const body = uls([
      "PptApi Call - Presentation.GetSlides",
      "Failed to restore selection after load content.",
      "OnServerFindSucceeded could not find target slide, time elapsed: 449 ms",
      "SendToWorkerApi with otel",
    ]);
    const window = fatalWindow(body);
    expect(window, "the batch carried the failure and this found nothing").toBeTruthy();
    // The lines BEFORE it are the point — the failure names itself, but only the
    // calls in flight say what the add-in was doing when the host gave up.
    expect(window).toContain("Presentation.GetSlides");
    expect(window).toContain("could not find target slide");
  });

  it("says nothing about a batch that carries no failure", () => {
    // Most batches are survey and activity noise. A window built from one of
    // those would read as evidence and be nothing of the kind.
    expect(fatalWindow(uls(["SIFB: LogActivityEndInternal AppUsageNPS"]))).toBe(null);
    expect(fatalWindow("not json at all")).toBe(null);
    expect(fatalWindow("")).toBe(null);
  });

  it("takes the NEWEST telemetry batches, and a bounded number of them", () => {
    // A wedged tab holds tens of thousands of requests and each body is its own
    // round trip, so an unbounded walk would outlast the round. The crash is
    // always at the end: the document's channel stops and telemetry keeps going.
    const requests = [
      "1. [POST] https://x/pt/RemoteUls.ashx => [200]",
      "2. [POST] https://x/pods/podedit.svc/GetUpdates => [200]",
      "3. [POST] https://x/pt/RemoteUls.ashx => [200]",
      "4. [POST] https://x/pt/RemoteUls.ashx => [200]",
    ].join("\n");
    expect(telemetryIndices(requests, 2), "newest first, and only two").toEqual([4, 3]);
    expect(telemetryIndices("", 5)).toEqual([]);
  });

  it("stops at the first batch that carries a failure", async () => {
    const opened: string[] = [];
    const sh = (cmd: string, arg?: string) => {
      if (cmd === "console") return "[ERROR] Failed to load resource: net::ERR_NETWORK_CHANGED";
      if (cmd === "requests")
        return [
          "8. [POST] https://x/pt/RemoteUls.ashx => [200]",
          "9. [POST] https://x/pt/RemoteUls.ashx => [200]",
        ].join("\n");
      opened.push(String(arg));
      return arg === "9" ? uls(["ErrorName: errorLocalChangeLostSingleUser"]) : uls(["nothing here"]);
    };
    const report = await collectCrashEvidence(sh as never, { at: "round 042" });
    expect(report).toContain("errorLocalChangeLostSingleUser");
    expect(report).toContain("ERR_NETWORK_CHANGED");
    expect(opened, "kept opening bodies after it had the answer").toEqual(["9"]);
  });

  it("never writes an empty host over a read that failed", async () => {
    // BOTH of round 044's crash reports said "(nothing)" for the document
    // channel and "no fatal entry in the last 0 telemetry batches" for the ULS.
    // Neither was true: `requests` had overflowed the spawn buffer, so nothing
    // was read at all — and "no fatal entry" is the sentence that DESCRIBES the
    // quiet form of the wedge. A failed read that reads as a diagnosis is worse
    // than a missing section.
    const state = { lastError: null as string | null };
    // The real `cli` records the failure on the state and still returns "" — the
    // empty string is exactly what made the two indistinguishable.
    const sh = Object.assign(
      (cmd: string) => {
        state.lastError = cmd === "requests" ? "overflow" : null;
        return cmd === "console" ? "[ERROR] something real" : "";
      },
      { state },
    );
    const report = await collectCrashEvidence(sh as never, { at: "round 044" });
    expect(report, "the console read worked and must still be there").toContain("something real");
    expect(report, "an overflowed read is named as one").toContain("more than the buffer holds");
    expect(report).not.toContain("(nothing)");
    expect(report, "0 batches scanned must never read as 0 failures found").not.toContain(
      "no fatal entry in the last 0",
    );
  });

  it("names a read that RAN and exited non-zero, not just one that could not be spawned", async () => {
    // THE SECOND HALF OF THAT FIX, and it shipped without a guard. `ask()`
    // honoured `lastError` — could not spawn, or overflowed — but not
    // `lastFailed`, which is a call that ran and came back non-zero. That is
    // exactly how a wedged tab answers `console` and `requests`, so the 18:06
    // wedge on 2026-08-15 produced a report reading "(nothing)" three times
    // over three failed reads, HOURS after the overflow half was fixed.
    //
    // Verified in production at the time and never pinned, which is how a fix
    // becomes a regression nobody notices. The distinction matters because
    // "(nothing)" is a claim about the HOST and this is a fact about the read.
    const state = { lastError: null as string | null, lastFailed: false };
    const sh = Object.assign(
      (cmd: string) => {
        // Ran, exited non-zero, returned the same empty string a healthy quiet
        // host would have.
        state.lastFailed = cmd !== "console";
        return cmd === "console" ? "[ERROR] something real" : "";
      },
      { state },
    );
    const report = await collectCrashEvidence(sh as never, { at: "a 30-minute wedge" });
    expect(report, "the console read worked and must survive").toContain("something real");
    expect(report, "a failed read must be named as one").toContain("could not be run");
    expect(report, "an empty host is a different claim from an unread one").not.toContain("(nothing)");
    expect(report, "and a scan of nothing must not read as a clean scan").not.toContain("no fatal entry in the last 0");
  });

  it("still produces a report when every read fails", async () => {
    // It runs on a host that has just died, so half of these are expected to
    // come back empty or throw. A forensics pass that took the driver down with
    // it would be worse than none — and the absence of a fatal entry is itself
    // the diagnosis: that is what the quiet form of the wedge looks like.
    const sh = () => {
      throw new Error("the browser is not open");
    };
    const report = await collectCrashEvidence(sh as never, { at: "round 042" });
    expect(report).toContain("could not read");
    expect(report, "the report is still a report").toContain("PowerPoint crash");
  });
});
