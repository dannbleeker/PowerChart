import { describe, expect, it } from "vitest";
// @ts-expect-error — plain .mjs tool, no types.
import { authPageKind } from "../scripts/round.mjs";

/**
 * WHAT THIS COST: a six-hour unattended run, stopped early, and a push
 * notification telling the owner he was needed when he was not.
 *
 * On 2026-08-23 the sole tab was `login.microsoft.com/consumers/fido/get`,
 * titled "Sign in to your account". The URL and the title both said sign-in, so
 * the loop stopped. The page rendered:
 *
 *     Sign in
 *     Sorry, but we're having trouble signing you in.
 *     AADSTS900561: The endpoint only accepts POST requests. Received a GET request.
 *     Request Id: ...
 *
 * No field, no button to proceed. An error page REPORTING a failed sign-in,
 * from a bad GET — and the profile still held a live session, which navigating
 * to OneDrive proved in one command.
 *
 * The SAME page had caused a false stop eight hours earlier. The test written
 * after that one was "check the tab title", and this is that test failing a
 * second time: a title can say "Sign in" on a page whose only content is a
 * request id. Read what is RENDERED, not what the tab is called.
 */
describe("telling a sign-in prompt from a sign-in error", () => {
  it("calls the real 2026-08-23 page an error, because nothing on it accepts input", () => {
    expect(
      authPageKind(
        "https://login.microsoft.com/consumers/fido/get?mkt=EN-US",
        'heading "Sign in" paragraph: Sorry, but we are having trouble signing you in. ' +
          "AADSTS900561: The endpoint only accepts POST requests. Request Id: bc99c67c",
      ),
    ).toBe("error");
  });

  it("calls a page with a field a prompt, and never touches it", () => {
    expect(
      authPageKind(
        "https://login.live.com/oauth20_authorize.srf",
        'heading "Sign in" textbox "Email, phone, or Skype" button "Next"',
      ),
    ).toBe("prompt");
  });

  it("resolves the ambiguous case toward PROMPT, because the two mistakes do not cost the same", () => {
    // An AADSTS code AND a way in. Treating this as an error would have the loop
    // navigate away from something a person may be part-way through; treating it
    // as a prompt costs one round. The asymmetry decides it.
    expect(
      authPageKind(
        "https://login.live.com/oauth20_authorize.srf",
        'AADSTS900561 heading "Sign in" button "Sign in" textbox "Email"',
      ),
      "an error page that also offers a way in was treated as recoverable",
    ).toBe("prompt");
  });

  it("defaults to PROMPT on an auth host it cannot read", () => {
    // Positive identification is required to recover. Anything unrecognised on a
    // login host is left alone.
    expect(authPageKind("https://login.live.com/x", "generic")).toBe("prompt");
    expect(authPageKind("https://login.microsoftonline.com/y", "")).toBe("prompt");
  });

  it("says nothing about pages that are not on an auth host at all", () => {
    expect(authPageKind("https://onedrive.live.com/", "Home - OneDrive")).toBe("not-auth");
    // Including one that merely MENTIONS the error — the host is what admits a
    // page to this judgement, not its text.
    expect(authPageKind("https://example.com/", "AADSTS900561")).toBe("not-auth");
  });
});

describe("a spawn that failed because the tool is missing, or because the box was busy", () => {
  /**
   * WHAT THIS COST: round 197, 2026-08-23, and the six retries that should have
   * absorbed it.
   *
   * The first call of the round was `eval () => String(window.innerWidth)` and
   * it came back `spawnSync … node.exe ETIMEDOUT`, while the full 124-file
   * vitest suite was running in the same session. playwright-cli was installed
   * and had driven six rounds that day.
   *
   * The driver exited 1 and told the reader to install the tool. Two faults in
   * one branch: the message named a CAUSE it could not know, and the stop
   * returned before any code was assigned, so `RECOVERABLE_STOPS` never saw it
   * and `--retry 6` never fired on a condition that clears by itself.
   */
  it("calls the real round-197 ETIMEDOUT transient", async () => {
    // @ts-expect-error — plain .mjs tool, no types.
    const { spawnFailureIsTransient } = await import("../scripts/round.mjs");
    expect(
      spawnFailureIsTransient(
        "spawnSync C:\\Users\\dann.pedersen\\AppData\\Local\\Microsoft\\WinGet\\Packages\\node.exe ETIMEDOUT",
      ),
    ).toBe(true);
  });

  it("keeps a genuinely missing binary FATAL, because retrying cannot conjure one", async () => {
    // @ts-expect-error — plain .mjs tool, no types.
    const { spawnFailureIsTransient } = await import("../scripts/round.mjs");
    expect(spawnFailureIsTransient("spawnSync playwright-cli ENOENT")).toBe(false);
  });

  it("defaults an unrecognised spawn error to FATAL, because a stop that loops is worse", async () => {
    // This loop runs unattended for hours. A wrong `transient` spins forever on
    // something no retry can fix; a wrong `fatal` costs one round and says why.
    // The asymmetry decides the default.
    // @ts-expect-error — plain .mjs tool, no types.
    const { spawnFailureIsTransient } = await import("../scripts/round.mjs");
    expect(spawnFailureIsTransient("something nobody has seen before")).toBe(false);
    expect(spawnFailureIsTransient(undefined)).toBe(false);
    expect(spawnFailureIsTransient("")).toBe(false);
  });

  it("puts the transient one where the retry loop can actually see it", async () => {
    // The whole defect was that this stop returned before a code was assigned,
    // so membership here is the half that makes the classifier do anything.
    // @ts-expect-error — plain .mjs tool, no types.
    const { RECOVERABLE_STOPS } = await import("../scripts/round.mjs");
    expect(RECOVERABLE_STOPS.has("cli-busy"), "cli-busy is not recoverable, so --retry still cannot see it").toBe(true);
  });
});
