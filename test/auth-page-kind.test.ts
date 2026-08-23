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
