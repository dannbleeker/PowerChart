// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from "vitest";
import { readFileSync } from "node:fs";

/**
 * Does the pane notice it is older than the site it came from?
 *
 * GitHub Pages serves the pane's HTML with `Cache-Control: max-age=600` and
 * gives us no way to set headers, so for ten minutes after a deploy PowerPoint
 * hands back the cached page — which names the PREVIOUS hashed bundle, still in
 * the browser's own cache even after that file 404s on the server. Old page,
 * old script, old build stamp, and until now nothing on screen saying so.
 *
 * It has cost two rounds. One ran a fix that was not in the build under test
 * and the result was read as evidence about it; the other went on hard-reloading
 * a pane that looked the same either way. The stamp was always there to read —
 * what was missing is that reading it only helps if you already know the number
 * it should be, and that number is on GitHub.
 */
const STAMP = "abc1234 · 2026-08-10 08:10Z";
const NEWER = "def5678 · 2026-08-10 09:30Z";

/** Boot the pane against the real markup, with a chosen answer from build.json. */
async function boot(build: { ok: boolean; body?: unknown } | "reject") {
  const parsed = new DOMParser().parseFromString(readFileSync("src/taskpane/taskpane.html", "utf8"), "text/html");
  parsed.querySelectorAll("script").forEach((s) => s.remove());
  document.body.innerHTML = parsed.body.innerHTML;
  vi.stubGlobal("__BUILD_STAMP__", STAMP);
  vi.stubGlobal(
    "fetch",
    vi.fn(async () =>
      build === "reject"
        ? Promise.reject(new Error("blocked by the host"))
        : ({ ok: build.ok, json: async () => build.body } as unknown as Response),
    ),
  );
  vi.resetModules();
  await import("../src/taskpane/app");
  // The check is fired and not awaited, so let its microtasks run.
  await new Promise((r) => setTimeout(r, 0));
  return document.getElementById("build-stamp")!;
}

afterEach(() => vi.unstubAllGlobals());

describe("a pane that is older than the site it came from", () => {
  it("says so, in the one place the build is already written", async () => {
    const el = await boot({ ok: true, body: { build: NEWER } });
    expect(el.textContent, "the pane did not name the newer build").toContain(NEWER);
    expect(el.textContent, "the pane stopped naming its OWN build").toContain(STAMP);
    expect(el.classList.contains("stale"), "nothing marked it as a problem").toBe(true);
  });

  it("stays quiet when the site agrees with it", async () => {
    // The failure mode that would make this feature worse than nothing: a
    // warning on every boot is one nobody reads, including on the boot that
    // matters. `build.json` is written from the stamp inside the built bundle
    // for exactly this reason — recomputing it would disagree by a minute or
    // two on every deploy and cry stale forever.
    const el = await boot({ ok: true, body: { build: STAMP } });
    expect(el.textContent).toBe(STAMP);
    expect(el.classList.contains("stale")).toBe(false);
  });

  it("says nothing at all when it cannot find out", async () => {
    // This is the add-in's FIRST outbound request, and the host may simply
    // refuse it — the pane runs inside an Office iframe with a CSP nobody here
    // controls. A convenience that breaks the pane it is diagnosing is worse
    // than the trap it replaces, so every way of not knowing is silent.
    // The non-OK case carries a body that WOULD warn, so the response check is
    // what has to reject it. With an empty body the later "is it a string"
    // guard caught it instead and `!res.ok` could be deleted with every test
    // still green — a 404 page that happens to parse is exactly the case it is
    // there for.
    for (const answer of [
      { ok: false, body: { build: NEWER } },
      { ok: true, body: {} },
      { ok: true, body: { build: "" } },
      "reject",
    ] as const) {
      const el = await boot(answer);
      expect(el.textContent, `answered "${JSON.stringify(answer)}" with a warning`).toBe(STAMP);
      expect(el.classList.contains("stale")).toBe(false);
    }
  });
});
