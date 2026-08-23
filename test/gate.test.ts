// @vitest-environment node
import { describe, expect, it } from "vitest";
import { readFileSync } from "fs";
// @ts-expect-error — plain .mjs tool, no types.
import { GATE_STEPS, NOT_COVERED, commandFor, runGate } from "../scripts/gate.mjs";

/**
 * WHAT THIS COST: a red CI on a push reported as green.
 *
 * On 2026-08-23 the local check was `prettier --check src test scripts docs` and
 * the claim made from it was "format clean". CI runs `prettier --check .`, the
 * driver had just archived a crash dump into `crashes/`, and `git add -A` swept
 * it in. The claim and the check had different scopes — the defect this repo
 * spends most of its comments on, arrived in the thing that verifies everything
 * else.
 *
 * The first version of the gate then failed twice more, and both are pinned
 * below: it called `isMain` with one argument where it takes two, so the script
 * printed nothing and exited 0 — a gate that silently passes everything — and
 * the test did not catch that, because it imported `runGate` and called it
 * directly, proving the FUNCTION worked while the SCRIPT did nothing.
 */
describe("the local gate", () => {
  it("takes its commands from package.json, so they cannot drift from CI's", () => {
    // THE REGRESSION THIS FILE EXISTS FOR. A hardcoded `prettier --check .`
    // here would be the same drift that caused the incident; reading the script
    // body means changing package.json changes the gate with it.
    const pkg = JSON.parse(readFileSync("package.json", "utf8")) as { scripts: Record<string, string> };
    for (const step of GATE_STEPS) {
      expect(commandFor(step), `gate step "${step}" does not match package.json`).toBe(pkg.scripts[step]);
    }
    // The specific one that broke: the gate must go through CI's format check,
    // and that check must still be repo-wide.
    expect(GATE_STEPS, "the gate no longer runs CI's format check").toContain("format:check");
    expect(pkg.scripts["format:check"], "format:check stopped checking the whole repo").toContain("prettier --check .");
    // A step CI does not have is a typo, and must fail loudly rather than be skipped.
    expect(() => commandFor("no-such-step")).toThrow(/no script/);
  });

  it("is wired to actually RUN when invoked as a script", () => {
    // The first version called `isMain(import.meta.url)` where it takes
    // (moduleUrl, argv1), so the guard was always false: `node scripts/gate.mjs`
    // printed nothing and exited 0. is-main.mjs's own docstring names that
    // failure — "every one of these CLIs printed nothing and exited 0" — and the
    // gate reproduced it. Asserted on the source, because the property is about
    // how the module is invoked, which an import cannot observe.
    const src = readFileSync("scripts/gate.mjs", "utf8");
    expect(src, "the isMain guard lost its argv argument and the gate is a no-op").toContain(
      "isMain(import.meta.url, process.argv[1])",
    );
  });

  it("stops at the first failure instead of reporting a partial pass", () => {
    const ran: string[] = [];
    const result = runGate({
      // `run` receives (command, step); keying on the step keeps this readable.
      run: (_cmd: string, step: string) => {
        ran.push(step);
        return step === "format:check" ? 1 : 0;
      },
      log: () => {},
    });
    expect(result.ok).toBe(false);
    expect(result.failedAt).toBe("format:check");
    expect(ran.at(-1), "the gate kept going after a failing step").toBe("format:check");
  });

  it("says out loud what it does NOT run, so green cannot be read as 'CI will pass'", () => {
    const lines: string[] = [];
    const result = runGate({ run: () => 0, log: (m: string) => lines.push(String(m)) });
    expect(result.ok).toBe(true);
    // The gate is a SUBSET and must admit it. This is why the list is exported
    // rather than being a comment.
    expect(NOT_COVERED.length, "a gate that claims to cover everything is the original defect").toBeGreaterThan(0);
    expect(lines.join("\n"), "the skipped steps were not printed on success").toContain("NOT run by this gate");
    // Coverage is the biggest gap: `npm test` does not enforce CI's thresholds.
    expect(NOT_COVERED.join(" "), "the coverage gap is not named").toMatch(/coverage/i);
  });
});
