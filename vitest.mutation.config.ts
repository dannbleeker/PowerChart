import { defineConfig } from "vitest/config";

/**
 * The suite Stryker runs, which is not the whole suite.
 *
 * Mutation testing copies the project into a sandbox and runs the tests there,
 * and a good third of this suite does not survive that: some read files by a
 * path relative to the working directory (the manifests, the tracker table, the
 * docs gates), some need `dist-lib/` to have been built, and one needs a
 * browser. None of them covers `src/core` — the only thing being mutated — so
 * excluding them costs no signal and is what makes the run possible at all.
 *
 * Keep this list honest. A test excluded because it is inconvenient rather than
 * because it cannot run in a sandbox is signal thrown away, and a mutation score
 * computed over a suite nobody can name is a number rather than a finding.
 */
export default defineConfig({
  test: {
    exclude: [
      "**/node_modules/**",
      "**/dist/**",
      "**/.claude/**",
      "**/.stryker-tmp/**",
      // Needs `dist-lib/`, which the sandbox has no build step to produce.
      "test/skill-scripts.test.ts",
      "test/showcase.test.ts",
      "test/visible-charts.test.ts",
      // Reads repo files by a cwd-relative path.
      "test/manifest.test.ts",
      "test/office-js-watch.test.ts",
      "test/skill-docs.test.ts",
      "test/manual.test.ts",
      "test/test-count.test.ts",
      "test/host-contract.test.ts",
      "test/triage.test.ts",
      "test/verify-deck.test.ts",
      "test/ooxml-validate.test.ts",
      "test/crashlog.test.ts",
      // The Office.js and task-pane layers. They exercise `src/render` and
      // `src/taskpane`, not the engine, and they are the slowest third of the
      // suite — mutating core and running these would measure the fake.
      "test/office-render.test.ts",
      "test/web-host.test.ts",
      "test/selftest.test.ts",
      "test/host-probe.test.ts",
      "test/pane-*.test.ts",
      "test/dom-pane.test.ts",
      "test/demo.test.ts",
      "test/excel-*.test.ts",
      "test/templates.test.ts",
      "test/preflight-smoke.test.ts",
      // Two 180k-value stress grids. They pass in seconds normally and do not
      // survive per-test coverage instrumentation, which is a fact about
      // Stryker rather than about them — and the extent/spread code they guard
      // is still mutated, just scored by the other tests that reach it.
      "test/agg.test.ts",
    ],
  },
});
