import { defineConfig } from "vitest/config";

export default defineConfig({
  test: {
    coverage: {
      provider: "v8",
      // src plus the one skill script that is pure and importable: pptx-paint.mjs
      // (the paint/node helpers of the headless pptx renderer). render-pptx.mjs
      // itself is a subprocess CLI — v8 can't measure it — so only its extracted,
      // in-process-tested core is gated here.
      include: ["src/**", "skill/scripts/pptx-paint.mjs"],
      // Type-only modules and the browser/Office-only entry files. app.ts is
      // NOT excluded any more: the pane-state suite boots and drives the real
      // module, so the file most likely to regress is measured rather than
      // hidden — a floor below its current level guards against silent drops.
      //
      // The *.html entries are the pane/excel HTML shells. v8 lists them as
      // uncovered source and then tries to PARSE them as JS, dumping a
      // RolldownError ("Unexpected JSX expression") stack per run that buries
      // real failures in the coverage log — exclude them so the log stays clean.
      exclude: ["src/demo/**", "src/core/types.ts", "src/index.ts", "**/*.html"],
      reporter: ["text", "html"],
      thresholds: {
        // A global backstop: glob-keyed thresholds only gate files they match, so
        // without this a NEW directory outside every glob below would be measured
        // but never asserted. Set well under the current aggregate.
        statements: 85,
        branches: 75,
        // The pure engine is the product — hold it to a high bar.
        "src/core/**": { statements: 95, branches: 88 },
        "src/render/**": { statements: 85, branches: 80 },
        // The task pane is driven end-to-end (pane-state.test.ts + the host
        // command handlers in pane-host-actions.test.ts, which drive Insert /
        // Same-scale / Load against a mocked host) — a regression floor, not the
        // engine's bar. Raised from 75/58 once the Office-backed command layer
        // was covered; still a few points under today's numbers for headroom.
        "src/taskpane/**": { statements: 80, branches: 65 },
        // The Excel data bridge — a regression floor under today's numbers (it was
        // live code changed by #141 but matched no glob, so it was ungated).
        "src/excel/**": { statements: 80, branches: 70 },
        // The headless pptx renderer's pure paint/node core (pptx-paint.mjs),
        // unit-tested in-process by pptx-paint.test.ts. A floor under today's
        // 100/88 — the CLI wrapper stays out of coverage as a subprocess.
        "skill/scripts/**": { statements: 95, branches: 80 },
      },
    },
  },
});
