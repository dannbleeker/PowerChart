import { defineConfig } from "vitest/config";

export default defineConfig({
  test: {
    coverage: {
      provider: "v8",
      include: ["src/**"],
      // Type-only modules and the browser/Office-only entry files. app.ts is
      // NOT excluded any more: the pane-state suite boots and drives the real
      // module, so the file most likely to regress is measured rather than
      // hidden — a floor below its current level guards against silent drops.
      exclude: ["src/demo/**", "src/core/types.ts", "src/index.ts"],
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
      },
    },
  },
});
