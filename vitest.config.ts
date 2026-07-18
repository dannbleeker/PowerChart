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
        // The pure engine is the product — hold it to a high bar.
        "src/core/**": { statements: 95, branches: 88 },
        "src/render/**": { statements: 85, branches: 80 },
        // The task pane is driven end-to-end (pane-state.test.ts), not
        // unit-tested — a regression floor, not the engine's bar.
        "src/taskpane/**": { statements: 75, branches: 58 },
      },
    },
  },
});
