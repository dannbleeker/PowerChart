import { defineConfig } from "vitest/config";

export default defineConfig({
  test: {
    coverage: {
      provider: "v8",
      include: ["src/**"],
      // Entry/bootstrap files that only run inside a browser or Office host,
      // plus type-only modules; everything with logic stays measured.
      exclude: ["src/demo/**", "src/taskpane/app.ts", "src/core/types.ts", "src/index.ts"],
      reporter: ["text", "html"],
      thresholds: {
        // The pure engine is the product — hold it to a high bar.
        "src/core/**": { statements: 95, branches: 88 },
        "src/render/**": { statements: 85, branches: 80 },
      },
    },
  },
});
