import { defineConfig } from "vite";
import { resolve } from "node:path";

// Library build for automation use: `npm run build:lib` → dist-lib/ssf-charts.js
export default defineConfig({
  build: {
    outDir: "dist-lib",
    lib: {
      entry: resolve(__dirname, "src/index.ts"),
      formats: ["es"],
      fileName: "ssf-charts",
    },
    rollupOptions: {
      // jszip stays a peer, not a passenger. Bundling it put a 133 KB copy
      // inside the library — and, because the import is dynamic, split the
      // build into chunks. The skill ships exactly one engine file and would
      // have loaded a chunk that was never packaged with it.
      external: ["jszip"],
      output: { inlineDynamicImports: true },
    },
  },
});
