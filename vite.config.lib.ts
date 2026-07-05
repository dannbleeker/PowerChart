import { defineConfig } from "vite";
import { resolve } from "node:path";

// Library build for automation use: `npm run build:lib` → dist-lib/powerchart.js
export default defineConfig({
  build: {
    outDir: "dist-lib",
    lib: {
      entry: resolve(__dirname, "src/index.ts"),
      formats: ["es"],
      fileName: "powerchart",
    },
  },
});
