import { defineConfig } from "vite";
import { resolve } from "node:path";

// Two entry points:
//  - index.html          → standalone demo gallery (SVG renderer, no PowerPoint needed)
//  - src/taskpane/taskpane.html → the Office.js task pane loaded inside PowerPoint
export default defineConfig({
  build: {
    rollupOptions: {
      input: {
        demo: resolve(__dirname, "index.html"),
        taskpane: resolve(__dirname, "src/taskpane/taskpane.html"),
        excel: resolve(__dirname, "src/excel/excel.html"),
      },
    },
  },
  server: {
    port: 3000,
    // Office add-ins must be served over HTTPS in production; for local sideloading
    // use `npx office-addin-dev-certs install` and point `https` at those certs.
  },
});
