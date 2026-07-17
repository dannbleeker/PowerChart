import { defineConfig } from "vite";
import { resolve } from "node:path";
import { execSync } from "node:child_process";

/**
 * A build stamp the pane can show.
 *
 * PowerPoint caches an add-in's pane hard, so "reload and try again" is an act
 * of faith: there is otherwise no way to tell a fixed build from the cached old
 * one, and a whole debugging session can be spent testing code the host never
 * fetched. The stamp makes the running build legible at a glance.
 *
 * Commit + build time, because either alone can mislead: the commit says which
 * code, the time says which deploy of it.
 */
function buildStamp(): string {
  let sha = "nogit";
  try {
    sha = execSync("git rev-parse --short HEAD", { encoding: "utf8", stdio: ["ignore", "pipe", "ignore"] }).trim();
  } catch {
    /* no git (a packaged build) — the timestamp still identifies it */
  }
  const when = new Date().toISOString().slice(0, 16).replace("T", " ");
  return `${sha} · ${when}Z`;
}

// Two entry points:
//  - index.html          → standalone demo gallery (SVG renderer, no PowerPoint needed)
//  - src/taskpane/taskpane.html → the Office.js task pane loaded inside PowerPoint
export default defineConfig({
  define: {
    __BUILD_STAMP__: JSON.stringify(buildStamp()),
  },
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
