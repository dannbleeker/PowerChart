/**
 * Was this module run as the CLI, or merely imported?
 *
 * Every script here that exports its decision logic for a test also has to know
 * not to run `main()` when that test imports it. Four of them answered the
 * question four different ways, and three of the four answers were wrong on
 * Windows:
 *
 *   import.meta.url.endsWith(process.argv[1].split("/").pop())
 *
 * `process.argv[1]` on Windows is `C:\repo\scripts\flaky.mjs`. Splitting that on
 * a forward slash yields the whole path back, and `file:///C:/repo/scripts/
 * flaky.mjs` does not end with a string full of backslashes — so the guard was
 * false every time and `main()` never ran. `node scripts/test-count.mjs
 * results.json` printed nothing and exited 0, which does not read as "this tool
 * did not run". It reads as a pass. The suite-shrink guard, the flake sweep and
 * the visual gate all did that on the owner's box, and the office-js watch did
 * until it was fixed in isolation — the same bug, found once and not swept for.
 *
 * The platform is a PARAMETER rather than a read of `process.platform`, and that
 * is the load-bearing part of this file. CI runs on ubuntu, where the broken
 * form works; a test that could only exercise the platform it runs on would go
 * green against the bug it exists to catch.
 */
import { fileURLToPath } from "node:url";
import fs from "node:fs";
import path from "node:path";

/**
 * @param {string} moduleUrl  the caller's own `import.meta.url`
 * @param {string|undefined} argv1  `process.argv[1]`, the path node was handed
 * @param {string} platform  `process.platform`; injected so BOTH cases are testable anywhere
 */
export function isMain(moduleUrl, argv1, platform = process.platform) {
  if (!moduleUrl || !argv1) return false;
  const windows = platform === "win32";
  let self;
  try {
    // The URL→path conversion and the path arithmetic both take their flavour
    // from `platform`, which is what lets an ubuntu runner exercise the Windows
    // case. Without that the regression this file exists for is unreachable in
    // CI, and a guard that cannot reach its bug is decoration.
    self = fileURLToPath(moduleUrl, { windows });
  } catch {
    // A non-file URL is not a path and cannot be the thing node was invoked
    // with. Not an error — under a bundler or a data: URL there is simply no CLI.
    return false;
  }
  const invoked = (windows ? path.win32 : path.posix).resolve(argv1);
  // Windows paths are case-insensitive and the drive letter's case is not
  // stable — the shell's `c:\` and node's `C:\` name one file and would
  // otherwise compare unequal.
  const same = (a, b) => (windows ? a.toLowerCase() === b.toLowerCase() : a === b);
  if (same(self, invoked)) return true;
  // A SYMLINK names one file by two paths, and the two sides of this comparison
  // come from different places: `import.meta.url` is always the resolved REAL
  // path, while `process.argv[1]` is whatever the shell handed over. Reached
  // through a link — a pnpm/workspace bin shim, a checkout under a linked
  // parent directory, `/tmp` on a mac — the strings differ and every one of
  // these CLIs printed nothing and exited 0. That is the exact failure this
  // file exists to stop, arriving by a second route.
  //
  // Only asked when the platform is the real one: an injected platform (which
  // is what lets an ubuntu runner exercise the Windows case) has no filesystem
  // to resolve against, and must keep answering from the strings alone.
  if (platform !== process.platform) return false;
  try {
    return same(fs.realpathSync(self), fs.realpathSync(invoked));
  } catch {
    // One of them does not exist on disk. Then it is not the file node ran.
    return false;
  }
}
