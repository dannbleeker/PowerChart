import { describe, expect, it } from "vitest";
// @ts-expect-error — a plain .mjs tool with no types.
import { isMain } from "../scripts/is-main.mjs";

/**
 * The predicate that decides whether a script is the CLI or a test's import.
 *
 * Every case here is a real invocation that used to answer wrong. Four scripts
 * carried four different answers to this question and three of them were false
 * on Windows for the same reason — `argv[1].split("/")` does not split a
 * backslashed path — so `node scripts/test-count.mjs results.json` printed
 * nothing and exited 0. That does not read as "the tool did not run". It reads
 * as a pass, from the guard against a silently shrinking suite, from the flake
 * sweep, and from the visual gate.
 *
 * The platform is injected on purpose. CI runs on ubuntu, where the broken form
 * works perfectly; a suite that could only exercise its own platform would go
 * green against the bug these cases exist to catch.
 */
describe("isMain", () => {
  const WIN_URL = "file:///C:/repo/scripts/flaky.mjs";
  const WIN_ARGV = "C:\\repo\\scripts\\flaky.mjs";
  const POSIX_URL = "file:///repo/scripts/flaky.mjs";
  const POSIX_ARGV = "/repo/scripts/flaky.mjs";

  it("says yes when Windows ran the script directly", () => {
    // The regression. `import.meta.url.endsWith(argv[1].split("/").pop())`
    // returns false here, which is how three CLIs came to be no-ops.
    expect(isMain(WIN_URL, WIN_ARGV, "win32")).toBe(true);
  });

  it("says yes when the drive letter's case differs", () => {
    // cmd, PowerShell and git-bash do not agree on the case of the drive
    // letter, and node keeps whichever one it was handed. Same file.
    expect(isMain("file:///c:/repo/scripts/flaky.mjs", "C:\\repo\\scripts\\flaky.mjs", "win32")).toBe(true);
    expect(isMain(WIN_URL, "c:\\REPO\\scripts\\flaky.mjs", "win32")).toBe(true);
  });

  it("says yes when posix ran the script directly", () => {
    expect(isMain(POSIX_URL, POSIX_ARGV, "linux")).toBe(true);
  });

  it("says no when something else was invoked — a test importing the module", () => {
    // The half the old guard got right, and the only half its tests could see:
    // vitest is argv[1] when a spec imports one of these tools, and a CLI that
    // ran on import would call process.exit in the middle of the suite.
    expect(isMain(POSIX_URL, "/repo/node_modules/vitest/vitest.mjs", "linux")).toBe(false);
    expect(isMain(WIN_URL, "C:\\repo\\node_modules\\vitest\\vitest.mjs", "win32")).toBe(false);
  });

  it("says no for a sibling script that merely shares a basename", () => {
    // A basename test — the other tempting shortcut — matches any script called
    // flaky.mjs anywhere on the machine.
    expect(isMain(POSIX_URL, "/somewhere/else/flaky.mjs", "linux")).toBe(false);
  });

  it("says no when there is no argv[1] and no usable URL", () => {
    expect(isMain(POSIX_URL, undefined, "linux")).toBe(false);
    expect(isMain(undefined, POSIX_ARGV, "linux")).toBe(false);
    // A data: or http: module URL is not a path. Not an error — under a bundler
    // there is simply no CLI to be.
    expect(isMain("data:text/javascript,0", POSIX_ARGV, "linux")).toBe(false);
    expect(isMain("https://example.com/flaky.mjs", POSIX_ARGV, "linux")).toBe(false);
  });

  it("resolves a relative argv[1], which is how npm scripts invoke these", () => {
    // `node scripts/flaky.mjs` hands node a relative path.
    const cwd = process.cwd();
    const url = new URL(`file://${cwd.replace(/\\/g, "/").replace(/^([A-Za-z]:)/, "/$1")}/scripts/flaky.mjs`).href;
    expect(isMain(url, "scripts/flaky.mjs", process.platform)).toBe(true);
  });
});
