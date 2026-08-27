import { describe, it, expect } from "vitest";
import { readFileSync, existsSync, readdirSync, statSync } from "node:fs";
import { createRequire } from "node:module";
import { join } from "node:path";

/**
 * The `image-size` stub, and why it is allowed to exist.
 *
 * Two HIGH advisories cover EVERY published version of `image-size` — both give
 * their range as `<=2.0.2` and 2.0.2 is the latest — so the `overrides` pin that
 * cleared `qs` has no target. npm's own remedy is a three-MAJOR downgrade of
 * pptxgenjs, the deck writer, to patch a parser that never executes.
 *
 * What makes replacing it safe is that nothing imports it: it appears in
 * pptxgenjs's `package.json` and nowhere in pptxgenjs's code, nothing else in
 * the tree requires it, and this project never calls `addImage` — the API the
 * real package exists to serve.
 *
 * THESE TESTS ASSERT THE PROPERTY, NOT THE LAYOUT. The first version resolved
 * `image-size` through a hardcoded `node_modules/pptxgenjs/index.js` base. That
 * passed on Windows and failed on Linux CI with "Cannot find module", because
 * where npm puts an overridden `file:` dependency is npm's business and it
 * differs by platform and by whether the tree was built with `install` or `ci`.
 * What actually matters is that the vulnerable code is not installed and that
 * the stub is loud — and neither of those needs to know where npm filed it.
 */
const STUB_DIR = "vendor/image-size-stub";

/** Every directory named `image-size` anywhere under node_modules. */
function installedCopies(root = "node_modules", found: string[] = []): string[] {
  let entries: string[];
  try {
    entries = readdirSync(root);
  } catch {
    return found;
  }
  for (const name of entries) {
    const path = join(root, name);
    let dir: boolean;
    try {
      dir = statSync(path).isDirectory();
    } catch {
      continue; // a broken link is not a copy of anything
    }
    if (!dir) continue;
    if (name === "image-size") found.push(path);
    else if (name === "node_modules" || !name.startsWith(".")) installedCopies(path, found);
  }
  return found;
}

describe("the image-size stub", () => {
  it("is still wired up in overrides", () => {
    const pkg = JSON.parse(readFileSync("package.json", "utf8"));
    expect(
      pkg.overrides?.["image-size"],
      "the override was removed — run `npm audit` before assuming that is safe",
    ).toBe(`file:${STUB_DIR}`);
  });

  it("THROWS rather than answering, so the assumption fails loudly", () => {
    // Loaded from OUR directory, which is in the repo and therefore always
    // there. Requiring it by package name would be testing npm's hoisting.
    const requireHere = createRequire(join(process.cwd(), "package.json"));
    const stub = requireHere(`./${STUB_DIR}/index.js`) as () => unknown;
    expect(() => stub()).toThrow(/stubbed in this project/);
    // It names where to go, because whoever hits this will have just added an
    // image to a deck and have no idea why it exploded.
    expect(() => stub()).toThrow(/vendor\/image-size-stub/);
  });

  it("installs no parser — the vulnerable code is nowhere in the tree", () => {
    // The advisories are in the ICNS and JXL/HEIF parsers. The point of the
    // override is that those files are never installed, not that they are
    // installed and unused. Whatever npm filed under whatever path, every copy
    // must be ours.
    // ZERO COPIES IS A PASS, and demanding otherwise is what made this test
    // fail on Linux CI while passing here. The property is "the vulnerable
    // parsers are not installed"; nothing installed satisfies it completely.
    // How npm represents an overridden `file:` dependency — a junction into the
    // repo on Windows, a symlink that may not even resolve on a `npm ci` tree —
    // is npm's business, and asserting a copy must EXIST was asserting that.
    //
    // The override staying wired is checked above, from `package.json`, which
    // is the portable place to check it.
    const copies = installedCopies();
    for (const dir of copies) {
      const manifest = join(dir, "package.json");
      expect(existsSync(manifest), `${dir} has no package.json`).toBe(true);
      const meta = JSON.parse(readFileSync(manifest, "utf8"));
      expect(meta.description, `${dir} is not our stub — the real package is installed`).toMatch(/Stub\./);
      expect(readFileSync(join(dir, "index.js"), "utf8")).not.toMatch(/icns|heif|jxl/i);
    }
  });
});
