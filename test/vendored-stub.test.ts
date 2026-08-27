import { describe, it, expect } from "vitest";
import { readFileSync } from "node:fs";
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
 * These tests guard the two halves of that: the override stays wired, and the
 * stub stays LOUD. A stub that quietly returned a plausible size would be worse
 * than the vulnerability, because a wrong dimension would flow into a generated
 * deck and nothing would say so.
 */
describe("the image-size stub", () => {
  it("is still wired up in overrides", () => {
    const pkg = JSON.parse(readFileSync("package.json", "utf8"));
    expect(
      pkg.overrides?.["image-size"],
      "the override was removed — check `npm audit` before assuming that is safe",
    ).toBe("file:vendor/image-size-stub");
  });

  it("THROWS rather than answering, from where pptxgenjs resolves it", () => {
    // Resolved through pptxgenjs on purpose. The override installs the stub
    // nested under the package that declares the dependency, so requiring
    // "image-size" from the repo root does not find it at all — a check written
    // from the root would pass for the wrong reason.
    const requireFromPptx = createRequire(join(process.cwd(), "node_modules", "pptxgenjs", "index.js"));
    const stub = requireFromPptx("image-size") as () => unknown;
    expect(() => stub()).toThrow(/stubbed in this project/);
    // It names where to go, because the person hitting this will be someone who
    // just added an image to a deck and has no idea why it exploded.
    expect(() => stub()).toThrow(/vendor\/image-size-stub/);
  });

  it("ships no parser — the vulnerable code is not on disk", () => {
    // The advisories are in the ICNS and JXL/HEIF parsers. The point of the
    // override is that those files are never installed, not that they are
    // installed and unused.
    const requireFromPptx = createRequire(join(process.cwd(), "node_modules", "pptxgenjs", "index.js"));
    const onDisk = readFileSync(requireFromPptx.resolve("image-size"), "utf8");
    expect(onDisk).not.toMatch(/icns|heif|jxl/i);
  });
});
