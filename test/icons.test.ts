import { describe, it, expect } from "vitest";
import { readFileSync } from "node:fs";
import { execFileSync } from "node:child_process";

/**
 * The add-in icons, and the palette they are allowed to use.
 *
 * These are generated (`scripts/build-icons.mjs`) rather than hand-exported,
 * because they are four sizes of one mark and the 16px one is where an export
 * goes wrong: a bar landing on a half pixel turns grey, and at 16px a grey bar
 * is an invisible one.
 *
 * The check below is the same shape as the showcase deck's in CI — regenerate,
 * compare to what is committed, fail on a difference — so the committed PNGs
 * cannot drift from the script that claims to produce them.
 */
describe("the add-in icons", () => {
  it("match what the generator produces", () => {
    // `--check` exits non-zero and names the stale file.
    expect(() => execFileSync("node", ["scripts/build-icons.mjs", "--check"], { encoding: "utf8" })).not.toThrow();
  });

  it("use only SSF palette colours", () => {
    // A PNG's palette is not readable without decoding it, so this reads the
    // generator's own constants — the thing that decides the pixels. The test
    // above is what ties those constants to the bytes on disk.
    const src = readFileSync("scripts/build-icons.mjs", "utf8");
    // NAMED CONSTANTS ONLY. The first version of this matched any three
    // adjacent `0x..` bytes and flagged `#89504E` — the PNG magic number in the
    // encoder below. A colour check that reads a file signature as a colour
    // would eventually be silenced rather than fixed.
    const hexes = [...src.matchAll(/^const [A-Z]+ = \[0x([0-9a-f]{2}), 0x([0-9a-f]{2}), 0x([0-9a-f]{2})\];$/gm)].map(
      (m) => `#${m[1]}${m[2]}${m[3]}`.toUpperCase(),
    );
    expect(hexes.length, "the generator declares no colours").toBeGreaterThan(0);
    // navy ground, white bars, one orange accent — nothing else.
    const allowed = new Set(["#00254C", "#FFFFFF", "#ED8936"]);
    for (const hex of hexes) expect(allowed.has(hex), `${hex} is not in the SSF palette`).toBe(true);
    // NO SEPARATE "the old blue must not come back" CHECK. There was one, as a
    // file-wide `/2a78d6/i`, and it failed on the generator's own comment
    // explaining why that blue was removed — a check that cannot tell a colour
    // declaration from a sentence about one. The constant scan above already
    // rejects it: putting `#2A78D6` back in NAVY fails with its own name.
  });

  it("are the three sizes the manifests actually reference, plus one", () => {
    // 16, 32 and 80 are wired into both manifests; 64 is generated but
    // referenced by neither. Kept because a store listing wants it and
    // regenerating one size later would mean rediscovering the geometry.
    const manifest = readFileSync("manifest.xml", "utf8");
    for (const size of [16, 32, 80]) {
      expect(manifest, `the manifest stopped referencing icon-${size}.png`).toContain(`assets/icon-${size}.png`);
    }
  });
});
