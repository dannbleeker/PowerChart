import { describe, expect, it } from "vitest";
import { readFileSync } from "fs";
import { fileURLToPath } from "url";

/**
 * The manifests, checked offline against the rules that matter.
 *
 * CI runs Microsoft's own `office-addin-manifest validate` in a job of its own,
 * and that is the authority. This is not a second copy of it: it pins the small
 * number of rules whose violation this project has actually met, so they cannot
 * come back when the validator is unreachable — which is not hypothetical. The
 * validator calls a Microsoft SERVICE, so it cannot run in a sandbox with no
 * route to it, and for the whole life of this repo that meant the manifests were
 * never validated at all.
 *
 * The first thing the CI job found on its first run: `<Version>0.1.0</Version>`.
 * "Manifest Version Too Low: The manifest has unsupported version number less
 * than 1.0" — an error, on all four manifests, since the day they were written.
 * It passed every test in this repo, because nothing here had ever looked.
 */
const MANIFESTS = ["manifest.xml", "manifest-excel.xml", "manifest-prod.xml", "manifest-excel-prod.xml"];

const read = (name: string) => readFileSync(fileURLToPath(new URL(`../${name}`, import.meta.url)), "utf8");

describe("the add-in manifests", () => {
  it.each(MANIFESTS)("%s declares a version Office will accept", (name) => {
    const version = /<Version>([^<]+)<\/Version>/.exec(read(name))?.[1];
    expect(version, `${name} has no <Version>`).toBeTruthy();
    const parts = version!.split(".").map((p) => Number(p));
    expect(
      parts.every((p) => Number.isInteger(p) && p >= 0),
      `${name}: "${version}" is not a dotted number`,
    ).toBe(true);
    expect(parts.length, `${name}: "${version}" needs at least major.minor`).toBeGreaterThanOrEqual(2);
    // The rule verbatim: "unsupported version number less than 1.0". A major of
    // zero is what the four shipped manifests had, and it is the one thing here
    // that has actually been wrong.
    expect(parts[0], `${name}: "${version}" is below 1.0, which Office rejects`).toBeGreaterThanOrEqual(1);
  });

  it.each(MANIFESTS)("%s keeps its own <Id>", (name) => {
    // Changing an add-in's GUID makes it a different add-in: every sideload is
    // orphaned and every deck's association with it is lost. `CLAUDE.md` says
    // "never do this"; this is what makes it a test rather than a hope.
    const id = /<Id>([^<]+)<\/Id>/.exec(read(name))?.[1];
    const expected = name.startsWith("manifest-excel")
      ? "c8a7e4b3-5d2f-4f9b-8a3c-8e6d1b2f7a54"
      : "b7f6d3a2-4c1e-4e8a-9f2b-7d5c0a1e6f43";
    expect(id, `${name}'s add-in GUID changed — every existing sideload would be orphaned`).toBe(expected);
  });

  it.each(["manifest-prod.xml", "manifest-excel-prod.xml"])("%s points at no localhost URL", (name) => {
    // `build-manifest.mjs --check` already gates staleness in CI, which is a
    // different question: a prod manifest can be perfectly current AND full of
    // localhost if the rewrite ever stops matching. This asks the thing that
    // makes the file worth sideloading.
    expect(read(name)).not.toMatch(/localhost/i);
  });
});
