import { describe, expect, it } from "vitest";
import { readdirSync, readFileSync, statSync } from "fs";
import { join } from "path";

/**
 * No source file carries a stray control byte.
 *
 * THE TRAP THIS GUARDS HAS BITTEN EIGHT TIMES, twice in one night. Writing code
 * through a python or bash heredoc consumes one level of escaping before the
 * text reaches disk, so a `\b` in a regex arrives as a BACKSPACE byte and a `\n`
 * in a string arrives as a real newline. What lands is not what was written, and
 * neither grep nor an editor nor a diff shows it — only `cat -A` does.
 *
 * Both failure modes are on record from 2026-08-19:
 *
 *   LOUD   a `\n` inside a JS string literal became a real newline, so
 *          `rounds-gate.mjs` died with `SyntaxError: Invalid or unexpected
 *          token` on a line that reads perfectly in an editor.
 *   SILENT a word-boundary escape in a regex became a backspace byte. No parse
 *          error, no warning, exit 0 — the instrument matched nothing and
 *          returned an empty list. A brand-new detector reporting "nothing
 *          found" looks exactly like a clean bill of health, which is what makes
 *          that one dangerous; it was caught only by checking the output against
 *          a case known to be positive.
 *
 * NO REGEX AND NO ESCAPES IN THIS FILE, deliberately. Two earlier drafts spelled
 * the character class with `\u` escapes and the writing tool turned them into
 * the very bytes this test rejects — so the test caught itself, twice, before it
 * ever caught anything else. A plain character-code comparison cannot be
 * mangled on the way to disk, which is the whole point.
 *
 * Tab, newline and carriage return are allowed. Nothing else below 32 has ever
 * been wanted in this repo's source.
 */
const ROOTS = ["src", "scripts", "test", "docs"];
const TEXT = /\.(ts|tsx|mjs|js|md|json)$/;

const TAB = 9;
const NEWLINE = 10;
const CARRIAGE_RETURN = 13;
const FIRST_PRINTABLE = 32;

/** The index of the first stray control character, or -1. */
export function strayControlAt(text: string): number {
  for (let i = 0; i < text.length; i++) {
    const c = text.charCodeAt(i);
    if (c === TAB || c === NEWLINE || c === CARRIAGE_RETURN) continue;
    if (c < FIRST_PRINTABLE) return i;
  }
  return -1;
}

function walk(dir: string, out: string[] = []): string[] {
  for (const name of readdirSync(dir)) {
    if (name === "node_modules" || name === ".git") continue;
    const p = join(dir, name);
    if (statSync(p).isDirectory()) walk(p, out);
    else if (TEXT.test(name)) out.push(p);
  }
  return out;
}

describe("no source file carries a stray control byte", () => {
  it("scans every text file under src, scripts, test and docs", () => {
    const files = ROOTS.flatMap((r) => walk(r));
    // A scan that found nothing to scan would pass forever — the same shape as
    // the bug it guards.
    expect(files.length, "the walk found no files, so this test proves nothing").toBeGreaterThan(100);

    const bad: string[] = [];
    for (const f of files) {
      const text = readFileSync(f, "utf8");
      const at = strayControlAt(text);
      if (at < 0) continue;
      const line = text.slice(0, at).split("\n").length;
      const code = "0x" + text.charCodeAt(at).toString(16).padStart(2, "0");
      bad.push(`${f}:${line} carries ${code} — almost certainly a heredoc escape; see this file's header`);
    }
    expect(bad, bad.join("\n")).toEqual([]);
  });

  it("would actually catch one", () => {
    // Or the scan could be matching nothing for a reason nobody notices, which
    // is precisely the silent failure it exists to prevent.
    expect(strayControlAt(`a regex ${String.fromCharCode(8)}bracketed by backspace`), "missed a backspace").toBe(8);
    expect(strayControlAt(`a NUL ${String.fromCharCode(0)} here`), "missed a NUL").toBe(6);
    expect(strayControlAt("plain source\twith\ttabs\r\nand newlines"), "flagged ordinary whitespace").toBe(-1);
    expect(strayControlAt(""), "an empty file is not a finding").toBe(-1);
  });
});
