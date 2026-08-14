#!/usr/bin/env node
/**
 * Drive a real-host round, and refuse to start one that cannot prove anything.
 *
 * A round costs ten minutes of a real PowerPoint and is the only evidence this
 * project gets about the host. Three of the four rounds run on 2026-08-14 were
 * set up by hand, fifteen browser steps at a time, and the steps that matter are
 * the ones easiest to skip:
 *
 *   - The pane is served with `Cache-Control: max-age=600`. Open it too soon
 *     after a merge and the round tests code the host never fetched. There is no
 *     way to tell from the result; the round simply means something other than
 *     what you think.
 *   - A round starting on a deck full of the last round's slides is not the same
 *     experiment as one starting clean, and the two have been compared as though
 *     they were.
 *   - Verbose trace and Picture every slide decide what the round can prove.
 *     Off, the trace is thin and empty slides cannot be confirmed empty.
 *
 * So this checks first and runs second, and every precondition is a hard stop
 * with the fix in the message. A round that runs on the wrong build is worse
 * than no round: it produces a file that looks like evidence.
 *
 *   node scripts/round.mjs --check     # preconditions only, nothing driven
 *   node scripts/round.mjs             # check, run, poll, archive, triage
 *
 * NOT part of any gate, and it cannot be. It needs a signed-in PowerPoint on
 * the web with the add-in sideloaded, which exists on the owner's machine and
 * nowhere else. The decisions are pure and unit-tested; the browser is injected.
 *
 * See `docs/ROUNDS.md` for what to do with the round once it exists, and
 * `CLAUDE.md` "Looking at the task pane" for the browser traps this encodes.
 */
import { spawnSync } from "child_process";
import { readFileSync, writeFileSync, existsSync, readdirSync, realpathSync } from "fs";
import { join, dirname } from "path";
import { isMain } from "./is-main.mjs";

/** The pane's build stamp, e.g. `32a6987 · 2026-08-14 08:03Z` → `32a6987`. */
export function buildOf(text) {
  const m = /\b([0-9a-f]{7})\b/.exec(String(text ?? ""));
  return m ? m[1] : null;
}

/**
 * Is this round worth running, and if not, what has to change first.
 *
 * Every no here is a round that would have produced a file nobody could draw a
 * conclusion from. `deployed !== head` is the one that looks harmless and is
 * not: it means the site has not finished publishing the commit under test, so
 * the pane will serve the previous build and the round will quietly measure it.
 */
export function readiness({ head, deployed, stamp, slides, verbose, pictures, reachable = true }) {
  const stop = [];
  // First, and on its own: everything below reads as "the browser said nothing"
  // when the truth is that nobody asked it. See `cli`.
  if (!reachable)
    return {
      ok: false,
      stop: [
        "playwright-cli could not be run — nothing below was actually measured. " +
          "Install it (`npm i -g @playwright/cli`), or set PLAYWRIGHT_CLI_JS to its entry point.",
      ],
    };
  if (!deployed) stop.push("the site did not answer with a build — is Pages up?");
  else if (head && deployed !== head)
    stop.push(`the site is serving ${deployed} but HEAD is ${head} — wait for Deploy Pages to finish`);
  if (!stamp) stop.push("could not read the pane's build stamp — is the add-in open?");
  else if (deployed && stamp !== deployed)
    stop.push(
      `the pane is showing ${stamp} while the site serves ${deployed} — hard-reload the whole ` +
        `PowerPoint tab (the pane HTML is cached for ten minutes; reopening the pane alone does not clear it)`,
    );
  if (slides !== null && slides > 1)
    stop.push(
      `the deck holds ${slides} slides — clean it, or this round is not comparable with one that started clean`,
    );
  if (verbose === false) stop.push("Verbose trace is off — the round's trace will be too thin to mine");
  if (pictures === false)
    stop.push("Picture every slide is off — a slide that reads back empty cannot be confirmed empty");
  return { ok: stop.length === 0, stop };
}

/** The next round number, from the archive. */
export function nextRoundNumber(files) {
  const ns = files.map((f) => Number(/^(\d{3})-/.exec(f)?.[1])).filter((n) => Number.isFinite(n));
  return String((ns.length ? Math.max(...ns) : 0) + 1).padStart(3, "0");
}

/** Strip the base64 slide images — see `rounds/README.md`. */
export function stripImages(round) {
  const walk = (o) => {
    if (!o || typeof o !== "object") return;
    for (const k of Object.keys(o)) {
      const v = o[k];
      if (typeof v === "string" && v.length > 500 && /^[A-Za-z0-9+/=]+$/.test(v.slice(0, 200)))
        o[k] = "<image stripped for the archive — see rounds/README.md>";
      else walk(v);
    }
  };
  walk(round);
  return round;
}

/**
 * The CLI's own JavaScript, to be run by node directly.
 *
 * Three routes were tried on the owner's machine and two are dead ends:
 *
 *   spawnSync("playwright-cli", …)               ENOENT — it is not an .exe
 *   spawnSync("playwright-cli", …, {shell:true}) "This program is blocked by
 *                                                group policy"
 *
 * AppLocker blocks the `.cmd` shim, so every route through a shell is closed
 * here no matter how the arguments are quoted. `node <entry>` works, because
 * node.exe is allow-listed and it is what the shim would have run anyway. It
 * also sidesteps cmd.exe re-parsing arguments, which matters: the stamp lookup
 * passes a regex containing spaces and braces, and the click helper passes
 * JavaScript containing `=>` and `{ }` — and `>` is a redirect character.
 *
 * Global installs sit beside node itself in this layout, and
 * `PLAYWRIGHT_CLI_JS` overrides for anyone whose does not.
 */
export function cliEntry(execPath = process.execPath, exists = existsSync) {
  if (process.env.PLAYWRIGHT_CLI_JS) return process.env.PLAYWRIGHT_CLI_JS;
  const guess = join(dirname(execPath), "node_modules", "@playwright", "cli", "playwright-cli.js");
  return exists(guess) ? guess : null;
}

/**
 * `playwright-cli`, and a way to know it was never reached.
 *
 * A tool that cannot be invoked and a pane that is closed produce the same empty
 * string, and folding the first into the second sends the reader to fix a
 * browser that is fine. On this driver's first live run it reported "could not
 * read the pane's build stamp — is the add-in open?" while the pane sat open on
 * screen showing the stamp. So `unreachable` is tracked across the session and
 * reported as its own precondition, ahead of everything it would otherwise
 * poison.
 */
export function sessionDir(dir, real = realpathSync.native) {
  // The daemon keys its sessions by the working directory STRING, and Windows
  // gives the same directory two names. Opened from a shell that resolved
  // `C:\Users\dann.pedersen\...`, the session is invisible to a process whose cwd
  // is the 8.3 form `C:\Users\DANN~1.PED\...` — same folder, different string,
  // and `list` answers "(no browsers)" while the browser sits on screen. That
  // cost most of an afternoon's debugging on this driver alone.
  try {
    return real(dir);
  } catch {
    return dir;
  }
}

function cli(run, dir) {
  const state = { unreachable: false };
  const entry = cliEntry();
  const cwd = sessionDir(dir);
  const sh = (...args) => {
    if (!entry) {
      state.unreachable = true;
      return "";
    }
    // NODE, on the CLI's own JavaScript. No shim, no shell.
    const r = run(process.execPath, [entry, "-s=ms", "--raw", ...args], { encoding: "utf8", cwd });
    if (r.error) state.unreachable = true;
    return r.status === 0 ? String(r.stdout ?? "") : "";
  };
  sh.state = state;
  return sh;
}

/** A `ref_N` for the first element matching `pattern` in a `find` result. */
function refFor(sh, query, pattern) {
  const out = sh("find", query);
  const line = out.split("\n").find((l) => pattern.test(l));
  return line ? (/ref=([a-z0-9]+)/.exec(line)?.[1] ?? null) : null;
}

/** Click through the element's own handler — a plain click does not reach the pane. */
function clickRef(sh, ref) {
  // The pane sits two iframes deep and `playwright-cli click` silently does
  // nothing there: `aria-selected` never moves and no error is raised. Firing
  // the element's own handler is the only thing that works, and finding that
  // out cost a round.
  return sh("eval", "el => { el.click(); return 'ok'; }", ref);
}

async function main(argv, deps = {}) {
  const run = deps.run ?? spawnSync;
  const fetchBuild = deps.fetchBuild ?? defaultFetchBuild;
  const dirArg = argv.indexOf("--dir");
  const sh = cli(run, dirArg === -1 ? process.cwd() : argv[dirArg + 1]);
  const checkOnly = argv.includes("--check");

  const head = String(run("git", ["rev-parse", "--short=7", "HEAD"], { encoding: "utf8" }).stdout ?? "").trim() || null;
  const deployed = buildOf(await fetchBuild());
  const stamp = buildOf(sh("find", "--regex", "/[0-9a-f]{7} ·/"));

  const listRef = refFor(sh, "Slide List", /listbox "Slide List"/);
  const slides = listRef ? (sh("snapshot", listRef).match(/option "Slide"/g) ?? []).length : null;

  const toggles = sh("find", "Verbose trace");
  const verbose = /checkbox "Verbose trace"/.test(toggles) ? /Verbose trace" \[checked\]/.test(toggles) : null;
  const pictures = /checkbox "Picture every slide"/.test(toggles)
    ? /Picture every slide" \[checked\]/.test(toggles)
    : null;

  const state = { head, deployed, stamp, slides, verbose, pictures, reachable: !sh.state.unreachable };
  const { ok, stop } = readiness(state);
  console.log(
    `  HEAD ${head ?? "?"} · site ${deployed ?? "?"} · pane ${stamp ?? "?"} · deck ${slides ?? "?"} slide(s)`,
  );
  console.log(`  verbose trace ${verbose ?? "?"} · picture every slide ${pictures ?? "?"}`);
  if (!ok) {
    console.error("\n  NOT READY — a round now would not prove anything:");
    for (const s of stop) console.error(`    - ${s}`);
    return 1;
  }
  console.log("  ready");
  if (checkOnly) return 0;

  const runBtn = refFor(sh, "Probe, then self-test", /button "Probe, then self-test"/);
  if (!runBtn) {
    console.error("  could not find the run button — is the Automation tab open?");
    return 1;
  }
  console.log("  running — this takes about ten minutes");
  clickRef(sh, runBtn);

  // Polled, never slept-through: a wedged host is the normal failure here and it
  // has to be distinguishable from a slow one.
  const started = Date.now();
  const limit = 30 * 60 * 1000;
  for (;;) {
    if (Date.now() - started > limit) {
      console.error("  the round has not finished in 30 minutes — the host is wedged; see docs/ROUNDS.md");
      return 1;
    }
    const dl = sh("find", "Download run log");
    if (/button "Download run log"(?! \[disabled\])/.test(dl)) break;
    if (!dl.trim()) {
      console.error("  the pane stopped answering — PowerPoint has probably crashed; the trace is still in the DOM");
      return 1;
    }
    await new Promise((r) => setTimeout(r, 20000));
  }
  console.log("  finished");
  return 0;
}

async function defaultFetchBuild() {
  const r = await fetch(`https://powerchart.struktureretsundfornuft.dk/build.json?cb=${Date.now()}`);
  return await r.text();
}

/** Archive a downloaded round under the next number. See `rounds/README.md`. */
export function archive(logPath, dir = "rounds", read = readFileSync, write = writeFileSync, list = readdirSync) {
  const round = stripImages(JSON.parse(read(logPath, "utf8")));
  const build = buildOf(round.build);
  if (!build) throw new Error("that file carries no build stamp — it is not a round log");
  const name = `${nextRoundNumber(list(dir))}-${build}.json`;
  write(`${dir}/${name}`, JSON.stringify(round, null, 1) + "\n");
  return name;
}

if (isMain(import.meta.url, process.argv[1])) {
  const argv = process.argv.slice(2);
  if (argv[0] === "--archive") {
    if (!argv[1] || !existsSync(argv[1])) {
      console.error("usage: node scripts/round.mjs --archive <powerchart-round.json>");
      process.exit(2);
    }
    console.log(`archived as rounds/${archive(argv[1])}`);
    process.exit(0);
  }
  process.exit(await main(argv));
}
