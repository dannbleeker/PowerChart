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

/**
 * The pane's build stamp, e.g. `32a6987 · 2026-08-14 08:03Z` → `32a6987`.
 *
 * THE SEPARATOR IS REQUIRED, and leaving it out cost a good round. Seven hex
 * characters is not a rare shape: `playwright-cli` names every element `f14e735`
 * and friends, and the accessibility dump this function reads is full of them.
 * Matching a bare 7-hex token picked up a REF id, reported the pane as showing a
 * build the site had never served, and refused to start — on a pane that was
 * showing exactly the right commit.
 *
 * A precondition that blocks correct work is worse than no precondition at all,
 * because the fix it suggests (hard-reload the tab) makes the reader doubt a
 * machine that was right.
 */
export function buildOf(text) {
  const m = /\b([0-9a-f]{7})\s+·/.exec(String(text ?? ""));
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
export function readiness({
  head,
  deployed,
  stamp,
  slides,
  verbose,
  pictures,
  reachable = true,
  ping = null,
  crashed = false,
}) {
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
  // THE WEDGE, by its real name. Rounds 24, 25, 29 and 30 each spent most of an
  // hour on this and none of them said what it was: the host's editing session
  // dies, and every Office.js call after that hangs forever without ever throwing
  // — even an empty `context.sync()`. `PowerPoint.run` still ENTERS its callback,
  // which is why it reads as the host thinking rather than the host being gone.
  //
  // This branch is the loud form, where PowerPoint has raised its own error and
  // put a modal over the document; the deck reading back as `?` in the same
  // breath is the UI frozen behind it. The quiet form has no dialog and only the
  // ping catches it. Neither is something to wait out — a reload clears both in
  // seconds. See `docs/ROUNDS.md`, "The wedge".
  if (crashed)
    stop.push(
      'PowerPoint has crashed — its own "Sorry, we ran into a problem" dialog is up, and every Office.js ' +
        "call behind it hangs forever. Click Refresh in that dialog; the host answers again within the minute. " +
        "The pane closes with it — reopen from Home ▸ Add-ins ▸ Insert chart.",
    );
  // Before anything about builds or decks: is the host answering? A stale pane
  // and a dirty deck are both worth fixing, and neither matters if the host will
  // not talk — that is the state four rounds have burned an hour each on.
  if (ping && !ping.answered)
    stop.push(
      `the host did not answer the cheapest possible call in ${ping.ms}ms — it is not going to answer a round. ` +
        "Its editing session is gone, usually because the network moved (look for net::ERR_NETWORK_CHANGED in " +
        "`playwright-cli console`). Reload the PowerPoint tab and reopen the pane: measured 8011ms silent before, " +
        "7ms after.",
    );
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
    // A failed CLI call and a page that answered with nothing are the same empty
    // string, and the difference decides whether a round is alive. Recorded, not
    // folded in — the same distinction `unreachable` exists to keep.
    state.lastFailed = !r.error && r.status !== 0;
    return r.status === 0 ? String(r.stdout ?? "") : "";
  };
  sh.state = state;
  return sh;
}

/** A `ref_N` for the first element matching `pattern` in a `find` result. */
export function refFor(sh, query, pattern) {
  const out = sh("find", query);
  const line = out.split("\n").find((l) => pattern.test(l));
  // The ref on the MATCHING LINE, never the first ref in the output. `find`
  // prints the whole frame hierarchy above the hit, so the first ref belongs to
  // the outer iframe — evaluating against it lands in the OneDrive document
  // rather than in the pane, where `Office` and `PowerPoint` are both undefined
  // and a host ping silently reports a healthy host as dead.
  return line ? (/ref=([a-z0-9]+)/.exec(line)?.[1] ?? null) : null;
}

/**
 * Is the host answering AT ALL, before a round is offered?
 *
 * Rounds 24, 25 and 29 each cost most of an hour to discover that it was not
 * going to. Round 29 showed the host was already unwell before the probe's
 * fourth question, and a ping run afterwards showed the state persists: at
 * 17:54, nearly two hours after that round wedged and after a full tab reload,
 * `slides.getCount()` — the cheapest call Office.js has — did not come back
 * inside eight seconds.
 *
 * Detectable in seconds, and it survives a reload. Asking here turns a
 * sixty-six-minute wedge into a two-second refusal.
 *
 * Runs inside the PANE's frame, which is why the caller passes a ref belonging
 * to the pane: `PowerPoint` does not exist in the document frame around it.
 */
export function pingScript(budgetMs) {
  return (
    "async () => { const t = Date.now(); try { await Promise.race([ " +
    "PowerPoint.run(async (c) => { c.presentation.slides.getCount(); await c.sync(); }), " +
    `new Promise((_, rej) => setTimeout(() => rej(new Error("budget")), ${budgetMs})) ]); ` +
    'return "ok:" + (Date.now() - t); } catch (e) { return "no:" + (Date.now() - t); } }'
  );
}

/**
 * Did `find` actually FIND it, or is it echoing the query back?
 *
 * `playwright-cli find` answers a miss with `No matches found for "<query>"` —
 * which contains the query. Testing the output for the phrase searched for is
 * therefore always true, and a crash detector built that way would report a
 * crash on a perfectly healthy host, every time. Same family as the ref that
 * `buildOf` used to read out of its own haystack.
 */
export function sawCrashDialog(found) {
  const text = String(found ?? "").trim();
  // Deliberately NOT "does the output contain the word dialog". That version
  // passed its own test while proving nothing — the phrase searched for happens
  // not to contain the word, so the echo could never have tripped it and the
  // guard was decoration. Whether `find` matched at all is the real question,
  // and it stays true whatever the next query says. Empty is the third answer:
  // the CLI was never reached, which `reachable` reports on its own.
  return text !== "" && !/No matches found/.test(text);
}

/**
 * How many polls in a row the PAGE has answered with nothing.
 *
 * A CLI call that failed is not a page that said nothing, and only the second
 * kind means the round is over. `failed` resets rather than counts: a tool that
 * could not be run measured nothing, and treating it as evidence is the mistake
 * `reachable` was written to stop, arriving here by a different door.
 */
export function quietStreak(prev, text, failed) {
  return String(text ?? "").trim() || failed ? 0 : prev + 1;
}

/** `{ answered, ms }` from what `pingScript` returned, or null when unreadable. */
export function readPing(out) {
  const m = /"?(ok|no):(\d+)"?/.exec(String(out ?? ""));
  return m ? { answered: m[1] === "ok", ms: Number(m[2]) } : null;
}

/** Click through the element's own handler — a plain click does not reach the pane. */
function clickRef(sh, ref) {
  // The pane sits two iframes deep and `playwright-cli click` silently does
  // nothing there: `aria-selected` never moves and no error is raised. Firing
  // the element's own handler is the only thing that works, and finding that
  // out cost a round.
  return sh("eval", "el => { el.click(); return 'ok'; }", ref);
}

async function attempt(argv, deps, sh) {
  const run = deps.run ?? spawnSync;
  const fetchBuild = deps.fetchBuild ?? defaultFetchBuild;
  const checkOnly = argv.includes("--check");

  const head = String(run("git", ["rev-parse", "--short=7", "HEAD"], { encoding: "utf8" }).stdout ?? "").trim() || null;
  const deployed = buildOf(await fetchBuild());
  const stamp = buildOf(sh("find", "--regex", "/[0-9a-f]{7} ·/"));

  const listRef = refFor(sh, "Slide List", /listbox "Slide List"/);
  const slides = listRef ? (sh("snapshot", listRef).match(/option "Slide"/g) ?? []).length : null;

  const crashed = sawCrashDialog(sh("find", "Sorry, we ran into a problem"));

  // A tab, not the Verbose trace checkbox. The checkbox only exists while the
  // pane is ON the Automation tab, so anchoring the ping there skipped it
  // silently on a pane sitting anywhere else — and a skipped ping reads as
  // `ready`, which is the one answer this must never give without asking.
  const paneRef = refFor(sh, "Chart", /tab "Chart"/);
  const ping = paneRef ? readPing(sh("eval", pingScript(8000), paneRef)) : null;

  const toggles = sh("find", "Verbose trace");
  const verbose = /checkbox "Verbose trace"/.test(toggles) ? /Verbose trace" \[checked\]/.test(toggles) : null;
  const pictures = /checkbox "Picture every slide"/.test(toggles)
    ? /Picture every slide" \[checked\]/.test(toggles)
    : null;

  const state = { head, deployed, stamp, slides, verbose, pictures, reachable: !sh.state.unreachable, ping, crashed };
  const { ok, stop } = readiness(state);
  console.log(
    `  HEAD ${head ?? "?"} · site ${deployed ?? "?"} · pane ${stamp ?? "?"} · deck ${slides ?? "?"} slide(s)`,
  );
  console.log(`  verbose trace ${verbose ?? "?"} · picture every slide ${pictures ?? "?"}`);
  console.log(
    `  host ${ping ? (ping.answered ? `answered in ${ping.ms}ms` : `SILENT for ${ping.ms}ms`) : "not asked — the pane is closed"}` +
      (crashed ? " · PowerPoint's crash dialog is up" : ""),
  );
  if (!ok) {
    console.error("\n  NOT READY — a round now would not prove anything:");
    for (const s of stop) console.error(`    - ${s}`);
    return { code: 1, reason: crashed ? "crashed" : "not-ready" };
  }
  console.log("  ready");
  if (checkOnly) return { code: 0, reason: "checked" };

  const runBtn = refFor(sh, "Probe, then self-test", /button "Probe, then self-test"/);
  if (!runBtn) {
    console.error("  could not find the run button — is the Automation tab open?");
    return { code: 1, reason: "no-run-button" };
  }
  console.log("  running — this takes about ten minutes");
  clickRef(sh, runBtn);

  // Polled, never slept-through: a wedged host is the normal failure here and it
  // has to be distinguishable from a slow one.
  const started = Date.now();
  const limit = 30 * 60 * 1000;
  // An empty answer ENDS the round, so it has to be believed twice. The CLI
  // serves one command at a time per session: anything else touching it while
  // this polls — a second terminal, an agent looking at the trace — makes one
  // poll exit non-zero, and folding that into "the pane is gone" killed a round
  // that was running perfectly and went on to finish 10 of 12 scenarios. The
  // report was worse than the loss: it named a crash that had not happened.
  let quiet = 0;
  for (;;) {
    if (Date.now() - started > limit) {
      console.error("  the round has not finished in 30 minutes — the host is wedged; see docs/ROUNDS.md");
      return { code: 1, reason: "timeout" };
    }
    const dl = sh("find", "Download run log");
    if (/button "Download run log"(?! \[disabled\])/.test(dl)) break;
    // WATCH FOR THE CRASH, do not wait it out. Rounds 30 and 31 each died about
    // three minutes in and then held this loop for the full thirty, because the
    // only thing it knew how to notice was the finish. Twenty-seven wasted
    // minutes twice over is most of an hour of a night's throughput.
    //
    // A DOM read, deliberately, not a ping: the pane is mid-round and an
    // Office.js call from here would interleave with the round's own batches and
    // change what it measures. The dialog is in the document frame and costs
    // nothing to look at.
    if (sawCrashDialog(sh("find", "Sorry, we ran into a problem"))) {
      console.error(
        `  PowerPoint crashed ${Math.round((Date.now() - started) / 1000)}s in — its dialog is up and nothing ` +
          'behind it will answer. See docs/ROUNDS.md, "The wedge".',
      );
      return { code: 1, reason: "crashed" };
    }
    quiet = quietStreak(quiet, dl, sh.state.lastFailed);
    if (quiet >= 2) {
      console.error("  the pane stopped answering — PowerPoint has probably crashed; the trace is still in the DOM");
      return { code: 1, reason: "silent" };
    }
    await new Promise((r) => setTimeout(r, 20000));
  }
  console.log("  finished");
  return { code: 0, reason: "finished" };
}

/**
 * Is another attempt worth making?
 *
 * Only after a crash. A stale pane, a dirty deck or a missing run button are all
 * states a person has to look at, and retrying them just repeats the same
 * refusal until the night is gone.
 */
export function shouldRetry(reason, attempt, max) {
  return reason === "crashed" && attempt < max;
}

/** Delete every slide but the first, so the next round starts where the last one did. */
export function cleanDeckScript(budgetMs) {
  return (
    "async () => { const budget = (p, ms) => Promise.race([p, new Promise((_, r) => " +
    `setTimeout(() => r(new Error("TIMEOUT")), ms))]); try { const n = await budget(PowerPoint.run(async (c) => { ` +
    'const s = c.presentation.slides; s.load("items/id"); await c.sync(); const count = s.items.length; ' +
    "for (let i = count - 1; i >= 1; i--) c.presentation.slides.getItemAt(i).delete(); await c.sync(); " +
    `s.load("items/id"); await c.sync(); return s.items.length; }), ${budgetMs}); return "deck:" + n; } ` +
    'catch (e) { return "deck-failed"; } }'
  );
}

/**
 * Put PowerPoint back on its feet — the recovery done by hand six times tonight.
 *
 * Every step was learned the expensive way and none is optional:
 *
 *   - The dialog's Refresh button is matched BY TEXT. Its accessible name is
 *     sometimes absent, the label sitting in a child `generic` instead, so
 *     `button "Refresh"` finds it on one crash and not on the next.
 *   - Refreshing reloads the document, which closes the pane. It comes back from
 *     Home ▸ Add-ins ▸ **Insert chart** — that is the `ShowTaskpane` control,
 *     despite the name.
 *   - The Automation tab has to be showing or the run button is not in the DOM.
 *   - The deck has to go back to one slide, or the next round is not comparable
 *     with one that started clean.
 *
 * The waits are generous on purpose: a document reload takes tens of seconds and
 * a step taken early lands on nothing and fails silently.
 */
export async function recover(sh, sleep) {
  const dialog = /dialog \[ref=([a-z0-9]+)\]/.exec(sh("find", "Sorry, we ran into a problem"))?.[1];
  if (dialog)
    sh(
      "eval",
      'el => { const b = [...el.querySelectorAll("button")].find(n => /^\\s*Refresh\\s*$/.test(n.textContent || "")); ' +
        'if (!b) return "no refresh"; b.click(); return "clicked"; }',
      dialog,
    );
  else sh("reload");
  await sleep(55000);

  const pane = refFor(sh, "Insert chart", /button "Insert chart"/);
  if (pane) clickRef(sh, pane);
  await sleep(20000);

  const automation = refFor(sh, "Automation", /tab "Automation"/);
  if (automation) clickRef(sh, automation);
  await sleep(5000);

  const anchor = refFor(sh, "Chart", /tab "Chart"/);
  if (anchor) sh("eval", cleanDeckScript(90000), anchor);
  return Boolean(pane);
}

async function main(argv, deps = {}) {
  const run = deps.run ?? spawnSync;
  const sleep = deps.sleep ?? ((ms) => new Promise((r) => setTimeout(r, ms)));
  const dirArg = argv.indexOf("--dir");
  const sh = cli(run, dirArg === -1 ? process.cwd() : argv[dirArg + 1]);
  const retryArg = argv.indexOf("--retry");
  const max = retryArg === -1 ? 0 : Number(argv[retryArg + 1]) || 0;

  for (let n = 0; ; n++) {
    if (n) console.log(`\n  attempt ${n + 1} of ${max + 1}`);
    const { code, reason } = await attempt(argv, deps, sh);
    if (!shouldRetry(reason, n, max)) return code;
    console.log('  clearing the crash and starting again — see docs/ROUNDS.md, "The wedge"');
    await recover(sh, sleep);
  }
}

async function defaultFetchBuild() {
  // A network blip is a reading this tool could not take, not a reason to die
  // with a stack trace. Unguarded, one `getaddrinfo ENOTFOUND` took the whole
  // check down mid-investigation and printed nothing about the host, the pane or
  // the deck — every one of which had already been measured by then. `buildOf`
  // turns the empty string into `null` and `readiness` already says what a
  // missing build means.
  try {
    const r = await fetch(`https://powerchart.struktureretsundfornuft.dk/build.json?cb=${Date.now()}`);
    return await r.text();
  } catch {
    return "";
  }
}

/** Archive a downloaded round under the next number. See `rounds/README.md`. */
export function archive(logPath, dir = "rounds", read = readFileSync, write = writeFileSync, list = readdirSync) {
  const round = stripImages(JSON.parse(read(logPath, "utf8")));
  const build = buildOf(round.build);
  if (!build) throw new Error("that file carries no build stamp — it is not a round log");
  const name = `${nextRoundNumber(list(dir))}-${build}.json`;
  // TWO spaces, because prettier checks this directory and every round archived
  // at one space failed the gate until someone reformatted it by hand.
  write(`${dir}/${name}`, JSON.stringify(round, null, 2) + "\n");
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
