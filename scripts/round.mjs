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
 *   node scripts/round.mjs             # check, run, poll, archive — then triage BY HAND
 *
 * It does NOT triage. This line said it did until 2026-08-19, and a session
 * planned around it: reading the archived round is `node scripts/triage.mjs
 * rounds/0NN-<build>.json` and `npm run rounds`, both separate commands. A
 * usage block that claims a step nobody runs is the same defect class as the
 * three stale slogans in `triage.mjs`.
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
import { collectCrashEvidence } from "./crash-forensics.mjs";

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
  unreachableAt = null,
  browserGone = false,
  ping = null,
  slideOk = null,
  crashed = false,
  loggedOut = false,
  authPopup = false,
  size = null,
  expectSize = null,
  wantDeck = null,
  deckFronted = true,
  canOpenPane = true,
  // Whether the ribbon carries the command AT ALL, apart from whether it can be
  // clicked. Defaults true so an absent argument cannot invent a missing add-in.
  commandPresent = true,
}) {
  const stop = [];
  /**
   * A CODE beside every message, because retrying has to be decided on what a
   * stop IS rather than on how it is worded.
   *
   * `shouldRetry` used to retry a crash and nothing else, so the driver stopped
   * dead on three states it already knows how to fix — and on 2026-08-15 a
   * person fixed exactly those three by hand, in the order `recover` does them:
   * the host was silent (the quiet wedge), the pane was a build behind, the deck
   * held eight slides. One reload and reopen cleared all three.
   *
   * Matching on the messages instead would tie the retry loop to prose that is
   * edited whenever a message is improved. See `RECOVERABLE_STOPS`.
   */
  const codes = [];
  const refuse = (code, message) => {
    codes.push(code);
    stop.push(message);
  };
  // First, and on its own: everything below reads as "the browser said nothing"
  // when the truth is that nobody asked it. See `cli`.
  if (!reachable)
    return {
      ok: false,
      stop: [
        "playwright-cli could not be run — nothing below was actually measured. " +
          "Install it (`npm i -g @playwright/cli`), or set PLAYWRIGHT_CLI_JS to its entry point." +
          // The call and the errno, when there is one. Without them this message
          // points at the install every time, and the install is almost never it.
          (unreachableAt
            ? `\n      the call that could not be spawned: \`${unreachableAt.args}\` — ${unreachableAt.error}`
            : ""),
      ],
    };
  // Before the sign-in check and everything under it: a browser that is not
  // there answers every read with nothing, and "is the add-in open?" is the
  // wrong question to send anyone to. See `noBrowser`.
  if (browserGone)
    return {
      ok: false,
      codes: ["browser-gone"],
      stop: [
        "there is no browser — the process died, taking the tab with it. The persistent profile still " +
          "holds the sign-in, so this is recoverable without a password: " +
          "`pw open --persistent --profile=C:/devtools/pw-profile --headed https://onedrive.live.com/`, " +
          "then open the deck, select its tab, and reopen the pane from Home ▸ Add-ins ▸ Insert chart.",
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
  // Before everything the browser could tell us, because if this is true then
  // nothing else was measurable and every message below is noise.
  //
  // UNLESS THE HOST IS ANSWERING, and that exception is the whole point. A
  // sign-in TAB is not a sign-in PROMPT. Chrome's "Restore pages?" brings back
  // every tab a crashed session had open, including an `oauth20_authorize`
  // popup that was already finished — so the tab list shows a login page beside
  // a deck that is working perfectly.
  //
  // Observed 2026-08-20: the check printed `host answered in 7ms · slide 1
  // resolved` and refused in the same breath, telling the owner to go and enter
  // a password they had already entered. The deck was fine; the tab was debris.
  //
  // A host that answers Office.js and resolves a slide IS an authenticated
  // session — an unauthenticated one cannot do either. So the ping is the
  // discriminator, and it is better evidence than the tab list because it asks
  // the thing we actually care about instead of looking at the furniture around
  // it. If the host is silent, the refusal stands and reads exactly as before.
  const hostAnswering = Boolean(ping?.answered) && slideOk !== false;
  if (loggedOut && hostAnswering)
    console.log(
      "  a Microsoft sign-in tab is open, but the host is answering — treating it as leftover from a " +
        "browser restore rather than a live prompt",
    );
  // AND WHEN THE HOST IS SILENT BUT THE DOCUMENT IS UP, SAY SO INSTEAD OF
  // GUESSING. `slides !== null` means the deck rendered its slide list, which a
  // signed-out browser cannot do — but it does NOT prove the session is fine,
  // because Office can raise a re-auth prompt beside a loaded deck, which is the
  // case this refusal was written for.
  //
  // So the two are genuinely indistinguishable from the tab list, and claiming
  // "sign in" is a guess that costs the owner a trip to the machine. A stale
  // `login.live.com` tab restored by Chrome after a crash is PERMANENT debris:
  // it never goes away on its own, so this refusal would fire on every future
  // round the moment the host went briefly quiet — which on this machine is the
  // commonest transient state there is.
  //
  // A silent host is already a recoverable stop with a reload behind it. Let
  // that run, and mention the tab rather than blaming it. If the host is still
  // silent after recovery, the round fails as `host-silent`, which is what was
  // actually observed.
  if (loggedOut && !hostAnswering && slides !== null)
    console.log(
      "  a Microsoft sign-in tab is open and the host is silent, but the deck's slide list still reads — " +
        "treating this as a silent host (recoverable) rather than a sign-in prompt. If a reload does not " +
        "fix it, the tab may be a real prompt after all.",
    );
  if (loggedOut && !hostAnswering && slides === null)
    return {
      ok: false,
      stop: [
        authPopup
          ? "Office has opened a sign-in prompt beside the deck — the deck tab is still there, but the host is " +
            "asking for credentials and nothing measured past this point can be trusted. " +
            "Sign in on the prompt, check the add-in pane is still open, then check again. " +
            "None of that is the agent's to do: it needs a password."
          : "the browser is on a Microsoft sign-in page — there is no document to run a round in. " +
            "Sign in, open the deck, sideload the add-in from Home ▸ Add-ins, then check again. " +
            "None of that is the agent's to do: it needs a password.",
      ],
    };
  if (crashed)
    refuse(
      "crashed",
      'PowerPoint has crashed — its own "Sorry, we ran into a problem" dialog is up, and every Office.js ' +
        "call behind it hangs forever. Click Refresh in that dialog; the host answers again within the minute. " +
        "The pane closes with it — reopen from Home ▸ Add-ins ▸ Insert chart.",
    );
  // Before anything about builds or decks: is the host answering? A stale pane
  // and a dirty deck are both worth fixing, and neither matters if the host will
  // not talk — that is the state four rounds have burned an hour each on.
  if (ping && !ping.answered)
    refuse(
      "host-silent",
      `the host did not answer the cheapest possible call in ${ping.ms}ms — it is not going to answer a round. ` +
        "Its editing session is gone, usually because the network moved (look for net::ERR_NETWORK_CHANGED in " +
        "`playwright-cli console`). Reload the PowerPoint tab and reopen the pane: measured 8011ms silent before, " +
        "7ms after.",
    );
  // AFTER the ping, because a host that answered `getCount` and then refused a
  // slide is a different state from one that answered nothing, and the fix is
  // the same only by coincidence. See `slideResolveScript`: this is the call the
  // 2s crash dies on, moved to where it costs two seconds instead of a round.
  if (slideOk === false)
    refuse(
      "slide-refused",
      "the host answered the cheap call but would not resolve slide 1 — this is the state the 2s crash " +
        "starts from (`OnServerFindSucceeded could not find target slide` in its own log). Reload the " +
        "PowerPoint tab and reopen the pane; an attempt that follows a recovery has never crashed.",
    );
  if (!deployed) refuse("no-build", "the site did not answer with a build — is Pages up?");
  else if (head && deployed !== head)
    refuse("site-behind", `the site is serving ${deployed} but HEAD is ${head} — wait for Deploy Pages to finish`);
  // `slides !== null` is the proof the DOCUMENT is up. Without it this fires on
  // a tab that is merely mid-reload — every `refFor` answers nothing there too —
  // and it is a refusal recovery is forbidden to retry, so a transient state
  // would end the night. Unknown is not the same as wrong; the 4:3 deck that
  // motivated this read its slide list perfectly well and simply had no add-in.
  if (!stamp && !canOpenPane && slides !== null && commandPresent)
    // PRESENT BUT UNUSABLE, which is the opposite conclusion from the one below
    // and was reaching it. A `Disconnected` document greys its whole ribbon —
    // `Insert chart`, `Add-ins`, everything — and that is transient: the tab
    // reconnects, or a reload fixes it. Recoverable on purpose.
    refuse(
      "host-disconnected",
      "the PowerChart command is in the ribbon but DISABLED — the document is not connected, " +
        "so nothing in the ribbon can be clicked. This clears on its own or with a reload; " +
        "it is not a missing add-in.",
    );
  else if (!stamp && !canOpenPane && slides !== null)
    // A DIFFERENT REFUSAL, because recovery cannot touch it. `recover` reopens
    // the pane from the ribbon's `Insert chart` control; a document that does
    // not offer that control has no add-in to open, and reloading it forever
    // will not produce one. Deliberately absent from `RECOVERABLE_STOPS`, so
    // this stops on the first attempt instead of the seventh.
    //
    // GUARDED ON THE NAME BEING ABSENT NOW, not merely on there being no ref.
    // Those are different states and this fired on both, turning a disconnected
    // document into a stop only the owner could clear.
    refuse(
      "addin-missing",
      "this document has no PowerChart command in its ribbon — the add-in is not loaded here, " +
        "and a reload will not bring it back. Sideload it into this deck (Add-ins ▸ Upload My Add-in), " +
        "or run against a deck that already has it.",
    );
  else if (!stamp) refuse("pane-closed", "could not read the pane's build stamp — is the add-in open?");
  else if (deployed && stamp !== deployed)
    refuse(
      "pane-stale",
      `the pane is showing ${stamp} while the site serves ${deployed} — hard-reload the whole ` +
        `PowerPoint tab (the pane HTML is cached for ten minutes; reopening the pane alone does not clear it)`,
    );
  if (slides !== null && slides > 1)
    refuse(
      "deck-dirty",
      `the deck holds ${slides} slides — clean it, or this round is not comparable with one that started clean`,
    );
  // BEFORE the size check, because it explains a wrong size rather than
  // repeating it. A cycle names the deck each leg should run against, and until
  // the driver actually fronted that tab the third leg measured whichever
  // document leg two left open — refusing with `wrong-size`, which reads as
  // "the owner set the slide size wrong" when nothing had asked for the right
  // deck at all. Naming the deck is a far better message than naming its size.
  if (wantDeck && !deckFronted)
    refuse(
      "deck-missing",
      `no open tab is the deck \`${wantDeck}\` — this round was told to run against it and cannot. ` +
        "Open it, or unset PW_DECK to run against whichever document is in front.",
    );
  // THE SIZE THE DECK ACTUALLY IS, when a profile was asked for. Only when both
  // are known: a host that would not answer has said nothing, and refusing on no
  // evidence is what `reachable` exists to prevent.
  if (expectSize && size && size !== expectSize)
    refuse(
      "wrong-size",
      `the deck is ${size} and this round was asked for ${expectSize} — set it in Design ▸ Slide Size and ` +
        "CHECK IT TOOK, because a click made while the document is loading is accepted and does nothing. " +
        "Filing a round under the wrong profile is worse than not running it.",
    );
  if (verbose === false) refuse("verbose-off", "Verbose trace is off — the round's trace will be too thin to mine");
  if (pictures === false)
    refuse("pictures-off", "Picture every slide is off — a slide that reads back empty cannot be confirmed empty");
  return { ok: stop.length === 0, stop, codes };
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

/**
 * Did this call RUN and say too much, rather than fail to run at all?
 *
 * `spawnSync` reports both as `error`, and the driver treated both as "the tool
 * could not be run". They are opposite facts: an overflow means the browser
 * answered, and every other spawn error means nobody asked. Round 044 spent two
 * attempts and two empty crash reports on the difference.
 *
 * Matched on the CODE, with the message as a fallback, because Node stamps
 * `err.code = "ENOBUFS"` on the overflow path while the message it prints names
 * the executable rather than the reason.
 */
export function isOverflow(err) {
  if (!err) return false;
  return err.code === "ENOBUFS" || /ENOBUFS/.test(String(err.message ?? err));
}

/** Which manifest a re-sideload uploads. The PROD one — see `sideloadAddIn`. */
export const MANIFEST_PATH = process.env.PW_MANIFEST ?? "C:/devtools/PowerChart/manifest-prod.xml";

/**
 * Has this process already tried to put the add-in back? One attempt, ever.
 *
 * A latch rather than a retry budget, because the failure it guards is not
 * transient: if the ribbon walk did not work the first time, it is because
 * Microsoft moved something, and walking it six more times finds the same
 * missing control six more times. `addin-missing` stays out of
 * `RECOVERABLE_STOPS`; this runs BEFORE readiness concludes, so a success is
 * simply a round that starts and a failure is the refusal the driver already
 * had. Per process, and the cycle runner spawns one per leg.
 */
let sideloadAttempted = false;

/** Test-only: let a suite arm the latch more than once. */
export function _resetSideloadLatchForTest() {
  sideloadAttempted = false;
}

/**
 * Put the add-in back after a browser death took the sideload with it.
 *
 * A web sideload does not survive the browser process. `recover` restores the
 * window, the deck and its tab — and then finds no PowerChart command in the
 * ribbon, because there is no add-in to open. On 2026-08-16 that ended a night
 * twice, about fifteen minutes each time.
 *
 * THE WALK IS THE ONE VERIFIED BY HAND on 2026-08-17, and the first step is the
 * one nobody would guess: the ribbon's `Add-ins` button ignores clicks until the
 * document surface has focus. Three attempts failed on that alone, all of them
 * reporting `aria-expanded=true` over a menu that was plainly shut in a
 * screenshot. The accessibility tree lied; the pixels did not.
 *
 * `manifest-prod.xml`, never `manifest.xml`: the latter points at
 * `https://localhost`, so it would sideload a pane that readiness then refuses
 * for disagreeing with the deployed site — a worse outcome than no pane, because
 * it looks like it worked.
 *
 * NO NATIVE FILE DIALOG EVER OPENS. Playwright intercepts the chooser, so
 * `upload` hands the path over directly. That is the only reason this is
 * automatable at all.
 *
 * Best-effort and self-cleaning. Every failure path dismisses whatever is open,
 * because a modal left over the document makes every later round refuse for a
 * reason that has nothing to do with the round.
 */
export async function sideloadAddIn(sh, sleep, manifest = MANIFEST_PATH) {
  const click = (ref) => sh("eval", "el => { el.click(); return 'ok'; }", ref);
  const step = (query, pattern) => refFor(sh, query, pattern);
  const giveUp = (why) => {
    console.error(`  could not put the add-in back: ${why}`);
    // ALWAYS, even when nothing looks open. A half-walked dialog is invisible
    // to `readiness` and fatal to every round after it.
    const cancel = step("Cancel", /button "Cancel"/);
    if (cancel) click(cancel);
    const close = step("Close", /button "Close"/);
    if (close) click(close);
    return false;
  };
  console.log("  the add-in is gone from this document — putting it back");
  // FIRST, and it is not optional: the ribbon will not open its menus while the
  // document surface is unfocused.
  const list = step("Slide List", /listbox "Slide List"/);
  if (!list) return giveUp("the deck's slide list is not readable — the document is not up");
  click(list);
  await sleep(2000);

  const addins = step("Add-ins", /button "Add-ins"/);
  if (!addins) return giveUp("no Add-ins button in the ribbon");
  click(addins);
  await sleep(5000);

  const seeAll = step("See all", /menuitem "See all installed add-ins"/);
  if (!seeAll) return giveUp("the Add-ins menu did not open");
  click(seeAll);
  await sleep(6000);

  const more = step("More Add-ins", /menuitem "More Add-ins"/);
  if (!more) return giveUp("no `More Add-ins` entry");
  click(more);
  await sleep(9000);

  const mine = step("MY ADD-INS", /tab "MY ADD-INS"/);
  if (!mine) return giveUp("the Office Add-ins dialog did not open");
  click(mine);
  await sleep(5000);

  const manage = step("Manage My Add-ins", /button "Manage My Add-ins"/);
  if (!manage) return giveUp("no `Manage My Add-ins` control");
  click(manage);
  await sleep(5000);

  const upload = step("Upload My Add-in", /menuitem "Upload My Add-in"/);
  if (!upload) return giveUp("no `Upload My Add-in` entry");
  click(upload);
  await sleep(6000);

  const browse = step("Browse", /button "Browse\.\.\."/);
  if (!browse) return giveUp("the upload dialog has no Browse button");
  click(browse);
  await sleep(2000);
  sh("upload", manifest);
  await sleep(3000);

  // THE BUTTON'S OWN STATE IS THE RECEIPT. It is disabled until a file is
  // accepted, so finding it enabled is the host confirming it took the
  // manifest — better evidence than the upload call not erroring.
  const go = refFor(sh, "Upload", /button "Upload"(?! \[disabled\])/);
  if (!go) return giveUp("the manifest was not accepted — Upload stayed disabled");
  click(go);
  await sleep(12000);

  const open = step("Insert chart", /button "Insert chart"/);
  if (!open) return giveUp("uploaded, but no PowerChart command appeared");
  click(open);
  await sleep(25000);
  return true;
}

/**
 * Bring the named deck's tab to the front, and say whether it is there at all.
 *
 * `PW_DECK` used to reach only `recover`, in the branch that reopens a browser
 * that died — so a cycle setting it per leg was choosing which deck a RECOVERY
 * would look for and nothing else. The ordinary path never selected a tab; it
 * ran against whatever happened to be fronted. The nightly cycle's third leg
 * therefore asked for a 4:3 deck, got the 16:9 one still open from leg two, and
 * refused with `wrong-size` every time — a stop no recovery addresses, so the
 * night ended there, reading as an operator's mistake when the runner had simply
 * never asked for the deck.
 *
 * Returns `true` when that deck is now fronted, `false` when no tab carries the
 * name. Never throws: a `tab-list` that could not be read is a reading not
 * taken, and `readiness` has its own, better-worded refusals for a browser that
 * is not answering.
 */
export function selectDeck(sh, deckName) {
  if (!deckName) return true;
  const line = sh("tab-list")
    .split("\n")
    .find((l) => l.includes(deckName));
  if (!line) return false;
  const n = /(\d+):/.exec(line)?.[1];
  if (!n) return false;
  sh("tab-select", n);
  return true;
}

/**
 * How long a single CLI call may take before it counts as a wedge.
 *
 * Env-overridable because the one thing that would make this wrong is a host
 * slower than any yet seen, and nobody debugging that at 2am should have to
 * edit a script to get past it.
 */
export const CLI_TIMEOUT_MS = Number(process.env.PW_CLI_TIMEOUT_MS) || 180_000;

export function cli(run, dir, entry = cliEntry()) {
  const state = { unreachable: false };
  const cwd = sessionDir(dir);
  const sh = (...args) => {
    if (!entry) {
      state.unreachable = true;
      return "";
    }
    // NODE, on the CLI's own JavaScript. No shim, no shell.
    //
    // `maxBuffer` because the default is 1 MiB and `requests` on a live
    // PowerPoint tab is bigger than that — the document channel alone runs to
    // hundreds of POSTs with query strings on them. Over the line, `spawnSync`
    // returns ENOBUFS and throws the output away, which is how round 044 lost
    // two crash reports and then the round itself.
    // BOUNDED, because the round's own deadline cannot bound this. The poll
    // loop checks a 30-minute limit at the TOP of each pass, which only ever
    // runs if the call below returned — so a CLI that wedges (a browser that
    // stopped answering CDP, a tab mid-crash) hangs the driver with the
    // deadline sitting there unreachable and nothing on screen. That is the
    // exact shape of an overnight run that is found dead in the morning having
    // printed nothing since hour one.
    //
    // Generous on purpose: the slowest legitimate call is an `eval` carrying a
    // 20s page-side budget, and `requests` on a live tab can return tens of
    // megabytes. Three minutes is far above both and far below a night. A
    // timeout arrives as `r.error`, which the existing branch below already
    // reads as "nothing was measured" — which is exactly what it is.
    const r = run(process.execPath, [entry, "-s=ms", "--raw", ...args], {
      encoding: "utf8",
      cwd,
      maxBuffer: 64e6,
      timeout: CLI_TIMEOUT_MS,
    });
    // A CALL THAT RAN AND SAID TOO MUCH IS NOT A TOOL THAT COULD NOT BE RUN, and
    // conflating them is what sent two rounds' debugging at a healthy install.
    // ENOBUFS means the browser answered and the answer did not fit; every other
    // spawn error means nothing was measured. Only the second is `unreachable`.
    // Per CALL, beside the two latches, so a reader one layer out can tell
    // whether the empty string it just got was an answer or a failure. The crash
    // report is that reader, and without this it wrote "(nothing)" over a read
    // that never happened — twice, on the only two crashes of round 044.
    state.lastError = isOverflow(r.error) ? "overflow" : r.error ? "spawn" : null;
    if (isOverflow(r.error)) {
      state.overflowed = args[0];
    } else if (r.error) {
      state.unreachable = true;
      // WHICH call, and what the OS said. "playwright-cli could not be run" sent
      // two rounds' worth of debugging at an install that was fine, because the
      // message named the tool and the tool was never the problem. A spawn
      // failure has a subcommand and an errno and both were being thrown away.
      state.unreachableAt ??= { args: args.join(" ").slice(0, 80), error: String(r.error?.message ?? r.error) };
    }
    // A failed CLI call and a page that answered with nothing are the same empty
    // string, and the difference decides whether a round is alive. Recorded, not
    // folded in — the same distinction `unreachable` exists to keep.
    state.lastFailed = !r.error && r.status !== 0;
    return r.status === 0 ? String(r.stdout ?? "") : "";
  };
  sh.state = state;
  // The session directory the CLI is rooted at, so callers can find what it
  // downloads without re-deriving it — `collectRound` needs the run log, and a
  // second `sessionDir()` call beside this one is a second place to get the 8.3
  // short-name normalisation wrong.
  sh.dir = cwd;
  /**
   * Start a fresh sweep, forgetting a spawn failure the last one saw.
   *
   * `unreachable` is deliberately sticky WITHIN a sweep: a call that never ran
   * and a page that answered with nothing are the same empty string, so once one
   * spawn has failed nothing that sweep read can be trusted. Across sweeps it is
   * a lie, and round 044 is what that costs.
   *
   * `--retry` survived the crash, `recover` reloaded the tab and reopened the
   * pane, and one of its ~8 spawns lost a race. The next attempt then measured a
   * perfectly healthy setup — `host answered in 4ms`, printed one line above —
   * and refused it with "playwright-cli could not be run — nothing below was
   * actually measured", which was false of every value on screen. A latch that
   * outlives its evidence turns a recoverable round into a stop, and this one
   * exited 0 while doing it.
   *
   * The same mistake as the poll that once ended a round on a single failed CLI
   * call (round 29), one call site further out. Fixed the same way: scope the
   * doubt to the thing it was actually about.
   */
  sh.startSweep = () => {
    state.unreachable = false;
    state.unreachableAt = undefined;
  };
  return sh;
}

/** A `ref_N` for the first element matching `pattern` in a `find` result. */
/**
 * Is the control THERE, whatever state it is in?
 *
 * `refFor` answers null for two situations that are not the same thing: no such
 * control, and a control that is present but DISABLED — a greyed-out button can
 * match the line and carry no ref.
 *
 * CAN, not does. "Playwright only issues a ref for something it could act on"
 * was the first version of this sentence and it is too strong: the add-in pane's
 * own `button "Use deck style" [disabled] [ref=f19e328]` has both, while the
 * PowerPoint ribbon's `button "Insert chart" [disabled]` has neither. Native
 * `disabled` and whatever Office marks its ribbon with are not treated alike.
 *
 * Which is the argument for asking the question directly rather than inferring
 * it: the rule to rely on is never "no ref means gone", in either direction.
 *
 * On 2026-08-19 that cost a round. PowerPoint's document went `Disconnected`
 * (the network again, the same fault that crashed round 089), which greys the
 * WHOLE ribbon. `canOpenPane` looked for `button "Insert chart"`, got no ref,
 * and concluded the add-in was gone — an un-retryable, owner-only refusal — for
 * a state a reload clears. The accessibility tree said `button "Insert chart"
 * [disabled]` the entire time, with the add-in loaded and fine.
 *
 * So: presence and usability are two questions, and the driver has to ask both.
 */
export function namePresent(sh, query, pattern) {
  return sh("find", query)
    .split("\n")
    .some((l) => pattern.test(l));
}

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
 * Touch a SLIDE before the round does, because the ping does not.
 *
 * THE EXPERIMENT, stated so a later reader can tell whether it worked.
 * PowerPoint crashed 2s into the FIRST attempt of rounds 043 and 044 — four for
 * four — and never into an attempt that followed a recovery, the difference
 * being about eighty seconds of reload and settling. Its own log says what it
 * was doing:
 *
 *     In OnDisconnect(), setting SlideViewNode.srcSlide to null
 *     Failed to restore selection after load content.
 *     OnServerFindSucceeded could not find target slide, time elapsed: 430 ms
 *     GlobalErrorHandler:DisplayErrorDialog: 5341289
 *
 * A document still settling when the round's first Office.js call lands. The
 * ping cannot see it: `slides.getCount()` is a COUNT, and this host answers it
 * in single-digit milliseconds while it is in exactly that state. Resolving a
 * slide is the cheapest call that goes down the path the host died on.
 *
 * The point is NOT to prevent the crash — it is to move it. If the theory holds,
 * this trips it here, where the check is two seconds and `--retry` recovers
 * before any round has been spent; today it costs a whole attempt plus the
 * recovery. If the theory is wrong this answers `ok` and the round crashes
 * anyway, which refutes it for the price of one extra call.
 *
 * `getItemAt(0)` and not the selection: a round starts on a one-slide deck, and
 * the selection API is itself one of the calls this host has wedged on
 * (`which selection call wedges the host`).
 */
export function slideResolveScript(budgetMs) {
  return (
    "async () => { try { return await Promise.race([ " +
    "PowerPoint.run(async (c) => { const s = c.presentation.slides.getItemAt(0); " +
    's.load("id"); await c.sync(); return "slide:" + (s.id || "?"); }), ' +
    `new Promise((_, rej) => setTimeout(() => rej(new Error("budget")), ${budgetMs})) ]); ` +
    '} catch (e) { return "slide-failed:" + (e && e.message ? String(e.message).slice(0, 80) : "?"); } }'
  );
}

/**
 * Ask the deck its slide size, so a round cannot lie about which one it ran.
 *
 * WHY THIS IS A PRECONDITION AND NOT A NICETY. Setting a deck to Widescreen
 * during the 2026-08-16 control run SILENTLY DID NOT TAKE — the click landed
 * while the document was in its greyed "Loading" state, the menu accepted it,
 * and nothing changed. It was caught only by reopening the menu and reading
 * which box was ticked.
 *
 * A round that believes it is 4:3 and is not proves exactly nothing, which is
 * the same harm as a round on a stale pane — and that is already a hard stop.
 * With a nightly cycle running 16:9 twice and 4:3 once, an unverified size means
 * filing a round under the wrong profile, which is worse than not running it.
 *
 * Points, via `pageSetup`, because that is what the add-in itself reads.
 */
/**
 * Set the deck's slide size, so a dedicated 4:3 deck can be made 4:3 unattended.
 *
 * `PowerPoint.PageSetup.slideWidth` and `slideHeight` are WRITABLE at
 * PowerPointApi 1.10, and round 096's `environment` line records this host
 * advertising 1.1 through 1.10. So the driver's `wrong-size` refusal — "set it
 * in Design ▸ Slide Size" — was asking a person to do something the API can do.
 *
 * OFF BY DEFAULT, AND THAT IS THE IMPORTANT HALF. Resizing the WRONG deck is a
 * quiet disaster: there is one 16:9 deck behind almost the whole archive,
 * `roundProfile` defaults to 16:9 for the 53 rounds that carry no size at all,
 * and `scenarioRegressions` compares within one profile — so a deck that
 * silently changed shape mid-series would split every comparison built on it.
 * A misaimed `PW_DECK` plus an automatic resize would do that to a real
 * presentation without anyone noticing.
 *
 * So it happens only when the owner asks for it with `PW_SET_SIZE=1`, and it
 * says so loudly when it does. The intended use is a deck that EXISTS to be 4:3,
 * where setting the size is idempotent and makes the deck what its name claims.
 */
export function setSlideSizeScript(size, budgetMs) {
  const [w, h] = size === "4:3" ? [720, 540] : [960, 540];
  return (
    "async () => { try { return await Promise.race([ " +
    "PowerPoint.run(async (c) => { const p = c.presentation.pageSetup; " +
    `p.slideWidth = ${w}; p.slideHeight = ${h}; await c.sync(); ` +
    'p.load("slideWidth,slideHeight"); await c.sync(); ' +
    'return "size:" + Math.round(p.slideWidth) + "x" + Math.round(p.slideHeight); }), ' +
    `new Promise((_, rej) => setTimeout(() => rej(new Error("budget")), ${budgetMs})) ]); ` +
    '} catch (e) { return "size-failed:" + (e && e.message ? String(e.message).slice(0, 60) : "?"); } }'
  );
}

export function slideSizeScript(budgetMs) {
  return (
    "async () => { try { return await Promise.race([ " +
    "PowerPoint.run(async (c) => { const p = c.presentation.pageSetup; " +
    'p.load("slideWidth,slideHeight"); await c.sync(); ' +
    'return "size:" + Math.round(p.slideWidth) + "x" + Math.round(p.slideHeight); }), ' +
    `new Promise((_, rej) => setTimeout(() => rej(new Error("budget")), ${budgetMs})) ]); ` +
    '} catch (e) { return "size-failed:" + (e && e.message ? String(e.message).slice(0, 60) : "?"); } }'
  );
}

/**
 * `"size:960x540"` → `"16:9"`; a failure or a silence → null.
 *
 * Null is NOT a mismatch. A host that would not answer has told us nothing, and
 * refusing a round on no evidence is the mistake `reachable` exists to prevent.
 */
export function readSlideSize(out) {
  const m = /size:(\d+)x(\d+)/.exec(String(out ?? ""));
  if (!m) return null;
  const w = Number(m[1]);
  const h = Number(m[2]);
  if (!w || !h) return null;
  const r = w / h;
  if (Math.abs(r - 16 / 9) < 0.01) return "16:9";
  if (Math.abs(r - 4 / 3) < 0.01) return "4:3";
  return `${w}x${h}`;
}

/** `"slide:287#62081387"` → true; a failure or a silence → false. */
export function readSlideResolve(out) {
  const s = String(out ?? "");
  if (/slide-failed/.test(s)) return false;
  return /slide:/.test(s) ? true : null;
}

/**
 * Is the browser sitting on a Microsoft sign-in page?
 *
 * Told apart from "the pane is closed" because the fix is completely different
 * and only one of them is the agent's to do. A signed-out browser answers every
 * pane read with nothing, which reads exactly like an add-in nobody opened —
 * and the check duly said "is the add-in open?" while the tab showed
 * `login.live.com`, sending the reader to look for a pane in a window that has
 * no document in it.
 *
 * Same distinction `reachable` exists for, a layer further out: there is a
 * difference between a question that was answered with nothing and a question
 * there was nobody to ask.
 */
export function signedOut(tabList) {
  return AUTH_TAB.test(String(tabList ?? ""));
}

const AUTH_TAB = /login\.live\.com|login\.microsoftonline\.com|login\.windows\.net/;

/**
 * Is that sign-in tab a POPUP beside a live deck, rather than the whole browser
 * sitting on a login page?
 *
 * `signedOut` is right to stop the round either way — if Office is asking for
 * credentials, nothing measured after it can be trusted — but the two look
 * completely different to the person who walks up to the screen, and until
 * 2026-08-16 they read the same sentence. What ended the overnight run of
 * 2026-08-15/16 was an auth popup that opened BESIDE a deck tab that was still
 * open, and the driver reported "the browser is on a Microsoft sign-in page".
 * The reader looked at a screen showing PowerPoint and a small dialog, and had
 * to work out for themselves that the driver was describing the dialog.
 *
 * A message that does not match what is on the screen is worse than a vague one:
 * it makes a reader doubt the whole report, which on an overnight run is the only
 * account of what happened.
 *
 * The signal is that some OTHER tab is still a document. Deliberately loose about
 * which — the driver's own recovery looks for `Presentation63` by name, and
 * hard-coding one deck's name into a diagnostic is how it would go quietly wrong
 * for the next deck.
 */
export function signInIsPopup(tabList) {
  const lines = String(tabList ?? "")
    .split("\n")
    .filter((l) => l.trim());
  return (
    lines.some((l) => AUTH_TAB.test(l)) &&
    lines.some((l) => !AUTH_TAB.test(l) && /officeapps\.live\.com|onedrive\.live\.com|sharepoint\.com|\.pptx/i.test(l))
  );
}

/**
 * Is there a BROWSER at all?
 *
 * The third thing that answers every pane read with nothing, after "the CLI
 * could not be run" and "the browser is signed out" — and the one that had no
 * name until it cost seven attempts on 2026-08-15. A round wedged, the browser
 * process died with it, and `recover` then reloaded and reopened a pane in a
 * window that did not exist, seven times, while the check reported "could not
 * read the pane's build stamp — is the add-in open?" The add-in was fine. There
 * was nothing to open it in.
 *
 * Worth its own precondition because the fix is specific and the loop can do it
 * unattended: reopen from the persistent profile, which still holds the
 * sign-in — **a dead browser is not a lost sign-in**, and believing otherwise
 * has now cost this project two separate stretches of hours.
 *
 * `pw list` answers `(no browsers)` in exactly this state, and that string is
 * what the daemon prints when it has no session for the working directory.
 */
/**
 * Where the browser that survives a session lives.
 *
 * `scripts/pw.sh` parks the SESSION in the repo and the PROFILE here, and the
 * profile is the half that holds the OneDrive sign-in across a browser death.
 */
export const PROFILE_DIR = process.env.PW_PROFILE_DIR ?? "C:/devtools/pw-profile";

/**
 * The deck a recovery reopens, when the browser has died and taken its tab.
 *
 * A DEFAULT rather than a constant, and overridable with `PW_DECK`, because the
 * deck's name changes and the old hard-coded `Presentation63` was already stale
 * on 2026-08-16 — the deck in use had become `Presentation64`. A sideload on
 * PowerPoint for the web is **per-document**, so a fresh deck is a fresh
 * sideload and a fresh name, and this will drift again.
 *
 * It fails in the worst possible way when wrong: `recover` reopens OneDrive,
 * finds no matching link, clicks nothing, and reports a closed pane — in exactly
 * the situation the function exists to rescue.
 */
export const DECK_NAME = process.env.PW_DECK ?? "Presentation64";

export function noBrowser(listOutput) {
  return /\(no browsers\)/i.test(String(listOutput ?? ""));
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

async function attempt(argv, deps, sh, healed = false) {
  const run = deps.run ?? spawnSync;
  const sleep = deps.sleep ?? ((ms) => new Promise((r) => setTimeout(r, ms)));
  const fetchBuild = deps.fetchBuild ?? defaultFetchBuild;
  const checkOnly = argv.includes("--check");
  // This sweep's reachability is about THIS sweep. See `sh.startSweep`.
  sh.startSweep?.();

  // BEFORE ANYTHING IS MEASURED, because every reading below is about whichever
  // tab is fronted. Only when a deck was actually asked for: with no `PW_DECK`
  // this is the behaviour it has always had, running against the open document.
  const wantDeck = process.env.PW_DECK || null;
  const deckFronted = wantDeck ? selectDeck(sh, wantDeck) : true;

  const head = String(run("git", ["rev-parse", "--short=7", "HEAD"], { encoding: "utf8" }).stdout ?? "").trim() || null;
  const deployed = buildOf(await fetchBuild());
  let stamp = buildOf(sh("find", "--regex", "/[0-9a-f]{7} ·/"));

  const listRef = refFor(sh, "Slide List", /listbox "Slide List"/);
  const slides = listRef ? (sh("snapshot", listRef).match(/option "Slide"/g) ?? []).length : null;

  const browserGone = noBrowser(sh("list"));
  // Read ONCE. Two `tab-list` calls could straddle the popup opening or closing
  // and produce a message describing neither state.
  const tabs = sh("tab-list");
  const loggedOut = signedOut(tabs);
  const authPopup = loggedOut && signInIsPopup(tabs);
  const crashed = sawCrashDialog(sh("find", "Sorry, we ran into a problem"));

  // A tab, not the Verbose trace checkbox. The checkbox only exists while the
  // pane is ON the Automation tab, so anchoring the ping there skipped it
  // silently on a pane sitting anywhere else — and a skipped ping reads as
  // `ready`, which is the one answer this must never give without asking.
  const paneRef = refFor(sh, "Chart", /tab "Chart"/);
  // CAN THE PANE BE OPENED AT ALL, if it is shut? `recover` reopens it from the
  // ribbon's `Insert chart` control, so a document that does not offer that
  // control is one recovery cannot help — and retrying it is pure waste.
  //
  // Measured on 2026-08-16: a 4:3 leg switched to a deck the add-in was not
  // registered for and spent SEVEN attempts, about fifteen minutes, rediscovering
  // that a pane it cannot open is not openable. Only asked when the pane is
  // actually shut, so a healthy round pays nothing for it.
  let canOpenPane = paneRef ? true : Boolean(refFor(sh, "Insert chart", /button "Insert chart"/));
  // THE SAME CONTROL, ASKED THE OTHER WAY. `canOpenPane` is "could I click it";
  // this is "is it there at all". A `Disconnected` document greys the whole
  // ribbon, so the first is false and the second is true — and reading only the
  // first turned a transient network state into `addin-missing`, which recovery
  // is forbidden to retry. Cost the round on 2026-08-19.
  const commandPresent = Boolean(paneRef) || canOpenPane || namePresent(sh, "Insert chart", /button "Insert chart"/);
  // PUT IT BACK, ONCE, rather than refuse a night over it. This is the state a
  // browser death leaves: the deck is up and readable and there is no add-in in
  // it, because a web sideload does not survive the process. `recover` restores
  // everything except that.
  //
  // Guarded on `slides` for the same reason the refusal below is: a tab that is
  // merely mid-reload answers nothing to every read and looks identical, and
  // walking ten ribbon steps against a loading document would leave a dialog
  // over it.
  let pane = paneRef;
  // NOT WHILE THE RIBBON IS GREYED OUT. The sideload walk needs to click the
  // Add-ins button, and on a disconnected document that button is disabled like
  // every other — so the walk fails at its first step and reports "no Add-ins
  // button in the ribbon", which reads as Microsoft having moved something. It
  // did exactly that on 2026-08-19, burning the one attempt the latch allows on
  // a document that simply had to reconnect.
  if (!pane && !canOpenPane && !commandPresent && slides !== null && !sideloadAttempted) {
    sideloadAttempted = true;
    if (await sideloadAddIn(sh, sleep)) {
      // RE-READ, because everything below was measured before the add-in
      // existed. The stamp especially: it is what says the pane is serving the
      // build under test, and a sideload that lands the wrong manifest must
      // still be caught by the ordinary check rather than assumed good.
      pane = refFor(sh, "Chart", /tab "Chart"/);
      canOpenPane = Boolean(pane) || Boolean(refFor(sh, "Insert chart", /button "Insert chart"/));
      stamp = buildOf(sh("find", "--regex", "/[0-9a-f]{7} ·/"));
      console.log(`  the add-in is back — pane ${stamp ?? "still not readable"}`);
    }
  }
  const ping = pane ? readPing(sh("eval", pingScript(8000), pane)) : null;
  // Only when the host is already answering: a slide resolve on a host that did
  // not survive `getCount` tells us nothing the ping has not, and costs 20s.
  const slideOk = pane && ping?.answered ? readSlideResolve(sh("eval", slideResolveScript(20000), pane)) : null;
  // Only when a profile was asked for — an unasked question costs a round trip
  // and answers nothing. `PW_EXPECT_SIZE=4:3` is how a nightly cycle says which
  // arm this round belongs to.
  const expectSize = process.env.PW_EXPECT_SIZE || null;
  let size = expectSize && pane && ping?.answered ? readSlideSize(sh("eval", slideSizeScript(15000), pane)) : null;
  // MAKE THE DECK WHAT THE LEG ASKED FOR, when the owner has opted in. See
  // `setSlideSizeScript` for why this is off by default: resizing the wrong deck
  // splits every comparison the archive rests on, and a misaimed `PW_DECK` would
  // do it to a real presentation silently.
  //
  // Only when the size is KNOWN and WRONG. A host that would not answer has said
  // nothing, and resizing on no evidence is the mistake `reachable` exists to
  // prevent — the same reason the refusal below needs both values.
  if (expectSize && size && size !== expectSize && process.env.PW_SET_SIZE && pane) {
    console.log(`  the deck is ${size} and this leg wants ${expectSize} — setting it (PW_SET_SIZE is on)`);
    const after = readSlideSize(sh("eval", setSlideSizeScript(expectSize, 20000), pane));
    console.log(`  slide size is now ${after ?? "unreadable"}`);
    // READ THE ANSWER BACK rather than assuming the write took. The driver's own
    // refusal text has warned about this for months — "CHECK IT TOOK, because a
    // click made while the document is loading is accepted and does nothing" —
    // and an API write deserves the same suspicion.
    size = after ?? size;
  }
  // RE-READ, after the slide touch rather than only before it. If the touch is
  // what trips the host, the dialog appears in the seconds that follow — and
  // reading the dialog only at the top of the sweep is how a crash the check
  // itself provoked would be carried into the round as `ready`.
  const crashedAfter = slideOk === false ? sawCrashDialog(sh("find", "Sorry, we ran into a problem")) : false;

  // THE AUTOMATION TAB HAS TO BE SHOWING BEFORE ANY OF THIS IS READ. `Verbose
  // trace`, `Picture every slide` and the run button all live on it, and a pane
  // does not open there — a freshly sideloaded or reopened one comes up on
  // `Chart`, where none of the three is in the DOM.
  //
  // 2026-08-20 is what this is for: the driver put the add-in back by itself for
  // the first time, reached `ready` with `verbose trace ?` in the same block,
  // and then stopped because the run button was not there either. Both readings
  // had the same cause and neither was acted on.
  //
  // Selected HERE rather than just before the click, so the two toggles are read
  // rather than guessed. A round that starts with verbose trace off produces a
  // trace too thin to mine, and `?` is not evidence that it is on.
  if (pane) {
    const autoTab = refFor(sh, "Automation", /tab "Automation"/);
    const already = /tab "Automation" \[selected\]/.test(sh("find", "Automation"));
    if (autoTab && !already) {
      console.log("  the pane is not on the Automation tab — selecting it");
      clickRef(sh, autoTab);
    }
  }
  const toggles = sh("find", "Verbose trace");
  const verbose = /checkbox "Verbose trace"/.test(toggles) ? /Verbose trace" \[checked\]/.test(toggles) : null;
  const pictures = /checkbox "Picture every slide"/.test(toggles)
    ? /Picture every slide" \[checked\]/.test(toggles)
    : null;

  const state = {
    head,
    deployed,
    stamp,
    slides,
    verbose,
    pictures,
    reachable: !sh.state.unreachable,
    unreachableAt: sh.state.unreachableAt ?? null,
    browserGone,
    ping,
    slideOk,
    crashed: crashed || crashedAfter,
    loggedOut,
    authPopup,
    size,
    expectSize,
    wantDeck,
    deckFronted,
    canOpenPane,
    commandPresent,
  };
  const { ok, stop, codes } = readiness(state);
  console.log(
    `  HEAD ${head ?? "?"} · site ${deployed ?? "?"} · pane ${stamp ?? "?"} · deck ${slides ?? "?"} slide(s)`,
  );
  console.log(
    `  verbose trace ${verbose ?? "?"} · picture every slide ${pictures ?? "?"}` +
      // Printed only when asked for, so an ordinary 16:9 round reads exactly as
      // it always has.
      (expectSize ? ` · slide size ${size ?? "?"} (want ${expectSize})` : ""),
  );
  console.log(
    `  host ${ping ? (ping.answered ? `answered in ${ping.ms}ms` : `SILENT for ${ping.ms}ms`) : "not asked — the pane is closed"}` +
      // Printed every round, pass or fail, because the experiment needs the
      // rounds where it says `resolved` as much as the ones where it does not.
      (slideOk === null ? "" : slideOk ? " · slide 1 resolved" : " · slide 1 REFUSED") +
      (state.crashed ? " · PowerPoint's crash dialog is up" : ""),
  );
  // SWEEP IT RATHER THAN REFUSE, when the deck is the only thing wrong. See
  // `onlyDirtyDeck` — this is the one stop the driver can clear better than a
  // person, and the person doing it by hand is what emptied a deck entirely.
  if (!ok && onlyDirtyDeck(codes) && !healed && sweepDeck(sh)) {
    // ONCE, never in a loop. The re-check reads the deck through the same call,
    // so a sweep that did not take refuses exactly as it would have — a failed
    // heal cannot read as a successful one.
    console.log("  the deck still holds the last round's slides — sweeping it rather than refusing");
    await sleep(8000);
    return attempt(argv, deps, sh, true);
  }
  if (!ok) {
    console.error("\n  NOT READY — a round now would not prove anything:");
    for (const s of stop) console.error(`    - ${s}`);
    return { code: 1, reason: state.crashed ? "crashed" : "not-ready", codes };
  }
  console.log("  ready");
  if (checkOnly) return { code: 0, reason: "checked" };

  // THE RUN BUTTON ONLY EXISTS ON THE AUTOMATION TAB, and a pane does not open
  // there. A freshly sideloaded or reopened pane comes up on `Chart`, so the
  // button is not merely hidden — it is not in the DOM at all.
  //
  // 2026-08-20: the driver put the add-in back by itself for the first time,
  // reached `ready`, and then stopped on `no-run-button` with the pane sitting
  // on Chart and `verbose trace ?` in the same block — the checkbox that reads
  // as `?` lives on the same tab as the button that was missing. Everything
  // needed to notice was on screen; nothing acted on it.
  //
  // So: select the tab, then look again. This is the cheapest possible recovery
  // and it was previously an un-retryable stop.
  // The Automation tab was selected before readiness read its toggles, so the
  // button should be here. One more try anyway — the pane can be re-rendered
  // between the two, and a second look is cheaper than a lost round.
  let runBtn = refFor(sh, "Probe, then self-test", /button "Probe, then self-test"/);
  if (!runBtn) {
    const automation = refFor(sh, "Automation", /tab "Automation"/);
    if (automation) {
      clickRef(sh, automation);
      runBtn = refFor(sh, "Probe, then self-test", /button "Probe, then self-test"/);
    }
  }
  if (!runBtn) {
    console.error("  could not find the run button, and the Automation tab did not bring it back");
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
  /** Consecutive polls whose CLI call could not be run — see `browserDiedMidRound`. */
  let failedPolls = 0;
  for (;;) {
    if (Date.now() - started > limit) {
      console.error("  the round has not finished in 30 minutes — the host is wedged; see docs/ROUNDS.md");
      await keepCrashEvidence(sh, "30 minutes with no finish");
      return { code: 1, reason: "timeout" };
    }
    const dl = sh("find", "Download run log");
    // CAPTURED HERE, against the call that produced `dl`. `sh.state.lastFailed`
    // belongs to the MOST RECENT call, and the crash-dialog read below overwrites
    // it before either counter is updated — so both were judging this read's
    // emptiness against a different call's success.
    //
    // That re-opened the exact regression the comment above this loop describes.
    // The `failed` argument exists because one poll exiting non-zero means
    // NOTHING was measured, and folding that into "the pane is gone" once killed
    // a round that went on to finish 10 of 12 scenarios. With the flag taken from
    // the crash-dialog read instead, a failed `dl` read paired with a successful
    // crash read counts as a genuine silence: two of those in a row and the
    // driver kills a healthy round, files a crash report for a crash that never
    // happened, and lets recovery reload the tab out from under it.
    const dlFailed = sh.state.lastFailed;
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
      const secs = Math.round((Date.now() - started) / 1000);
      console.error(
        `  PowerPoint crashed ${secs}s in — its dialog is up and nothing behind it will answer. ` +
          'See docs/ROUNDS.md, "The wedge".',
      );
      // NOW, before recovery reloads the tab and takes the request log with it.
      // This is the only window in which the host's own account of the crash
      // exists, and three separate hand passes were spent reaching it.
      await keepCrashEvidence(sh, `${secs}s into a round`);
      return { code: 1, reason: "crashed" };
    }
    // THE BROWSER, not just the pane. Checked before the quiet counter because
    // the counter cannot see this state at all — see `browserDiedMidRound`.
    failedPolls = dlFailed ? failedPolls + 1 : 0;
    if (failedPolls >= DEAD_BROWSER_POLLS && browserDiedMidRound(failedPolls, sh("list"))) {
      const secs = Math.round((Date.now() - started) / 1000);
      console.error(
        `  the browser died ${secs}s into the round — the process is gone, taking the tab with it. ` +
          "The persistent profile still holds the sign-in, so this is recoverable without a password.",
      );
      return { code: 1, reason: "browser-gone" };
    }
    quiet = quietStreak(quiet, dl, dlFailed);
    if (quiet >= 2) {
      console.error("  the pane stopped answering — PowerPoint has probably crashed; the trace is still in the DOM");
      await keepCrashEvidence(sh, `${Math.round((Date.now() - started) / 1000)}s into a round, pane silent`);
      return { code: 1, reason: "silent" };
    }
    await new Promise((r) => setTimeout(r, 20000));
  }
  console.log("  finished");
  // FINISH THE JOB, rather than hand back a round somebody has to collect. The
  // three steps below were done by hand after every one of the 48 archived
  // rounds, and two of them are where the mistakes were: a stale log filed as a
  // fresh round (039 was byte-identical to 038), and a deck cleared by hand down
  // to zero slides.
  //
  // Each is best-effort and none can fail the round. The round is DONE by this
  // point — its evidence is in the pane either way, and a driver that turned a
  // good round into a non-zero exit over housekeeping would be worse than the
  // housekeeping.
  const roundFile = await collectRound(sh, stamp, sleep);
  return { code: 0, reason: "finished", roundFile, build: stamp, size };
}

/**
 * Download the run log, archive it, and leave the deck ready for the next round.
 *
 * `stamp` is the build the PANE was serving, checked against the log's own build
 * before anything is filed — see `archive`. That is what stops the previous
 * round's file being archived as a new one when a wedge left the download button
 * disabled and the click did nothing.
 *
 * Returns the name it filed, or null. The caller needs it for the outcome
 * receipt: a cycle runner that has just been told a round finished has no other
 * way to name WHICH round, and guessing "the newest file in rounds/" is the
 * assumption that already produced one wrong overwrite diagnosis in this repo.
 */
async function collectRound(sh, stamp, sleep) {
  let filed = null;
  try {
    const dl = refFor(sh, "Download run log", /button "Download run log"/);
    if (!dl) return null;
    clickRef(sh, dl);
    await sleep(12000);
    const logPath = `${sh.dir ?? "."}/.playwright-cli/powerchart-run-log.json`;
    if (!existsSync(logPath)) {
      console.error("  the run log did not arrive — archive it by hand once it does");
      return null;
    }
    filed = archive(logPath, "rounds", readFileSync, writeFileSync, everyRoundEverFiled, stamp);
    console.log(`  archived as rounds/${filed}`);
  } catch (err) {
    // Named, never swallowed. A round whose log was not filed is a round that
    // will be filed by hand, and the person doing it needs to know why.
    console.error(`  could not archive this round: ${err instanceof Error ? err.message : String(err)}`);
  }
  // The deck last, so a failed archive still leaves it ready — the two are
  // independent and coupling them would cost the next round for the sake of
  // this one's paperwork.
  // SAY WHICH, because the two lead to different mornings. A sweep that did not
  // clean the deck leaves the next round to refuse with `deck-dirty`, and this
  // line used to claim the opposite whatever happened.
  if (sweepDeck(sh)) console.log("  deck swept — the next round starts clean");
  else console.error("  the deck was NOT swept — the next round will refuse until it is");
  return filed;
}

/**
 * The stops a reload-and-reopen actually clears — which is to say, the stops
 * `recover` was written for.
 *
 * Derived from that function rather than from a judgement about which refusals
 * feel transient: `recover` clicks Refresh or reloads, waits out the reload,
 * reopens the pane from the ribbon, clicks the Automation tab, and cleans the
 * deck. Every code here is undone by one of those five steps, and every code
 * NOT here survives all of them.
 *
 * `wrong-size` is deliberately absent and must stay that way. `recover` could
 * set a deck's slide size — it is two clicks — and doing so would CHANGE WHAT
 * THE ROUND MEASURES rather than restore it, which is the one thing recovery is
 * not allowed to do. A deck in the wrong profile is a setup error for a person,
 * not a transient state to clear.
 *
 * `addin-missing` is absent for the same kind of reason and was learned the same
 * way. `recover` reopens the pane from the ribbon's `Insert chart` control, so a
 * document that does not carry that control has nothing for recovery to click;
 * on 2026-08-16 the driver spent seven attempts and about fifteen minutes
 * proving that twice in one night. It is not transient either — a browser death
 * takes the sideload with it, and only a person can put it back.
 *
 * Deliberately absent, and each for its own reason: `site-behind` and `no-build`
 * are waiting for Pages and a reload does not make it deploy faster;
 * `verbose-off` and `pictures-off` are choices a person made in the pane and
 * silently re-making them would change what the round measures; the sign-in and
 * unreachable-CLI states never get here because they return before the codes do.
 */
export const RECOVERABLE_STOPS = new Set([
  "browser-gone",
  "crashed",
  "host-silent",
  "slide-refused",
  "pane-closed",
  "pane-stale",
  "deck-dirty",
  // A greyed-out ribbon on a `Disconnected` document. Transient by nature — the
  // tab reconnects or a reload clears it — and it used to be reported as
  // `addin-missing`, which recovery is forbidden to retry.
  "host-disconnected",
]);

/**
 * Is another attempt worth making?
 *
 * IT USED TO BE A CRASH AND NOTHING ELSE, and the cost of that showed up on
 * 2026-08-15 in two places on one afternoon. Mid-round, the QUIET form of the
 * wedge exits as `silent` — the host stops answering with no dialog — and the
 * driver went home, though `docs/ROUNDS.md` says in as many words that a reload
 * clears both forms and `recover` already does exactly that. At check time, a
 * person then hand-fixed a silent host, a pane one build behind and an
 * eight-slide deck, in the same order `recover` does them, because a `not-ready`
 * was never retried either.
 *
 * So the question is no longer "was it a crash" but "does recovery address
 * everything that refused". A `not-ready` whose codes are all recoverable is
 * worth another attempt; one carrying a single stop recovery cannot touch is
 * not, and stopping on it is the behaviour the old comment was right about — a
 * round that retries a stale build until the night is gone measures nothing.
 *
 * `codes` is optional so the two reasons that carry none (`crashed`, `silent`)
 * read the same as they always did.
 */
/**
 * What the recovery about to run is actually recovering FROM.
 *
 * One line, and it is read at the worst moment — mid-loop, by someone deciding
 * whether the round is worth watching. Saying "clearing the crash" when the deck
 * was merely dirty is how a debugging session starts by hunting a crash that
 * never happened.
 */
export function recoveryFor(reason, codes) {
  if (reason === "crashed") return "clearing the crash and starting again";
  if (reason === "silent") return "the host went quiet — reloading and starting again";
  if (reason === "timeout") return "the round wedged and did not finish — reloading and starting again";
  const named = (codes ?? []).filter((c) => RECOVERABLE_STOPS.has(c));
  if (!named.length) return "recovering and starting again";
  const say = {
    crashed: "a crash dialog",
    "host-silent": "a silent host",
    "slide-refused": "a host that would not resolve slide 1",
    "pane-closed": "a closed pane",
    "pane-stale": "a stale pane",
    "deck-dirty": "a dirty deck",
  };
  return `recovering from ${named.map((c) => say[c] ?? c).join(" and ")}, then starting again`;
}

/**
 * Has the BROWSER died under a round that is still polling?
 *
 * The hole this closes, and it cost 24 minutes on 2026-08-16. `quietStreak`
 * resets to zero whenever a CLI call FAILED, deliberately: one failed call means
 * nothing was measured, and treating it as "the pane is gone" once killed a
 * healthy round that went on to pass 10 of 12. That protection is right and is
 * left alone here.
 *
 * But a dead browser makes every call fail, permanently — so the quiet counter
 * can never reach its threshold, the crash dialog cannot be read either, and the
 * loop polls a corpse until the thirty-minute limit. `pw list` said
 * `(no browsers)` outright while the driver sat there.
 *
 * Two conditions, and both are needed. A STREAK of failures, so ordinary
 * contention (a second terminal, an agent reading the trace) cannot trigger it —
 * three consecutive misses is a minute of silence, far past any blip. And then
 * an affirmative `(no browsers)`, so the round is never ended on the absence of
 * evidence. Either alone would re-make the mistake the other guards against.
 */
export const DEAD_BROWSER_POLLS = 3;
export function browserDiedMidRound(failedStreak, listOutput) {
  return failedStreak >= DEAD_BROWSER_POLLS && noBrowser(listOutput);
}

/**
 * Is a dirty deck the ONLY thing standing in the way?
 *
 * A dirty deck is not a fault — it is the last round's slides, and the driver
 * already knows how to sweep them (`cleanDeckScript`, which `recover` runs every
 * time it fires). Refusing over it makes a person do by hand the one step the
 * machine does better: hand-clearing took a deck to ZERO slides on 2026-08-16
 * and produced `slide 1 REFUSED`, the state the 2s crash starts from.
 *
 * ONLY, and that word is the whole guard. A deck that is dirty AND on a stale
 * pane is a round that would measure the wrong build, and healing the cheap half
 * moves it one step closer to running while still being wrong. Anything beyond a
 * dirty deck refuses out loud, as before.
 */
export function onlyDirtyDeck(codes) {
  return Array.isArray(codes) && codes.length === 1 && codes[0] === "deck-dirty";
}

export function shouldRetry(reason, attempt, max, codes) {
  if (attempt >= max) return false;
  // A WEDGE MID-ROUND IS RECOVERABLE, and leaving `timeout` out of this stopped
  // an unattended run dead in its first hour. The quiet wedge that produces it is
  // the same state `silent` names — `docs/ROUNDS.md` says a reload clears both
  // forms — and `recover` already does exactly that. The argument for excluding
  // it was cost: a wedge has already burned thirty minutes and another may burn
  // thirty more. But that cost is bounded by `--retry N`, which the caller chose,
  // and the alternative is a ten-hour run that ends at hour one with the host
  // sitting idle and recoverable.
  // `browser-gone` joins them: the profile keeps the sign-in, so reopening needs
  // no password and `recover` already does it. A round that ends this way has
  // burned a minute, not thirty — it is the cheapest of these to retry.
  // `threw` joins them: an unexpected exception used to kill the process
  // outright, so `--retry` never saw it. Retrying is bounded by `max` — a
  // deterministic bug fails that many times and stops — and the alternative is
  // a night that ends on its first surprise with recovery never attempted.
  if (
    reason === "crashed" ||
    reason === "silent" ||
    reason === "timeout" ||
    reason === "browser-gone" ||
    reason === "threw"
  )
    return true;
  if (reason !== "not-ready") return false;
  // An EMPTY list is not a licence. It means nothing was recorded about why the
  // check refused, and retrying on no evidence is how a loop spins.
  return Array.isArray(codes) && codes.length > 0 && codes.every((c) => RECOVERABLE_STOPS.has(c));
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
export async function recover(sh, sleep, profile = PROFILE_DIR) {
  // NO BROWSER AT ALL comes first, because everything below reloads and clicks
  // inside a window that is not there. Not hypothetical: on 2026-08-15 a round
  // wedged, the browser process died with it, and this function then reloaded
  // nothing and reopened nothing seven times while the check reported "could not
  // read the pane's build stamp — is the add-in open?"
  //
  // Reopening is the loop's to do, not the owner's. The persistent profile still
  // holds the sign-in — a dead browser is not a lost sign-in — so it needs no
  // password, and the alternative is a ten-hour run ending in its first hour.
  if (noBrowser(sh("list"))) {
    sh("open", "--persistent", `--profile=${profile}`, "--headed", "https://onedrive.live.com/");
    await sleep(15000);
    // The deck, and then ITS tab. Clicking the file opens a NEW tab while the
    // CLI stays on the old one, and skipping that is how a healthy setup reads
    // as a closed pane.
    // THE DECK'S NAME IS NOT A CONSTANT, and hard-coding it here was a trap that
    // fired the day it was written about. This said `Presentation63` while the
    // deck in use had become `Presentation64` — a new document, because a web
    // sideload is per-document — so a browser death would have reopened OneDrive
    // and then failed to find anything, silently, in exactly the situation this
    // function exists for. `PW_DECK` overrides it for a deck named anything else.
    const deckName = process.env.PW_DECK ?? DECK_NAME;
    const deckPattern = new RegExp(`link "${deckName}`);
    const deck = refFor(sh, deckName, deckPattern);
    if (deck) clickRef(sh, deck);
    await sleep(25000);
    const line = sh("tab-list")
      .split("\n")
      .find((l) => l.includes(deckName));
    const n = line ? /(\d+):/.exec(line)?.[1] : null;
    if (n) sh("tab-select", n);
    await sleep(20000);
  }
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

  sweepDeck(sh);
  return Boolean(pane);
}

/**
 * Delete every slide but the first, through the pane's own Office.js context.
 *
 * Pulled out of `recover` so the readiness check can use it without the reload
 * and reopen that surround it there — see `onlyDirtyDeck`. Same call, same
 * budget, ONE implementation: a second sweep written beside this one is how the
 * hardcoded deck name and the three stale slogans in `triage.mjs` happened.
 */
export function sweepDeck(sh) {
  const anchor = refFor(sh, "Chart", /tab "Chart"/);
  if (!anchor) return false;
  // READ WHAT IT ANSWERED. This threw the result away and returned `true`
  // unconditionally, so "deck swept — the next round starts clean" printed
  // whether the sweep had cleaned the deck, failed outright (`deck-failed`), or
  // left slides behind — and the next round then refuses with `deck-dirty` for a
  // reason the previous round's output said could not have happened.
  //
  // `cleanDeckScript` returns `deck:N`, the slides remaining. One is clean: the
  // loop stops at index 1 on purpose, because a deck cannot have zero slides and
  // a fixed-count delete once took one to exactly that.
  const out = sh("eval", cleanDeckScript(90000), anchor);
  const left = /deck:(\d+)/.exec(out)?.[1];
  if (left === undefined) return false;
  return Number(left) <= 1;
}

/** Where the driver leaves its account of how the round ended. */
export const RECEIPT_PATH = ".round-outcome.json";

/**
 * What just happened, as structure rather than prose.
 *
 * The driver's exit code is BINARY — 0 for a round that finished, 1 for
 * everything else — so anything downstream that wants to know WHY a round
 * stopped has exactly two options: parse the console output, or be told. The
 * console output is prose written for a person at 2am and edited whenever a
 * message is improved, and `rounds-gate.mjs` already refused to parse prose once
 * for the same reason: a gate that reads sentences is a gate that breaks when
 * someone fixes a sentence.
 *
 * So the driver writes down what it knows. `reason` and `codes` are the same
 * values `shouldRetry` judges, which means a reader can apply `RECOVERABLE_STOPS`
 * itself rather than reimplementing the judgement — and there must only ever be
 * one implementation of "is this worth another attempt".
 *
 * `roundFile` matters as much as the reason. A caller told only that a round
 * finished would have to guess which file it produced, and "the newest file in
 * rounds/" is precisely the assumption that produced a wrong overwrite
 * diagnosis in this repo once already.
 *
 * Pure, and separate from the writing, so the shape can be tested without a
 * filesystem.
 */
export function outcomeReceipt({ reason, codes, roundFile, build, size, threw, at }) {
  return {
    reason: reason ?? null,
    // Always an array. A reader doing `codes.includes(...)` on a `finished`
    // round should get `false`, not a crash on undefined.
    codes: Array.isArray(codes) ? codes : [],
    roundFile: roundFile ?? null,
    build: build ?? null,
    // THE PROFILE STRING, which is what `readSlideSize` actually returns —
    // "16:9", "4:3", or "960x540" for anything else. This was written as
    // `{ width: size.width, height: size.height }` against a shape the driver
    // has never produced, so every cycle leg recorded `"size": {}` and the
    // field that says WHICH ARM a round belonged to said nothing at all. It
    // read as correct because its test passed an object in, which tested the
    // assumption rather than the driver.
    size: size ?? null,
    // Only on the path that has one — see the catch in `main`. Present means
    // the round failed in a way nothing anticipated, which is a different
    // thing from every named refusal and should not have to be inferred.
    ...(threw ? { threw } : {}),
    // Whether the reason is one recovery ADDRESSES, decided here by the same set
    // the driver retries on — never re-derived downstream.
    // THE SAME QUESTION `shouldRetry` ANSWERS, asked the same way. This read
    // only the codes, and most stops carry none: a crash, a silent host, a
    // wedge, a dead browser and an unexpected throw are all retried on their
    // REASON alone. So a night that recovered from a crash six times ended by
    // printing "crashed is not something recovery addresses — it needs a
    // person", directly under six lines saying "clearing the crash and starting
    // again". The receipt contradicted the console in the same output.
    recoverable: shouldRetry(reason, 0, 1, codes),
    at: at ?? new Date().toISOString(),
  };
}

async function main(argv, deps = {}) {
  const run = deps.run ?? spawnSync;
  const sleep = deps.sleep ?? ((ms) => new Promise((r) => setTimeout(r, ms)));
  const dirArg = argv.indexOf("--dir");
  const sh = cli(run, dirArg === -1 ? process.cwd() : argv[dirArg + 1]);
  const retryArg = argv.indexOf("--retry");
  const max = retryArg === -1 ? 0 : Number(argv[retryArg + 1]) || 0;

  const write = deps.write ?? writeFileSync;

  for (let n = 0; ; n++) {
    if (n) console.log(`\n  attempt ${n + 1} of ${max + 1}`);
    // THE UNKNOWN FAILURE IS A FAILURE TOO. `attempt` is ~200 lines with one
    // try/catch in it, and everything it does not anticipate arrived here as an
    // unhandled rejection: the process died with a stack trace, `--retry`
    // covered none of it, and no receipt was written — so a night that had six
    // attempts left ended on the first one, and whatever ran it could not even
    // say why.
    //
    // Retried like any other reason, and bounded by the same `--retry` the
    // caller chose. A deterministic bug will simply fail `max` times and stop;
    // a transient one — the kind `recover` exists for — gets the same second
    // chance a crash dialog does.
    let outcome;
    try {
      outcome = await attempt(argv, deps, sh);
    } catch (err) {
      console.error(`  the round threw where nothing expected it to: ${err?.message ?? err}`);
      outcome = { code: 1, reason: "threw", codes: [], threw: String(err?.message ?? err) };
    }
    const { code, reason, codes, roundFile, build, size, threw } = outcome;
    if (!shouldRetry(reason, n, max, codes)) {
      // ONLY THE OUTCOME THAT STANDS. Writing a receipt per attempt would leave
      // a caller reading the state of a round that recovery went on to fix,
      // which is the opposite of what the file is for.
      //
      // Best-effort, exactly like archiving: a round that ran is not undone by
      // a failed write, and turning one into a non-zero exit would make the
      // paperwork more important than the evidence.
      try {
        write(RECEIPT_PATH, JSON.stringify(outcomeReceipt({ reason, codes, roundFile, build, size, threw }), null, 2));
      } catch (err) {
        console.error(`  (could not write ${RECEIPT_PATH}: ${err?.message ?? err})`);
      }
      return code;
    }
    // NAMED FOR WHAT ACTUALLY HAPPENED. This said "clearing the crash" whatever
    // the reason was, and the moment the retry covered more than crashes it
    // started lying: round 047 refused on a dirty deck alone and was told a
    // crash was being cleared. A recovery line is read while someone is
    // debugging, and one that invents a crash sends them looking for it.
    console.log(`  ${recoveryFor(reason, codes)} — see docs/ROUNDS.md, "The wedge"`);
    await recover(sh, sleep);
  }
}

/**
 * Write the crash report beside the rounds, and never let it cost the driver.
 *
 * Named for what it is: the evidence outlives the tab only if something keeps
 * it. `crashes/` rather than `rounds/` on purpose — these are not rounds, and
 * everything downstream pools that directory.
 */
async function keepCrashEvidence(sh, at, write = writeFileSync) {
  try {
    const report = await collectCrashEvidence(sh, { at });
    const stamp = new Date().toISOString().replace(/[:.]/g, "-").slice(0, 19);
    const path = `crashes/${stamp}.md`;
    write(path, report);
    console.error(`  the host's own account of it is in ${path}`);
  } catch (err) {
    console.error(`  (could not keep the crash evidence: ${err?.message ?? err})`);
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
export function archive(
  logPath,
  dir = "rounds",
  read = readFileSync,
  write = writeFileSync,
  // THE DEFAULT, because the caller that mattered was passing the unsafe one.
  // `everyRoundEverFiled` was written after two different rounds were both
  // filed as 064, and it was wired to the `--archive` subcommand — the path a
  // person uses by hand, occasionally. `collectRound`, which archives EVERY
  // round automatically, went on passing `readdirSync`: the working tree alone,
  // blind to any round committed on a branch that is not checked out. That is
  // precisely the collision it was written to stop, on the path that runs a
  // hundred times more often.
  //
  // Defaulted rather than fixed at the call site so the next caller cannot make
  // the same omission. Tests pass their own lister and are unaffected.
  list = everyRoundEverFiled,
  expectBuild = null,
) {
  const round = stripImages(JSON.parse(read(logPath, "utf8")));
  const build = buildOf(round.build);
  if (!build) throw new Error("that file carries no build stamp — it is not a round log");
  // THE LOG MUST BE THE ROUND THAT JUST RAN, and this is the guard the twin
  // check cannot be. `Download run log` is DISABLED while a round is running, so
  // clicking it after one that wedged does nothing at all — and the previous
  // round's file is still sitting at the same path, waiting to be filed under a
  // new number. That happened: round 039 was archived byte-identical to 038, a
  // whole round of evidence that never took place.
  //
  // The twin check below catches that case only when the two logs are IDENTICAL.
  // A stale log that merely differs — an older round, a different build — sails
  // past it. Comparing the log's own build stamp to the pane's is what closes
  // it, and it costs nothing: the driver already read the stamp to decide the
  // round was worth running.
  if (expectBuild && build !== expectBuild)
    throw new Error(
      `that log is build ${build} and the pane is serving ${expectBuild} — it is the PREVIOUS round's file, ` +
        "which is what sits at that path when a round wedges and the download button does nothing. Nothing archived.",
    );
  // THE DOWNLOAD IS A FILE ON DISK THAT IS ONLY SOMETIMES REPLACED. `Download
  // run log` is disabled while a round is running, so clicking it after a round
  // that WEDGED does nothing at all — and the previous round's log is still
  // sitting at the same path, waiting to be archived a second time under a new
  // number. That happened: round 039 was filed as byte-identical to 038, a
  // whole extra round of evidence that never took place, and only a checksum
  // caught it.
  //
  // A fabricated round is the worst thing this directory can hold. Everything
  // downstream pools these files — verdict histories, the rasterise arms, the
  // scenario flip detector — so one duplicate quietly doubles the weight of
  // whatever the real round happened to say.
  const body = `${JSON.stringify(round, null, 2)}\n`;
  const twin = list(dir)
    .filter((f) => /^\d{3}-.*\.json$/.test(f))
    .find((f) => {
      // A NAME THIS LISTER RETURNS NEED NOT BE A FILE. `list` defaults to
      // `everyRoundEverFiled`, whose entire purpose is to name rounds that are
      // NOT in the working tree — committed on a branch nobody has checked out —
      // so that `nextRoundNumber` below cannot hand their number out twice. That
      // function only ever inspects the NAMES. This check opens them, and it
      // inherited the new lister as collateral when the default changed.
      //
      // Unguarded, that threw ENOENT on the healthy path and only there: `.find`
      // walks the on-disk names first and reaches a git-only name only when
      // nothing matched, which is exactly the case of a genuinely new round. The
      // guard was perfectly inverted — a duplicate was caught, a real round
      // crashed — and `collectRound` swallowed the throw, so the driver exited 0
      // reporting a finished round with nothing written to rounds/ and the next
      // leg's download then overwrote the evidence.
      //
      // A file that cannot be opened is not a twin of anything.
      try {
        return read(`${dir}/${f}`, "utf8") === body;
      } catch {
        return false;
      }
    });
  if (twin)
    throw new Error(
      `that log is byte-identical to ${twin} — the pane never wrote a new one, ` +
        "which is what happens when the round did not finish. Nothing archived.",
    );
  // THE NUMBER COMES FROM WHAT `list` CAN SEE, and on 2026-08-16 that was not
  // everything. Round 064 was archived on `main`, its findings committed to a
  // branch, and `git checkout main` then REMOVED the file from the working tree
  // — it is tracked only on the branch. The next round was archived from `main`,
  // where the directory ends at 063, and was numbered 064 as well.
  //
  // Nothing was overwritten (`nextRoundNumber` is max+1, so it never lands on a
  // file it can see) and no evidence was lost. What it produces is two DIFFERENT
  // rounds both called 064 in two git contexts, which collides the moment either
  // is merged — and a pooled report that silently reads one of them twice, or
  // not at all.
  //
  // So the caller passes what GIT knows as well as what is on disk; see the
  // `--archive` branch below. `archive` itself stays pure and takes the union.
  const name = `${nextRoundNumber(list(dir))}-${build}.json`;
  // TWO spaces, because prettier checks this directory and every round archived
  // at one space failed the gate until someone reformatted it by hand.
  write(`${dir}/${name}`, JSON.stringify(round, null, 2) + "\n");
  return name;
}

/**
 * Every round this repo has ever filed — on disk AND in git — so the next number
 * cannot reuse one that is committed on a branch nobody has checked out.
 *
 * `readdirSync` alone sees the WORKING TREE, and a round's file lives on the
 * branch its findings were committed to until that merges. Check out `main` and
 * the file disappears; archive from there and the number is handed out twice.
 * That happened on 2026-08-16 and produced two different rounds both called 064.
 *
 * `git ls-files` covers the current branch's index; `git log --all` covers every
 * branch, which is what actually matters here — the collision is with a round
 * committed somewhere else. Falls back to the directory alone if git is not
 * available or the call fails, because refusing to archive a real round over a
 * numbering nicety would be the worse trade.
 */
function everyRoundEverFiled(dir) {
  const onDisk = readdirSync(dir);
  try {
    const inGit = spawnSync("git", ["log", "--all", "--name-only", "--pretty=format:", "--", dir], {
      encoding: "utf8",
      maxBuffer: 64e6,
    });
    if (inGit.status !== 0 || !inGit.stdout) return onDisk;
    const names = inGit.stdout
      .split("\n")
      .map((l) => l.trim())
      .filter(Boolean)
      .map((l) => l.slice(l.lastIndexOf("/") + 1));
    return [...new Set([...onDisk, ...names])];
  } catch {
    return onDisk;
  }
}

if (isMain(import.meta.url, process.argv[1])) {
  const argv = process.argv.slice(2);
  if (argv[0] === "--archive") {
    if (!argv[1] || !existsSync(argv[1])) {
      console.error("usage: node scripts/round.mjs --archive <powerchart-round.json>");
      process.exit(2);
    }
    console.log(`archived as rounds/${archive(argv[1], "rounds", readFileSync, writeFileSync, everyRoundEverFiled)}`);
    process.exit(0);
  }
  process.exit(await main(argv));
}
