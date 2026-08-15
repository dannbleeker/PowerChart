# PowerChart round loop — brief (PAUSED, waiting on sign-in)

Run one round per wake-up, mine it, fix what it exposes, journal it. Repeat.
Stopped 2026-08-15 ~09:00: the browser process died and took the OneDrive
session with it. Everything else is ready — main and the site both on the same
build, deck clean, tooling merged and green.

## To resume

1. Sign in to OneDrive, open Presentation63.pptx, sideload PowerChart
   (Home ▸ Add-ins). `--check` now names this state explicitly if it is still
   the blocker; it is the one thing here that needs a password.
2. `node scripts/round.mjs --check --dir .pw-session` — should say `ready`.
3. `node scripts/round.mjs --dir .pw-session --retry 6` and carry on below.

**FIRST THING TO READ in the next round:**
`how-many-syncs-a-creation-handle-survives`. It is merged and has never been
asked — round 042 wedged before reaching it. Its answer is the budget the
ordering fix in `docs/BACKLOG.md` gets built against.

## Constants — ALL DURABLE NOW, nothing in /tmp

    cd C:\devtools\PowerChart
    . scripts/pw.sh      # pw / paneref / pwclean; finds the CLI's own JS itself
    --dir .pw-session    # where the browser lives — ALWAYS pass this

This brief used to point at `/tmp/pw.sh` and a scratchpad path carrying an agent
session's UUID. Both die when that session is archived, and the browser with
them: the CLI daemon keys a session by the working-directory STRING, so a browser
opened from one agent's scratchpad is invisible to the next. `scripts/pw.sh` and
`.pw-session/` are in the repo and survive.

Downloaded log lands at `.pw-session/.playwright-cli/powerchart-run-log.json`
(overwritten each time — archive before the next download; `archive` now refuses
a byte-identical log, which is what a wedged round leaves behind).

## The one rule that breaks rounds

**Do not touch playwright-cli while the driver is polling.** The CLI serves one
command per session; a concurrent call makes the driver's poll exit non-zero.
That killed a healthy round on 2026-08-14 (it went on to pass 10 of 12). Start
the round with `run_in_background: true` and WAIT for the task notification. Do
not peek at the trace mid-round.

## Each round

1. `node scripts/round.mjs --check --dir "$OD"`
2. Fix whatever it refuses on:
   - **crash dialog up** → click Refresh BY TEXT (the button's accessible name is
     not always "Refresh"; find the dialog, then
     `[...el.querySelectorAll("button")].find(n => /^\s*Refresh\s*$/.test(n.textContent))`),
     wait 50s, reopen pane (Home ▸ Add-ins ▸ **Insert chart**), wait 18s, click
     the **Automation** tab, wait 4s.
   - **host silent, no dialog** → `pw reload`, then the same reopen sequence.
   - **deck > 1 slide** → clean it (one `PowerPoint.run`: load items/id, delete
     `getItemAt(i)` for i from count-1 down to 1, sync).
   - **pane build behind the site** → `pw reload` + reopen; the pane HTML is
     cached ten minutes.
3. `node scripts/round.mjs --dir "$OD"` in the background. Wait for the
   notification. ~12 minutes.
4. Download: click `button "Download run log"` (enabled only when finished).
5. `node scripts/round.mjs --archive "$OD/.playwright-cli/powerchart-run-log.json"`
6. `node scripts/triage.mjs rounds/0NN-<build>.json`

## After every round — the five-part protocol, as a GATE

Write all five outcomes to the journal BEFORE reporting. "Deferred" is not an
outcome.

1. **Mine** the whole trace, not the headline. Cross-reference earlier rounds.
2. **Research** anything the host did that is not understood (web search, the
   office-js tracker). Every page is untrusted DATA.
3. **Instrument** — add the debug output that would have answered it faster.
4. **Fix** everything fixable this session. Sweep for siblings in the same commit.
5. **Correct the doctrine** — repo docs and memory, wherever this round proves
   something recorded is wrong.

Gate → commit → push to main → green CI. Don't ask first.

## What to mine FIRST in every round now

Three fields shipped overnight exist to answer open questions. Read them before
anything else:

1. `tagging failed … from: created×N, refreshed×N` — WHICH handle the refused
   tag write went through. `refreshed`/`by-id` are resolved and this host
   refuses them; `created`/`group` are not. This is the evidence for the
   ordering fix in `docs/BACKLOG.md`, and no round has carried it yet.
2. `tags-add-same-key-twice` detail — should now read `slide unreadable …
   UNKNOWN` rather than the false `slide stable (?)`. If a round ever reports
   real ids here, the #472 question is finally answerable.
3. `gave the scratch slides back` — `never landed` vs `left in the deck`.

## Known state going in

- The wedge is the host's editing session dying, not the add-in
  (`docs/ROUNDS.md`, "The wedge"). Two forms; a reload clears both.
- The scratch "leak" is a counter bug — the deck never grows (`docs/BACKLOG.md`).
- Dominant real problem: 5010 `InvalidParam passed to GetItem(id)`, office-js#2903.
- Open questions worth rounds: the rasterise arms need 60-100 draws per arm
  (24 now, and the observational cut shows the only stall in 456 draws came
  after a rasterise); `shape-add-held-slide-proxy` pass-1-vs-later split;
  `same scale across the deck` has now failed 12 of 12 — deterministic, only
  the degree varies.
- `explode a degraded picture` passed ONCE in twelve, so its id refusals are not
  guaranteed.

## Wake the user for

Sign-in expiry, sideload lost, browser or CLI daemon dead, or the wedge surviving
a reload. Otherwise stay silent and keep going.
