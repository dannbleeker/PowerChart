# Round loop journal — started 2026-08-14 20:50, hard stop 2026-08-15 04:50

Brief: `loop-brief-v2.md` beside this file. One round per wake-up, five-part
protocol as a gate, journal before reporting.

## Round 29 — 3c09b6a — 10/12 pass (before the loop, kept for continuity)

Two attempts; first wedged 8s into question 4 (crash dialog), cleared in 90s.
Archived `rounds/029-3c09b6a.json`, merged as #488.

- **Mine** — `left=73 deckBefore=1 deckAfter=1 stillListed=0`: the scratch slides
  never landed. Id namespaces differ (`4123571114#…` vs `287#62081387`).
- **Research** — office-js#2903 confirmed as the 5010 source (18 hits).
- **Instrument** — none needed; the handback line already carried the answer.
- **Fix** — driver poll no longer ends a round on one failed CLI call;
  `archive()` writes 2-space JSON; `rounds/*.json` in `.prettierignore`.
- **Doctrine** — `docs/BACKLOG.md` scratch entry corrected: the counter is the
  bug, not the deck.

## Round 30 — 6d6e278 — started 20:57

Preconditions: host was silent (form 2, no dialog) AND the pane was a build
behind; one `pw reload` + reopen fixed both. Deck cleaned 7 → 1. Check green:
`host answered in 9ms · ready`.

WEDGED. Driver hit its 30-minute limit; the new poll behaved correctly (no false
"pane stopped answering"). Crash dialog up, deck still 1, `--check` named it.

- **Mine** — died at `shape-add-fresh-slide-proxy`, the fourth round to do so,
  but the recorded answer is `no-scratch-slide ms=90135`: the probe never got as
  far as asking. Three 90s timeouts on `listing the deck's slides` followed, then
  it gave up after 3 silent questions.
- **Research** — pulled the ULS batch (request 47998). Same signature as the two
  earlier crashes, and it corrects my own framing: the `PptApi Call` sequence in
  flight is GetSlides / GetId / GetItemOrNullObject — READS, the scratch-slide
  resolution. No `addGeometricShape` anywhere. `Failed to restore selection after
  load content` appears in all three.
- **Instrument** — none added; the ULS route is now documented and took two
  commands. The probe already records the scratch id and deck length.
- **Fix** — nothing add-in-side to fix: the host's log says the write never
  happened.
- **Doctrine** — `docs/ROUNDS.md` corrected: the crash fingerprint is a read
  sequence, and "the round's first write" was wrong.

Recovery: clicked Refresh by text, reopened pane, cleaned deck 1→1, check green
(`host answered in 4ms`). Round 31 started 21:3x.

## Round 31 — 6d6e278 — started 21:36

WEDGED, same signature, ~3 minutes in. Crash dialog up, deck 1.

- **Mine** — identical shape to round 30: `shape-add-fresh-slide-proxy` →
  `no-scratch-slide` after 90s, then three more 90s timeouts, then gave up.
- **Research** — nothing new to look up; the signature is now understood.
- **Instrument** — THE ONE THAT MATTERED. Rounds 30 and 31 each died at ~3
  minutes and held the driver's poll for the full 30. The poll now looks for the
  crash dialog every iteration and stops when it appears, reporting how far in.
  A DOM read, not a ping — a ping would interleave with the round's own batches.
  Merged as #490.
- **Fix** — same commit.
- **Doctrine** — already corrected after round 30 (#489).

## Round 32 — abcb7c5 — started 22:07

Preconditions green after a reload to the new build; deck cleaned to 1;
`host answered in 3ms`. First round to run with the mid-round crash watch
in the driver — though the watch only takes effect for rounds started after
#490 merged, so this one still has the old 30-minute behaviour.

WEDGED ~3 minutes in, same signature. Started before #490 merged, so it still
burned the full 30 minutes.

## Round 33 (first try) — 8b287e3 — 22:48

WEDGED **2 seconds in** — and the new crash watch caught it, live-verified: the
dialog was genuinely up and the host genuinely silent. 30 minutes saved.

Crashes are accelerating: 3 min, 3 min, 3 min, 2 s. Five of six attempts tonight.

- **Mine** — nothing new in the trace; it never got far enough to have one.
- **Research** — checked whether a fresh document would help. Dozens of scratch
  decks exist in OneDrive but the list gives no clean refs, and sideloading the
  add-in into a new one needs the file chooser, which is the owner's step. Not
  pursued.
- **Instrument + Fix** — `--retry N` (#491): the driver now clears the crash and
  starts again itself, encoding the fifteen-step recovery done by hand six times.
  Only a crash is retried; mutation-proven three ways.
- **Doctrine** — the recovery's traps are in the code's own comments: Refresh
  matched by TEXT, pane reopened after the reload, Automation before the run
  button, deck cleaned last.

## Round 33 (retrying) — cabb357 — started 22:57

`--retry 6`, unattended. Check green: `host answered in 3ms`.

FINISHED first try. Archived `rounds/030-cabb357.json` — 10 of 12.

## Round 34 — cabb357 — 23:10, archived as 031

Two crashes at 2s each, both cleared and restarted by `--retry` with no hand on
it, third attempt finished. First fully unattended recovery. 10 of 12.

## The thread worth pulling — `same scale across the deck`

Three rounds, same failure, and it is a real feature breaking:

    029  5 of 8 carry the scale; 2×no-config;                  flipped at 6
    030  4 of 8 carry the scale; 2×no-config;                  flipped at 5
    031  4 of 8 carry the scale; 1×unknown-shape, 1×no-config; flipped at 4

Every loss is `InvalidParam passed to GetItem(id)` (5010, office-js#2903) at
`writing the chart's config tag`. "Flipped" = two consecutive losses and the
scenario stops, so the score is a floor.

- **Mine** — traced it to `finishCharts` writing the tag through
  `shapes.getItemOrNullObject(id)` with an id read a sync earlier.
- **Research** — office-js#2903 already catalogued; nothing new to find.
- **Instrument** — `tag-through-refetched-shape` (#492): asks whether the id
  round trip is what the host refuses. `tags-on-fresh-shape` already says the
  fresh shape's own `.tags` works, so this is the untested half.
- **Fix** — deliberately NOT yet. A rewrite of this path on a theory is exactly
  what has been reverted before. The question decides it.
- **Doctrine** — PENDING_QUESTIONS entry states what each answer means.

## Round 35 — 360309f — started 00:05

First round carrying `tag-through-refetched-shape`. `--retry 6`, unattended.

FINISHED first try. Archived 032. 10 of 12.

**The new question answered immediately: `no-id`, every pass.** The host will not
report a freshly created shape's id at all, so the re-fetch was never reached.

Read with its neighbours that IS the answer — `tags-on-fresh-shape yes`,
`shape-proxy-survives-one-sync unreadable`, `shapes-items-count-honest
unreadable`. Wrote it up in RESEARCH.md as "a shape is reachable only in the
batch that created it", and added the partner `tag-the-creation-proxy-a-sync-later`
because production falls back to the creation proxy when the id is unreadable —
exactly the case this host produces. #493.

## Round 36 — 86eaf65 — 00:40, archived as 033

FINISHED first try. 10 of 12.

**The partner answered `yes`** — a tag CAN be written through the creation handle
a sync later. That corrects what I wrote an hour earlier:

    write through an aged creation handle   works
    read  through an aged creation handle   unreadable
    re-fetch by id                          no id to re-fetch by

- **Mine** — five rounds now, `same scale` fails identically every time.
- **Research** — none needed; the host answered directly.
- **Instrument** — the partner question, which paid for itself in one round.
- **Fix** — still not written. The evidence now points at "stop reaching for the
  id, keep the handle" rather than "tag at creation time", which is what the
  overstatement would have built.
- **Doctrine** — RESEARCH.md corrected in #494, with the wrong version kept
  visible rather than quietly replaced.

## Round 37 — 4d68e73 — started 01:20

Third sample of the `no-id` / `yes` pair. `--retry 6`, unattended.

FINISHED first try. Archived 034. 10 of 12.

The decisive field, found by reading the trace rather than the headline: every
settle-pass failure reports `withId: 0`. The repair that exists to rescue a lost
config tag resolves shapes BY ID, and this host never yields one.

- **Fix attempted and abandoned correctly** — wrote the obvious theory (the
  positional `load` shares a batch with the tag writes) as a test expecting to
  FAIL. It passed on the first run: the insert path already survives that. Kept
  the test, labelled behaviour-not-guard. One wrong theory caught by writing the
  test first. #495.

## Round 38 — 0e230db — 03:00, archived as 035

FINISHED first try. 10 of 12. Third sample: `no-id` ×4 rounds,
`tag-the-creation-proxy-a-sync-later: yes` ×3. The resample bar is met.

**Where the fix belongs, corrected:** NOT the settle pass. It opens a fresh
`PowerPoint.run`, so the creation handle does not exist by the time it runs —
its only routes are an id and a collection read, and this host refuses both.
The retry belongs in the DRAWING context, before the handle is thrown away.

Same goal as `binding-names-shape-later` (`unreadable` here, four rounds); the
creation-handle write is a second route that answers yes and needs no 1.8
surface. #496.

## Round 39 — cfbf434 — started 04:05

Final round of the night. `--retry 4`.

FINISHED first try. Archived 036. 10 of 12. Fourth sample of the pair:
`no-id`, `tag-the-creation-proxy-a-sync-later: yes`.

## Night's tally — 2026-08-14 20:50 to 2026-08-15 04:30

Nine rounds attempted, **six landed** (029-036 archived, less the wedged ones),
all 10 of 12. Twelve PRs merged, every one gated on exit codes.

Rasterise arms: 24 draws per arm pooled over 12 rounds. Still needs 60-100.

What the night actually bought, in order of what it cost to learn:

1. The wedge is the host's session dying — PowerPoint's own crash dialog or a
   severed session — not the add-in. A reload clears both, in seconds.
2. The scratch "leak" is a counter bug; the deck never grew.
3. A shape can be WRITTEN through its creation handle a sync later, but never
   read back, and its id never round-trips.
4. `same scale across the deck` fails because the settle pass resolves by id and
   this host has none — `withId: 0`, every failure, six rounds.
5. Therefore the retry belongs in the drawing context, not the settle pass.

Three wrong theories were killed by measurement rather than argument: the
burnsTheSlide question, the shared tag/load batch, and "both routes out of the
creating batch are closed".


## Resumed window — 2026-08-15 07:20 to 12:20

### Round 40 — 3c163c4 — archived as 037 — 10/12

**All three overnight instruments delivered on their first real outing.**

- `from: created×1` ×7 and `created` ×4. **Not one `refreshed`, `group` or
  `by-id`.** The suspect that started this — the pre-grouping re-read taking the
  tag target — is cleared on the real host. Swapping the target changes nothing;
  the ordering is what has to change.
- A SECOND fault was hiding under the first: 4× `Cannot read properties of
  undefined (reading 'add')` on a `created` handle. `.tags` GONE, not refused.
- Slide identity now reads `unreadable … UNKNOWN (?, ?, ?)` — the false
  `stable (?)` is gone.
- Scratch handback now reads `79 never landed — the host took the add and the
  deck never listed them`.

Instrument built in response: `how-many-syncs-a-creation-handle-survives` (#501).
`tags-on-fresh-shape` (creating batch) and `tag-the-creation-proxy-a-sync-later`
(one sync) both answer yes; production's handles are older because the renderer
chunks a chart across batches. The question ages a handle one sync at a time and
reports where it breaks — a NUMBER the ordering fix can be built against.

### Round 41 — 3c163c4 — archived as 038 — 10/12

Identical `from` distribution to 037. Two rounds agreeing is what the ordering
conclusion now rests on. #502.

### Round 42 — 5a2522e — WEDGED, then the browser died

Wedged in the QUIET form — no crash dialog, so the 30-minute timeout caught it
rather than the crash watch. ~~The lifetime question was never answered.~~
**Wrong, corrected 2026-08-15 14:10 — it was answered twice and the answer sat
in the pane's storage for six hours.** See round 43 below: the pane still had
the run on offer, `Download the crashed run` handed over all 270 steps, and
`how-many-syncs-a-creation-handle-survives` had answered `survives-8` at 50.5s
and again at 67.2s before the wedge. A wedged round is not an empty one, and
what it reached is worth downloading before the next run buries it.

Then the archive step INVENTED A ROUND. `Download run log` is disabled while a
round runs, so the click did nothing and the previous round's file was still at
the same path; it was filed as 039, byte-identical to 038. A checksum caught it.
Deleted, and `archive` now refuses a byte-identical log by name (#503).

**BLOCKED — the browser process is gone and the sign-in with it.** `pw list`
answers `(no browsers)`, no chrome processes. Reopening works, but the fresh
browser has no cookies: onedrive.live.com redirects to login.live.com. Signing in
is the owner's step and mine to stay out of.

TO RESUME: sign in to OneDrive, open Presentation63.pptx, sideload PowerChart
(Home ▸ Add-ins), then `node scripts/round.mjs --check --dir "$OD"` and carry on
from the brief. The loop was stopped rather than left waking into a blocked
state.

Open question still unanswered: `how-many-syncs-a-creation-handle-survives`. It
is merged and will be asked by the first round that runs.

## Resumed 2026-08-15 14:00 — the sign-in was never gone

### Round 43 — 04510e2 — archived as 039 — 10/12

**The lifetime question is answered: `survives-8`, and it is not the
constraint.**

Setup, because the blocker recorded above was not the one that existed: the
browser process was dead, but `--persistent --profile=C:/devtools/pw-profile`
still carried the OneDrive session, so no password was needed. The deck opened
into a THIRD tab and the CLI stays on the tab it opened — `pw tab-list` then
`tab-select 2` is the step that was missing, and without it `--check` reads a
healthy setup as a closed pane. First attempt crashed 2s in; `--retry` cleared
it and attempt 2 landed clean.

1. **Mine.** `how-many-syncs-a-creation-handle-survives` answered `survives-8`
   on all three passes, `stable: true`, and all three were sampled in the
   `collection-refused` regime — the handle took a tag write at every one of
   eight successive syncs while the host was refusing collection reads. Five
   samples now exist across two builds (three here, two recovered from round
   42). **The budget is not the constraint: an UNRESOLVED creation handle does
   not age out.**

   Production failed anyway, and the same way as 037 and 038: every
   `tagging failed` carried `from: created×1` — no `refreshed`, no `group`, no
   `by-id`, three rounds running. The host's own statement dump names the
   mechanism with the ids in it:

       var slide  = slides.getItem("288#2569279682") /* originally getItemOrNullObject(...) */;
       var shapes = slide.shapes;
       var shape  = shapes.getItem("27") /* originally addTextBox(...) */;   ← 5010
       var tags   = shape.tags;

   `27` is the title box, drawn in the first batch and written to in the last —
   the same id the comment at `powerpoint.ts:6956` already quotes from the
   2026-08-07 log. That line is the resolver: `created[k].load("id")` on each
   batch's own sync resolves EVERY created shape including the anchor, so by the
   time the tag is written there is no unresolved handle left. `ungroupedFallback`
   is not the culprit — it slices the anchor off deliberately.

   Rest of the sheet: `tags-add-same-key-twice` still `other` with the honest
   `slide unreadable … UNKNOWN (?, ?, ?)`; `scratch-slides-returned: some` —
   1 of 84 deleted, **83 never landed**, deck 7 slides with 6 added and 0 blank,
   so the counter bug reading holds; 37 draws, ZERO stalls, slowest batch 17.1s,
   one population.

2. **Research.** Nothing new to look up. The 5010 is office-js#2903 (closed, not
   planned) and triage names it inline; no unexplained host behaviour this round.

3. **Instrument.** The three overnight fields all delivered again, and the one
   gap this round exposes is in the probe rather than the pane: the lifetime
   question asks through `ctx.scratch()`, a slide resolved in the batch, while
   production's parent is a slide handle Office has rewritten from
   `getItemOrNullObject(id)` to `getItem(id)`. The probe therefore cannot see a
   refusal that comes from the PARENT path. A partner question that varies only
   the parent would settle it. Not built here — see below.

4. **Fix.** Nothing shipped in the renderer. The ordering change this evidence
   supports is the open BACKLOG item, and the repo's own note on it says it
   "wants a session of its own"; starting it inside a round-mining pass is how a
   wrong theory gets shipped. What the round buys it is the headroom: keep one
   handle unresolved and there are at least eight syncs to spend.

5. **Doctrine.** Round 42's "never answered" corrected above. `docs/BACKLOG.md`
   ordering item now carries the measured budget. This brief's FIRST THING TO
   READ is spent and replaced.

Verdict grid is now 15 rounds: `same scale across the deck` FAIL 15 of 15 —
deterministic, only the degree varies (4 of 8 charts carried the shared scale
here, host flipped at chart 5 of 8). `explode a degraded picture` is the one
scenario that has ever flipped, and its single pass is still 1 in 15.

Rasterise arms: 30 draws per arm pooled over 15 rounds, both 0.0%. Still short
of the 60-100 an arm it needs.

Recovered round-42 log kept at `.pw-session/crashed-5a2522e.json` — not filed
under `rounds/`, since a wedged partial has no self-test and would read as a
16th round in the ledger.

### Round 044 — b6582c7 — two attempts lost to the DRIVER, not the host

Both attempts died identically, and the message was false:

    PowerPoint crashed 2s in — its dialog is up …
    clearing the crash and starting again

    attempt 2 of 7
    HEAD b6582c7 · site b6582c7 · pane b6582c7 · deck 1 slide(s)
    host answered in 6ms
    NOT READY — playwright-cli could not be run — nothing below was actually measured.

Every value on that sheet was measured, by the tool it says could not be run,
one line above the refusal. Exit code 0 while doing it.

**Root cause, reproduced away from the round rather than guessed at.** Running
`collectCrashEvidence` directly against the live browser printed it in one pass:

    spawn errors: 2
      - requests => spawnSync …\node.exe ENOBUFS

`sh("requests")` on a live PowerPoint tab is bigger than `spawnSync`'s 1 MiB
default `maxBuffer`. Over the line Node throws the output away and returns an
`error`, and `cli` read any `error` as "the tool could not be run".

Three defects, all fixed here, each with a mutation-proven guard:

1. **The buffer.** `maxBuffer: 64e6`. The default is invisible — nothing about
   the call says 1 MiB, and the failure names the executable rather than the size.
2. **An overflow is not an unreachable tool.** They are opposite facts: one means
   the browser answered and the answer did not fit, the other means nobody asked.
   `isOverflow` splits them.
3. **The doubt outlived its sweep.** `unreachable` is deliberately sticky within
   one readiness sweep and was sticky for the whole PROCESS, so a failure during
   attempt 1's crash handling refused attempt 2 forever. `sh.startSweep()` at the
   top of each attempt. This is the round-29 poll bug — "do not end a round on one
   failed CLI call" — one call site further out, unswept.

**And the damage was worse than two lost attempts: it emptied both crash
reports.** They read

    ## Document channel, last 12
    (nothing)
    ## PowerPoint's own error log (ULS)
    (no fatal entry in the last 0 telemetry batches — the quiet form of the wedge
    looks like this)

Nothing was measured, and the second sentence names a diagnosis. Round 043's
report from the same crash carried twelve real channel entries and a ULS window.
`collectCrashEvidence` now reports a failed read AS a failed read; verified
against the real browser, the report went 1834 → 3167 characters.

**Why 043 survived this and 044 could not**, which is the part worth carrying:
the overflow is a function of TAB AGE. Round 043's crash was collected at request
768 of a freshly opened tab and fitted; 044's third attempt crashed at request
**16748** of a tab that had been alive for two hours. So the bug only bites a
long-lived tab — which is to say, only during a round LOOP, which is the only
thing this driver is for. A defect that cannot reproduce on the first run of the
day is worth naming as such.

### Round 044, third attempt — b6582c7 — archived as 040 — 10/12

Attempt 1 crashed 2s in as usual; attempt 2 reached `ready` and ran, which is the
driver fix above doing its job on its first outing.

1. **Mine.** Two things, and the second is the one that closes an argument.

   **`survives-8` replicated a third time** — three passes here, `stable: true`,
   all in the `collection-refused` regime, on a third build. Eight samples across
   three builds now say a creation handle nobody resolved does not age out.

   **`tag-through-refetched-shape` is NEW and answers `no-id`, stable ×3** —
   "the fresh shape would not report an id". Taken with the above, the last
   alternative to the ordering change is gone: you cannot re-fetch the tag target
   by id, because on this host a fresh shape has no id to re-fetch it BY. Not
   "the id is refused" — there is no id. Route by route: `refreshed` is resolved
   and refused, `by-id` cannot be built, `created` works and is the one the pass
   throws away. The ordering change is not the best option left, it is the only
   one.

   Otherwise identical to 039: 10 of 12, the same two failures, `same scale` 4 of
   8 with the host flipping at chart 5, 37 draws and ZERO stalls, slowest batch
   18.4s, scratch handback 82 of 83 never landed.

2. **Research.** None needed; the one unexplained thing this round is the 2s
   crash, and the host now explains it itself — see below.

3. **Instrument.** The crash report is the instrument, and it worked the moment
   it was fixed. From `crashes/2026-08-15T12-49-04.md`, PowerPoint's own log:

       319 | In OnDisconnect(), setting SlideViewNode.srcSlide to null
       321 | Failed to restore selection after load content.
       389 | OnServerFindSucceeded could not find target slide, time elapsed: 430 ms
       390 | GlobalErrorHandler:DisplayErrorDialog: 5341289

   **PowerPoint has now crashed 2s into the FIRST attempt four times running**
   (043, and all three 044 attempts) and never into an attempt that followed a
   recovery. The difference between them is ~80 seconds of reload and settling.
   The host's own account is that it could not find the target slide after
   restoring content — a document still settling when the round's first Office.js
   call lands. That is a theory with four samples and no test.

   **Built, and it is an experiment rather than a fix.** `--check` pinged the
   host but the ping touches no slide — `getCount` is a count, and this host
   answers it in single-digit ms while in exactly the state it then dies from.
   The pre-flight now resolves slide 1 and the check line carries `slide 1
   resolved` or `slide 1 REFUSED` every round, pass or fail, because the rounds
   where it says `resolved` are what make the others mean anything. The point is
   to MOVE the crash, not prevent it: trip it in a two-second check that
   `--retry` recovers from, instead of an attempt plus 80s. If it answers
   `resolved` and the round crashes anyway, the theory is refuted for the price
   of one call — which is the outcome worth having either way.

   The crash dialog is now re-read AFTER the slide touch as well as before it. A
   crash the check itself provoked would otherwise be carried into the round as
   `ready`.

4. **Fix.** The three driver defects above.

5. **Doctrine.** `docs/BACKLOG.md` ordering item records the closed route.

### Round 045 — ca866e3 — archived as 041 — 10/12, and the first clean start in five

    host answered in 4ms · slide 1 resolved
    ready
    running — this takes about ten minutes
    finished

**No crash, on the first attempt.** Rounds 043 and 044 crashed 2s into every
first attempt — four for four — and this is the first that did not.

**What that is worth, stated before it gets quoted as more.** n=1, and the
experiment is confounded BY DESIGN: the check resolves a slide, so a clean run is
equally consistent with "the touch settled the host" and with "the host was fine
today". The reading that would have been decisive is the other one — `resolved`
followed by a crash anyway would have refuted the theory outright — and it did
not happen. So this is one point of weak support for a theory that still has no
test, and the brief already says to record the word every round for exactly this
reason. Three or four more clean starts make it an effect; one does not.

1. **Mine.** Everything else is the fourth consecutive repeat, which is itself
   the finding — this host is now boringly reproducible:
   - `survives-8`, stable, on a fourth build. Eleven samples.
   - `tag-through-refetched-shape: no-id`, stable, replicated.
   - Every tag failure `from: created×1`. Four rounds, not one `refreshed`,
     `group` or `by-id`.
   - `same scale` 4 of 8, host flips at chart 5 of 8. 16 of 16 failures.
   - 37 draws, ZERO stalls, slowest batch 16.9s, one population.

   **Two things that are NOT repeats, both in the deck rather than the trace:**
   - **One added slide came back blank and the rasterise agrees** (7 added, 6
     carry shapes). Rounds 039 and 040 both had zero. One witness plus its
     picture is the standard this repo set for calling a slide empty, and this
     one meets it.
   - **Two of the seven new slide ids are `256#0` and `257#0`** against
     `288#3603562595` and friends for the rest. An id whose second half is `0` is
     not the shape this host gives a slide it has finished adding, and the round
     also reports `delete-by-id left slides behind and this host does not list the
     ids they were added under`. Those two facts are probably the same fact.

2. **Research.** Nothing unexplained beyond the `#0` ids, which are worth a
   search of the office-js tracker next round rather than a guess now.

3. **Instrument.** The next question is already framed: does a slide whose id
   reads `NNN#0` correspond to the blank one? The round log carries both halves —
   `deck.inventory` and `deck.newSlides` — but nothing joins them, so it took a
   hand query to notice. Joining them in `triage` is the cheap next instrument.

4. **Fix.** None. Nothing this round exposed is fixable without first knowing
   whether the `#0` id and the blank slide are the same event.

5. **Doctrine.** The brief's "first thing to read" now points at the experiment's
   own word, and this entry records what a single clean start does and does not
   license.

### Round 046 — a54401c — archived as 042 — 10/12, and it settled the ordering fix

`slide 1 resolved`, clean first attempt again — two in a row now, still not an
effect.

Getting there needed a hand recovery worth recording: `pwclean` answered
`deck-failed` and `--check` said why — the QUIET wedge, host silent for 8016ms
with no dialog, plus a pane still on `ca866e3` and an eight-slide deck. One
reload cleared all three. That is `recover`'s own sequence, done by hand, because
the driver would not do it — see the fix below.

1. **Mine. `collection-read-poisons-the-creation-handle` answers `yes`** — three
   passes, `stable: true`, every one taken while the host was refusing collection
   reads. **The pre-grouping re-read does NOT poison the creation handle.** The
   fake was wrong, the ordering fix is viable, and it was rebuilt the same hour.

   Otherwise the fifth consecutive repeat: `survives-8`, `tag-through-refetched-shape:
   no-id`, every tag failure `from: created×1`, `same scale` 4 of 8 flipping at
   chart 5, 30 draws with zero stalls. One added slide blank again, and the same
   two `#0` ids as round 041 — `256#0` and `257#0`, identical strings across two
   rounds, so they are not per-round noise.

2. **Research.** None owed; the one open question was put to the host and
   answered.

3. **Instrument.** `triage` now names added slides whose id ends `#0` and joins
   them to the confirmed blanks, because round 041's finding took a hand query
   and nearly went unseen. **It earned itself immediately, in the negative
   direction: 0 of the odd-id slides are the blank one, in both rounds.** The
   coincidence is dead and nobody spends a round on it.

4. **Fix.** Two, neither of them the renderer:
   - The driver stopped on states it knows how to fix. `shouldRetry` retried a
     crash and nothing else, so the quiet wedge (`silent`) ended a round and a
     `not-ready` was never retried at all — which is why the recovery above was
     done by hand. Every readiness stop now carries a CODE, and a check is
     retried when recovery addresses all of them and refused when even one stop
     it cannot touch is present (`site-behind`, `verbose-off`).
   - `origin tag lost` is its own trace line. Losing it costs drag tracking, not
     re-editability, and reporting it as "charts are not re-editable until
     repaired" was about to become the common case.

5. **Doctrine.** `docs/BACKLOG.md` records the answer and what it licenses.

### Round 047 — d2ca1c9 — archived as 043 — 10/12, and it does NOT validate the fix

The first round carrying the tag-anchor change, and the honest answer is that one
round cannot tell. Recorded before anything else because the temptation to read
it as a win was real.

**The driver fix worked, visibly.** Attempt 1 refused on `deck-dirty` alone,
retried, `recover` cleaned the deck 8 → 1, attempt 2 ran and finished. Under the
old rule that was a dead stop, and the two rounds before it were hand-recovered.
(Wart: it still prints "clearing the crash and starting again" when nothing
crashed. The message predates the change.)

**The renderer fix changed nothing measurable, and here is why that is not a
verdict.** Fault counts, three consecutive rounds:

    round  build     tags-undefined  cfg-tag-5010  group-5010  no-queue  tagging-failed
    041    ca866e3        1               6            1          1           4
    042    a54401c        5               6            5          5           8
    043    d2ca1c9        5               6            5          5           8

**042 and 043 are identical, and 041 → 042 is a five-fold jump with NO renderer
change between them** — a54401c added a probe and documentation and nothing else.
So the counts are dominated by the host's regime, they swing by 5× on their own,
and a single round comparing two builds is measuring mood. `same scale` moving
from 4-of-8 to 3-of-8 sits inside that same noise.

What IS informative: the config write still fails six times, exactly as before.
If the anchor were reaching the write unresolved, that number should have moved.
And `origin tag lost` — the line that exists for the case where the config lands
and only the origin fails — fired ZERO times, which says the same thing from the
other side.

**So the fix is unvalidated, not refuted, and the next hypothesis is better than
the last one.** `addGroup` fails 5× with 5010 in the same pass, and a failed sync
poisons its own context — this project already knows that and works around it
elsewhere. `no chart's tag could be queued` firing 5× is a context-level symptom,
not a handle-level one: the QUEUE is refusing, before any handle is exercised. If
the grouping attempt poisons the context the tag write sits in, then no choice of
handle can help and the whole ordering effort has been aimed one level too low.

That is testable and cheap: tag in a context that has not just tried to group.

**Instrument owed, and it is the lesson of this entry:** the rasterise arms are
pooled across rounds because one round could never answer them, and the tag
failure counts need exactly the same treatment. Comparing two builds by eye in a
regime that swings 5× is how a fix gets declared good. `npm run rounds` should
pool tag failures per build the way it already pools draws per arm.

Probes: `which-end-a-short-read-drops` answered `unreadable` — "the collection
would not list its items", which is this host's usual answer to anything asking a
collection to enumerate itself. The trade it prices stays unpriced; the question
is right and the host is not answering it yet.
