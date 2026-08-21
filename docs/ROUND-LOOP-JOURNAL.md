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

### Round 048 — 355a37c — archived as 044 — 10/12, and the question missed

**The driver recovered the whole three-stop state on its own**, which is what
this morning cost three hand-recoveries:

    NOT READY — a silent host, a pane on d2ca1c9 while the site served 355a37c,
                and a seven-slide deck
    recovering from a silent host and a stale pane and a dirty deck, then starting again
    attempt 2 of 7 — ready

Both fixes from the last two commits doing their job on their first outing, and
the recovery line naming all three rather than announcing a crash.

1. **Mine. `does-a-failed-group-poison-the-tag` answered `no-refusal`, three
   times** — "the host grouped, so the question was never put". The probe was
   honest about missing rather than inventing an answer, and the miss is itself
   the finding: **this host groups two fresh shapes from one batch perfectly
   well.** What it refuses is PRODUCTION's grouping, whose members and slide
   handle are several syncs old by the time `addGroup` is called. Grouping is not
   refused as such; it is refused for aged handles.

   The rest is the sixth consecutive repeat, and the new table says so at a
   glance — `355a37c` scores 5/6/5/5/8, identical to `a54401c` and `d2ca1c9`.
   Three builds, one of them carrying the tag-anchor change, all indistinguishable
   against a noise floor of 1-vs-5 measured within a single build.

2. **Research.** None owed.

3. **Instrument.** The question now ages its handles two syncs before grouping,
   which is what it should have done first time. **And it comes with a warning
   that the fake found:** under `strictTags` the aged version answers
   `refused-after-group` for a reason that has nothing to do with grouping — the
   age rule refuses the write on its own. On a real host the same word could mean
   either thing, so it must be read beside `tag-the-creation-proxy-a-sync-later`,
   which answers `yes` here and excludes age. A question that can be right for
   the wrong reason is worth shipping only with the control named next to it.

4. **Fix.** None owed; the renderer is where it was.

5. **Doctrine.** `PENDING_QUESTIONS` carries the miss, the reason for the ageing,
   and the read-it-with-the-control rule.

## The long run — 2026-08-15 evening

### Pair 1 — 56eb477 × 2 — archived as 045 and 046

**THE FIRST SAME-BUILD PAIR, and it settles the tag-anchor question.** Both
rounds scored `tags-undefined 5 · cfg-tag-5010 6 · group-5010 5 · no-queue 5 ·
tagging-failed 8` — **identical, zero spread.** Four consecutive builds now score
the same, and one of them predates `tagAnchorIndex`:

    a54401c  5 6 5 5 8   ← before the anchor change
    d2ca1c9  5 6 5 5 8   ← anchor change lands
    355a37c  5 6 5 5 8
    56eb477  5 6 5 5 8   (twice, no spread)

Five rounds, four builds, no within-build variance on the current one. **The
anchor change did nothing.** That is not the ambiguous single-round reading of
043; it is a measurement with its own noise floor attached. Whether to revert it
is the owner's — it is merged renderer code.

1. **Mine — and the finding did not need a round at all.** Joining each chart's
   grouping outcome to its tag outcome, across the whole archive:

       grouped      64 chart(s),  1 lost the tag =  2%
       NOT grouped  62 chart(s), 41 lost the tag = 66%

   Per round it is nearly mechanical: three grouped and none lost, two or three
   ungrouped and two lost, eleven rounds running. **A chart that gets grouped
   keeps its config; one that cannot loses it two times in three.** When grouping
   succeeds the tag goes onto the GROUP — a handle made in that batch, never
   resolved — and it lands. When grouping is skipped the tag falls back to a
   `created` handle and is refused.

   So the entire handle question — four rounds, a merged renderer change, a fake
   corrected, three probes — has been about the LOSING path. It was aimed one
   level too low, and the evidence had been in the archive for eleven rounds
   unqueried. `npm run rounds` prints it every round now.

   `not grouping: no member handle this host will accept` carries `refreshed: 0`
   every time, so what decides a chart's config is whether the PRE-GROUPING
   RE-READ returned anything. That is where the next attempt belongs.

2. **Research.** None owed; the archive answered it.

3. **Instrument.** The three lines that decide a chart's fate now carry the
   slide (`not grouping`, `a chart's tag could not even be queued`, `tagging
   failed`). Rounds 043-045 each lost a chart's config and the one line that
   names a slide named `257#0` every time — an id whose second half is `0`, which
   is not the shape this host gives a slide it has finished adding. Next round
   says outright whether the charts that lose their config are the ones on those
   slides. If they are, this is one SLIDE fault rather than three tag faults.

4. **Fix — four, and the run exposed every one of them.**
   - **A mid-round wedge was not retryable.** Round B timed out at thirty
     minutes and the driver stopped; in an unattended run that is nine idle
     hours. `timeout` now recovers like a crash, bounded by `--retry`.
   - **Crash reports were STILL empty**, hours after the overflow fix, and this
     one was mine: `ask()` honoured `lastError` (spawn, overflow) but not
     `lastFailed` (ran, exited non-zero) — which is exactly how a wedged tab
     fails. Three sections read "(nothing)" over three failed reads.
   - **A dead browser read as a closed pane.** The wedge killed the browser
     process; `recover` then reloaded and reopened a pane inside a window that
     did not exist, SEVEN times, while the check asked "is the add-in open?"
     `noBrowser` names it, and recovery now reopens from the persistent profile
     — no password, because a dead browser is not a lost sign-in.
   - **`npm run rounds` reported the oldest round** in the archive rather than
     the newest.

5. **Doctrine.** `docs/BACKLOG.md` now LEADS with grouping-saves-the-config, and
   says in as many words that the three sections under it are about the losing
   path.

**Cost of the pair:** one wedge, one dead browser, seventy-one slides of litter
to clean. All four fixes above came out of it, which is the trade this loop
exists to make.

### Pair 2 — 5856d1d × 2 — archived as 047 and 048

**The `NNN#0` slide lead is dead, and it died from the archive rather than from
the instrument built for it.** The draw trace already carried `chart` and
`onSlideKey`, so the join was available on rounds already filed:

    1/8  slide 256#0            GROUPED   config kept    <- a #0 slide
    2/8  slide 256#0            GROUPED   config kept    <- a #0 slide
    3/8  slide 256#0            GROUPED   config kept    <- a #0 slide
    4/8  slide 257#0            no group  CONFIG LOST    <- a #0 slide
    5/8  slide 288#1168146411   no group  CONFIG LOST    <- an ordinary id

Identical in 043-047. A `#0` slide carries the three best-behaved charts AND a
failure; an ordinary slide fails too. **The id shape predicts nothing**, and the
instrument shipped this build confirmed it independently: every failure line in
047 named `291#567483212` and `288#230020768`, both ordinary.

**What does predict it is POSITION IN THE DECK-WIDE UPDATE**, and six rounds now
show the same decay curve rather than a coin:

    charts 1-3   re-read matches all 24     grouped, config kept
    chart 4      re-read matches 20 of 24   partial, thrown away, config lost
    chart 5      re-read returns NOTHING    config lost
    charts 6-8   never attempted

`same scale across the deck` says it itself — "the host flipped at chart 4 of 8"
— seventeen rounds running.

1. **Mine.** Grouping-saves-the-config strengthened with the pair: **73 grouped,
   1 lost (1%); 68 ungrouped, 47 lost (69%)**.
2. **Research.** None owed.
3. **Instrument.** `how-many-collection-reads-a-context-survives` — how many
   times one context can re-read a slide's shapes before the answer goes short.
   That is the mechanism the decay curve implies, asked directly.
4. **Fix.** A comment, and it was asserting something false: the partial-match
   branch claimed an ungrouped chart "is still tagged, still re-editable". The
   69% above refutes it. **The branch itself is deliberately left alone** — the
   alternative strands shapes, and choosing between two harms is not a call to
   make from a trace.
5. **Doctrine.** `docs/BACKLOG.md` carries the refuted lead, the decay curve, and
   the experiment with its prediction.

**The suspect is our own perf work.** `updateChartsInSlides` was deliberately made
ONE context, four syncs, flat in N — the right shape for a host that can hold a
context, and this one appears to degrade as a context is used. `#112` already made
the opposite call for the demo deck. **Not acted on**: the fix is a 390-line
restructure of the live update path, and this project has three shipped-broken
fixes on record from changing that path on a theory. The probe decides it for the
price of one round.

### Pair 3 — 3f04df2 × 2 — archived as 049 and 050

**`how-many-collection-reads-a-context-survives` answered `unreadable-at-1`**,
three samples, stable: the scratch slide would not enumerate its collection even
ONCE, so there was never a baseline to degrade from. The question was built to
decide the strongest lead this project has, and it could not be put.

**That is the third question to die on the same harness limit** —
`shapes-items-count-honest` and `which-end-a-short-read-drops` are the other two,
both `unreadable` — and three is a pattern rather than three misses.
**The finding is about the harness, not the host: this scratch slide is strictly
worse at collection reads than a real one**, whose collection enumerates fine for
the first three charts of a deck update, every round. Collection questions cannot
be asked from a probe here. They have to be instrumented in production.

1. **Mine.** Grouping-saves-the-config, pooled over 26 rounds: **79 grouped, 1
   lost (1%); 72 ungrouped, 51 lost (71%)**. Six rounds tonight and the effect
   has only sharpened.
2. **Research.** None owed.
3. **Instrument — moved from the probe into the production path.** `boundedSync`
   now counts syncs per context (a WeakMap, so the count dies with the context),
   and the two re-read failure lines carry `contextSyncs`. The deck-wide update
   runs every chart through ONE context and the re-read decays chart by chart, so
   this is the x-axis the decay curve has never had. Next round says at which
   sync of a context the re-read starts going short — which is exactly what the
   probe could not ask.
4. **Fix.** None owed; the renderer is where it was, deliberately.
5. **Doctrine.** `PENDING_QUESTIONS` records the harness limit and tells the next
   reader not to wait on that question.

**A repair worth recording, because it was mine and the tables caught it.** The
patch that wrote the finding into `PENDING_QUESTIONS` hit `FAKE_BASELINE`
instead, replacing an answer with prose and inventing a
`…-ORIGINAL` key for the value it displaced. `host-contract.test.ts` failed on a
baseline entry for a question that does not exist — which is precisely the check
that exists so a table cannot quietly drift from the probes it describes.

### Pair 4 — 1a1e05f × 2 — archived as 051 and 052 — THE ROOT

**`contextSyncs` answered on its first outing, and it refuted the hypothesis it
was built for.** Both failing charts show `contextSyncs=1`: the re-read that goes
short is the FIRST sync of its context, not the thirtieth. The context is fresh.
**Chunking `updateChartsInSlides` would change nothing**, and that 390-line
restructure is ruled out before anyone started it — which is exactly what the
instrument was for.

**Then the same trace said what DOES separate them**, and it is a switch:

    chart 1/8  slide 256#0           32 shapes already on it  -> re-read OK, GROUPED
    chart 2/8  slide 256#0           40 shapes already on it  -> re-read OK, GROUPED
    chart 3/8  slide 256#0           40 shapes already on it  -> re-read OK, GROUPED
    chart 4/8  slide 257#0            0 shapes already on it  -> re-read SHORT 20/24
    chart 5/8  slide 288#807648421    0 shapes already on it  -> re-read EMPTY

Pooled over the whole archive:

    slide already had shapes   82 chart(s), 81 grouped = 99%
    freshly added, empty       74 chart(s),  1 grouped =  1%

**A chart drawn onto a slide this run has just added does not get grouped, and a
chart that is not grouped loses its config.** That is the root, and everything
this project has done to the tag path for four rounds sits one level above it.

It is also not new. `shape-add-held-slide-proxy` answers `threw`, a web-new-slide
id does not round-trip, and the #108-#111 saga was four attempts at drawing on a
freshly added slide. The evidence has been in every round since; nobody had
joined `onSlide` to the grouping outcome.

1. **Mine.** As above. `256#0` — the id flagged as malformed two pairs ago — is
   the slide that WORKS, because it is the established one.
2. **Research.** None owed; the archive answered it.
3. **Instrument.** `npm run rounds` prints **WHICH SLIDE THE CHART LANDED ON**
   every round now.
4. **Fix.** None shipped, deliberately. The matcher already knows the honest rule
   — "the positional rule is still right for a slide this run added blank" — but
   that branch is reachable only when NOTHING matched, so a chart matching 20 of
   24 falls past it. On a slide this run added blank we KNOW our shapes are the
   only ones there, so a short read is a host lie we can detect rather than obey.
   Contained, but still surgery on the grouping path, which carries three
   shipped-broken fixes on record. It wants a person awake.
5. **Doctrine.** `docs/BACKLOG.md` now leads with the root and lists what was
   ruled out: context wear, the `#0` ids, and the poisoned-context theory.

**Two of tonight's driver fixes proved themselves in this pair.** Round B wedged
at thirty minutes and the driver recovered and finished on attempt 3 — before
tonight that was a dead stop. And the crash report from that wedge reads "(could
not read — the call could not be run)" three times rather than "(nothing)",
which is the `lastFailed` fix refusing to report an empty host as evidence.

### Pair 5 — a004711 × 2 — archived as 053 and 054 — a question CLOSES

**`does a rasterise poison the next draw` is answered: NO.**

    after a rasterise     0 stalled /  60 drawn = 0.0%
    after a cheap read    0 stalled /  60 drawn = 0.0%

Sixty draws per arm is the bar the tool itself set, and the "NOT an answer yet"
line is gone. Open since round 33, counterbalanced the whole way, and it closes
in the negative — a rasterise is not what makes a draw stall. Nothing to fix,
which is the point: the suspicion is retired rather than carried.

The root cause holds and sharpens with two more rounds:

    slide already had shapes   91 chart(s), 90 grouped = 99%
    freshly added, empty       80 chart(s),  1 grouped =  1%

**A slip of mine worth recording.** Round B was launched with a shell `&`
instead of the harness's background mechanism, so its console output was lost and
no completion signal existed. The process survived and the round completed — the
run log and the archive are what matter, and both were intact — but the recovery
was luck rather than design. A waiter that watches for the process to exit
restored the signal without touching playwright-cli. Launch a round the same way
every time.

### Pair 6 — 61f431b × 2 — archived as 055 and 056

Accumulation, and the two things worth writing down are both about what has
STOPPED happening.

**The 2s crash has not recurred since the pre-flight shipped.** Four crash
reports say "2s into a round" and all four are from a forty-five-minute window
before `slideResolveScript` existed. In the twelve rounds since, `slide 1
resolved` has printed on every check and not one round has crashed 2s in:

    before the pre-flight   4 first attempts, 4 crashed
    after                  12 rounds,        0 crashed

**Observational, not counterbalanced**, and it must be read that way: the
pre-flight resolves a slide, which may settle the host rather than merely
measure it, and the host's mood is not controlled. But four-in-four against
zero-in-twelve is the strongest thing this experiment can say without an arm that
deliberately skips the touch — and skipping it would cost a round to learn
something the loop already gets for free.

**The root cause is unmoved by more data**, which is what a real effect does:

    slide already had shapes   97 chart(s), 96 grouped = 99%
    freshly added, empty       84 chart(s),  1 grouped =  1%

Nothing else new. `same scale` failed the same way, the same charts, in the same
order — which after twelve rounds is itself the finding: this host is
deterministic here, and the remaining work is a decision rather than a
measurement.

### Pair 7 — 2a11873 × 2 — archived as 057 and 058

Round B wedged at thirty minutes and the driver recovered and finished on attempt
3 — the second time tonight that fix has saved a round which, this morning, would
have been a dead stop.

    slide already had shapes  103 chart(s), 102 grouped = 99%
    freshly added, empty       88 chart(s),   1 grouped =  1%

Fourteen rounds and the ratio has not moved. Nothing new from the host.

**The work this pair actually bought was in the documents, not the rounds.**

- **`chooseGroupMembers`'s contract carried the same refuted premise** the
  partial-match branch did — "strictly worse than an ungrouped chart that is
  still re-editable". The conclusion survives (a throw costs every chart in the
  batch, grouping nothing costs one) but the premise does not, so it now says
  plainly that this is a choice between two losses and that the way out is
  upstream. Two other candidates in `PUBLISHING.md` were checked and are sound —
  both describe charts the rescue actually grouped.
- **The brief's queue was stale in all five items**, every one answered during
  this run, and item 5 said something now false: "it moves when the tag path
  moves". It moves when the fresh-slide re-read is fixed. Rewritten around what
  is settled, with nine questions listed as answered and their results, so the
  next session cannot spend a round re-deriving any of them.

**And the honest note at the top of it:** the mechanism is settled and more
rounds do not add to it. What is left is two decisions, both the owner's. A loop
that keeps running past the point where measurement helps is manufacturing
activity, and saying so is worth more than another pair.

### Pair 8 — 85277f5 × 2 — archived as 059 and 060

    slide already had shapes  109 chart(s), 108 grouped = 99%
    freshly added, empty       92 chart(s),   1 grouped =  1%

Sixteen rounds, unmoved. Both rounds clean, no wedge, nothing new from the host.

**Doc drift caused by this run's own fixes, found and corrected:**

- **Three rotting counts.** `docs/BACKLOG.md` (twice) and `host-probe.ts` asserted
  `same scale` had failed "seventeen rounds running". It is 34 of 34 now, and
  would be wrong again tomorrow — so they say "every round it has run" and point
  at `npm run rounds` for the number. A count frozen in prose is a comment with a
  half-life.
- **`ROUNDS.md` described a driver that no longer exists.** Its `--check` list was
  missing the slide-resolve and no-browser refusals, and said nothing about the
  eight states `--retry` now recovers on its own — including two it deliberately
  does not, which is the more useful half. Rewritten against
  `RECOVERABLE_STOPS`, which is derived from what `recover` actually does.

**The audit is out-yielding the rounds, and that is the honest read of where this
is.** Two refuted premises and four stale statements in three pairs, against
sixteen rounds that agree with each other. The host has said what it has to say
about this mechanism; the documents had not caught up.

### Pair 9 — 9dc3bc8 × 2 — archived as 061 and 062

    slide already had shapes  115 chart(s), 114 grouped = 99%
    freshly added, empty       96 chart(s),   1 grouped =  1%

Eighteen rounds. Both clean, no wedge, nothing new.

**`CLAUDE.md` was the last document pointing at the wrong problem.** It said the
red here is "the shape collection dying part-way through", which sends a reader
looking at POSITION — and position is what the archive refuted. It now carries
the split above and names the freshly-added slide, with the note that four rounds
of tag-path work aimed one level above it before anyone joined `onSlide` to the
grouping outcome. That matters more than another round: `CLAUDE.md` is what a
fresh session and the parallel stream read first.

Four documents now agree with the measurement — `CLAUDE.md`, `docs/BACKLOG.md`,
this brief and `docs/ROUNDS.md` — where at the start of the night three of them
described the tag writer as the problem.

### Pair 10 — e68147f — round A archived as 063; round B stopped on SIGN-IN

**The run ends here, on the one state that is the owner's.** Round B wedged at
thirty minutes, the driver recovered as it now does, and attempt 3 found a
`login.live.com` popup open beside the deck — the session's token had expired.
The check named it exactly:

    the browser is on a Microsoft sign-in page — there is no document to run a
    round in. … None of that is the agent's to do: it needs a password.

Nothing was touched: the popup is left exactly as Office opened it, because
dismissing an authentication prompt is not a workaround and the owner may want to
complete that prompt as it stands. **The readiness work paid for itself one last
time** — this stopped in one attempt with the right diagnosis, where the same
class of state cost seven attempts and an "is the add-in open?" earlier tonight.

**A precision note for whoever reads this next:** `signedOut` fires on any tab
matching `login.live.com`, and what actually happened was a POPUP beside a deck
tab that was still open. Treating that as signed-out is the right call — if
Office is asking for credentials, a round cannot be trusted — but the message
says "the browser is on a sign-in page", which is not quite what a reader will
see on screen.

**Coverage the audit found, and it was mine.** Two of tonight's own fixes
shipped without guards, on paths a green suite could not have covered because
nothing tested them at all:

- The `lastFailed` half of the crash-forensics fix — verified in production on
  the 18:06 wedge and never pinned.
- `syncsOf` / `contextSyncs` — the number that ruled out a 390-line restructure.

Both guarded and mutation-proven now. That is the third class of defect this run
produced, after round findings and doc drift: **fixes of mine, unguarded.** Worth
checking for deliberately, because the suite was green before I looked.

## Round 064 (`bcd5773`, 2026-08-16) — the retry works, and it exposed the next link

**The prediction was staked before the round and it held.** Charts 4 and 5 of
`same scale across the deck` — the two that had failed identically for five
rounds running, one on a short re-read and one on an empty one — **both grouped
for the first time.** The trace is causal rather than statistical, which is why
one round is enough to believe this much of it:

    4/8   re-reading the slide's shapes again after a settle delay  waitedMs=1500
    4/8   grouped the chart's shapes                                by=ids
    5/8   re-reading the slide's shapes again after a settle delay  waitedMs=1500
    5/8   grouped the chart's shapes                                by=ids

`the settled retry repaired 2`, on the scenario's own verdict line. **No chart in
the round went ungrouped** — `NOT grouped 0`, against a pooled 98 charts at 79%
tag loss. Freshly-added slides went 2 of 2 grouped where the archive had 1 of 98.

**AND THE SCENARIO STILL FAILED**, 4 of 8. The chain did not break; it moved one
link along, and the new link could never have been seen before, because these
charts had never once got as far as having a group:

    4/8   tagging failed   from: group×1   slide 258#4111159134   5010
    5/8   tagging failed   from: group×1   slide 259#3844610554   5010

**THE DOCTRINE THIS CORRECTS — read it before quoting the 2% again.** This file
and `docs/BACKLOG.md` have said "grouping is what saves a config" on the strength
of:

    grouped      123 chart(s),   3 lost the tag = 2%
    NOT grouped   98 chart(s),  77 lost the tag = 79%

That number was measured on a population that **excluded freshly-added slides by
construction** — a fresh-slide chart could not group, so it could never appear in
the grouped column. The moment one does, it loses its tag anyway: this round's
own split is `grouped 5, 2 lost = 40%`. The rule was never "a group saves the
config"; it was "a slide that was already there saves the config", and grouping
was standing in for it.

**Why the group handle is refused too, and it is already written down from the
other side.** A shape proxy carries its PARENT's object path. The group is made
in the grouping batch, but its slide handle is by then rewritten to
`slides.getItem(id)` — and a freshly-added slide's id does not round-trip on this
host (`shape-add-held-slide-proxy: threw`, the #108-#111 saga). The round says so
in as many words: *the host grouped through a slide handle two syncs old, so the
refusal was never provoked.* The members were never too old. The PARENT was.

**So the next question is the slide handle, not the shape handle** — and unlike
every previous "next question" here, it is not a guess: the failure names the
slide, both times, and both are freshly added.

Two other tag failures in the round, neither new and neither on this path:
`from: created×1` on slide `260#1686107471` beside an empty settle re-read, and
five × `Cannot read properties of undefined (reading 'add')` on `262#3659873566`
— the documented `target.tags` guard.

**Not yet a pair.** The causal lines (retry → group, same chart) are strong
enough to stand alone; the counts are not, and `40%` in particular is five
charts. A same-build second round is what makes the tag-loss number mean
anything.

## Round 065 (`bcd5773`, 2026-08-16) — the pair, and it replicates exactly

Second round on the same build, run without a merge between the two, which is
what makes it a pair rather than two sheets. **It is structurally identical to
064, chart for chart:**

    1-3/8   grouped by=ids                             no retry needed
    4/8     settle delay 1500ms → grouped by=ids        tagging failed  from: group×1
    5/8     settle delay 1500ms → grouped by=ids        tagging failed  from: group×1
    260     from: created×1, beside an empty settle re-read
    262     5× no chart's tag could be queued

`the settled retry repaired 2` in **both** rounds. The slide ids differ between
them (`258#4111159134` vs `258#2150477121`), so what repeats is the POSITION and
the freshness, not an id.

**This is the strongest replication this project has.** The noise floor — 1
versus 5 for the same fault with nothing changed — is about counts. These are not
counts: the same two charts take the same retry, group, and are refused through
the same handle kind, twice. Nothing here needs a third round.

**What still moved between the two, and it is the downstream half:** `same scale`
scored 4 of 8 then 3 of 8, and re-editable went 7 then 6. That is the part that
was always noisy and it stays noisy. The mechanism is deterministic; the score is
not, and reading the score as the result is the mistake this pair exists to
prevent.

**So both conclusions from 064 stand on two rounds now:**

1. The settled retry works, deterministically, on exactly the charts it was built
   for.
2. `from: group×1` is refused on a freshly-added slide — so **grouping does not
   save a config**, and the old 2%-vs-79% split was an artifact of a population
   that could not contain these charts.

**The next thing to build is the slide handle**, and it now has two witnesses
naming it rather than one.

### A numbering trap, found by walking into it

The second round archived as `064` as well. **Nothing was overwritten and no
evidence was lost** — `nextRoundNumber` is max+1, so it can never land on a file
it can see. What happened is subtler: round 064's file was committed to a branch,
`git checkout main` removed it from the working tree, and archiving from `main` —
where the directory ends at 063 — reissued the number.

The damage is not destruction, it is two different rounds both called 064 in two
git contexts: a merge collision, and a pooled report that reads one of them twice
or not at all. Every number in `npm run rounds` is a pool over that directory.

`--archive` now numbers from the union of the working tree and `git log --all`,
so a round committed on an unmerged branch still holds its number. Falling back
to the directory alone when git is unavailable, because refusing to archive a
real round over a numbering nicety is the worse trade.

**Worth recording how the wrong diagnosis was caught:** the first guard written
for this was "refuse to overwrite an existing file", and its test would not go
red. That is what proved the overwrite theory wrong — `nextRoundNumber` cannot
produce a colliding name. A guard that cannot fail is evidence about the
diagnosis, not a spare safety net.

## Rounds 066 + 067 (`d8ba7df`, 2026-08-16) — the prediction failed, and the instrument was wrong too

**The staked prediction was: single-batch grouping moves off 24% toward the
multi-batch 100%, and `tags-undefined` falls with it. Half of that happened, and
the half that did not is the more useful half.**

The pair is identical on every number, which is what makes it worth reading:

    round   1-batch draws  grouped  not-grouping  group-5010  tags-undefined  cfg-tag-5010
    065           5           0          0            5             5              3
    066           5           0          5            0             0              8
    067           5           0          5            0             0              8

**Grouping did not improve. It was already zero.** Widening `refreshShapes` did
not turn single-batch charts into grouped charts; it turned a doomed `addGroup`
into an honest decline. Ten spurious errors a round — five 5010s and five
`tags-undefined` — became five `not grouping` lines. `cfg-tag-5010` rose from 3
to 8 because those five charts now reach the tag write and are refused there
instead of failing before it. **For the user the outcome is unchanged; for anyone
reading a round it is much clearer.**

**Why it did not work, which is the next thread.** The single-batch charts show
**no settle-delay line**. Their re-read did not come back empty — it came back
with items that matched none of our ids, and the retry only fires on an empty or
partial read. A zero-match falls straight through to `use: "none"`.

A single-batch chart loads its shape ids in the same sync that creates them and
re-reads on the very next one, so the ids are plausibly not resolvable yet —
which is the same settling story that DID work for charts 4 and 5. **Retrying on
a zero-match is the obvious next experiment**, and it is small.

### THE 100% WAS MINE, AND IT WAS WRONG

`poolBatchSpanVsGroup` reported `333 of 333 = 100%` for the multi-batch arm. That
number is a bug in the pooling function, not a fact about the host. It looked for
the grouping verdict within a fixed few trace entries and did not count
`not grouping` among the outcomes — so **every honest decline was dropped from
both arms**, and only draws that either grouped or threw were counted at all.

Corrected, the archive says **353 of 452 (78%) against 49 of 214 (23%)**. The
separation is real and still the sharpest thing here; the absolutes were not.

Two things made it worse than an ordinary slip. It was quoted into a commit
message, a PR body, `docs/BACKLOG.md`, a code comment and two tests before anyone
checked it. And the re-read this very change added pushed the verdict further
from the draw, so the report showed **zero single-batch draws** in round 066 — the
instrument went blind in exactly the place it was being used to measure.

It was caught because the number was implausible, not because anything failed.
That is the same lesson as the starved-questions sweep, from the other side:
**pool your instrument's own behaviour, and distrust a clean 100%.**

## Rounds 068 + 069 (`2c7dcd8`, 2026-08-16) — the pair did NOT replicate, and that is the result

**Reading 068 alone would have shipped a false conclusion**, and this pair is the
case for the discipline added the same day.

    round   1-batch draws  grouped  not-grouping  retries  unmatched(traced)
    067           5           0          5           2           0
    068           5           4          0           7           4
    069           5           1          4           7           5

**068 read as a clean win: 0 to 4 of 5.** 069, same build, no merge between,
scored **1 of 5**. The retry fired identically in both (7), and the zero-match
persisted in both (4 and 5) — what moved is whether the charts grouped at all.

**One instrumented field explains the whole spread.** `listed` — how many shapes
the host named — on the traces that survived the retry:

    068    listed 10, 9, 16, 17     all >= drawn  ->  positional fallback fires  ->  4 grouped
    069    listed  1, 1, 1, 1, 31   four short    ->  fallback cannot fire       ->  4 declined

Four reads returned **ONE shape** for a slide holding seven to nine, and every
one of them carries `afterRetry: true` — the host read that short *after* a
1.5-second settle. More waiting is not the answer.

**So the mechanism is now fully named, and it is not what the fix assumed:**

1. The pre-grouping re-read **never matches our ids** on these charts.
   `withOwnId` is 7 of 7 and 9 of 9, so our handles are fine — the host lists the
   slide under ids that are not the ones it gave us at creation.
2. Grouping therefore depends entirely on the **positional fallback**, which is a
   guess ("the last N on the slide") and only legal when nothing matched.
3. That fallback needs `listed >= drawn`, and whether the host lists enough is
   **moody**: 9-17 one round, 1 the next.

**What the change is actually worth.** The old build grouped these charts 0 times
in two rounds (066, 067). The new build grouped them 4 and 1. That is a real
improvement — it never happened before and now sometimes does — but it is
unreliable, and it arrives through a positional guess rather than through the id
match the retry was meant to repair. **Recorded as partial, not as a win.**

Scenario level is unchanged in both: 10 of 12, `same scale` 4 of 8, and `the
chart is actually visible` still ends "the host would not name the chart
afterwards, so it carries no config". Grouping improved; the config did not
follow, which is consistent with the earlier correction that grouping is not what
saves it.

### The web search, recorded including its null result

No upstream issue reports this symptom — a collection listing shapes under ids
that do not match the ones returned at creation. The near relatives are all
already cited here: #5022 (sync hangs after add-then-read; delay workaround,
applied), #2903 (applied), #6498 (inserts not reflected instantly), #6363
(properties unavailable after sync, ten failed attempts). **Written down so the
next session does not spend the same twenty minutes finding the same absence.**

**The internal parallel is stronger than any of them.** This project already
measured the SLIDE version: scratch slides came back as `4123571114#123571113`
while the deck listed `287#62081387` — a freshly added slide reporting an id in a
namespace the deck will not answer to. Round 068/069's shape traces are the same
signature one level down.

**And the search surfaced something worth reopening**: Microsoft's shape-BINDING
documentation, unprompted, in two of three searches. `bindings.add` takes the
live Shape proxy inside the batch that created it — no id round trip, no
collection read, which is exactly the two things failing here. This repo parked
that route as *unanswerable from the probe* (`binding-names-shape-later` asked
eight times, never once reached its own question). That was a fact about the
PROBE. The retry proved the other route: measure it in production, where the
question can actually be asked.

## Rounds 070 + 071 + 072 (`01f3607`, 2026-08-16) — `same scale` passes, and not for the reason it was built

**THE SCENARIO THAT HAD FAILED 35 ROUNDS OUT OF 35 NOW PASSES, THREE TIMES:**

    070   12 of 12 scenarios   same scale 8 of 8, all 8 re-editable
    071   12 of 12 scenarios   same scale 8 of 8, all 8 re-editable
    072   12 of 12 scenarios   same scale 8 of 8, all 8 re-editable

`explode a degraded picture` passes in all three as well, and `the chart is
actually visible` has lost its "the host would not name the chart afterwards"
caveat. Twelve of twelve is the first clean sheet this project has recorded.

**THE FIX WORKED FOR A REASON IT WAS NOT DESIGNED FOR, and that is worth more
than the win.** The binding was added to give the SETTLE a durable handle. The
settle route has never once executed:

    round   grouped  notG  tagFail  cfg5010  origin5010  bind ok/fail  retries  unmatched
    069        11      4       8       8         0          0 / 0         7         5
    070        18      1       0       0         7          0 / 0        10         0
    071        19      0       0       0         7          0 / 0        10         0
    072        14      4       1       0         3          0 / 0        10         3

`bind 0/0` in every round: nothing fails, so the settle has nothing to rescue.
What changed is upstream of it. **The zero-match re-read is gone** — the defect
rounds 068/069 named exactly (`withOwnId` 7 of 7 with zero matches, the host
listing shapes under ids it never gave us). Calling `bindings.add` on the live
proxy in the drawing batch appears to **stabilise the shape's identity**, so the
re-read afterwards reports ids that match ours and the config tag write lands.

Attribution is clean: the only difference between `2c7dcd8` and `01f3607` is
that one commit.

**THE HOST IS STILL MOODY UNDERNEATH, and that is the honest reading of the
third round.** 072 grouped 14 where 071 grouped 19, declined 4 where 071 declined
0, and had 3 unmatched re-reads where 070 and 071 had none. Retries fired 10
times in all three. So the wobble did not go away — **the OUTCOME became robust
to it.** `same scale` is 8 of 8 whether the internals wobble or not, which is a
better result than a quieter host would have been.

`cfg5010` is **0 in all three** post-change rounds against 8 before. That is the
cleanest single signal in the set.

**What is now the top failure, and it is the one nothing here can check.**
`writing the chart's origin tag` 5010s 3-7 times a round. The origin tag is the
drag-delta round trip, and `docs/PUBLISHING.md` says it "needs a real drag and so
cannot be scripted at all". The failures have migrated to precisely the feature
no automated round reaches — see the backlog for the proposal to make the origin
ARITHMETIC scriptable even though a mouse drag is not.

### A driver gap this run exposed and closed

A browser that died UNDER a running round was invisible for 24 minutes.
`quietStreak` resets to zero on a failed CLI call — deliberately, so contention
cannot end a healthy round — but a dead browser makes every call fail forever, so
the counter never reaches its threshold and the loop polls a corpse to the
thirty-minute limit. `browserDiedMidRound` now needs a STREAK of failures AND an
affirmative `(no browsers)`; either alone re-makes a bug the other prevents.

## Rounds 075 + 076 (`aeefb93`, 2026-08-16) — the origin tag is fixed, and the binding is proven

**The prediction was staked before the pair and it held in both rounds:**

    round   scenarios  origin5010  charts that cannot follow a drag  cfg5010  grouped
    074        12/13        8                    8                     0       17
    075        13/13        0                    0                     0       19
    076        13/13        0                    0                     0       16

**Thirteen of thirteen, twice** — the first perfect sheets this project has
recorded, and the first with the drag check in the list at all.

**What was wrong.** The origin tag is the one write that cannot dodge a resolved
handle. The config tag is queued BEFORE `load("id,left,top")` and lands; the
origin tag needs that loaded position, so it went out after the load had
resolved the target into `shapes.getItem(id)` — the handle this host refuses.
Before the binding change the CONFIG tag took the 5010s and the origin tag none;
after it the config tag took none and the origin tag took eight or nine a round.
**The refusal never went away. It moved to the only write still going through a
resolved proxy**, and stayed invisible because a scenario samples one chart while
half the population was failing.

**And this refutes the alternative I staked against it.** `binding.getShape()` is
NOT `shapes.getItem(id)` under another name. A write through the binding lands
where the identical write through a resolved proxy was refused eight or nine
times a round. That closes the question the probe could never reach —
`binding-names-shape-later` asked eight times across eight rounds and never once
got to its own question, which was a fact about the probe.

**The host is still moody underneath and the outcome no longer cares.** Grouping
was 19 in one round and 16 in the other on the same build. Every scenario passed
in both. That is the same shape as the settled retry: the wobble did not go away,
the result became robust to it.

### What this leaves

`same scale across the deck` has now passed on **070, 071, 072, 073, 074, 075
and 076** — seven rounds across four builds, after failing 35 straight. It is no
longer a claim that needs a pair.

**The pooled `CHARTS THAT CANNOT FOLLOW A DRAG` count is now history rather than
news**: 34 charts across 5 rounds, all of them 072-074. It will not climb again
unless something regresses, which is exactly what it is there to notice.

**The settle-by-binding route has still never executed.** `settledByBinding` is 0
across seven rounds — the settle only runs when the config tag fails, and the
config tag stopped failing. It is now a fallback that has never fired, kept
deliberately: it costs nothing at rest and it is the only route left if the
config tag ever starts failing again. Recorded so nobody mistakes it for tested
code.

## Round 077 (`357632b`, 2026-08-16) — the first 4:3 round, and it is CONFOUNDED

**The first round in 53 not run at 16:9.** It scored **10 of 13** against 13 of
13 twice on the same build, and the failure is a class this project has never
seen:

    round   scenarios  UnexpectedError  tagging-failed  cfg5010  orig5010
    076        13/13          0               0            0        0
    077        10/13         52              18            0        0

    36 × UnexpectedError  at=writing the chart's config tag
    16 × UnexpectedError  at=settling the config tag through its binding

**The 5010s stayed fixed.** `cfg5010`, `orig5010` and `origLost` are all zero —
everything landed this week held. What appeared instead is `UnexpectedError`,
mostly on a write through the GROUP handle (`from: group×1`, 17 of 18).

**AND IT EXERCISED CODE THAT HAD NEVER RUN.** `settledByBinding` was 0 across
seven rounds because the settle only fires when the config tag fails and the
config tag had stopped failing. Here it fired sixteen times — and the binding
route was refused too, with the same `UnexpectedError`. So the fallback is no
longer untested; its first test says it does not save this case.

### WHY THIS IS NOT YET A FINDING ABOUT 4:3

**Two variables changed at once**, which is the exact mistake `docs/ROUNDS.md`
records about rounds 24 and 25 — "differed only in this, and were compared as
though they did not".

1. **The aspect ratio**, which is the thing under test.
2. **The deck.** Round 077 ran on `Presentation65`, not the `Presentation64` that
   carries all 52 previous rounds. That deck was created during a browser crash
   earlier the same day, and its whole ribbon — Present, Share, Slide Layout,
   Slide Size — was greyed out until a reload. A document whose session was
   already damaged is not a clean instrument.

`UnexpectedError` is also a different SHAPE of error from everything in this
archive: it is Office.js's generic failure, not the specific
`InvalidParam / 5010` that names a refused handle. That is as consistent with a
sick document as with an aspect ratio.

### THE CONTROL THAT SEPARATES THEM, and it costs one round

**Switch `Presentation65` back to 16:9 and run it again.**

    back to 13/13   the deck is fine and 4:3 is the cause
    still ~10/13    the DECK is the cause and 4:3 is exonerated

One variable moves, and it is the cheap one. The alternative — switching the
proven `Presentation64` to 4:3 — tests the same question but risks the deck
carrying 52 comparable rounds, which is not a trade worth making while a
cheaper control exists.

**Nothing about 4:3 support should be written down until that control has run.**

## Round 078 (`8e9227c`, 2026-08-16) — the control, and 4:3 is EXONERATED

**Same deck, same build, 16:9 instead of 4:3. One variable moved.**

    round   deck             size   scenarios  UnexpectedError  tagFail
    076     Presentation64   16:9      13/13          0            0
    077     Presentation65   4:3       10/13         52           18
    078     Presentation65   16:9      10/13         52           18

**077 and 078 are identical on every number.** The failure follows the DECK, not
the aspect ratio. `Presentation65` was created during a browser crash, has gone
into an unusable greyed ribbon state twice, and turned up at the start of this
round holding **100 slides**. It is a damaged document, and every one of those 52
`UnexpectedError`s belongs to it.

**Had the control not been run**, "4:3 breaks the config tag write" would have
gone into the docs on the strength of a single round, and every subsequent 4:3
result would have been read through it.

**Two things it also proved in passing.** The driver swept a 100-slide deck
rather than refusing — the heal added earlier the same day, on a deck far past
anything it was designed against. And `settledByBinding` fired sixteen times
after seven rounds of never running: the fallback is no longer untested, and its
first test says it does not save this case.

### What the 4:3 question still needs

**Nothing has yet been learned about 4:3.** Round 077 measured a sick deck. The
question needs a CLEAN 4:3 deck, which is now the first item of the nightly work
below.

## Rounds 079-081 (`17a8204`, 2026-08-16) — the first full cycle, and 4:3 is fine

**One 4:3 validation round and a 16:9 pair, all on one build.**

    round   deck             size      scenarios
    079     Presentation66   720x540     13/13
    080     Presentation64   960x540     12/13   explode a degraded picture
    081     Presentation64   960x540     13/13

**4:3 SCORED BETTER THAN 16:9 IN ITS OWN CYCLE.** Round 077's 52
`UnexpectedError`s belonged entirely to `Presentation65`, a deck created during
a browser crash — round 078 reproduced them at 16:9 on that same deck. On a
clean 4:3 deck the failure does not exist. **Nothing about 4:3 is broken.**

Both new mechanisms worked on their first real use: the size is recorded
(`{"width":720,"height":540,"source":"pageSetup"}`) and readiness confirmed the
profile before each round (`slide size 4:3 (want 4:3)`).

### THE DIVERGENCE CHECK GOT ITS FIRST OUTING WRONG, and it is fixed

It reported:

    explode a degraded picture — passed at 4:3, failed at 16:9

True of the worst reading and **wrong about the cause**. The 16:9 pair disagrees
with ITSELF: 080 failed and 081 passed. Collapsing a profile to its worst
outcome turned a flaky scenario into a slide-size difference, and would have sent
someone to investigate an aspect ratio for a scenario that is simply unreliable.

`profileDivergence` now keeps both outcomes per profile and separates them:

    DIVERGED between slide sizes      one profile passed, the other failed
    UNSTABLE WITHIN a slide size      one profile did both — flaky, not different

The gate says the second in its own words and does not call it divergence.
Mutation-proven.

**Worth noting what caught it**: not a test, but reading the first real output
and finding it did not match the rounds it was describing. A check built the same
afternoon, wrong on its first live data, is the ordinary case rather than a
surprise.

### `explode a degraded picture` is now instrumented rather than guessed at

Five failures across eight rounds and **no evidence**: the scenario reaches a
null from `updateChartInSlide` and the trace held nothing between the deck scan
and the verdict. There is a line now, carrying `asPicture` first because the
picture path is the one that fails intermittently.

**Upstream has no match.** office-js#3698 (image insertion refused while a shape
is selected) is the nearest and is a different API — `setSelectedDataAsync`, not
`ShapeFill.setImage` — with no evidence here of a selection. #225 (large base64
fails mid-insert) leaves a half-inserted image where ours leaves the chart
intact; #5022 hangs where ours returns. **A promising chain via #5443 was killed
by reading it**: that issue is Excel on Mac, not PowerPoint shape selection.

So no workaround was applied. Guessing one now would produce a change nothing
could evaluate; the next round that reproduces this will say which step failed.

## Rounds 082-087 (`9f175a8` → `95170cf`, 2026-08-16/17) — six perfect rounds, and nine bugs in the instruments watching them

    round   build      scenarios   note
    082     9f175a8      13/13     first perfect round in the archive
    083     9f175a8      13/13     the pair
    084     600ea90      13/13     the orphan instrument's first non-zero reading — a phantom
    085     fbdc49b      13/13     survived a real browser death, recovered unaided
    086     2c2547f      13/13     settle guard caught two phantoms, missed a third
    087     95170cf      13/13     4s settle held; zero contradicted by the deck

**Six consecutive 13/13 rounds, zero `UnexpectedError` in any of them.** `explode
a degraded picture` passed all six, having failed 47 of its first 57.

### The product was fine. The instruments were not.

Two product defects were found across the whole stretch — a blank slide shipping
in the finished deck, and an update dying on a refused id. **Nine defects were in
the reporting**, five of them introduced the same day they were found.

### The orphan instrument was wrong four times, each differently

1. **Mismatched units.** It subtracted inner chart shapes, delete CALLS and
   top-level slide shapes from one another and reported "283 stranded" beside its
   own `before: 3, after: 3`.
2. **An algebraic identity.** `shortfall` and `unexplained` summed to zero on
   every line, so the reading was the same on healthy and broken data — and its
   control test passed vacuously.
3. **One stale host read.** "92 shapes grew" was PowerPoint lagging an
   `addGroup` our own sync had already resolved; the deck showed one grouped
   chart per slide.
4. **Two agreeing stale reads.** With a 1.5s settle added, both reads landed
   inside one lag and agreed on the stale number, so the reading was marked
   `settled` and was still wrong.

**Every one was caught by the same second source** — `deck.inventory`, already in
the round file, taken long after any lag. It has never been wrong. It is now
cross-checked on every reading, and disagreements are counted rather than
discarded, because an instrument's own error rate belongs in its report.

The settle delay was then sized from data that had been in every round file for
weeks: the gap from a group commit to the count that followed it is stale up to
**3193ms**, so 1.5s was never enough. No new instrumentation was needed to learn
that — the timestamps simply had not been differenced.

### The gate spent fifteen rounds blaming nine innocent builds

`traceNovelty`'s "NEW BEHAVIOUR" bucket took its median over EVERY prior round,
so a signature stayed new until it appeared in more than half the archive — while
the denominator grew. `re-reading the slide's shapes again after a settle delay`
debuted in round 064 and has sat at 10-11 since; it was announced as new in
fifteen rounds and attributed to nine different builds, the last a commit that
only changed a slide counter. The signature round 086 had actually changed went
unmentioned.

Now a five-round window, and each entry names the build it first appeared in. On
the real archive the bucket reports **nothing**, which is the truth.

### A blank slide of ours was shipping in the finished deck

One `slides.add()` can land TWO slides. The branch that fires when a landed slide
cannot be named reported one — `after - before` is 2 in all ten archive
occurrences, never 1 — so the sweep's clamp ran one short and left the remainder.
Round 085's inventory carries it: `257#3837665135`, zero shapes, listed in
`newSlides`. Nine rounds are affected and five of the last thirteen began their
scenarios on a deck already dirty by one slide, invisible because the scenarios
measure GROWTH.

The comment above that branch cites a real 2026-08-11 measurement; what went
stale was the inference drawn from it, contradicted ten times since by data
nobody re-read.

### The stranding question has three observations

Round 087 finally failed to group something (15 grouped, 4 not). Pooled across
every round that can answer: **three at-risk charts, all zero growth.** Three is
not five. A round that groups everything cannot answer this either way, and the
report now says so rather than reading zero as an all-clear.

### The browser died three times, and the cause is not a crash

`ERR_NETWORK_IO_SUSPENDED` in the console tail, and Windows Event 507 — entering
connected standby — four times in one morning. The machine was sleeping. Recovery
survives it, and did so unaided in round 085. **A sideload does not always
survive** the browser process: once it was lost and once it was not, and an
`Upload My Add-in` appears to register against the account rather than the
session.

### Six vacuous tests

Written and caught in this stretch, all the same shape: built from what the
author expected rather than from what the code does. A callback passed as a
fourth argument to a three-parameter function; a fixture asserting a lookup
happened rather than a click; `ref=rCancel` against a `[a-z0-9]+` matcher; a
fixture modelling "rare" as "absent"; a control that could not fail; and a
novelty fixture with the signature in every prior, where the old and new readings
agree. **Mutation caught all six. Reading caught none.**

## Round 088 — 3056f91 — 12/13 — and a prediction that could not be answered

First round after the 14-commit merge (#576-#590). Browser was dead at `--check`;
the profile still held the sign-in, so it was reopened without a password and
**the sideload survived** — see the doctrine note below, because this file said
it would not. `HEAD 3056f91 · site 3056f91 · pane 3056f91 · deck 1 slide`.

The reading was pre-registered before the host was touched, which is the only
reason the round is legible. All three calls came out.

- **Mine.** `same scale across the deck` PASSED — its tenth consecutive pass, not
  its first. #586's branch was never entered: it runs only on a re-read that is
  SHORT after the settled retry, and round 088 recorded none. Across all 64
  rounds there are 42 short re-reads and `afterRetry: true` on **exactly none**
  of them — every one predates `REREAD_RETRY_MS`. What the round did record is
  the failure #586 deliberately does not rescue: 2 empty re-reads and 1
  zero-match, all after the retry, on charts that then hit `not grouping: no
  member handle this host will accept` and a 5010 on the id readback.

  Those two charts kept their config anyway. Tempting, and not a finding: the
  pooled table already prices an ungrouped chart at 32% keeping its tag, so
  two-for-two is inside the noise. Recorded so the next session does not spend a
  round on it.

  **The one real finding is the denominator.** The scenario ran **6 charts, not
  8** — the first time in 64 rounds. Its verdict is `scaled === charts.length`
  against a population it DISCOVERS (`probeCharts`, guarded only at `< 2`), so
  `6 of 6` is a pass and the gate compared PASS to PASS and said "no scenario
  regressed". Cause, consistent with four fields and not proven by one round:
  the upstream `insert onto a slide that already has content` stalled mid-draw
  (`shapes 1-10 of 16, 45s`), so fewer probe charts existed to find. Deck slide
  count was 7 — identical to nine rounds that found 8 — so deck size is NOT the
  cause; that hypothesis was formed and refuted against the archive.

- **Research.** office-js **#5022** (open, under investigation, product bug):
  `context.sync()` hangs after delete-then-re-read on the web; the reported
  workaround is "a timer of 1-2 seconds between the `shape.delete()` and the next
  `await context.sync()`" and the reporter says it **remained unreliable**. That
  is `REREAD_RETRY_MS` at 1500ms, independently corroborated as a mitigation
  rather than a cure — which is exactly round 088's two empty re-reads surviving
  it. **#6498** (open, needs attention, 2026-02-09): shapes inserted on the web
  not reflected without a refresh; same family, no workaround offered. Nothing
  upstream we are not already doing. Recorded so it is not rediscovered.

- **Instrument.** `poolScenarioPopulations` in `triage.mjs`, printed by the gate:
  it names a scenario passing on a smaller population than it usually runs. On
  round 088 it prints `same scale across the deck — 6 this round, usually 8 over
  63 prior round(s) (and it still reports PASS)`. Median, not mean, so one small
  round cannot lower the bar for the next — the mutation that turns it into a
  mean fails its own test.

- **Fix.** Three defects in the judge, each with a test that goes red when the
  fix is reverted:
  1. `scenario-passes` recorded a SKIPPED scenario as `FAILED`, while
     `probe-answers` and `probe-detail-matches` beside it already answered
     `undetermined` — and this file's doctrine says "a skip is not a flip". Round
     088 made it live: `insert onto a slide that already has content` came back
     `skipped: true, ok: false`.
  2. `roundToJudgeOn` took the FIRST round on a build, so with the cycle's three
     rounds per build a prediction could be judged on a sibling of its own
     control. Now the last.
  3. `stampDate` read the stamp down to the DAY and compared with strict `>`, so
     a round taken the same day a prediction was staked could not judge it. Round
     088 is that case exactly. Reading `HH:MM` too makes it work by lexicographic
     order with no extra branch.

  **The first draft of test 2 was vacuous** and mutation caught it: the fixture
  put a later build after the siblings, and both readings return that later
  build. The siblings have to be the newest rounds or the test proves nothing.

- **Doctrine.** Four documents asserted things this archive contradicts:
  - `docs/ROUNDS.md` said "a group is deleted whole", so only an ungrouped chart
    can strand shapes. #586 ended that: it groups a SUBSET and leaves the
    remainder out of the parts tag on purpose. **`atRisk` reads the host's shape
    type, so a subset group reports `group` and is counted SAFE** — the
    instrument is blind to the only stranding the code now creates deliberately.
    Latent today (the branch has never run), a landmine tomorrow.
  - `docs/round-prompt.md` and the #586 ledger entry both stated "failed 34 of 34
    rounds" and "chart 4 matched 20 of 24 in every round on record" as current
    fact. Neither is true; both corrected in place with the archive numbers.
  - `docs/BACKLOG.md` asserted the shipped change and its opposite, in two live
    sections of one file.
  - **A browser death does NOT always take the sideload.** It did on 2026-08-18
    and did not on 2026-08-19; the ribbon group was intact after the relaunch.

  The prediction is recorded `undetermined` on `088-3056f91`, not held. Once the
  judge was fixed it began printing **held** — a false positive that two bugs had
  been hiding between them: a claim that could not discriminate, behind a judge
  that refused to judge. Fixing one without the other manufactures the artifact.

## Round 089 — 3056f91 — 13/13 — the pair, and what it took back

Same build, no merge between, which is the discipline. **PowerPoint crashed 324s
into attempt 1**; the driver caught the dialog, wrote the host's own account to
`crashes/2026-08-19T14-46-31.md`, cleared it and ran again. Attempt 2 finished
clean. That machinery had been built and never yet met the thing it was for.

- **Mine — the pair takes back round 088's loudest reading.** 089 ran `same scale
  across the deck` at **8 of 8**, not 6. So the shrunken population does NOT
  replicate: it is what happens WHEN an upstream scenario is skipped, not a new
  steady state. Reported alone, round 088 would have read as "the scenario
  shrank". This is what pairs are for, and it is the second time the same
  denominator has needed one.

  The deck inventories say the same thing from the other side: 088 ended
  `0,2,2,11,24,24,1` — two slides carrying 24 LOOSE shapes, the two charts whose
  re-read came back empty — and 089 ended `0,4,2,11,1,1,1`, one shape per slide,
  every chart grouped.

  **#586 was not exercised in EITHER round.** 0 short re-reads in both. Two
  observations now, on one build, and 42 short re-reads in the archive with
  `afterRetry: true` on none of them.

  **The zero-match is the live one:** 1 in 088, 2 in 089, 3 in 087. It is the
  failure #586's strict-majority bound deliberately does not rescue, and it is
  the only one still happening.

- **Research — the crash was the network, not the host and not us.** The captured
  console is `ERR_NETWORK_CHANGED`, `ERR_NAME_NOT_RESOLVED`, `ERR_HTTP2_PING_FAILED`,
  `ERR_CONNECTION_CLOSED`, and a `wss://…augloop.office.com` that would not
  resolve. Same family as the dead browser at the start of this session, and the
  same window in which `gh` could not reach `api.github.com` and `curl` could not
  resolve a host from this machine. Nothing to fix in this repo; recorded so the
  next crash file is not read as a PowerPoint bug.

- **Instrument — the novelty view earned its keep.** It named **one trace
  signature never seen in 64 prior rounds**: `error|reading the deck's style`.
  That is #583, merged the same morning and never before run against a host.

- **Fix — a failed read was reporting an absence.** `reading the deck's style`
  hung for its full 90s budget while the host kept answering other calls
  throughout it (see the correction under round 090 — this line first had the
  ordering backwards). `readDeckStyle` caught that and returned `null` — the same value a deck
  carrying no style returns — and `style-from-deck` AWAITS it, so the pane would
  tell the user their deck is unbranded and switch them to the browser's style on
  the strength of a read that never happened. `readDeckStyleWithReason` now
  separates them; the fire-and-forget pane-load caller keeps the narrow contract
  it wants. The fake could not express a host that has the API and will not
  answer the READ — `refuseCustomXmlReads` closes that, beside the write switch
  that already existed.

  **This is the same defect this repo keeps finding in new places:** unreadable
  reported as negative. `collectDeckEvidence`'s `beforeUnknown`, "no history is
  not a spike", "a miss is not a failure", and now a deck style. Fourth time.

- **Doctrine.** The crash-recovery path works on a real crash, first evidence.
  The round-089 population is the counter-example the brief needs beside round
  088's: a scenario's denominator moves with what ran before it, so quote the
  denominator or quote nothing.

## Round 090 — f6cd580 — 13/13 — and the deck-style read fails every time

Four attempts, and the driver drove all of it: a stale pane after the merge, then
a browser death 51s into attempt 3, then a clean run. No password anywhere. The
round before it never started at all — see the refusal below, which was the
session's largest finding and was not in the round file.

- **Mine.** 13 of 13, `same scale` 8 of 8, and the cleanest re-read round on
  record: **0 short, 0 zero-match, 0 empty**, 20 grouped lines with `partial:0`.
  Deck ended `0,4,2,5,1,1,1` — one shape per chart slide, everything grouped.
  Nothing changed in the draw path between 089 and 090, so a clean round is
  weather: 2 zero-matches to 0 is inside the noise floor and is not a result.

  **#586 still not exercised — three rounds now.** 0 short re-reads after the
  settled retry in 088, 089 and 090.

- **THE DECK-STYLE READ FAILS ON THIS HOST, 2 FOR 2.** Round 089 flagged it as a
  signature never seen in 64 rounds; round 090 reproduced it exactly:

      089  afterMs 90000, idleMs -89057, "listing the deck's slides" answered in 44ms
      090  afterMs 90000, idleMs -89xxx, "listing the deck's slides" answered in 56ms

  **CORRECTED: those two lines were first read backwards, here and in #595.** The
  44ms and 56ms are how long that OTHER call took, not how long before the stall
  it ran. `idleMs` is `issued - lastAnswered` and `lastAnswered` is sampled when
  the deadline fires, so a NEGATIVE value means the host answered something else
  89 seconds INTO this call's wait — one second before it was abandoned.

  Two rounds, both since #583 landed, which is every chance it has had at that
  point. The host is alive throughout and the read never answers.

  **CORRECTED after round 092 — "has never once succeeded here" was wrong.** It
  succeeds in 428ms. See the round 092 entry: the failure is INTERMITTENT, and
  this line was written from the only two observations that existed.

- **Research — a null result, recorded as one.** No upstream issue describes a
  PowerPoint-web custom-XML READ that hangs. office-js #2937 ("CustomXMLPart
  can't sync") is Word on the desktop and bibliography-specific; #3936 is Excel
  on iPad and needs an unsaved file. Neither is this. Worth reporting upstream —
  the owner's call, not ours to post.

- **Instrument.** The pooled report now says, in as many words, that #586's
  subset branch has not run in the pool, instead of printing a zero at-risk count
  that reads as an all-clear. The population check stayed quiet, correctly: 8 of
  8 is the usual denominator.

- **Fix.** The deck-style read gets its own 10s budget instead of the 90s
  readback one. **The number is not sized from a successful read, because this
  host has never produced one** — it bounds the damage rather than fitting a
  distribution, and both callers degrade safely: the pane keeps the browser's
  style, the button says it could not read. `PW_DECK_STYLE_TIMEOUT_MS` overrides.
  What the 90s was costing: a `PowerPoint.run` context held open for a minute and
  a half on EVERY pane load, and a person waiting that long before the button
  admitted defeat.

- **Doctrine — the refusal that stopped the round before it started.** `--check`
  said this document had no add-in and that a reload would not bring it back: the
  one stop recovery may not retry. The tree said `button "Insert chart"
  [disabled]` and `status: Disconnected` the whole time. `refFor` returns the ref
  on the matching line and Playwright issues refs only for actionable elements,
  so ABSENT and DISABLED arrived as the same null and the driver chose the
  permanent explanation for a transient state. Now `host-disconnected`, and
  recoverable — proven on the same document, which then reloaded and came back.

  **Fifth time this codebase has reported UNREADABLE as NEGATIVE**, after
  `beforeUnknown`, "no history is not a spike", "a miss is not a failure", and
  the deck style. It is the house defect.

### After round 090 — the number nobody could read was the answer

Not a round. Mining round 090's own trace harder, before staking the next one.

**`idleMs: -89057` is not a glitch, it is the finding.** It reads as nonsense, so
two rounds carried it and nobody used it. `idleMs` is `issued - lastAnswered`,
and `lastAnswered` is sampled when the DEADLINE fires — so a negative value means
the host answered something else AFTER this call went out, while it was still
waiting. PowerPoint answered `listing the deck's slides` in 44ms, **89 seconds
into the deck-style read's 90-second wait**, one second before it was abandoned.

So the deck-style stall is **one call stuck on a healthy host**, not a host that
went quiet. Those are different faults wanting different fixes, and this project
had no way to say which it was looking at.

`stallShape` now says it in words on every stall, because the sign of a number is
not something a reader notices — and the comment on `lastStall` had assumed
sequential code, where a negative gap cannot happen. The pane's deck-style read
is fire-and-forget, so it breaks that assumption, which is exactly why this is
the call that exposed it.

**A cheaper question, asked only after the read has already failed.** The
namespace lookup, `getOnlyItemOrNullObject`, the `load` and `getXml` all go into
ONE batch, so one sync covers four calls and the trace can only report the sync.
`getCount()` on the same scoped collection needs the namespace lookup and nothing
else, so its answer splits the two explanations:

    it answers    -> the collection is reachable; the fault is in
                     getOnlyItemOrNullObject / load / getXml
    it also hangs -> the whole customXmlParts surface is dead here and #583
                     cannot work on this host at all

The next round decides it, which two rounds so far could not.

## Round 091 — de76543 — 13/13 — and the feature nobody could click

Two attempts. The driver cleared a silent host and a stale pane by itself, which
is the pair the previous session left it: `8012ms silent` before, `3ms` after.

- **Mine.** 13 of 13. The staked prediction came back **`undetermined — neither
  line appeared`**, and the ledger was right to say so rather than call it a
  refutation. That is what `insteadOf` was added for, one round earlier.

  **The reason is my own fix.** The deck-style read fires from `Office.onReady`,
  and dropping its budget from 90s to 10s moved its failure out of the round's
  traced window: round 090 recorded it at entry 135 of 528, round 091 has not one
  deck-style line in 529. The bug did not go away, the instrument stopped being
  able to see it. **A fix that hides its own subject is worse than the cost it
  removed.** The conclusion is remembered now and replayed at round start, so the
  next round can judge the claim it was staked for.

- **AND THE FEATURE WAS UNREACHABLE ANYWAY.** Chasing why the probe never ran, in
  the live pane after the round:

      Office.context.host  "PowerPoint"    <- isPowerPointHost() would say true NOW
      style-from-deck      disabled: true
      style-to-deck        disabled: true

  `app.ts` asked `isPowerPointHost()` at MODULE SCOPE, where `Office.context`
  does not exist yet, so it answered false on every load inside PowerPoint and
  disabled both buttons for the whole session. Nothing ever asked again. **#583's
  entire user-facing feature could not be clicked on PowerPoint web**, and
  thirteen scenarios could not see it because none of them clicks a button.

  Swept: `Use deck theme` has the same gate and `renderOptions()` is also called
  at module scope, so it was disabled for the same reason — hidden only by
  incidental re-renders, which makes it harder to diagnose, not safer. Both are
  under one `syncHostOnlyButtons()` now, called again from `Office.onReady`.

- **Research.** Nothing new to search: the question this round was staked to ask
  was never put. Recorded rather than skipped.

- **Instrument.** `stallShape` shipped this round and had nothing to describe —
  no stalls. The replay above is the round's real instrument change.

- **Doctrine — a correction to yesterday's correction.** "Playwright hands out a
  ref only for something it could act on" was too strong. Same session, same
  page: the ribbon's `button "Insert chart" [disabled]` has no ref, the pane's
  `button "Use deck style" [disabled] [ref=f19e328]` has one. Native `disabled`
  and whatever Office marks its ribbon with are not treated alike. The fix
  (`namePresent`) stands; the explanation did not, and is corrected in
  `round.mjs` and in memory.

## Round 092 — 0aa6f91 — 13/13 — the bug that stopped happening

Second attempt again; the driver cleared a silent host and a stale pane by itself.

- **Mine.** 13 of 13. The staked prediction came back `undetermined — neither
  line appeared` for the SECOND round running, and the reason is different this
  time. 091 could not ask because the 10s budget moved the failure out of the
  traced window. **092 did not ask because the failure did not happen**: zero
  deck-style entries, and the replay correctly stayed silent because there was no
  verdict to carry — which is the behaviour its second test pins.

- **THE READ WORKS. I said it never had, and that was wrong.** Driving the pane's
  own button on the real host after the round:

      click 1   428ms   "This deck carries no style — using yours."
      click 2-4  <1s    same

  That is the `isNullObject` path completing normally, not the timeout path. So
  the deck-style read is **intermittent**, not broken: 90s and never in rounds 089
  and 090, under half a second now. "#583's deck-style read has never once
  succeeded on this host" was written from the only two observations that existed
  and is corrected in the journal, in `powerpoint.ts` and in the ledger.

  **The 10s budget is now sized from evidence rather than from harm** — 428ms
  observed, so ten seconds is ~23x the success and still fails fast. Not tightened
  further: one sample from a healthy host is not a distribution.

  **A hypothesis at n=4, offered as one.** The two rounds that hung (089, 090) are
  also the two with a host crash and a browser death. The two that did not (091,
  092) booted cleanly. Four rounds is not an effect. The test is cheap though:
  when it hangs again, look at whether that round had a crash.

- **THE FEATURE IS REACHABLE FOR THE FIRST TIME, verified on the real host.**
  Before the fix, `button "Use deck style" [disabled]`; after it,
  `button "Use deck style" [ref=f25e328] [cursor=pointer]`, and clicking it
  answers. Everything above was only measurable BECAUSE that fix landed — the
  button could not be clicked at all until this round.

- **Research.** None this round: the question the ledger is staked on was not put,
  and inventing a search to fill the slot would be the padding this protocol
  exists to prevent.

- **Instrument.** The replay shipped and correctly emitted nothing. A test that
  only ever fires is not a test; the silent case is pinned too.

## Round 093 — 0aa6f91 — 13/13 — THE FIRST REAL PAIR, and it does not replicate

First attempt, no recovery needed. **And the first time in this stretch that two
rounds ran on ONE build** — 089 through 092 each had their own, because fixes
were landed between every round. The brief says not to do that, in bold, and the
reason is exactly what this pair shows.

- **Mine — same build, nothing changed, and the internals swing hard:**

      round   grouped  refused  errors  5010  deck inventory
      092        20       0        2      5   0,4,2,5,1,1,1
      093        15       4        9     11   0,4,2,17,24,24,24

  Three slides ended holding **24 shapes each instead of one**: three charts that
  did not group, **seventy-two shapes left loose on the deck**. And both rounds
  report `13/13` and the byte-identical verdict line `8 of 8 charts carry the
  shared scale (max=105 …); 8 still re-editable`.

  The trace and the deck agree exactly — 15 grouped and 4 refused in the trace,
  three big slides in the inventory — so both are trustworthy here. This is the
  authority confirming the instrument rather than contradicting it, which is the
  first time in this journal it has gone that way round.

  **The scenario is not lying.** It asks whether the config survived, and it did:
  the ungrouped fallback keeps the tag. It simply cannot see grouping — which is
  what #586, #520 and most of the last month have been about. So
  `scenarioRegressions` compares PASS to PASS across a round that left seventy-two
  loose shapes and one that left none.

  **Read as noise, not as a fall.** 0 and 4 refusals on the same build is inside
  this project's own floor (1 vs 5, nothing changed). The finding is not "093 is
  worse"; it is "the verdict cannot see the difference at all".

- **Research.** None: the pair answered a question about our own instrument, not
  about the host, and there is nothing to look up for it. Said rather than padded.

- **Instrument.** `poolGroupingOutcome`, printed by the gate every round —
  charts grouped, charts refused, the median refusals of the archive, and the
  deck line beside it. Counted from the trace, which is exact; the deck is
  printed as corroboration rather than used as the measure, because turning slide
  shape-counts into "ungrouped" needs a threshold and a guard sized by guesswork
  is how this instrument has been wrong before.

- **Fix — process, not code.** Stop landing between the two rounds of a pair.
  089-092 were four rounds on four builds and none of them can be compared with
  its neighbour. This pair took twenty minutes and produced the clearest result
  of the night.

## Rounds 094 + 095 — ab5d730 — 13/13 and 13/13 — the blindness replicates

A second pair, on a second build, run back to back with nothing landed between.

      build      round   grouped  refused  deck inventory
      0aa6f91     092      20        0     0,4,2,5,1,1,1
      0aa6f91     093      15        4     0,4,2,17,24,24,24
      ab5d730     094      20        0     0,4,2,5,1,1,1
      ab5d730     095      19        1     0,4,2,5,1,1,24

**Four rounds, two builds, every one 13/13 with a byte-identical verdict line** —
and two of them left charts ungrouped as loose shapes on the deck. The finding
from the first pair is not a one-off: `same scale` asks whether the config
survived, which it does either way, so grouping is invisible to it.

What the second pair adds is the SHAPE of the noise. Refusals are not rare and
not constant: 0, 4, 0, 1 across four rounds on two builds. Any single round's
refusal count is uninformative, and the difference between two rounds on one
build is not evidence of anything about the build. `poolGroupingOutcome` prints
it every round now precisely so nobody has to infer it from a verdict that cannot
see it.

**This is what pairing buys.** Four rounds run as two pairs answered a question
that eight rounds run singly could not have, because the comparison a single
round invites — against the previous round, on a different build — is exactly the
comparison the noise floor forbids.

## Rounds 096 + 097 — 9e81c14 — two staked claims held, and one of them closes a question

The richest pair of the run. Both predictions were written down before the host
was touched, and both came out.

- **`the-round-file-can-finally-say-which-host-it-ran-on` — HELD, both rounds.**
  The first two rounds in 71 to carry it:

      host PowerPoint · platform OfficeOnline · version 0.0.0.0
      requirementSets 1.1 … 1.10 · canInsertSlidesFromBase64 · canInsertPicture
      slide size 960x540, source documentFile

  So the trace mark was the whole story and there is no second slice. Seventy
  rounds were archived without any of this.

- **`the-deck-style-namespace-is-reachable` — HELD, both rounds, identically.**
  `the namespace IS reachable — the fault is further in`, `parts: 0`,
  `observedBeforeTheRound: true` — the replay carrying it across the mark, which
  is the whole reason it exists.

  **The expensive half is refuted**: `customXmlParts` is not dead on this host,
  so #583 needs a fix rather than a product decision. `getCount()` answers; it is
  `getOnlyItemOrNullObject` / `load` / `getXml` that hangs. And `parts: 0` names
  the call precisely — **asking for the only item of an EMPTY collection**, which
  is the state every unbranded deck is in, so every pane load in the wild.

  Fixed by counting first and never making that call unless the count is 1. It
  costs one round trip on a deck that carries a style and buys not hanging on a
  deck that does not.

  It also settles the intermittency from round 092: the manual clicks that
  answered in 428ms hit the same empty namespace and did not hang, so this is a
  RACE rather than a property of the empty case. The guard removes the call, not
  the race — worth remembering if it ever hangs at `count === 1`.

- **The honesty fixes are visible on the real host.** 096 came back **12/13**,
  skipping `the chart is actually visible` — the first honest reading that
  scenario has ever produced. 097 passed it. That difference is the fix working:
  the control render is available sometimes and not others, and the verdict now
  says which rather than reporting green either way. A lower score for a truer
  reason.

  `stallShape` also had its first outing, and correctly declined to over-read two
  stalls: `issued immediately after the previous answer — an ordinary sequential
  gap, says little`.

- **Grouping, third pair running: 19/0 then 15/4.** Pooled across three pairs the
  refusal counts are 0, 4, 0, 1, 0, 4. Neither rare nor constant, and invisible
  to a verdict that reports 13/13 either way.

## Rounds 098 + 099 — 3da2ef6 — a first-ever failure that the pair dissolved

      round  score  failed scenario                    scratch-unresolved  same scale  refused
      098    12/13  edit a chart on the visible slide          14            7 of 7        0
      099    13/13  —                                           0            8 of 8        4

Same build, nothing landed between them, and **none of round 098 replicated**.

`edit a chart on the visible slide` had passed 73 of 73 rounds and failed here for
the first time. The gate caught it and exited 1 — `had passed the previous 3
rounds running` — which is exactly its job. Then 099 passed it and the gate went
quiet again.

**The 14× spike is the round's real story, and it is weather.** `scratch slide
landed but its id will not resolve` and `could not remove the unusable scratch
slide` both ran 14 times against a baseline of 1, and both were 0 in the very next
round. The 7-chart population is downstream of that same churn — the probe charts
are discovered from a deck that fourteen unremovable scratch slides had polluted.

**I had a live reason to suspect my own change**, since 098 was the first round on
the count-first deck-style read. The pair says no: nothing in 099 differs except
the host's mood. That is the whole argument for pairing, and it is the second time
tonight a single round would have sent someone after a regression that does not
exist.

### The staked prediction failed, and the instrument was why

`counting-first-removes-the-hang` — FAILED. The probe line is still there, so the
read still fails with the guard in place.

**But the probe was lying, and it was mine.** Its `meaning` field was the fixed
sentence `getOnlyItemOrNullObject/load/getXml is what hung` — true while the read
was one batch, false the moment the read gained a `getCount` of its own. Round 098
printed it on a build where the count runs FIRST, describing a call it had not
observed. So the verdict refutes the prediction and not the diagnosis: the
diagnosis was never tested, because the instrument built to test it asserted its
answer.

The field now carries the operation the bounded sync names (`at=<what>`), read off
the actual failure. The count-first change is kept rather than reverted — it takes
one call off the path every pane load walks, costs a round trip only on a branded
deck, and nothing here shows it made anything worse. Its premise is simply
unconfirmed, and the next round can say.

### A note on what the gate can and cannot do

The gate judges the NEWEST round, so 098's alarm disappeared the moment 099
passed. That is correct for a regression gate and worth writing down anyway: an
unattended loop can erase its own alarm by continuing, and the only durable record
of round 098's failure is this entry and the archive. Not changed — a gate that
keeps firing about a round two back would be reporting history, and the pooled
per-scenario view already carries it.

## Rounds 100 + 101 — 504033c — the honest probe names the fault, and it is ONE call

Both 12/13, both skipping `the chart is actually visible`, and both reporting the
same thing byte for byte the first time the probe was allowed to measure rather
than assert:

    failedAt: "The value of the result object has not been loaded yet. Before reading the value"

**That is not a timeout.** `boundedSync` rejects on timeout, so this says the sync
RESOLVED and the value was still unloaded — a different fault from the 90-second
hangs, on the same line of code.

**Six rounds, three costumes, one fault.** Put side by side, what looked like two
separate bugs is one:

    089, 090   first call = getOnlyItemOrNullObject + load + getXml   hung, 90s
    096, 097   same                                                  hung, 10s
    098, 100, 101   first call = getCount (counting moved to front)  resolved,
                                                                     value NOT loaded
    the probe — always the SECOND call, fresh context                answered, always
    round 092, manual clicks on a long-lived pane                    428ms

**The first customXmlParts call after a pane loads does not work. The second
does.** The failure takes the shape of whichever call happens to be first, which
is exactly why moving `getCount` to the front changed the error message and not
the outcome.

And the diagnostic probe had been demonstrating the cure since the day it was
written, without anyone reading it that way: it opens a fresh context, asks the
same question, and has answered on every round it ran — immediately after the
read it was diagnosing had just failed.

So the read spends the bad call: attempt, and on failure attempt once more. Not a
timeout to tune.

- **Fix.** `attemptDeckStyleRead` twice. Guards proven red both ways — dropping
  the retry, and retrying without preserving `unreadable`, which would turn a
  genuinely dead read into a silent "no style" and is the exact lie that flag
  exists to prevent. The fake gained `refuseCustomXmlReadsOnce`, because a
  boolean cannot express "once" and without it a retry is indistinguishable from
  a read that gave up.

- **Instrument.** This is what the honest probe bought, on its first outing. The
  hardcoded `meaning` had been printing a conclusion about `getOnlyItemOrNullObject`
  on builds where that call no longer ran first.

- **The visibility skip is now in both rounds** where 096/097 split one-each. The
  control render's availability moves; the verdict reports which, which is all it
  was ever supposed to do.

## Rounds 102 + 103 — f3e3941 — the retry is taken back by the rounds that tested it

Both 12/13. The two scenario problems did not replicate — 102's
`insert onto a slide that already has content` and 103's visibility skip are one
each, which is weather. **The deck-style regime replicated exactly**, and it is
worse than the one before it.

    rounds     build      probe says
    096-099    pre-retry  the namespace IS reachable
    100, 101   504033c    the namespace IS reachable   (count-first, NO retry)
    102, 103   f3e3941    the namespace is UNREACHABLE too, at `counting the
                          deck's style parts`

**Two rounds on each side of one commit, and that commit is mine.** The probe had
answered on every round it ever ran — that was the whole evidence for "the second
call works" — and it stopped the moment the read began making two calls of its own
ahead of it.

The plain reading: the read's retry **spends the good second call**. The probe,
now third, gets the same nothing the first one got. A trace line appearing where
none did is one of the two readings this project treats as real, and this one has
a pair on each side.

**So the retry is reverted.** I wrote, one round earlier, "the retry stays; it is
simply not the whole cure" — that was too generous to my own change. It is not
merely not a cure: on this host it moved the failure from one call to the whole
surface, and on the interactive path it made a person wait two budgets instead of
one before being told.

What survives is the diagnosis it was built on — the first customXmlParts call
after a pane loads is the one that fails — and the instrument honest enough to
show the retry backfiring. Spending the bad call is still the right shape; doing
it INSIDE this read is not, because this read is what the probe measures.
Somewhere earlier in pane start, once, is the version worth trying next.

**And the ledger said HELD for round 102 before I read the round.** The claim
named one of the probe's two failure lines and the round produced the other, so
the named line was genuinely absent and the entry read as a cure while the read
had failed both attempts. A false HELD, in my own claim, of exactly the kind this
file spends its time finding elsewhere. `trace-line-present` takes a `scope` now,
so an absence claim can name the whole family, and the verdict says which of the
two it matched rather than quoting the one it did not use.

## Rounds 104 + 105 — 9b6a78f — the revert holds, and the diagnosis is now clean

    102, 103   WITH my retry     probe: the namespace is UNREACHABLE too
    104, 105   retry removed     probe: the namespace IS reachable   (both)

Two pairs bracketing one commit. The probe answers again on both rounds with the
retry gone, so it really was being starved of the good second call, and the
revert was not wasted motion.

**The pair also gives the cleanest statement of the fault the archive has
produced:** on both rounds the read fails at `counting the deck's style parts` —
its FIRST call — and the probe counts the same namespace successfully moments
later. First call fails, second works, twice, on one build.

105 skipped `the chart is actually visible` where 104 passed it; that is the
control render's availability moving, which is what the skip is there to report.
Grouping 20/0 then 17/2 — the sixth pair, and the refusal counts still swing on
one build.

- **Fix.** `warmCustomXmlSurface` spends the bad call before the read, in the
  pane's boot chain, where nothing waits on the answer and nothing else measures
  it. One thrown-away `getCount`. It cannot throw — failing is what it is for —
  and the guard for that is a mutation that rethrows, which takes the whole
  pane-load chain down.

- **Instrument.** The fake's once-fault only fired on the ITEM read, so it could
  not express "the first call, whatever it is" — and the warm-up sailed straight
  through it while the test failed for the wrong reason. It covers `getCount` now,
  which is what the real fault does.

- **The gate caught a real one.** Adding a renderer export the pane calls means
  adding it to the pane's MOCK in the same change: without it `app.ts` throws at
  boot and 110 tests die at once. Second time this session — the same thing
  happened with `readDeckStyleWithReason`.

## Rounds 106 + 107 — fa6448a — the first quiet pair, and why quiet is not an answer

    106   12/13   deck-style entries: 0
    107   13/13   deck-style entries: 0

**The first two rounds in this archive where the deck-style read recorded no
failure at all**, and the staked claim came back `held` on exactly that.

**It is weaker than it looks, and the claim was the wrong shape.** The probe
writes only when the read FAILS and the warm-up traces nothing, so an empty scope
is equally consistent with the read never having run — and the pane-load window
sits before the round's trace mark, so nothing in the file can settle it. A cure
and a silence have the same shape here.

That is this project's house defect pointed the other way: an absent failure line
read as evidence of success. Round 102 had pushed me toward claiming an ABSENCE
after a message-shaped claim misfired, and the right answer there was a POSITIVE
claim, not a broader negative one. Two mistakes about the same instrument, in
opposite directions, three rounds apart.

- **Fix.** The read records success as well as failure — `the deck-style read
  answered`, carrying the part count and the call it got past — and the verdict
  replays at round start like the failure one. A round can now SAY the read
  worked instead of leaving a gap that has to be interpreted.

- **Re-staked positively** as `the-read-answers-and-says-so`. HELD means the read
  is genuinely cured on this host and the round says so in a line that only
  exists if it completed. FAILED — `round starting` present, success line absent
  — means the read is not running in the pane-load chain at all, which would make
  this pair's quiet a worse story than a cure rather than a better one.

- **The warm-up is kept and is not yet credited.** It is cheap, it cannot throw,
  and nothing here argues against it. But two silent rounds do not show it is the
  reason they were silent, and saying so costs nothing.

Grouping, seventh pair: 20/0 then 17/2. 106 skipped `the chart is actually
visible` and 107 passed it — the control render moving again, which is what the
skip exists to report.

## Rounds 108 + 109 — f47a6d4 — the deck-style read is cured, and the round SAYS so

Both rounds carry `the deck-style read answered`, `parts: 0`, at `counting the
deck's style parts` — the call that had failed on every measured round since 089.
Positive evidence, twice, on one build.

**The whole arc, from the archive:**

    089, 090   —          hung 90s, no probe existed yet
    096-101    5 builds   probe: namespace reachable, read still failing
    102, 103   my retry   probe: UNREACHABLE too
    104, 105   reverted   probe: reachable again
    106, 107   warm-up    silent — cure or silence, indistinguishable
    108, 109   + success  THE READ ANSWERED

Twenty rounds. The fault was one sentence long the whole time: **the first
`customXmlParts` call after a pane loads does not work, and the second does.**

### What it cost to find, and why

Three wrong turns, each caught by a round:

1. **A retry inside the read** (#602), reverted by the next pair (#603). The read
   is what the probe measures, so a retry there spends the very call the
   measurement depends on. Two pairs bracketing one commit said so.
2. **A false HELD** — the claim named one of the probe's two failure lines and
   the round produced the other, so a genuine failure read as a cure.
3. **An absence claim** that could not tell a cure from a read that never ran,
   which is the same defect as (2) pointed the other way.

And the instrument only became useful when it stopped asserting: its `meaning`
field was a hardcoded sentence naming `getOnlyItemOrNullObject`, and it kept
printing that on builds where the count ran first. The moment it reported the
operation it had actually observed, the diagnosis fell out in one round.

**Nothing here fixes the host.** The bad first call is still bad; the warm-up
makes it the throwaway rather than the one a person waits on. Remove the warm-up
and the failure returns — that is what the guard pins and what 089-105 record.

Grouping, eighth pair: 20/0 then 18/2. Both rounds skipped `the chart is actually
visible`; the control render has been unavailable on this build across four
rounds now, which is worth watching but is exactly what the skip is for.

## Round 110 — 48cd380 — the cure holds, and the visibility gate's blindness has one cause

`the deck-style read answered` for the third consecutive round. The cure is not a
two-round coincidence.

**And the visibility gate's blindness is fully determined.** Across 15 rounds the
correlation is exact:

    blind rounds    8 (now 9)  — each with exactly ONE `rasterising a slide`
                                 stall at its full 20000ms budget
    sighted rounds  6          — zero stalls
    rasteriser actually unstable: 0 times in 85 rounds

The chart draws visibly in every one of them — the render moves ~1000 bytes each
time. What is lost is the CONTROL: the SECOND `getImageAsBase64` of the same
slide hangs. The first answered, which is where `before` came from.

Note the shape. The deck-style fault was the FIRST call failing and the second
working; this is the mirror — the first works and the second hangs. Two different
host surfaces, opposite orders, same class of fault.

- **Instrument.** The blind verdict names the stall now instead of shrugging "an
  unstable rasteriser cannot be ruled out" — true, but unhelpful when the archive
  has pinned the cause fifteen times. It says nothing when there is no stall to
  point at, which is the guard.

- **NO FIX SHIPPED, deliberately.** A retry is the obvious move and there is no
  pair behind it. The last obvious retry this session made things worse and the
  rounds took it back within two rounds — and this one would sit on the very
  call whose second attempt is the thing failing. It wants its own pair, not a
  05:00 guess.

  What that costs: the project's only mechanical evidence that a drawn chart is
  visible is unavailable in more than half of all rounds. Worth someone's next
  session.

## Round 111 — d56cf96 — 13/13 — the cold read never stopped failing

The first round carrying four new instruments, and one of them overturns a story
this project has been telling itself for twenty rounds.

**THE COLD RE-READ FAILS ELEVEN TIMES A ROUND.**

    cold re-reads that fell short   11   (1 short, 4 empty, 4 zero-match, 2 short)
    settle-delay retries fired      11
    post-retry failures              0
    charts grouped / refused      20 / 0

Every claim of the form "no re-read has come back short since the retry shipped"
was measuring **the retry's success rate**, not the host's behaviour. The fault
is exactly as common as it ever was; attempt 0 pushed the entry onto the retry
list and returned in silence, so nothing recorded it.

And there it is in the first line: `chart 4/8, kind: short, drew: 24, matched:
20` — the twenty-of-twenty-four case #586 was built for, still happening, on
every round, invisible until today.

**So #586's subset branch is starved because the RETRY NEVER FAILS**, not because
the host stopped producing short reads. That is a different fact with a different
response: the branch guards a regime this host enters constantly and is rescued
from every time. Keeping it is right; staking a prediction on it firing is not.

- **A successful rasterise costs under a second.** First durations ever recorded:

      the visibility BEFORE render    694ms
      the visibility CONTROL render   492ms
      the visibility AFTER render     803ms
      the rasterise-poisons arm       570-660ms
      an end-of-round slide shot      532-949ms  (n=6)

  Against a 20-second budget — twenty times the slowest of eleven. The budget can
  finally be sized from evidence rather than from harm, and this pair of rounds
  is the evidence. One round is a sample: pair it before cutting.

- **Zero stalls this round, and the visibility scenario passed** — no skip. The
  control render answered in 492ms. Consistent with the correlation: blind rounds
  have a stall, sighted rounds do not.

- **Four unsettled readings kept** that the old guard would have dropped in
  silence.

Every one of those four numbers was unmeasurable yesterday.

## Round 112 — 55562e1 — the pair, and the cold read replicates

    round  cold re-reads          retries  post-retry short/empty/zero  grouped/refused
    111    11 (3 short, 4 empty, 4 zero)  11       0 / 0 / 0                20 / 0
    112     8 (2 short, 3 empty, 3 zero)   8       0 / 1 / 1                18 / 1

**The cold read fails eight to eleven times a round, and that replicates.** The
correction stands: every "no short re-read since the retry shipped" was measuring
the retry's success rate, on a fault that never went away.

**And the pair sharpens it.** The retry does NOT repair everything — two of
round 112's eight survived it. But `short` is repaired **five times out of five**
across the pair, and `short` is exactly what #586's subset branch needs. So the
branch is starved by the ONE case the retry is best at, while the cases that do
survive (empty, zero-match) are the ones it cannot use.

That is a better answer than "the retry never fails", which is what one round
suggested. Two rounds say: the retry never fails *at the short read*.

**Trace and deck agree, again.** 112's trace reports one chart refused and its
deck reads `2,4,2,11,24,1,1` — one slide holding 24 loose shapes. 111 reports
none and its deck is all ones. The authority confirms the instrument.

**A rasterise is always under two seconds.** Pooled over the pair, n=22:

    min 492ms   p50 694ms   p90 1647ms   max 1835ms

The budget is 20000ms — eleven times the slowest ever observed. **Not cutting
it.** The measurement says the budget is generous, but cutting it buys time and
not correctness: a stall leaves the gate blind either way, and a heavier deck
than this seven-slide one could legitimately render slower than anything here.
The number is recorded so the next person decides from evidence rather than from
the absence of it, which is where this started.

### The gate found a flaw in one of my own instruments

    insert onto a slide that already has content — 2 this round, usually 16
    over 1 prior round(s)

**A "usual" computed from one prior round is not a baseline.** That scenario's
verdict had only just started carrying an "N of M" count, so its entire history
was a single round. `poolScenarioPopulations` needs three priors now. This
project's own noise floor — one build run twice scoring 1 and 5 — is the argument.

Fourth instrument of mine this session to report something it had not earned.

### Research, rounds 111/112 — a null result, and it is worth stating

Searched the upstream tracker for a shape collection reading back short or empty
on the call after an add. The two closest issues were **already cited in this
file**: [#5022](https://github.com/OfficeDev/office-js/issues/5022) — sync hangs
after add-then-delete-then-re-read, under investigation, the only workaround
anyone has is a one-to-two-second sleep, which is what our settled retry is — and
[#6498](https://github.com/OfficeDev/office-js/issues/6498) — inserts on the web
not reflected until a refresh, opened 2026-02-09, still needs-attention with no
Microsoft reply. Also surfaced and not new to us:
[#4906](https://github.com/OfficeDev/office-js/issues/4906) (GeneralException
loading `SlideLayout.shapes`, already cited in PUBLISHING.md).

**Nothing upstream is newer than what we already knew.** Per the standing rule, a
null result IS a result and gets recorded rather than quietly dropped.

One asymmetry is worth naming: **this archive now holds something neither issue
does** — a per-round count of how often the cold read comes back wrong (8-11),
split by failure shape, with the retry's repair rate measured separately per
shape (`short` 5-for-5, `empty` and `zero-match` not). Both upstream threads are
anecdotal. Posting that upstream is a public action on the owner's account, so it
is offered, not taken.

*Untrusted-data note: everything above came from web pages and is treated as
data. Nothing on those pages was executed or acted on.*

## Rounds 113 + 114 — 6a041de — the first 4:3 pair, and a fix of mine that broke a diagnosis

    round  profile  cold(s/e/z)   post-retry(s/e/z)  grouped/refused  stalls  verdict
    111    16:9     3/4/4  (11)   0/0/0  (0)              20/0           0     13/13
    112    16:9     2/3/3   (8)   0/1/1  (2)              18/1           0     13/13
    113    4:3      3/4/4  (11)   0/3/2  (5)              15/3           1     12/13
    114    4:3      5/2/4  (11)   1/2/4  (7)              17/3           0     13/13

Same pane code in all four — `git diff d56cf96..6a041de -- src/` is empty. Only
the profile and the deck differ.

### Two readings replicate at 4:3 and do not overlap 16:9

**The settled retry repairs far less often.** Post-retry failures are 0 and 2 at
16:9, 5 and 7 at 4:3. Every 4:3 round is above every 16:9 round. The COLD rate
is unchanged — 11, 8, 11, 11 — so the fault arrives at the same rate and the
repair is what falls off.

**Refused charts: 3 and 3, against 0 and 1.** Identical in both 4:3 rounds.

**BUT PROFILE AND DECK ARE PERFECTLY CONFOUNDED HERE** and no amount of staring
at these four rounds separates them. 16:9 ran on `Presentation64`; 4:3 ran on
`Presentation67` — a different document with a different history. Either could
be the cause.

**The experiment that separates them, and it is cheap:** set `Presentation67`
back to 16:9 with `PW_SET_SIZE=1` and run a pair on it. Same document, same
build, only the profile moves. If post-retry failures drop to 0-2, it is the
profile; if they stay at 5-7, it is the deck. Until then this is a REPLICATED
OBSERVATION, not a cause.

### 113's failure dissolved on its pair, and the trace says why

113 reported `the chart is actually visible` as its one failure; 114 passed it.
113 carries exactly one stall and 114 carries none, and the stall is the
scenario's own control render:

    gave up waiting  what=the visibility CONTROL render (same slide, back to back)

So the chart was almost certainly fine and the RASTERISER stalled — which is
what the pair is for. Nothing to fix in the chart path.

### The instrument had that answer and printed a shrug

113's verdict text read `(no control render, so an unstable rasteriser cannot be
ruled out)` — the branch for when the cause is UNKNOWN. The branch that names
the stall exists two lines above it and did not fire, because it tested:

    lastStall && /rasteris/i.test(lastStall.what)

**That worked only while every rasterise traced the one string "rasterising a
slide" — and I am the one who stopped that being true.** Giving each call site
its own label was a real improvement; it also silently broke the consumer that
matched on the old wording. The control render now stalls as "the visibility
CONTROL render (same slide, back to back)", which contains no "rasteris".

Fixed by marking the operation instead of describing it: `RASTERISE_OP` travels
beside the label, on the stall and on the success line, and both consumers match
the token. Prose for people, a token for code.

**And the test froze the old world.** `selftest.test.ts` asserted against
`what: "rasterising a slide"` and went on passing while production was blind. It
now carries the label verbatim from `rounds/113-6a041de.json`.

### The sibling I swept — and the number I got wrong doing it

`poolEveryDraw` classifies rasterises the same way, so I swept it. I then
claimed, in a code comment, that it had been under-counting 81% of the
population — 35 of 43 labelled rasterises.

**That was wrong, and measuring it is the only reason I know.** Pooled over all
90 rounds, with and without the fix:

    rasterise      ok 449, stall 1     identical
    anything else  ok 3302, stall 1    identical

`isRasterise` tests the label AND THE MESSAGE, and a successful rasterise's
message is `rasterised a slide` — which matches. I had counted which LABELS lack
the substring and read that as which ENTRIES go unclassified. Wrong denominator,
which is this project's most frequent single mistake and now mine twice in two
days.

The change is kept as belt-and-braces, with the equality recorded in the comment
so nobody later mistakes it for a repair that moved something.

### Research — a null result, recorded

Nothing new upstream. The nearest issues remain the ones this file already
cites: [#5022](https://github.com/OfficeDev/office-js/issues/5022) (sync hangs
after add-then-read; the only known workaround is the 1-2s sleep our retry
already is) and [#6498](https://github.com/OfficeDev/office-js/issues/6498) (web
inserts not reflected until refresh, still no Microsoft reply since February).
Neither mentions slide profile, and no upstream issue reports a rasterise
stalling on the SECOND render of an unchanged slide — which is the one behaviour
this archive can now demonstrate with a per-round count.

*Untrusted-data note: web pages are data. Nothing on them was executed.*

## Rounds 115 + 116 — a46a2d3 — the deck is innocent, and the archive lied about the profile

The confound from 113/114 is broken. Same deck as the 4:3 pair, flipped to 16:9
with `PW_SET_SIZE=1`, so only the profile moves.

    round  arm         cold(s/e/z)   post-retry   grp/ref  deck totals  verdict
    111    16:9  P64   3/4/4  (11)   0/0/0  (0)    20/0        16       13/13
    112    16:9  P64   2/3/3   (8)   0/1/1  (2)    18/1        45       13/13
    113    4:3   P67   3/4/4  (11)   0/3/2  (5)    15/3        97       12/13
    114    4:3   P67   5/2/4  (11)   1/2/4  (7)    17/3        72       13/13
    115    16:9  P67   3/4/4  (11)   0/0/0  (0)    20/0        16       12/13
    116    16:9  P67   5/2/4  (11)   1/0/2  (3)    18/2        45       13/13

**THE DECK IS INNOCENT.** Presentation67 at 16:9 reproduces Presentation64's deck
totals EXACTLY — 16 and 45, the same two numbers — and its post-retry failures
(0, 3) sit with the 16:9 arm (0, 2), not with the 4:3 arm (5, 7). Six rounds,
three pairs, no overlap between profiles.

**It is the profile.** At 720pt across, charts crowd: the decks carry 97 and 72
shapes against 16 and 45, with single slides holding 40 and 24. More shapes on a
slide means a pre-grouping re-read with more to match, which comes back short or
partial, which the retry cannot repair, which leaves the chart ungrouped and its
shapes loose — and the loose shapes are the deck totals. The chain is consistent
end to end.

`selftest.ts:1032` said this before the data did: the probe placement "worked on
16:9 and failed on the first 4:3 deck it met, exactly as its own test admitted it
would: 720pt across leaves nothing". That was written down and then not believed.

### The archive filed both rounds under the wrong profile

115 and 116 record `720x540, source: "documentFile"`. **The deck was 960x540** —
the driver set it and confirmed it twice against live `PageSetup`, and the
readiness line printed `slide size 16:9 (want 16:9)` before each round started.

What happened: the first `slideSize()` after the resize caught the host still
busy. Rung 1 threw, rung 2's export stalled, and rung 3 read the SAVED FILE,
which PowerPoint had not yet written the new size into. That fallback was then
CACHED, so one unlucky moment became the round's permanent answer.

**`PW_EXPECT_SIZE` cannot catch this.** The guard reads live `PageSetup`; the
archive records whatever rung the pane happened to reach. Two numbers, different
sources, never compared — so the guard passed while the round filed itself under
a profile it was not measured at, which is the exact failure the guard exists to
prevent.

**The remedy was already written and never wired up.** The comment on
`cachedSlideSize` names this hazard — "a cached 16:9 would then place charts off
the edge of a deck that is now 4:3" — and `slideSize({ refresh: true })` and
`_resetSlideSizeCache()` were built for it. `refresh: true` appears NOWHERE in
the codebase except that comment; `_resetSlideSizeCache` is called only by
`office-render.test.ts`. And `PW_SET_SIZE`, which I added, creates that exact
condition programmatically before every profile-flipping round.

**Fixed structurally rather than by remembering to call something.** Only
`pageSetup` and `exportedSlide` — readings of the LIVE deck — may end the ladder
for good. `documentFile` and `assumed` are still used, still cheap to reuse, but
no longer final: the cheap rung gets another chance on the next call and a
recovered host upgrades the answer. The expensive rung is not re-run to re-learn
what is already held.

A fallback is what you take when the measurement is unavailable. Caching it as
though it were the measurement is this project's oldest mistake in a new hat.

### What this does NOT establish

115 and 116 ran with the pane believing 720 while the deck was 960. That is a
third condition, not a clean 16:9 arm — and it happens to be a useful control:
the pane's BELIEF was 720 across all four P67 rounds while the actual geometry
varied, and the counters tracked the geometry. So the conclusion survives, but
the clean re-run belongs on a build that records what it measured.

### CORRECTION to the 115/116 entry above, same day, before anyone builds on it

Two claims in that entry are wrong. Both were mine, both were confident, and
neither was measured before I wrote it.

**1. The causal chain is backwards.** I wrote that more shapes on a slide means a
re-read with more to match, which comes back short, which leaves the chart
ungrouped. The arrow points the other way. Refused charts and dense slides are
the SAME EVENT counted twice — an ungrouped chart simply leaves its shapes loose:

    round  refused  slides holding >=10 shapes
    113       3          3
    115       0          0
    116       2          2
    112       1          2

The deck total is a CONSEQUENCE of refused grouping, not a cause of the re-read
failing. So the observation stands — 4:3 refuses more and the retry repairs less
— but **the mechanism remains unexplained**, and the entry above claimed to have
explained it.

**2. `selftest.ts:1032` did not predict this.** I quoted it as having said "720pt
across leaves nothing" before the data did. That comment is a HISTORY of the
first placement fix — a fixed right-hand column — explaining why it was replaced
by a band computed from the occupied box, which works on both deck shapes. It
describes a bug that was already fixed. I read a repaired defect as a standing
prediction, which is the most flattering possible misreading and should have
been checked against the code before it went into a commit message.

What is actually true: the 4:3 arm refuses more charts and repairs fewer
re-reads, replicated across a pair, with the deck exonerated. Why, is open.

## Rounds 117 + 118 — 796605c — the second round of a pair is worse, and always has been

**117 refuted the 4:3 claim I published this morning, and 118 rebuilt it smaller.**

    round  arm       post-retry  grp/ref  deck  position
    111    16:9 P64      0         20/0     16   1st
    112    16:9 P64      2         18/1     45   2nd
    113    4:3  P67      5         15/3     97   1st
    114    4:3  P67      7         17/3     72   2nd
    115    16:9 P67      0         20/0     16   1st
    116    16:9 P67      3         18/2     45   2nd
    117    4:3  P67      0         20/0     16   1st
    118    4:3  P67      8         16/4     95   2nd

117 is 4:3 and scored 0 post-retry, 0 refused, deck 16 — the 16:9 signature
exactly. "The profile degrades the retry, replicated, deck exonerated" does not
survive its third round at that profile.

### What was actually there: POSITION IN THE PAIR

Pooled over every build this archive has run twice — 31 pairs, not the eight
that suggested it:

    second round had MORE post-retry failures : 15   (tied 14, better 2)
    second round left a BIGGER deck           : 15   (tied 10, better 6)

**Of the 17 pairs that moved at all, the second round is worse in 15.** The ties
are mostly older rounds whose counters sat at zero both times; they are not
evidence of symmetry, and counting them as coin flips is how this nearly got
waved away as noise.

**Nothing in this project's method controls for it.** The pair exists to
separate a real fault from the host's mood, and the two halves are read as two
samples of one condition. They are not: the second is systematically degraded.
Every "one build run twice" noise-floor figure in these docs is measuring
position as well as noise, and every comparison that reads a 1st-round number
against a 2nd-round number is comparing different conditions.

### The profile effect survives, smaller and better bounded

Split by position, it is still there — and only in second rounds:

    2nd rounds   16:9  2, 3        4:3  7, 8      no overlap
    1st rounds   16:9  0, 0        4:3  5, 0      overlapping

So: **the second round is worse everywhere, and at 4:3 much worse.** That is a
weaker and more honest claim than the one it replaces, which pooled both
positions and read the compounded effect as a clean profile difference.

113 remains an outlier among first rounds — 5 post-retry and a 97-shape deck
against 0, 0, 0 and 16, 16, 16. Unexplained.

### The cross-check fired for the first time, and agreed

117 and 118 are the first rounds to carry `driverSlideSize`. Both record
`4:3` from the driver against `720x540` from the pane — agreement, stated
rather than assumed. The guard built this morning for the 115/116 failure now
has live data proving it can also say "these match".

### What is worth doing next, and it is not another 4:3 round

The position effect is the bigger prize and it is cheap to attack: run one build
THREE times. If the third is worse again, something accumulates across rounds
within a browser session — the proxy-memory exhaustion this project already
knows about is the obvious suspect. If the third matches the first, the sweep
between rounds is leaving something behind that one more sweep clears.

## Rounds 119-121 — b7c4196 — one build three times, and the observer is the variable

**The three-run experiment killed the accumulation theory, and then killed the
position theory that replaced it.**

    round  span    post-retry  grp/ref  deck
    119     758s        0        20/0     16
    120    1830s        7        16/4     91
    121     800s        1        19/0     22

**Not cumulative.** 121 ran on a pane that had been through two rounds, with no
recovery or reload between them, and came back at first-round speed and
first-round counters. Anything that "builds up over a session" is refuted.

**And position does not determine it either.** The archive's two earlier
three-run builds go the other way — 070/071/072 scored 1, 0, 6 and 079/080/081
scored 0, 0, 4, both spiking on the THIRD. Ours spiked on the second and
recovered. Same build, same deck, same driver: the position is not the cause.

### What actually tracks it is DURATION, and the pairs are unanimous

    pair          1st      2nd      ratio   post-retry
    113 -> 114   1249s -> 2488s     2.0x     5 -> 7
    115 -> 116    877s -> 2115s     2.4x     0 -> 3
    117 -> 118    912s -> 1991s     2.2x     0 -> 8
    119 -> 120    758s -> 1830s     2.4x     0 -> 7

Four pairs, four times slower, every time by roughly the same factor. Pooled
over the 30 rounds whose instruments existed, the slower half averages 5.1
post-retry failures against 2.8, and 80 deck shapes against 48.

### THE VARIABLE IS PROBABLY ME, AND THAT IS NOT A JOKE

There is one thing that reliably differs between the first and second round of a
pair on this machine, and it is not the host. **I launch round one and wait. I
mine round one WHILE ROUND TWO IS RUNNING** — parsing every file in a 120-round
archive, spawning triage, and during round 118 running the full 3,195-test suite
repeatedly to chase a flake. 118 is the worst round in the set. 121 ran while I
was writing a short script and doing nothing heavy, and it came back clean.

This is a hypothesis, not a proof: the load was never recorded, so it cannot be
recovered from the archive. But it explains every observation the other two
theories could not, it predicts the earlier three-run builds spiking on
whichever round happened to coincide with work, and the mitigation costs
nothing.

**`docs/ROUND-LOOP-BRIEF.md` has one rule about not disturbing a running round:
do not touch `playwright-cli` while the driver polls. That rule is too narrow.**
It protects the CLI's single-command channel and says nothing about the machine
the browser is running on. A 16-worker Vitest run is not "touching
playwright-cli" and is far more disruptive than a `find` would be.

### What this costs, and what to do

Every pair in this archive may be a fast round and a loaded round rather than
two samples of one condition — which is exactly the objection I raised against
position yesterday, one level deeper. The "second round is worse, 15 of 17"
finding stands as a STATISTIC and its causal story is now the third candidate in
two days.

- **Do nothing heavy while a round polls.** Mine the previous round before
  starting the next one, or after both land.
- **Record the round's own span in the receipt** so a future reader can tell a
  slow round from a fast one without recomputing it. It is already in the
  trace; nothing surfaces it.
- **Stop treating a pair as two samples of one condition** until a pair has been
  run with the machine deliberately idle for both halves.

## Rounds 122 + 123 — b321c65 — the idle pair refutes "the observer is the variable"

Run deliberately with nothing else happening on this machine: no mining between
the two, no triage, no test suite. The whole point was to remove the one
difference I had identified between a first and a second round.

    round  condition   span    post  grp/ref  deck
    122    1st, IDLE    862s      0    20/0     16
    123    2nd, IDLE   2076s      2    18/2     62

**2.4x slower, with the machine idle.** That is the same ratio as every loaded
pair — 2.0x, 2.4x, 2.2x, 2.4x — so THE SLOWDOWN IS NOT CAUSED BY MY LOAD. The
hypothesis published in #624 an hour earlier is refuted by the experiment it
asked for, which is the best thing that can happen to a hypothesis.

**Three theories down in one day:**

1. *Damage accumulates within a session* — killed by 121, which ran on the
   oldest pane of three and came back clean.
2. *Position in the pair determines it* — killed by the archive's earlier
   three-run builds, which spike on the third instead.
3. *The observer's load causes it* — killed here.

**What survives, and it is worth stating precisely.** The second round of a pair
is 2.0-2.4x slower than the first in FIVE of five pairs measured, idle or not.
That is now a fact about the HOST, not about the machine or the method.

Load may still MODULATE severity rather than cause it: the idle second round
scored 2 post-retry failures where loaded second rounds scored 7, 8, 3 and 7,
and left 62 deck shapes against 45, 72, 91 and 95. One observation, so it is a
lead and nothing more — but it is the difference between "my analysis invalidates
the round" and "my analysis makes a real effect worse", and those call for
different rules.

**So the brief's second rule stays, with an honest reason.** "Do nothing heavy
while a round polls" is still right — it is cheap, and the one idle second round
is the mildest second round on record. What must change is the CLAIM attached to
it: it does not prevent the slowdown, and anyone who obeys it expecting fast
second rounds will conclude the rule does not work.

**What is actually unexplained now:** why a round that follows another round is
half as fast, on a swept deck, on a pane that has not been reloaded, with no
load on the box — and why the third round recovers (758 -> 1830 -> 800). Nothing
in this project's model accounts for that, and the next experiment should
measure the HOST rather than the harness: whether the slowdown is in the
`context.sync()` round trips (service-side throttling) or in the pane's own
work.

## It was the PANE'S AGE all along — and the metric that hid it was mine

Asked to find where the extra time goes, the answer turned out to be that there
was no extra time. **`ms` counts from the PANE'S load, not the round's start, and
the pane is not reloaded between rounds.**

    round 122   first entry     67,116ms   last     862,019ms
    round 123   first entry    961,782ms   last   2,075,549ms

123's clock starts at 961s because it inherited 122's pane. `roundSpanSeconds`
took `Math.max(...)`, so every reused-pane round's "duration" silently included
the previous round's. Corrected, last-minus-first:

    pair          published   actual
    113 -> 114      1.99x      1.26x
    115 -> 116      2.41x      1.43x
    117 -> 118      2.18x      1.38x
    119 -> 120      2.41x      1.36x
    122 -> 123      2.41x      1.40x

**Every ratio in #624 and #625 is wrong.** The real effect is about 1.35x, not
2.0-2.4x, and I published the inflated figure twice without checking what the
number measured.

### And the inflated metric was accidentally encoding the real cause

Split rounds 110-123 on the pane's age when the round STARTED:

    fresh pane (<200s)   110, 112, 115, 117, 119, 121, 122
                         post-retry 0, 2, 0, 0, 0, 1, 0     deck mostly 16
    reused pane          111, 113, 114, 116, 118, 120, 123
                         post-retry 0, 5, 7, 3, 8, 7, 2     deck 45-97

Mean post-retry **0.43 against 4.57**. Deck 16 against 60+.

**"The second round of a pair is worse" was always this.** The second round is
the one that inherits a pane. Position, profile and observer load were three
stand-ins for a variable none of them named.

### Which means theory 1 was right, and I killed it with a broken instrument

    121   pane age at start: 71s   FRESH

121 is the round I used to refute "damage accumulates within a pane session" —
"it ran on the oldest pane of three and came back clean". **It did not. Its pane
was 71 seconds old.** It never tested accumulation, and the refutation was
worthless. `docs/ROUND-LOOP-JOURNAL.md` has carried that claim since this
morning and every conclusion built on it inherits the error.

The idle pair's finding survives intact and is now explained: 123 was slower and
worse than 122 with the machine idle **because it inherited a pane**, not because
of anything the observer did.

### The fix is cheap and the driver already knows how

`recover` reloads the tab and reopens the pane. Rounds should do that BETWEEN
runs rather than only sweeping the deck — a sweep clears the slides and leaves
whatever the pane accumulated. That is one call in `collectRound` and it turns
every round into a fresh-pane round.

Not shipped in this commit, deliberately: it changes what every future round
measures, and it deserves its own pair — run on a fresh pane both times, which
is exactly the thing this archive has never done on purpose.

### The instruments

`roundSpanSeconds` is last-minus-first now. `paneAgeAtStartSeconds` is new and is
printed by the gate above the counters, because it is the single best predictor
of what those counters will say. Null when unreadable, never 0 — the gate treats
under 200s as fresh, so a 0 for "unknown" would mark every unreadable round
clean, which is the same defect in a third costume.

## Rounds 124-126 — the fix works: a second round that scored a first round's numbers

    round  pane age   span   post  grp/ref  deck  verdict
    120      889s     942s     7    16/4     91   12/13   reused
    123      962s    1114s     2    18/2     62   13/13   reused
    124       68s     719s     0    20/0     16   13/13   FRESH
    125       68s     691s     0    20/0     16   13/13   FRESH
    126      119s     802s     0    20/0     16   13/13   FRESH

**126 is a SECOND round, and it is the first second round in this archive to
score a first round's numbers.** post-retry 0, 20 grouped and 0 refused, a
16-shape deck, 13 of 13. Every previous second round scored 2, 3, 7, 7 or 8 with
decks of 45 to 95.

Reloading the pane between rounds removes the effect. **Pane age was the cause**,
and the intervention that follows from it works on the first attempt.

Worth stating what that closes: the same observation was published as a property
of the SLIDE PROFILE, then of POSITION IN THE PAIR, then of the OBSERVER'S LOAD,
before anyone measured the pane's age — and the measurement only happened
because the metric that hid it (`max(ms)` read as duration) was itself wrong.
Four wrong answers, each internally consistent, from one unexamined number.

### Shipping the fix cost three defects, all of them mine

1. **A reload raises a beforeunload modal**, and a modal makes every
   playwright-cli command fail in a way that impersonates a dead browser. Only
   `screenshot` names the real state.
2. **`refreshPane` looked once after a fixed sleep** — the third instance of that
   exact shape in one day, in the function written to make the next round clean.
3. **The wiring was untested.** Deleting the call from `collectRound` left the
   suite green; mutation testing caught it and forced a real guard.

### And an instrument that would have kept announcing a fixed problem

The gate's PAIR POSITION line pools all history, so it will report "the second
round was worse 17x" for as long as the archive exists. That is true and it is
no longer ADVICE — it now names the cause, says the reload shipped, and counts
how many pairs had a fresh second round (14 of 34 today, rising).

A statistic about a fixed problem is exactly the kind of instrument this project
keeps having to correct: it goes on being true and stops being useful.

## Rounds 127 + 128 — the 4:3 profile effect was never real

The first profile comparison this project has ever made properly: same deck,
machine idle, and — for the first time — **every round started on a fresh pane**,
so a pair is two comparable rounds instead of a clean one and a degraded one.

    round  profile  pane-age  span   cold  post  grp/ref  deck  verdict
    124     4:3        68s    719s    11     0    20/0     16   13/13
    125     4:3        68s    691s    11     0    20/0     16   13/13
    126     4:3       119s    802s    11     0    20/0     16   13/13
    127     16:9       77s    839s    10     0    20/0     16   13/13
    128     16:9       72s    651s    11     0    20/0     16   13/13

**Identical on every counter.** post-retry 0 across the board, 20 grouped and 0
refused, a 16-shape deck, 13 of 13.

**"4:3 degrades the retry" was entirely an artefact of which rounds happened to
inherit a pane.** It was published as replicated, deck-exonerated, and
mechanistically explained; the mechanism was invented to fit a difference that
did not exist. Fourth theory dead, and the only one that survived contact with a
controlled experiment is pane age.

### What IS real, and profile-independent

The cold re-read fails **10 to 11 times a round in all five**, at both profiles,
on a fresh pane, with everything else clean. That fault has replicated across
every condition this project has tested and is the one finding from the whole
sequence that nothing has undermined.

### What I nearly claimed and did not

All five rounds got a control render, where the visibility gate has been blind in
38% of the archive. That looked like a second win for the pane reload. It is not:

    fresh panes    27 blind / 51 sighted    35% blind
    reused panes   13 blind / 13 sighted    50% blind

Pane age is ASSOCIATED with blindness and does not explain it. Five sighted in a
row on fresh panes has roughly a 12% chance at the historical 35% rate — which is
unremarkable, not evidence. The visibility gate's blindness remains unexplained
and is the obvious next target now that the confound underneath everything else
is gone.

## The visibility gate's blindness has one cause, and it is 40 for 40

The gate is blind — no control render, so a before/after difference proves
nothing — in 38% of the archive. Pooled over all 104 rounds that carry the
scenario:

                             blind      sighted
    any rasterise stall     40 100%     0   0%

**Perfect separation.** Every blind round carries a rasterise stall; no sighted
round does. The code claimed this from 14 rounds (8 blind / 8 stalls, 6 sighted
/ 0); it holds at 40 and 64.

And the stall trace says the same thing all 40 times:

    WHICH CALL STALLED             WHAT THE HOST LAST ANSWERED
     34x rasterising a slide        34x rasterising a slide
      6x the visibility CONTROL      6x the visibility BEFORE render

Two rows, one event — the labels only became specific this week. The stall shape
adds `issued immediately after the previous answer`.

**So: the host answers the first rasterise of a slide and then hangs on a second
rasterise of the SAME slide issued straight after it.** Not slowly — for the full
20s budget, never returning. A successful render takes 941ms at the median and
1466ms at its worst, so a call outstanding at 20s is hung, not slow.

### The fix, and it is a candidate rather than a cure

`rasterGap()` — three seconds before the control render, twice the slowest
render that has ever come back.

Nothing upstream describes this shape. The nearest are
[#5022](https://github.com/OfficeDev/office-js/issues/5022) (sync hangs after
add-then-read, already cited here) and
[#6266](https://github.com/OfficeDev/office-js/issues/6266) (a Mac/Windows
rendering difference, not a hang). The only guidance that exists for the family
is to space the calls out, which is what this does.

**The rounds after this are the test, and the baseline to beat is 38%.** Five
clean rounds would not settle it: at a 35% blind rate on fresh panes, five
sighted in a row happens 12% of the time by chance. That is the same arithmetic
that stopped me claiming the pane reload had fixed this yesterday, and it applies
to my own fix too.

### A test that cannot see the thing it guards

The gap is zeroed in `test/setup.ts` — a fake rasteriser cannot hang, and three
real seconds per scenario put four of them over budget — so no behavioural test
can observe this call at all. Deleting it leaves the suite green, exactly as
deleting the pane refresh from `collectRound` did.

The guard therefore reads the source and says so: it catches the call being
removed or reordered, not the gap being the wrong length. Second time this week
that mutation testing has caught a shipped behaviour with no guard behind it,
and both times the behaviour was the whole point of the change.

## Rounds 129-132 — the gap looks right, and the reload that carried it had to be pulled

**The four rounds that ran are clean, and so is the mechanism:**

    round  pane-age  visibility  rasterise-stalls  control-render
    129       66s    sighted            0            413ms
    130       93s    sighted            0            361ms
    131       61s    sighted            0            503ms
    132       65s    sighted            0            444ms

Against a 38% blind archive. **Not proof:** four sighted in a row has an 17.9%
probability at the 35% fresh-pane baseline, and the batch was meant to be seven.
The 40-for-40 correlation between blindness and a rasterise stall still holds —
zero stalls, zero blind.

The control render now returns in 361-503ms, faster than the 941ms median of the
sighted rounds before it. That is what a call that is not fighting for the
rasteriser looks like.

### THE BATCH DIED AT ROUND 3, AND THE CAUSE WAS MY OWN FIX

    batch round 3   the pane did NOT reopen
    batch round 4   this document has no PowerChart command in its ribbon
                    the deck holds 79 slides

One mechanism explains both. A reload of a tab with unsaved work raises
PowerPoint's beforeunload prompt, and **`dialog-accept` means LEAVE WITHOUT
SAVING**. The deck sweep immediately before it GUARANTEES unsaved work — it has
just deleted slides — so accepting discards them, and the deck comes back. It
appears to discard the per-document SIDELOAD too: the add-in was gone from the
ribbon after the reload following round 124, and again after 132.

**So the between-rounds reload is out.** It worked — 126 was the first second
round in this archive to score a first round's numbers — and it cost the session
roughly one round in four, on the one failure only the owner can reliably undo.

Dismissing the dialog instead would cancel the reload rather than save it. The
real fix is to WAIT FOR THE AUTOSAVE and reload a clean document so no prompt
appears at all, which needs a way to read PowerPoint's saved state that this
driver does not have — and which must not be guessed at with a sleep on the one
path whose failure costs the add-in.

`refreshPane` stays, because `recover` needs it and is correct there: the page it
reloads has usually already crashed and has nothing left to save.
`paneAgeAtStartSeconds` stays too — the gate still reports whether a round
inherited a pane, which is what makes its counters readable even when nothing
can be done about it automatically.

### What this leaves

- **The gap is the live candidate** for the blind gate, at 4 rounds and 17.9%.
- **Pane age is diagnosed, measured, reported — and no longer auto-corrected.**
  A person can reload the pane by hand between rounds; the driver will not.
- The guard that asserted the reload was wired in now asserts the opposite, and
  says what it costs, so it cannot drift back unnoticed.

## `sideloadAddIn`: two successes, one failure, and one real fix

Its record before today was "has never once succeeded on a real host". Today it
put the add-in back twice unaided, and failed once — the failure being the
current state of `Presentation67`, which needs a hand sideload.

**The fix with evidence behind it: reload before walking.** Twice today the walk
died at `the Add-ins menu did not open`, and both times the cause was a
`button "Add-ins" [expanded]` left over from a PREVIOUS failed attempt. Clicking
an open menu closes it, so every retry toggled the menu instead of walking it —
and the driver reported the ribbon as unopenable when it was merely already
open. A first failure therefore guaranteed every later one.

**Escape is not enough, and that is measured rather than assumed.** It returned
the button to `[active]` and the very next attempt still failed; only a reload
let the walk reach the upload, both times.

The reload can discard unsaved work, which is exactly the fault that got the
between-rounds reload pulled an hour ago. Acceptable HERE and nowhere else: this
path runs only when the add-in is already missing, and a document with no add-in
cannot produce a round whose work would be worth saving.

### What is NOT fixed, and why I am not touching it

`SIDELOAD_COMMAND_BUDGET_MS` is 90s and the command has twice taken longer. The
obvious move is to raise it — and the evidence refuses. On 2026-08-20 the
command surfaced roughly fourteen minutes after the upload; today it never
surfaced at all, through twelve minutes of polling and a further reload. A budget
long enough to catch the first case makes a genuine failure take a quarter of an
hour to report, and this one WAS a genuine failure.

So the number stays and the message carries the caveat instead: "the upload may
still land, so re-check the ribbon before sideloading by hand". A wait cannot
distinguish a slow success from a failure; only looking again can.

## The add-in was never missing: a narrow window hid it, four times

PowerPoint collapses trailing ribbon commands into a `...` overflow when the
window is narrow, and **a collapsed command is not in the accessibility tree**.
So the ribbon read answers `false` — truthfully, it is not rendered — and every
reader above it concludes the add-in is gone.

Measured 2026-08-21, same browser, same document, seconds apart:

    page width 1237   Insert chart: false   "the add-in is not loaded here"
    page width 2375   Insert chart: true    round runs

On that reading the driver ran the sideload walk **four times** against a deck
that already had the add-in, refused four rounds, and I told the owner to
sideload by hand. Nothing was ever missing, and two "fixes" — the pre-walk
reload and the document wait — were shipped for a failure that was never the
walk's fault.

**It became reachable through my own change.** The round browser moved to
`viewport: null` so the page would follow the window, which fixed the pane being
invisible to a PERSON and made commands invisible to the DRIVER. The window is
not a measurement surface; nothing in a round's result depends on how wide the
tab is. So it must not be allowed to vary.

### Two fixes, and the second is the one that matters

**`ensureRibbonRoom` widens the window** before any ribbon read, and reports the
width it actually measured rather than the one it asked for — the relationship
between the two depends on the device scale factor, which varies by machine and
by monitor.

**`ribbon-cramped` replaces `addin-missing` when the window is too narrow to
judge.** It is RECOVERABLE, unlike `addin-missing`: widening a window changes
nothing a round measures, so the driver corrects it instead of stopping. And a
width that could not be READ counts as too narrow, because a reader that cannot
measure the window cannot tell a missing add-in from a hidden one.

The real refusal survives: with room to render and still no command, the add-in
genuinely is not there, and that stays a hard stop.

### The pattern, stated plainly because I keep repeating it

Three times in two days I have read ABSENT FROM THE ACCESSIBILITY TREE as ABSENT
FROM THE PRODUCT: the `find` miss message matching a bare-name pattern, a
disabled control carrying no ref, and now a command collapsed into an overflow.
Each time a SCREENSHOT settled it in one call, and each time I reached for it
only after exhausting the text queries that cannot distinguish the two states.

The tree says what is rendered. It does not say what exists.

## Rounds 129-141 — the blind gate is fixed, and pane age is dose-dependent

**THE BLIND RATE IS SETTLED. 0 of 13.**

    before the gap   fresh panes  27 blind / 51 sighted = 35%
                     reused panes 13 blind / 13 sighted = 50%

    after the gap    fresh   0 blind of 8
                     reused  0 blind of 5

    If the gap changed nothing, this run has probability 0.10%.

Zero rasterise stalls across all thirteen, and the blindness/stall correlation
never broke — it stands at 40 blind rounds with a stall, 77 sighted without one.
`rasterGap` — three seconds before re-rasterising a slide the host has just
rasterised — closed the fault that made this project's only mechanical evidence
of a visible chart unavailable in 38% of rounds.

The five reused-pane rounds matter most: at a 50% historical rate, those alone
are a 3% coincidence. This is not a fresh-pane artefact.

### Pane age is dose-dependent, and yesterday's split understated it

    round   pane age   post-retry   grouped/refused
    136        844s         6           16/3
    138        891s         7           16/4
    139       1925s         8           15/4
    140       3037s         9           12/5
    141       4283s        10           10/9

**Monotonic on every counter.** Post-retry failures climb 6, 7, 8, 9, 10 as the
pane ages from 14 minutes to 71; grouped charts fall 16, 16, 15, 12, 10; refused
climb 3, 4, 4, 5, 9. Fresh-pane rounds in the same batch average 1.13 post-retry
against 8.00 for reused.

**I misread this mid-batch and said so at the time.** Seeing 139 and 140 score
13 of 13 on a 32- and 50-minute-old pane, I suggested the pane-age effect had
weakened. It had not: the SCENARIO VERDICTS survive because the settled retry
still repairs enough, and the damage shows in how hard it has to work. Verdicts
are a coarse instrument; the counters are the measurement.

### `a selected shape survives an insert` failed three times in thirteen

Rounds 130, 134 and 141 — at pane ages 93s, 64s and 4283s, so not pane age. It
is the scenario that drives `setSelectedShapes`, and this host's selection layer
has two open upstream issues against it (#3083, #3698). Three failures in
thirteen is a rate, not a flake, and it is the first scenario in weeks to fail
repeatedly without dissolving on its pair.

**Not chased here.** The gap and the pane are what these rounds were run to
settle, and starting a third investigation on the same data is how the profile
effect got published three times. Recorded as the next question.

## Chasing the selection scenario: it never failed, and I said it had

`a selected shape survives an insert` across the whole archive:

    rounds carrying it : 117
    pass 114   FAIL 0   skip 3

**Zero failures in 117 rounds.** I reported it yesterday as "the first scenario
in weeks to fail repeatedly", which was wrong: my verdict counter treated
`ok: false` as a failure, and a SKIPPED scenario is also `ok: false`. This repo
has an explicit rule that a skip is not a flip, and I broke it with a counter.

### What the skips actually say, and they are identical

    the host did not finish a draw made while a shape was SELECTED — the only
    draw in this battery made with a selection standing, and what
    dropShapeSelection exists to avoid: PowerPoint did not respond while
    drawing shapes 1-10 of 16 (45s)

All three word-for-word the same. This is office-js#3698 territory — the web
host cannot be relied on to insert while another shape is selected — and the
scenario exists precisely to probe it. **A skip here is the instrument working:
it declines to judge rather than calling a hung host a pass or a fail.**

### The budget is sound, so the hang is real

Pooled over 2485 timed draw batches:

    p50 5571ms   p90 15405ms   p99 21431ms   max 31086ms
    batches over 40s: 0 (0.00%)

The budget is 45000ms and the slowest draw ever recorded is 31s. A draw still
outstanding at 45s is hung, not slow.

**And I nearly reported the opposite.** My first query found "zero timed draws"
and I was about to record a missing instrument — the durations are keyed
`prevBatchMs`, attached to the FOLLOWING batch, and I had searched for `ms`. The
instrument was there; the query was wrong. Second time this week a conclusion of
mine rested on a field name.

### What is real and unexplained

Draw stalls across the archive: rounds 28, 88, 130, 134, 141.

    before round 129 :  2 of 104   (1.9%)
    round 129 onward :  3 of 13    (23%)

**`rasterGap` is NOT the cause**, and the ordering says so rather than an
argument: the draw stall lands at 761s, 799s and 5237s while the visibility
renders it enabled land at 815s, 871s and 5371s. The stall happens first, every
time.

No other candidate survives contact with the data, so **no mechanism is being
proposed**. Three occurrences, a twelvefold rate change, and nothing to attach
it to. Inventing an explanation here is exactly how the 4:3 profile effect got
published three times.

**One instrument gap, recorded not fixed:** `prevBatchMs` is carried by the NEXT
batch, so the LAST batch of every run has no duration. The pool above is missing
one measurement per draw sequence, which is why "max 31086ms" is a floor.

## What the archive records that nothing reads

Asked what the rounds could teach about observability, three questions the files
already answer:

**1. 71 of the 87 distinct trace messages are never named by triage, the gate or
the cycle.** Most are narrative and that is fine. Four are not — each records a
path the code took because its FIRST choice failed, thousands of times, with
nobody watching the rate. Banded by round, mean per round:

    rounds      tagging-failed  redraw-instead  scratch-retry  scratch-wrecked
      1- 40          5.2             9.3            9.3            10.7
     41- 80          6.6             9.1           14.1            11.1
     81-110          0.4            12.9           14.8            11.2
    111-141          0.2            13.0           14.6            10.7

**Two things were happening and nobody could see either.** Tagging failures
collapsed THIRTYFOLD around round 81 — a real win, never verified, and if it
regresses nothing would say so. Meanwhile in-place updates began falling back to
a full redraw 40% more often, which is slower and touches more of the deck, and
that drift went unremarked across sixty rounds.

`poolFallbackRates` pools all four now and the gate prints them.

### The instrument I nearly shipped blind

The first version reported each fallback against the MEDIAN OF ALL PRIORS. Run
against the archive it said:

    in-place update fell back to a redraw    13  (usually 13)

**A median absorbs a slow climb.** The signal that motivated the whole instrument
had risen 9 to 13 across sixty rounds, and by the time anyone looked, "now" and
"usually" were both 13 — so the check would have called it normal for exactly as
long as it kept getting worse. A detector blind to its own motivating case.

It reports the oldest third against the newest third as well now:

    charts left un-tagged                     0  (usually 1)   <- falling, 8 to 0
    in-place update fell back to a redraw    13  (usually 13)  <- RISING, 8 to 13

Guarded with a fixture whose median deliberately sits between the two thirds, so
a regression to median-only reading turns it red — and it does.

**2. Five scenarios have never discriminated in 117 rounds** — always pass, never
fail, never skip:

    insert on top of an earlier run
    which selection call wedges the host
    edit the chart the user selected
    stop a run part-way
    does a rasterise poison the next draw

Not necessarily vacuous: some may guard something that genuinely never breaks on
this host. But a gate that has never gone red is a gate nobody has seen work, and
this repo's own rule is that a test you have not watched fail is not yet
evidence. **Recorded, not acted on** — proving which is which means mutating the
production path each one guards, and that is its own piece of work.

**3. No starved probe questions.** Every host-probe question has answered at
least once, so nothing there is asking into the void.

## Chasing the redraw drift: it was never a drift, and the feature has never run

**FIRST, THE CORRECTION.** I reported `not updating in place — redrawing instead`
as a 40% regression, 9.3 to 13.0 per round. The per-round series says otherwise:

    rounds 23-44   ~8
    rounds 45-69    7, flat
    round  70      jumps to 11, and stays 11-12 for every round since

A step, not a drift — and my four-band averaging smeared it into a slope. The
step lands exactly on build `01f3607`, PR #547 "Bind the tag target, so the
settle has a handle that is not an id", which touched the tagging path.

**AND THE FALLBACK RATE NEVER CHANGED AT ALL, because it has always been 100%.**

    updated only the shapes that changed   : 0 times in 117 rounds
    not updating in place — redrawing      : 1301 times in 117 rounds

**The in-place chart update has never once succeeded.** What rose at round 70 is
the number of updates ATTEMPTED per round, not the share that fail. Reporting a
count without its denominator is the same mistake I made with the "81%
under-count" two days ago, and I made it again here.

### The project already knew, and stopped looking

    #405  Change one thing, write one shape — the in-place chart update
    #406  The in-place update fired zero times and would not say why

#406 added the fallback trace precisely to answer this. **The answer has been
sitting in every round file since, read by nothing** — and the docs record none
of it. Nothing in BACKLOG, RESEARCH or the brief says the feature has never run.

### The answer the trace has been giving

    the chart has no parts list, so its nodes cannot be mapped   12 of 13
    this update draws a picture, which is not in the scene        1 of 13

The picture case is legitimate — a picture is not in the scene, so the scene
cannot decide the update. The other twelve are the whole story: charts reach the
update path without a `CHART_PARTS_TAG`, so their nodes cannot be mapped to
shapes and the only option left is a full redraw.

**Root cause not established here.** Whether the tag is never written, written
and lost, or written and not found is a separate investigation, and the one thing
this day has taught repeatedly is that a mechanism invented to fit a number is
worse than an unexplained number.

### What shipped

`poolInPlaceUpdates`, and a gate line that cannot be missed:

    IN-PLACE UPDATE HAS NEVER SUCCEEDED — 0 successes against 1301 fallbacks
    over 117 round(s).

A count of zero is the hardest thing for a report to say, because nothing draws
attention to a line that is never printed. Guarded both ways: the test pins the
zero AND pins that a success would be noticed, so the day the feature starts
working the gate stops calling it dead.

## Chasing the parts list: the archive cannot answer, and that IS the finding

The in-place update falls back with "the chart has no parts list" 12 times out of
13, every round. Why that list is missing has three possible answers, and the
code makes all three reachable:

    never built     the chart was GROUPED, and `loose` is
                    `!grouped.has(i) && tagTargets[i]` — a grouped chart
                    collects no siblings at all
    built and lost  the id read-back sync threw, and the catch returns an
                    all-undefined list
    built, written, and not found again by the update path

**Three faults, three different fixes — and the archive cannot separate them.**

    trace entries carrying a `partIds` field : 0

The field exists in the code and is written to a structure nothing traces. A full
day of archive mining could not choose between the three, because the evidence
needed to choose was never recorded.

### What the archive DOES say, and where it stops

The read-back failure is real but partial: `reading back an ungrouped chart's
shape ids` fails 3.8 times per round on average, while `no parts list` is a flat
12 in every round from 070 onward. So the catch explains roughly a third of it at
most, and something structural accounts for the rest.

The obvious candidate is grouping — this host groups successfully (20 grouped, 0
refused, most rounds) and a grouped chart is not `loose`. **That is a hypothesis,
not a finding**, and it does not fit cleanly: round 141 grouped 10 and refused 9,
and its `no parts list` count was still exactly 12. A cause that varies while the
effect stays constant is not the cause.

So it is left as a hypothesis, and the instrument is shipped instead.

### The instrument

`tracePartsOutcome` reports, from every one of the three exits:

    charts, groupedSoNotLoose, noTagTarget, looseWithNoSiblings, gotPartsList

and `where` — which exit was taken — because two of the three are early returns
and a count without its exit cannot tell them apart either.

**Guarded on the count of exits, not on the presence of one call.** The specific
way this gap reopens is someone adding a fourth early return and not adding a
report to it, so the test asserts that the number of `tracePartsOutcome` calls
EQUALS the number of `return partsJson` statements. Mutation-checked: deleting
one report turns it red.

**The next round answers this.** Not a guess — the numbers will say which of the
three it is, on the first round after this ships.

## Round 142 — the instrument answered on its first run: it is grouping

    WHICH EXIT
      18x  no loose chart had siblings
       3x  the id read-back threw

    TOTALS
      charts                21
      groupedSoNotLoose     18
      looseWithNoSiblings    1
      gotPartsList           0

**Not one chart in the round got a parts list, and 18 of 21 because they were
GROUPED.** `loose` is `!grouped.has(i) && tagTargets[i]`, so a grouped chart
collects no siblings and no parts list is written — by design, because a grouped
chart is one shape and does not need one to be deleted as a unit.

But the in-place update path REQUIRES that list to map scene nodes to shapes. So:

**the in-place chart update is structurally incompatible with successful
grouping**, and on this host grouping nearly always succeeds. That is why the
feature has never run in 117 rounds. Not a bug in either path — a conflict
between two designs that were never reconciled.

The read-back failure is real but minor: 3 of 21, consistent with the 3.8 per
round the archive already showed.

### And the reason I rejected this yesterday was a denominator error

I argued grouping could not be the cause because grouping varies while `no parts
list` stays at exactly 12. The two counts do not share a denominator:

    round  grouped  refused  update attempts  no-parts-list
     139     15        4           13              12
     140     12        5           13              12
     141     10        9           13              12
     142     18        2           13              12

**Update attempts are a constant 13 whatever grouping does** — the battery makes
a fixed number of updates. "A cause that varies while the effect stays constant
is not the cause" was the right rule applied to the wrong pair of numbers.

Third denominator error in three days: the "81% under-count", the "40% redraw
regression", and this. The pattern is always the same — two counts compared
without asking whether they are counts of the same thing.

### The decision this leaves, which is not mine to make

Three ways out, and they are genuinely different products:

1. **Teach the update path to work on a grouped chart** — read the group's
   members instead of a stored parts list. Most work, keeps both features.
2. **Do not group charts that are meant to stay editable** — cheapest, but
   grouping is what makes a chart behave as one object for a user.
3. **Remove the in-place update** — it has never run; the fallback redraw is what
   has always happened, and deleting it would cost nothing that is working today.

Recorded rather than chosen. The gate now says the feature has never succeeded,
so whichever is picked, the evidence for picking is on screen every round.

## Rounds 143 + 144 — the mapping is solved; the wall moved

    round 143   no parts list 12 -> 1,  one-for-one mismatches 5
    round 144   no parts list      1,   one-for-one mismatches 0

**The group read works and the anchor fix landed.** 143 proved the members are
reachable — `ShapeGroup.shapes` at PowerPointApi 1.8, against office-js#3014's
2022 note that they "cannot be reached" — and it was off by exactly one, five
times out of five, because the anchor was being sought by the TAGGED SHAPE'S id
and a group is never among its own members. 144, with `items.slice(1)`, has zero
mismatches.

**13 of 13 both rounds, and that matters more than the counters here.** This is
the first change in this sequence that WRITES to a chart rather than measuring
one. The drawing-order assumption — that the group answers in the order the
parts tag would have recorded — is not provable from the type definitions, and a
wrong order would scramble a chart rather than refuse. Two rounds of `edit a
chart on the visible slide` and `an update follows a moved chart` passing is the
evidence that it holds on this host.

### Four of eleven declines are now the differ working

    6x  the chart carries no scene fingerprint
    3x  too much of the chart changed to be worth writing shape by shape
    1x  this update draws a picture, which is not in the scene
    1x  no parts list and no readable group members

"Too much changed" and the picture case are correct refusals — the fast path
declining work a redraw does better. Before this week none of these could even
be reached.

### The fingerprint blocker is not new, it was masked

    rounds 139-142   0 fingerprint declines
    rounds 143-144   6 each

Zero before, because the parts check fired first and short-circuited. Charts now
get past it and meet the next condition, which was always there.

**And it has a shape:** charts 1/8, 2/8 and 3/8 carry a fingerprint; 4/8 through
8/8 do not. Per-item, not per-batch — which points at the options rather than at
the tag write, since `sceneTag` is defaulted as `{ sceneTag: sceneFingerprint(scene), ...opts }`
and a caller-supplied `opts` overrides it.

**Not chased here.** Recorded with the numbers that will start it.

## The fingerprint blocker, chased — and the guess in the section above was wrong

Round 144's entry closed by saying the per-item split "points at the options
rather than at the tag write, since `sceneTag` is defaulted as
`{ sceneTag: sceneFingerprint(scene), ...opts }` and a caller-supplied `opts`
overrides it."

**That was the wrong mechanism.** No caller was overriding anything. The split
was per-item because the eight charts come from TWO DIFFERENT WRITERS, and only
one of them had ever written the tag:

    charts 1-3   insertSceneIntoSlide  ->  powerpoint.ts, stamps sceneTag
    charts 4-8   the generated deck    ->  pptx-deck.ts -> ooxml.ts, never did

`sceneFingerprint` was not imported in `pptx-deck.ts` at all. The default that
looked like the suspect is on the path that already worked.

**The lesson is about the shape of the evidence.** "Per-item, not per-batch"
was read as "the items differ in their options" when it equally meant "the items
took different code paths" — and the second reading was cheap to check and never
made. One grep for `sceneFingerprint` in `src/render/` would have settled it
before the theory was written down.

### The decline message was lying about itself

    "the chart carries no scene fingerprint - it was drawn by an older build"

These charts were drawn by the CURRENT build, which had never written one. The
message named a cause it could not know, and named it confidently enough that
the journal repeated the framing. A refusal that reports a condition should
report the condition, not a story about how the condition arose.

### Present is not fixed

`tryInPlaceUpdate` does `sceneFingerprint(buildChart(JSON.parse(tags.config)))`
and refuses on a mismatch, so writing SOME value would have moved the refusal
one reason along and looked like progress — the same way the parts check hid
this one for 117 rounds. The guard therefore asserts the value equals what the
update path recomputes, and was mutation-checked with a plausible-looking wrong
fingerprint, which a presence check would have passed.

Landed as #645. What round 145 should show: the six `no scene fingerprint`
declines gone, and — for the first time — `updated only the shapes that changed`
on a chart from the deck.

## Round 145 — the fingerprint fix held; the positive claim did not

The prediction was staked before the round landed, in two halves, because this
fix had a failure mode that looks like success.

    rung  4  the chart carries no scene fingerprint      6 -> 0   as predicted
    rung  7  no longer renders to the scene that was drawn  0 -> 0   as predicted
    rung  9  too much of the chart changed to be worth it  3 -> 8
    in-place SUCCESSES                                     0 -> 0   PREDICTION FAILED

**The half that mattered held.** Rung 7 is the mismatch rung: if the value
written into the deck were not the value `sceneFingerprint(buildChart(config))`
produces at update time, all six declines would have migrated there and the
total would not have moved. They did not migrate. The stored fingerprint is the
one the update path recomputes, and #645 is a real fix rather than a cosmetic
one.

**The positive claim failed outright.** "At least one `updated only the shapes
that changed`" was staked precisely so a smaller round could not satisfy it, and
it came back zero. The denominator is sound — 15 declines in each of 143, 144
and 145, and 578 / 571 / 580 trace entries — so this is a real zero, not a
different population.

### Where the six went, and what was waiting there

Five of the six landed on rung 9, which is the differ correctly declining work a
redraw does better: `changed 18 of 24` and `changed 9 of 16`, three quarters and
better than half of the chart. Those are right.

**The sixth kind reached the host and the host refused it** — three charts, all
`changed 1 of 24`, all with the same error:

    InvalidArgument | errorLocation=Shape.textFrame
    statement: var textFrame = shape.textFrame;

`changed 1 of 24` is the differ working perfectly: it isolated a single changed
node and went to write only that. The write went to the wrong object. For a
grouped chart `shapes = [old, ...parts]` uses the GROUP as node 0, and node 0 is
the group's first member. A `ShapeGroup` has `fill` and `lineFormat` — both
navigations succeeded in the host's own statement list — and no `textFrame`.

### The safety argument in #643 was wrong, and precisely wrong

That PR justified the positional mapping like this: *"the one-for-one guard still
refuses anything that doesn't line up"*. It does not. **A count guard checks how
many, and this was an error in which element means what.** `parts.length + 1`
came to 24 against 24 nodes both before and after the defect — the guard was
satisfied by the broken mapping and would have been satisfied by any permutation
of it.

The lesson generalises past this bug: a cardinality check is not an alignment
check, and quoting one as protection for the other is how a wrong mapping gets
called safe in writing.

### The suite had a test for this and it was green

`updates a GROUPED chart through its group members` edits the title — node 0,
the only node that was mis-targeted — and passed the whole time. The fake host
lets a group accept a name and a text frame, so the wrong target was invisible
in jsdom while the real host refused it three times out of three.

**The double was more permissive than the thing it doubles**, which is the same
shape as the fake that populated values on request and hid a 56-round defect. A
test that exercises the exact defect and passes is worse than no test: it was
counted as coverage.

The new guard asserts the TARGET rather than the outcome — after a title-only
update the group must still be named `PowerChart` — because the outcome was
indistinguishable in a host that accepts everything.

### And one underneath it

Fixing the mapping made every grouped write refuse with *"the host would not
confirm every shape that had to change"*. `isLive` asks `isNullObject === false`,
which is the protocol of an `…OrNullObject` lookup; a shape handed back as an
item of a loaded collection has no such flag and reads `undefined`. The guard
called a shape dead because it had arrived alive by a different route.

Landed as #646. What round 146 should show: the three `Shape.textFrame`
refusals gone, and the first `updated only the shapes that changed` in 119
archived rounds.

## Round 146 — the in-place update ran, for the first time in 119 rounds

    round 145   13 update attempts   0 in place   3 refused at Shape.textFrame
    round 146   13 update attempts   3 in place   0 refused at Shape.textFrame

    WINS: changed 1 of 24, saved 47   (x3)

Both halves of the staked prediction held. The three charts that the host had
refused are the same three that now succeed, and each one wrote a single shape
where a redraw would have written 48.

**The population is provably identical.** Thirteen update attempts in each
round — the constant established weeks ago — and the three did not appear from
anywhere, they moved out of the decline column. 145 was 13 declines and 0
successes; 146 is 10 and 3.

### The unclassified bucket was the interesting one, for two rounds

My decline breakdown bucketed by `data.why` and let everything without one fall
into a `?` bucket that I read as noise:

    round 143   2 unclassified
    round 144   4 unclassified   <- THREE of these were Shape.textFrame refusals
    round 145   5 unclassified   <- and these

I opened the bucket at round 145 and found the host refusals in it. **They were
in round 144's trace too, and I had reported that round without looking.** The
entries had no `why` for the ordinary reason that they had not declined by a
rule — they had THROWN, and a throw carries an `error`, not a `why`.

So the mining lesson is specific: **a residual bucket in a breakdown is not
noise until it has been opened once.** The named categories are the ones already
understood; the interesting thing is, almost by definition, the one that does
not fit them. Two rounds of a host-side defect sat in plain sight in a bucket I
had labelled and skipped.

**And the count in that bucket is not the count of anything.** Two of round
146's "unclassified declines" are a by-id `GeneralException` that RECOVERED
(`asked 1, recovered 1`) — matched only because the message contains the word
"refused". Excluding them is what makes the arithmetic exact at 13 and 13. A
filter that greps prose will pick up prose.

### What is left is the differ declining correctly

    8x  too much of the chart changed     18 of 24, 9 of 16
    1x  this update draws a picture
    1x  no parts list and no readable group members

Three quarters and better than half of a chart changed: a redraw genuinely does
those better. The remaining two are correct by construction.

13 of 13 scenarios passed. Friction — 3 errors, 2 id refusals, 1 general
exception — all recovered.

## Round 147 — the pair confirms it

    145   ok 0 | declined 10 | host-refused 3 | attempts 13 | scenarios 13/13
    146   ok 3 | declined 10 | host-refused 0 | attempts 13 | scenarios 13/13
    147   ok 3 | declined 10 | host-refused 0 | attempts 13 | scenarios 13/13

An exact repeat, reason breakdown included. `9ef98af` differs from `2521d23`
only by a comment and an archive, so this is the second run of a pair rather
than a new build — and the second run is usually the WORSE one. An identical
result is therefore the strongest form this evidence takes: three successes is
the rate, not a lucky round.

Thirteen attempts in all three rounds, which is the constant. Nothing here
depends on a denominator that moved.

### Both of these rounds were read through an instrument that could not see the failure

`poolInPlaceUpdates` — the pool behind the gate's "A FEATURE THAT HAS NEVER RUN"
banner — counted two of three outcomes:

    updated only the shapes that changed        -> ok
    not updating in place — redrawing instead   -> fell
    in-place update refused — redrawing instead -> COUNTED AS NEITHER

The third is the THROW path, which carries an `error` and no `why`. So every
host-side refusal registered as neither a success nor a fallback and left no
trace in the gate at all: the three `Shape.textFrame` refusals in round 145 and
two more in 144. The tool was structurally incapable of showing the thing that
was missed by hand, which is why looking harder would not have helped.

Pooled over the archive it now reads:

    in-place update: 3 succeeded, 1358 declined, 5 refused by the host over 122 round(s)
      1197x  the chart has no parts list, so its nodes cannot be mapped to shapes
       121x  this update draws a picture, which is not in the scene
        19x  too much of the chart changed to be worth writing shape by shape
         5x  InvalidArgument at Shape.textFrame

And the residual is NAMED rather than tallied. A bucket reported as a number
reads as noise and gets skipped; reported as a line it gets opened.

### A measurement fault found by making it twice

`attempt` re-read `git HEAD` on every retry, so committing in this clone while a
round retried moved the build under test mid-round — the driver would then
refuse its own round as `site-behind` against a commit that had never been
deployed and was not what it was measuring. Both times it was harmless only by
luck.

The habit is the operator's; the fault is the driver's. **A round tests one
build**, so `main` now reads HEAD once and hands it to every attempt.
