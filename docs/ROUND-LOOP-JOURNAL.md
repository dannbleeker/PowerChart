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
rather than the crash watch. The lifetime question was never answered.

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
