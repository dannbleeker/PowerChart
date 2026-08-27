# PowerPoint web silently drops `SlideCollection.add()` under load

## Environment

- **Host**: PowerPoint on the web (Office Online), accessed via
  office.com.
- **Office.js requirement set**: PowerPointApi 1.3+ (`slides.add()` itself
  requires only 1.3; the SSF Charts add-in's manifest requests up through
  1.10 for other features, but this bug reproduces using nothing beyond the
  base `slides.add()` / `slides.getCount()` API surface).
- **Add-in**: SSF Charts, sideloaded via `manifest-prod.xml` (Phase 2 of
  `docs/PUBLISHING.md`), hosted at
  `https://ssf-chart.struktureretsundfornuft.dk/`.
- **Client**: browser-hosted Office Online session, not a desktop build.

## Observed behavior

Two independent Phase-2 validation runs of the SSF Charts add-in against a
real PowerPoint-web deck both show slides going missing after
`context.presentation.slides.add()` calls that completed their
`context.sync()` with no thrown error.

### Run 1 — `Presentation_3.pptx` (smoke subset, 2026-07-29)

- Smoke subset: 12 items (Title, Contents, then 8 charts + 2 elements).
- The harness issued **20** `slides.add()` calls in total: 12 first
  attempts plus 8 retries after the harness's own self-check detected
  missing slides.
- The deck grew by only **10** slides.
- **10 adds vanished silently** — no exception, no rejected promise; the
  `context.sync()` that queued each add returned successfully.
- Harness self-check reported: `addsIssued: 20, slidesAdded: 10`.
- Wall clock for the run: **773.4 seconds** (~65s average per item), with
  long stalls between adds — timeouts in the harness appear to race with
  when the host actually commits a sync.
- Pattern: loss correlates with the host being "under load" in the same
  session — adds issued early in the run tend to succeed, adds issued
  later in the same session fail more often.

### Run 2 — `Presentation_2.pptx` (full deck, 2026-07-28)

- Full deck: 37 items, deck grew by 63 slides.
- Different failure shape: rather than vanishing outright, some
  `slides.add()` calls landed content at the **wrong position** — the
  content intended for item *N* appeared on slide *N+1*, leaving slide *N*
  empty.
- **5 of 63** slides were misplaced this way.

### Run 3 — `Presentation_4.pptx` (smoke subset, 2026-07-30)

The first run whose output deck was read back at the XML level rather than
trusted to the harness's own report. The two disagree, and the disagreement
is the most useful evidence in this document.

- Smoke subset: 12 items. The harness issued **19** adds; the deck grew by
  **10**. Wall clock **680.2s** (~113s per item).
- The harness reported: `Inserted 6 of 12 … Host failed on: Stacked,
  Scatter, Bubble, Gantt, Combo, Heatmap … 4 charts landed ungrouped …
  3 slides came back BLANK: slide 9, slide 10, Agenda (slide 11)`.
- The deck says otherwise:
  - **Gantt did not fail — it landed twice.** Slides 6 and 7 both hold the
    same 31-shape Gantt chart; slide 7 additionally wears the harness's
    `NOT COMPLETE` banner. The same is true of Line on slides 4 and 5
    (36 shapes each).
  - **The "BLANK" Agenda slide holds all 13 of its shapes.** Its content is
    in `slide11.xml`; the web client simply never painted it — which is the
    behaviour of office-js#2699, below.
  - **Four `NOT COMPLETE` banners sit on top of complete content** (the
    contents slide, both duplicate charts, and the agenda).
  - Only **one** of five charts kept its `POWERCHART_CONFIG` tag.

The mechanism is visible in the timing: the harness's timeout fires at 45s,
it stamps the slide and issues a retry, and the abandoned `sync()` then
commits anyway — so both the original and the retry end up on the deck, one
of them wearing a banner that its own content contradicts. `waitForLateSync`
gives the abandoned call 5 seconds to report back; the observed commits
arrive far later than that, so every readback taken during the run is a
snapshot of a host still mid-flight.

### Run 4 — the client itself, 2026-07-31

With the add-in's one-call insert path *disabled* — the **full 37-item
deck**, drawn shape by shape as in Runs 1-3 — PowerPoint on the web did not
stall. It **crashed**, roughly five seconds in:

> Microsoft PowerPoint — Sorry, we ran into a problem. Please try again.
> [Refresh]

The twelve-item smoke subset on the same path, same session, did start
rendering — so the crash is a function of the volume of queued work, not of
the API being touched at all. It is the same load-dependence Run 1 records
for silently dropped adds, one step further along: enough of it and the
client does not lose an add, it goes down.

The same decks inserted instead as a single generated `.pptx` through
`insertSlidesFromBase64`, on the same host in the same session, land in
**4.1 seconds** for the twelve and **6.4 seconds** for all thirty-seven,
with every chart grouped and tagged and nothing lost. That contrast is the
sharpest evidence in this document: the shape-by-shape API surface is not
merely slow on the web client, at volume it is capable of taking the client
down, while the file-insert path on the identical host and deck is not.

Taken together, the four runs show at least three related failure modes
under the same root cause (adds and commits not landing where/when the
caller expects them to): outright silent loss (Run 1), off-by-one placement
(Run 2), and commit-after-abandonment producing duplicate slides plus false
failure reports (Run 3).

## Expected behavior

Every queued `slides.add()` whose `context.sync()` resolves without error
should result in exactly one new slide appended (or inserted at the
requested index) in the presentation. A successful `sync()` should be a
reliable signal that the operation took effect — callers should not need
to re-read `slides.getCount()` afterward to find out whether it actually
did.

## Minimal repro

```js
PowerPoint.run(async (context) => {
  const before = context.presentation.slides.getCount();
  await context.sync();

  for (let i = 0; i < 20; i++) {
    context.presentation.slides.add();
  }
  await context.sync();

  const after = context.presentation.slides.getCount();
  await context.sync();

  console.log(`before=${before.value}, after=${after.value}, expected +20`);
});
```

Run this block roughly 10 times in a row within the **same** PowerPoint-web
session (same deck, same browser tab, don't reload between runs). Some
runs will report `after - before < 20` with no error ever thrown along the
way. The failure rate appears to increase as the session accumulates more
slides / more prior operations ("under load"), matching the Run 1 pattern
where later adds in a long-running session fail more often than early
ones.

## Workaround

SSF Charts' harness works around this by verifying the settled slide
count after each batch of adds and re-issuing `slides.add()` for whatever
is missing:

1. Record `slides.getCount()` before issuing a batch of adds.
2. Issue the adds, `sync()`.
3. Re-read `slides.getCount()` after the sync settles.
4. If the observed growth is less than the number of adds issued, retry
   the shortfall (this is where Run 1's 8 retries came from).
5. For the misplacement mode seen in Run 2, the harness additionally
   verifies slide *content* (not just count) matches the expected item at
   each index before moving on, since a count-only check would not catch
   content landing one slide off from where it belongs.

Run 3 showed the limit of doing this inline: every check above races a
commit whose deadline the host does not publish, and losing that race
writes fiction into the deck (a duplicate slide, a banner contradicting its
own content, a chart reported failed that landed twice). SSF Charts
therefore now also runs a **settled reconciliation pass after the run
finishes** — `src/core/reconcile.ts` plus `reconcileDeck` in
`src/render/powerpoint.ts`. It re-reads the added range once the host has
stopped moving, pairs each slide with the item its `POWERCHART_DEMO_SLOT`
tag names, and repairs the difference: delete provably redundant twins,
clear stale banners, re-group charts left loose. The same pass is exposed
as a **Repair deck** button so a deck damaged by an earlier session can be
fixed without re-inserting anything.

This is a workaround, not a fix — retries still cost the ~65s/item wall
clock seen in Run 1, and a false-negative verify could in principle mask a
different bug. The PR that hardens the Phase-2 harness with this
retry-and-verify logic should be linked here once opened
(`PowerChart` repo, TBD — not yet filed as of this writing).

## Related

The web client losing track of committed state is a **pattern**, not an
isolated report. Four other issues describe the same class of failure on the
same host, three of them still open (checked 2026-08-01):

- [OfficeDev/office-js#6363](https://github.com/OfficeDev/office-js/issues/6363) —
  properties that were loaded and synced come back "not available" on
  PowerPoint web. Labelled a **regression and a product bug**; the reporter
  tried ten loading strategies, none worked. This is the read-side twin of the
  write-side loss reported here: the same deck, read twice, answers differently.
- [OfficeDev/office-js#5022](https://github.com/OfficeDev/office-js/issues/5022) —
  `context.sync()` hangs indefinitely after shapes are added, deleted and
  re-read. Under investigation; the only known workaround is a 1–2 second sleep.
- [OfficeDev/office-js#4272](https://github.com/OfficeDev/office-js/issues/4272) —
  `context.sync()` hangs once more than ~51 items are queued in one `load()`.
  Directly shapes how much a caller can safely read back at once.
- [OfficeDev/office-js#2903](https://github.com/OfficeDev/office-js/issues/2903) —
  content added to a newly created slide lands on the wrong slide on PowerPoint
  web, with `InvalidParam passed to GetItem(id)`. **Closed as "not planned"**
  after inactivity, which is a decision about the report rather than a fix for
  the behaviour.

Taken together with this report, the common thread is that the web client
acknowledges a mutation, then fails to reflect it in the model a subsequent
call reads — silently, and without an error the caller can branch on.

- [OfficeDev/office-js#2699](https://github.com/OfficeDev/office-js/issues/2699) —
  shapes created successfully but not painted on PowerPoint web until a
  zoom/repaint (see `docs/repro/ellipse-web-repro.yaml` in this repo for a
  minimal isolation of that bug). It's plausible this is a related class of
  bug: both involve the web client's rendering/commit pipeline losing track
  of state under some conditions, rather than the Office.js API layer
  itself rejecting the call. Worth checking whether the "missing" slides in
  Run 1 are truly absent from the deck's underlying model, or present but
  unpainted/unlisted the way #2699's shapes are.

## How to submit

1. Copy this document's content into a new office-js issue.
2. Attach screenshots from `Presentation_2.pptx` and `Presentation_3.pptx`
   (the two decks uploaded to the current Claude Code session) showing the
   missing/misplaced slides described above.
3. Submit at
   <https://github.com/OfficeDev/office-js/issues/new/choose>, picking the
   bug-report template.
4. Link back the resulting issue number here and from the retry-workaround
   PR once it exists, per the lockstep convention in `CLAUDE.md`.
