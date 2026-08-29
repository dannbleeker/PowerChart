# Backlog

Curated candidate work, from a research sweep (July 2026) comparing every
existing chart kind against think-cell, Excel, Highcharts, Datawrapper, and
Mekko Graphics, plus a chart-type survey across the Zelazny / FT Visual
Vocabulary taxonomies and competitor add-ins (Zebra BI, Vizzlo, UpSlide).

**This is the only backlog document.** Items graduate from here into PRs and are
deleted when they ship — what has shipped is recorded by the README feature
table and by git, not here. Rejected ideas stay in §2 so they aren't
re-proposed, and §3 keeps what ~290 rounds against the live host established,
because a finding outlives the fix that answered it. **§1 opens with the whole
open list; if a thing is not on that list it is not open.**

Feasibility is judged against the live-add-in constraint: rects, lines, text,
ellipses, and any of PowerPoint's 177 preset geometries (all of them
PowerPointApi 1.4), plus polygon *outlines* — no freeform curves, and no
images. The SVG and skill-pptx renderers additionally have filled polygons and
patterns.

## 1. Open

**Everything actually open, as of 2026-08-29.** The sections below carry the
evidence; this list carries the state. Anything not on it is either shipped,
refused, or a finding rather than a task.

**Five of the eight closed on 2026-08-29, and the three left are all the
owner's.** Two are product decisions and one goes out under his GitHub identity:

    1  where a unit label belongs when the band above the plot cannot hold three
       (SMALLER as of 2026-08-29: the option is a short unit, and clipping it to
       40% of the chart took the total 1,327 -> 1,014 with no new shapes)
    3  whether a crowded slide should get a picture instead of native shapes
    5  filing this project's host measurements to the office-js tracker

Nothing on that list is waiting on engineering. Items 2, 4, 6, 7 and 8 were
closed by shipping, by measuring the remedy and declining it, or by the host
refusing the mechanism outright — each says which, in its own row.

| | What | Where | State |
|---|---|---|---|
| 1 | **`valueAxisTitle` runs over the legend and the totals row** — **916 of the 1,014** remaining pairs (was 1,208 of 1,327), and still the whole of what is left bar a 98-pair tail | §3, *Text drawn over text* | **PART SHIPPED 2026-08-29, the rest is the owner's call.** The option is documented as a short unit — "e.g. `€m`" — and the two uses in the shipped deck are `€m` and `$m (log)`, but it was drawn with NO width bound, and the sweep tests it with a 27-character sentence. Clipping it to 40% of the chart takes 1,327 → **1,014 with no new shapes**, and keeps author text where three earlier remedies deleted it. Flooring its `y` under the title was measured in the same hour and **refused**: 1,156, and 310 `value-axis-title / category#` that did not exist before. What is left — where a unit belongs when the band cannot hold title, unit and top tick — is unchanged in kind, just smaller. (Was "287 of 317" — a figure from the uncrossed sweep, not comparable. See the note on the three totals.) |
| 2 | ~~**The slow-insert offer is not wired**~~ | §1, *The slow-insert offer* | **SHIPPED 2026-08-29.** All five steps, five pane tests, four mutants dead. The blocker this row named was a stale finding: a fresh slide's id IS usable once the slide settles, and the probe answer flipped at round 254 because of our own commit. Kept in §1 for the lesson, not the task. |
| 3 | **Adding to an occupied slide costs ~24s; a fresh one ~0.75s** | §1, *Adding a chart to an existing slide* | Measured over 2,917 batches. Item 2 (shipped) is the user-facing half. **A fourth option is now costed, 2026-08-29**: draw the chart as ONE picture — ~43s becomes rasterise 705ms plus a single-shape insert, a 7–20× saving, and every part of the mechanism already ships. **Still the owner's call**, and a bigger one than the offer: it changes what the user RECEIVES, not just what they are told. |
| 4 | ~~**The picture fast-path**~~ | §1, *The largest product cost was in the FAST path* | **CLOSED 2026-08-29 on evidence, both halves.** The cost it was filed under had moved: redraws are 14.5% and exactly 2 a round, and the real cost was inside the FAST path, which took a median of 15.7s because it wrote every property of every changed node. It now writes only what differs — `same scale across the deck` 146–156s → **103s**, below the IQR of 275 rounds. And the picture feature itself is **not worth building**: a picture redraw is ~40ms of delete plus ~770ms of picture against 15–50s to draw the same shapes natively, so a fast path could save about five percent of an operation that is already twenty to sixty times cheaper than its alternative. |
| 5 | **File this project's host measurements to the office-js tracker** | §1, *Report what this project has measured* | **Owner-gated** — it goes out under his GitHub identity, so nothing is filed without his word on that specific issue. Three are written and ready. |
| 6 | ~~**The dual-axis gutter**~~ — 14 pairs, and the remedy costs 1,394 axis readings | §3, *Text drawn over text* | **MEASURED AND DECLINED 2026-08-29.** Not "30 pairs": 14, of 1,327, and none of them a tick number. Applying the merge to the secondary case fixes 8 and DELETES 1,394 secondary tick numbers — 174 readings lost per pair gained, where the original decision recorded about ten. The trade got seventeen times worse as the rest of the engine improved. The design is written and stays written; on these numbers nobody should build it. |
| 7 | ~~**Positional group-member mapping → `Shape.creationId`**~~ | §3 | **CLOSED 2026-08-29 — the host refuses it.** This host reports PowerPointApi **1.10** and does not populate `Shape.creationId`: `absent` / `no-creation-id` on all three probe questions, 25 of 25 rounds. So it was never a matter of gating on 1.10. The positional mapping stays, still inferred, still guarded by the node-0 anchor test — which is now the permanent answer rather than a stopgap. |
| 8 | ~~**`slideSize()` rung 1 times out in EVERY round that reads a size**~~ | §3 | **DONE 2026-08-29**, and measured on both sides. The bound is now its own constant at 1,500ms rather than the shared selection budget, and rounds 303 and 305 are a clean A/B on the same fallback path: `exportedSlide ms=4404` → `ms=2108`, **2.3 seconds back** per cold insert, with rung 1 still answering in 259ms when it answers. Remaining, and much smaller: whether a warm-up call removes the stall altogether. The measurement that got here: Re-measured 2026-08-29 over 270 rounds: 157 of 157, always the full 4000ms — not "about twice as often as it answers". **And now timed**: a SUCCESSFUL rung-1 read costs 246–270ms (rounds 297–299), so the stall and the answer are different calls and the bound is not buying the answer. **Round 300 then broke the "157 of 157" entirely** — no stall, 138ms, on a round that followed a crash recovery and a full tab reload; round 301, on a tab left idle 35 minutes, stalled again. **And 301 measured rung 2 at ~280ms**, which is what this was blocked on — the export costs what the read costs, so a 500ms bound would take a stalled run from 4,280ms to under 800. No longer blocked on a measurement; only on choosing between a lower bound and a warm-up. |

Two standing costs that are not tasks: the host crashes during evidence
collection after a long round (§3 — it costs the evidence, not the run), and
`test/overlap-budget.test.ts` holds **1,327** overlapping pairs whose per-shape
table is the live list. Three figures in one day, and only the last step was the
engine: **537** was the uncrossed sweep, **4,010** the same engine measured by a
sweep that finally crossed option variants with data shapes, and **1,327** what
was left once `buildMultiples` stopped building grids with no room. Do not
compare any of them with an earlier number without naming the sweep.

**Found and fixed 2026-08-29: `buildMultiples` built grids with no room.** A
panel height of −0.4 was handed on, `clampDim` rewrote it to
`DEFAULT_SIZE.height`, and ten full 300-point charts were stacked 9.6 points
apart inside a box 60 points tall. 2,683 of the sweep's pairs; all of them gone,
total now **1,327**. **Validated**: rounds 298 and 299 at 16:9 and 300 at 4:3,
all 14 of 14, plus a byte-identical showcase across 123 configs — see §3.

**Newly open, 2026-08-29: `duplicateSlot` judges a window it cannot fill.** The
scenario sizes its reconcile from a slide count this host reports LATE, so a
round where an insert answers "nothing landed" for slides that did land gets a
hard FAILED verdict about two slides out of four — and the cycle's fatal check
stops the night on it. Round 297 is the worked example; the fix is a few lines
and is deliberately not taken at 2am, because it changes a verdict rule with a
273-round series behind it. See §3.

**Newly reopened, 2026-08-29: PowerPointApi 1.8 bindings.** §2 retires them as
"unanswerable on this host", on the strength of the host rejecting the batch
that carries a binding. Over the 25 rounds to 2026-08-28 that face appears
ONCE, against `yes` 15 times and `silent` 8. It is the last untried way out of
the id refusals and the evidence that closed it has weakened. Not put on the
list above because nothing has been designed — the `silent` third is its own
problem — but it is no longer refuted, and `binding-names-shape-later` is back
on the probe's resample shortlist so the next rounds keep counting.

### The largest product cost was in the FAST path, not the redraws — 2026-08-29

**Looking for the picture feature below, I measured the redraws and found they
had stopped being the cost.** Over the last 30 rounds:

    in-place 341 · redrawn 58 = 14.5% redrawn   (was 20.9%)
    29  this update draws a picture          — legitimate, cannot be diffed
    29  no parts list and no group members   — the addressable half

Exactly two redraws per round, every round. But the same query printed the
number nobody had looked at:

    in-place update ms: n=341 median 15,661 · max 59,000

**The fast path was taking sixteen seconds.** `same scale across the deck` runs
entirely through it and costs 166s, 39% of the battery, which is what "the
largest product cost" was actually made of.

**What it was spending it on.** `applyNodeInPlace` wrote every property of every
changed node unconditionally. Counted by a recording proxy, a whole text node is
exactly **20 statements** and a rect **7** — 20 being the figure `powerpoint.ts`
already quoted from the host's own statement list, which is how the count is
known to be measuring the right thing. So a retitle, the single edit this path
exists for, sent twenty statements to change one string.

`planSceneUpdate` holds BOTH scenes and therefore already knew which of the
twenty were stale. It just never said. It now returns `parts` beside `changed`
— box, fill, line, text, font, align — and the applier writes only those.

**Three measurements, all agreeing.** Statements first:

    retitle                 20 ->   2   (90% fewer)
    recolour               152 ->  44   (71% fewer)
    rescale, same scale    272 -> 180   (34% fewer)

Then per-update milliseconds on the live host, like-for-like by node count,
against the five preceding builds:

    changed=1     ~2,100-2,500ms  ->   1,117ms
    changed=9    ~11,100-12,100ms ->   4,825ms
    changed=18   ~15,700-16,600ms ->  12,042ms

Then the scenario itself:

    299 4275306  147s      302 aa459dd  156s
    301 e0b82ef  146s      303 81ab4a4  103s

**103s against a pooled median of 166s and an IQR of 151-211s over 275 rounds**
— below the whole interquartile range of the archive.

Time falls by less than statements do, consistently, and that is the expected
shape: `syncMs` shows a fixed floor of roughly a second on every sync, so a
retitle that sends a tenth of the statements still pays two syncs.

**One round for the scenario figure**, n=12 for the per-update figures. Three
independent size classes all moved the same way by a lot, and the baseline is
eight builds deep, but it is one round and should be read as one.

**What is left of the original entry** is the picture feature — and rounds 304
and 305 close it. **It is not worth building.**

The two halves of a picture redraw were timed on 2026-08-29 and the split is
lopsided:

    deleting the chart being replaced    13-38ms    (removed=1)
    drawing the chart as one picture     742-799ms  (nodes=24)

About **810ms in total**, against 15 to 50 seconds for the same twenty-four
shapes drawn one at a time — three batches at the 5,013 to 16,579ms this archive
measures per batch. So the picture path is already twenty to sixty times cheaper
than the thing it replaces.

A "picture fast path" could at best reuse one existing shape as the host and
save the delete: **40ms of 810, about five percent**, for a feature that has to
teach a differ about a fill it cannot see. The argument for expecting that was
written down before the measurement and the measurement agrees with it, which is
the only reason to trust either.

The `removed=1` is why the delete is so cheap and is worth keeping in mind: a
grouped chart is one shape to delete however many it holds. A chart that failed
to group would cost twenty-four deletes, and this host groups reliably now.

**So item 4 is closed on evidence rather than deferred.** The redraws that
remain are correct, cheap, and not worth removing.

### Two of every three redraws are a grouped chart the update cannot map — 2026-08-26

**CLOSED 2026-08-27. The redraw rate is at its design floor — both remaining
redraws are correct, and neither is a missed optimisation.**

Rounds 275 and 276 each redraw exactly twice, and the two reasons are these:

    this update draws a picture, which is not in the scene the differ compares
    the chart has no parts list and no readable group members

The FIRST is a deliberate refusal and removing it would reintroduce a silent
wrong answer. `render: "image"` does not produce a picture — the renderer takes
that path only when handed `pictureBase64` — so collapsing a chart to a picture
builds the SAME scene it already has. The differ compared 24 nodes to 24
identical nodes, said "nothing changed", wrote nothing and reported success,
while the slide kept its 24 native shapes. Worse, the auto-picture fallback is
what the add-in reaches for when the host has ALREADY failed to draw shapes, so
the one path that rescues a struggling host was the one being skipped. Teaching
the fast path to handle it is a real feature — it writes a closed set of `rect`
and `text` properties and a picture fill is neither — not a small fix.

The SECOND is settled: see "ANSWERED 2026-08-27" below. Both routes into a
refused group are shut inside the batch, the redraw is the right answer, and what
rescues the group is the next batch.

So the remaining work here is not "stop redrawing". It is either the picture
feature above, or nothing. Leaving the original analysis below because the
reasoning that got here is worth more than the conclusion.

**The largest product cost still on the table, and the obvious fix is wrong.**

An update either writes only the shapes that changed, or redraws the chart
whole. A redraw is far more expensive: it is why `same scale across the deck`
costs 167s, 38% of the battery. Measured over the last 30 rounds:

    in-place 325 · redrawn 86 = 20.9% redrawn

    56  the chart has no parts list AND no readable group members
    30  this update draws a picture, which is not in the scene the differ compares

The picture case is legitimate — a picture cannot be diffed against a scene. The
other 56 are the whole of the addressable cost, and they are all the same shape:
a chart that GROUPED, whose group the update could not read — not because the
host refuses, as this said until 2026-08-26, but because a session-wide latch had
turned the read off after one refusal.

**How it got here is worth stating, because each step was right.** Grouping used
to fail, so charts were loose, so they carried a parts list and the update mapped
nodes to shapes through it. Grouping now succeeds — `fresh-slides-group` reads
96/98 — and a grouped chart is not "loose", so it collects no parts list. The
update was given a different route for those: `shape.group.shapes` (#643),
written through in #646. That route is what fails:
`group-children-via-getcount` answers `unreadable` in 34 of 40 rounds.

So a grouped chart has neither route, and redraws.

**`grouped-child-by-id-from-slide` is no longer moot.** It was retired
2026-08-21 with the reasoning that the by-id route "is not used and does not need
an answer". The route that made it moot is the route that is failing. A question
is only moot while the thing that made it moot works.

**The naive fix was tried on 2026-08-26 and REVERTED.** Collecting a parts list
for grouped charts as well as loose ones looks free — the ids are already
populated by the drawing batch, so it reads a property rather than making a call
— and the suite killed it in one run:

    reports no growth for a GROUPED chart — "a group goes in one delete": expected 14 to be 1

The parts tag is not only the in-place MAPPING, it is also the DELETE list. Hand
one to a grouped chart and a single group delete becomes fourteen shape deletes,
which is the "chart grows by a whole chart on every edit" bug the tag exists to
prevent. `loose` gated it for a reason that is not written down anywhere.

### ANSWERED 2026-08-26, and the fix is dead

`grouped-child-by-id-from-slide` finally has an answer. It starved for 125
rounds, was retired as moot, and was re-opened above because the route that made
it moot is the route that fails. Asked directly, three times, in under two
seconds each:

    grouped-child-by-id — no-such-shape: child 2 of group 5 read back as undefined
    grouped-child-by-id — no-such-shape: child 5 of group 9 read back as undefined
    grouped-child-by-id — no-such-shape: child 9 of group 13 read back as undefined

**A grouped child is not addressable by id off the slide on this host.** Not
"unreadable", not refused — the slide's shape collection does not contain it at
all. Grouping takes the children out of the slide's collection, and the way back
is `shape.group.shapes` — which DOES enumerate on a real slide, `yes` both fresh
and re-resolved. Only the by-id route is closed.

So the mapping-tag idea is DEAD, and so is anything else built on reaching a
grouped chart's shapes BY ID. Capturing child ids at draw time would have
produced a tag full of ids that resolve to nothing.

### CORRECTED THE SAME DAY — the redraw was NOT unavoidable

The paragraph that stood here said "the redraw is not a missed optimisation, it
is the only correct behaviour available for a grouped chart on this host". That
was wrong, and it was wrong for the reason everything else on this page was
wrong: it rested on a probe answer measured on the scratch slide.

Two routes to a grouped chart's shapes exist and they are not the same question:

    by id            `slide.shapes.getItemOrNullObject(childId)`   — DOES NOT WORK
    by enumeration   `shape.group.shapes` loaded for `items/id`    — WORKS

The first is settled above, three runs, `no-such-shape`. The second was asked on
a real slide the same afternoon and answered `yes` both fresh and re-resolved by
id — where the probe's scratch-slide version, `group-children-via-getcount`,
answers `unreadable` in 34 of 40 rounds.

So the update path had a working route all along. What was taking it away was a
LATCH: `groupReadRefused` held for the whole SESSION, and one refusal ended
in-place updates for the rest of the round. Across rounds 254-261, in-place
updates after the first refusal: zero, eight rounds running. Resetting it per
batch (`55630e7`) gives:

    baseline 255-261   11 in-place / 3 redrawn = 21.4%   seven identical rounds
    with fix 262-264   12 in-place / 2 redrawn = 14.3%   three identical rounds

`a-group-refusal-no-longer-ends-the-round` watches it at 3/3.

**ANSWERED 2026-08-26, and the answer is no.** The latch is deliberately kept
WITHIN a batch, because the refusal arrives at the sync rather than at the
property access and poisons the batch it lands in. So charts after the refusal in
the SAME batch still redraw, which is why the gain is one chart rather than
three.

That is not a limitation to be relaxed. Rounds 265-267 spent three attempts
re-asking inside the batch and produced no in-place update at any of them. Both
routes are shut: the recovered proxy's `.group.shapes` answers and then its sync
throws (round 267 in production), and a fresh by-id proxy fails because the by-id
LOOKUP is what poisons these syncs — the same refusal that sent the chart down
this path. The retry has been deleted.

One premise of the paragraph this replaces was also wrong: **the fake CAN model a
sync-time refusal**, and does — `faults.refuseShapeById` sets `refuseThisSync` at
`test/helpers/office-host.ts:1722`, and `pendingHostError` throws once and clears
rather than poisoning everything after it. Two of the three rounds were spent
believing the suite could not be asked.


**How it was answered matters as much as the answer.** Not by a round — by
`scripts/experiment.mjs`, in under two seconds, four iterations apart. The first
three attempts were wrong in ways this archive has documented for months: asking
a creation proxy for its id, reading it without a load, and asking on a SCRATCH
slide whose shape collection this host answers `unreadable` for in 92% of
rounds. Each mistake cost a minute instead of a round.

### Adding a chart to an existing slide costs ~24s; making a new one costs ~0.75s — MEASURED 2026-08-23, not fixed

Pooled over 2,917 timed batches in 169 rounds. Every batch in the archive is
ten shapes, so this is a curve in slide load, not in batch size:

    shapes this run already drew on the target    median batch
      0                                              3886ms
      1-20                                           5490ms
      21-50                                         13995ms
      51-100                                        18074ms

**A run slows itself 4.7x by piling onto one slide.** This is the effect
`shapesDrawnOn` already documents ("about +0.44s per shape present … quadratically")
and it is now measured across the whole archive rather than in one round.

What a user sees, from the two self-test arms:

    insert on top of an earlier run           3s   4 charts onto NEW slides    ~0.75s each
    insert onto a slide that already has …   48s   2 charts onto an EXISTING   ~24s each

**About 32x, and it is structural rather than a bug.** The two arms separate
cleanly on the path they take — counted from one round's trace:

    insert on top of an earlier run          0 batches   2 file handovers
    insert onto a slide that already has …   4 batches   0 file handovers

The fast arm draws **no shapes at all**: it hands the host a generated file. That
is also why 4 charts in 3s is not a contradiction of the curve above — a single
ten-shape batch has a floor of 3886ms, so 4 charts drawn shape-by-shape could
not finish in three seconds and never do.

You cannot hand the host a file to add shapes to a slide that already exists, so
that path draws shape by shape and pays the curve. **SSF Charts is fast exactly
when it CREATES slides and slow exactly when it ADDS to one.**

**Why this is a product decision and not a fix.** The side-by-side layout — two
or three small charts on one slide — is the everyday think-cell workflow, and
it is the slow case. Nothing in the pane tells the user that, and nothing
should be changed unilaterally: the cost is the host's, the mitigation
(`leastLoadedChart`) is correctly scoped to the self-test, which picks a chart
arbitrarily and so should pick wisely. A user picks the chart they mean, so
there is no choice for the product to make on their behalf.

**A FOURTH OPTION, COSTED 2026-08-29: draw it as ONE PICTURE.** The three below
all decide what to SAY about the cost. This one removes most of it, and it is
the only route that can — the cost is per shape drawn, so the way out is to draw
fewer shapes, and a chart can be one.

Re-derived from `batch issued`'s own `prevBatchMs` and `onSlide`, over the whole
archive rather than the 169 rounds above:

    shapes already on the slide      median ten-shape batch
      1-20        n=2,405                    5,013ms
      21-50       n=  904                   13,971ms
      51-100      n=  241                   16,579ms

    rasterising one slide            n=2,093    705ms

So a 24-node chart onto a crowded slide is three batches at ~14.3s — **about
43 seconds** — against a rasterise of 705ms plus a single-shape insert. Even
pessimistically that is under ten seconds, so the saving is roughly **7x to
20x**, and every part of the mechanism already ships: `pictureBase64`, the
auto-picture fallback for a host that has failed to draw shapes, and the
`explode a degraded picture` scenario that turns one back into native shapes.

**It stays the owner's call, and it is a bigger call than the offer was.** A
picture is not a chart PowerPoint can edit — no nudging a bar, no restyling
through the ribbon — and the add-in's own explode path is the only way back.
That is a change to what the user RECEIVES, where the three options below only
change what they are TOLD. Costed here so the decision is informed; not taken.

The options are all product calls: say nothing; show progress honestly on the
slow path; or offer "put this on its own slide" when the target is already
loaded. Worth a decision, not worth guessing at.

**What this is NOT.** Not a batching problem — an empty Office.js round-trip is
**5ms** (the driver's readiness ping, `slides.getCount()` plus one sync), so
round-trips are ~0.1% of a batch and raising `SHAPES_PER_SYNC` cannot help; the
cost is per-shape drawing on the host. And not a visibility problem: an earlier
reading of this same data claimed 2.3x for "the slide the user is looking at"
and was retracted in `0638b6a` — it split on a sentinel meaning "slide id not
loaded yet". Nothing in this archive distinguishes on-screen from off-screen.

### The slow-insert offer — SHIPPED 2026-08-29, and the blocker was a stale finding

**Owner's decision, taken 2026-08-27:** show progress on the slow path, and when
the estimate exceeds **14 seconds**, quietly offer to put the chart on its own
slide instead. Not a modal — it interrupts the press the user just made — but it
does gate this insert, because it is a question about THIS chart.

All five steps are done. `insert-cost.ts` prices the insert from the archive's
four measured medians; `worthOwnSlide` gates the offer on the slowness being the
SLIDE'S fault rather than the chart's; `offerSentence` quotes both costs;
`addSlideForChart` adds the slide through the existing retry path; and the pane
recomputes the placement for the blank slide instead of carrying over a cascade
that dodged obstacles which are not there. Five pane tests cover it end to end,
four mutants dead and the fifth recorded as equivalent.

**THE THING WORTH KEEPING IS WHY IT LOOKED BLOCKED.** This entry said the
wiring could not be finished because a fresh slide's id is useless on this host,
citing `insertSceneIntoSlide`'s own comment, which cited the 2026-08-08 probe
sheet: `shape-add-fresh-getitem-slide` answers `threw`, so a freshly-added slide
"does not work, by any route".

That was true when written and had stopped being true. Over 269 archived rounds
the answer is `threw` 227 times and `yes` 41 — and **every `threw` is round 253
or earlier**. From 254 on it is `yes` in 38 of 40.

**The flip is OURS, not the host's.** The commit on the boundary is 77f9ca4,
which stopped the probe HOLDING the id `slides.add()` hands back and made it
re-read the slide's id positionally once the add had settled. The two are
different id spaces, not near-misses — `4123571114#123571113` at add time
against `256#2587447327` a moment later, for the same slide. So the old answer
was never a fact about `getItem`. It was a fact about a stale id, recorded under
a question whose name says `getitem`.

The corrected rule, now in `host-baseline.mjs` and at the call site:

> A new slide's id is not durable until the slide settles, and is durable
> afterwards. Re-read it positionally after the add and `slides.getItem(id)`
> resolves it.

`addSlideForChart` was already built that way — it goes through `addSlides`,
which does not return until a FRESH context has confirmed the deck grew, then
reads the id off a positional handle. So the wiring worked, and only the
documentation said it could not.

**Two lessons, and the second is the expensive one.**

- A probe question is named for the CALL it makes, and answers about whatever is
  actually broken. `shape-add-fresh-getitem-slide` spent three weeks reading as
  a verdict on `getItem` while measuring a stale id, and the fix for the stale
  id silently rewrote the verdict.
- **A finding copied into a comment does not get re-measured.** The sheet is
  regenerated every round; the comment quoting it is not, and it was the comment
  that stopped this work. Where a comment leans on an archived answer it should
  say which question, so the next reader can re-run it in one command instead of
  believing it.

Still open here: making an insert onto a crowded slide actually faster. The
offer routes around the cost; it does not remove it. See the entry above.


### Text that overlaps text on data the samples do not carry — MEASURED 2026-08-19, not fixed

Superseded 2026-08-28. Its 75-pair figure was measured with a cruder ink rule and was never comparable; re-measuring gave 2,148. See **Text drawn over text** in §3, and `test/overlap-budget.test.ts` for the live per-shape table.

### The crash is in the deck SCAN, not the screenshots — located 2026-08-27

Superseded by **The crash, counted across the whole archive** in §3, which counts the whole archive by build rather than by record and finds no dominant phase.

### Report what this project has measured to the office-js tracker

**Researched:** 2026-08-06. **Owner-gated — nothing may be filed without the
owner's word**, because it goes out under his GitHub identity.

Reading the tracker turned up five open defects sitting under code this add-in
ships, and the weekly sweep (`scripts/office-js-watch.mjs`) now keeps that
current. The traffic has been one-way. This project holds measurements of
PowerPoint on the web that are **not in the tracker at all**, and a fixed host
retires a workaround permanently where a guard only routes around it.

**The three worth filing, in order of how much the evidence adds.**

1. **A freshly-added slide's handle is good for exactly one `context.sync()`.**
   Not "GeneralException happens sometimes" — the probe asks three questions that
   isolate the cause. `shape-add-fresh-slide-proxy` (resolve and use inside one
   sync) → **yes**. `shape-add-held-slide-proxy` (same slide, same id, proxy one
   sync older) → **threw**, `GeneralException`,
   `errorLocation: SlideCollection.getItem`. `shape-add-positional-slide-proxy`
   (by index instead of id) → **yes**. So it is the *holding* that fails, not the
   id, not the slide, and not `getItem`. Most reports of this reach the exception
   and stop; the trio is the contribution.

2. **A non-empty `setSelectedShapes([id])` wedges the whole selection
   subsystem.** The call itself is taken — no error, no refusal, the sync
   resolves — and every selection call after it goes silent: neither resolving
   nor rejecting. Measured, twice: `getSelectedShapes` ran out a 90-second
   budget, and the `setSelectedSlides` behind it did too. This is a **third**
   claim, distinct from the two already filed — #3083 is `setSelectedShapes([])`
   failing to clear, #3698 is the empty call never resolving plus the picture
   interaction. The self-test's ladder produces the exact rung and the last one
   the host answered, which is the shape of evidence those two threads lack.

3. **Tag writes through a shape proxy several syncs old.**
   `InvalidParam passed to GetItem(id)`, code 5010, **46 times in one 38-item
   run**, leaving charts on the slide carrying no config. Related to #2903, so
   this is a corroborating comment with a volume and a reproduction rather than
   a new issue.

**What makes this a day of work rather than an afternoon.** Microsoft's issue
template wants a Script Lab repro, and an issue without one is triaged slowly or
not at all. Each finding has to be reduced to a minimal snippet a stranger can
paste — the probe is the wrong artifact to hand over, being a whole add-in.
That reduction is the work; the findings themselves are already written down.

**Attach the answer sheet.** `test/fixtures/host-answers-web.json` is a
reproducible record with the build stamp and requirement sets in it, which is
more than most reports carry. Check it before sending: every chart in this repo
is invented dummy data, but a sheet carries slide ids and raw host error text
from a real deck, and the owner should read what goes out under his name.

**Do not open duplicates.** Search the tracker first — the sweep's
`KNOWN_ISSUES` table is the list of what has already been read, and #2903,
#3083 and #3698 are corroboration targets, not new issues.

**Priority:** medium. High leverage and very slow: the issues this project
depends on have sat open for one to nine years, so nothing here should be
planned around a fix arriving. File it because a fixed host helps everyone
writing a PowerPoint add-in, not because it unblocks us.

### Drive PowerPoint on the web from Playwright, so the manual run is a command

Shipped. It is the round loop: `npm run cycle`, `scripts/round.mjs`, `scripts/cyclePlan.mjs`, and 290-odd archived rounds under `rounds/`. See docs/ROUNDS.md.

### ~~Take more than two draws per arm~~ — MOOT, the question it served is closed

Struck through by its own author: the question it served closed. See git.

### ~~READ THIS FIRST: a chart on a FRESHLY ADDED SLIDE cannot be grouped~~ — FIXED, measured 2026-08-25

36 of 36 charts on freshly added slides now group. Struck through by its own author; see git for the eight-commit route.

### THE WORK QUEUE that comes out of all this — 2026-08-16

Stage 0 of the matcher fix was done first, deliberately, because it decides the
shape of everything under it. **It dissolved the blocker it was meant to
confirm**, and split one item into two with very different prospects.

**What Stage 0 asked:** the fix wants to treat a short re-read on a slide we own
as a host lie rather than an instruction. Who can honestly promise the slide is
ours? `Grouping.refreshShapes` documents itself as exactly that guarantee — "the
caller can guarantee the target N shapes are the last N on the slide" — but every
call site sets it from `spansBatches()`, i.e. "this chart spanned sync batches".
**The field's contract and its use have diverged** (see the separate item below).
And the obvious substitute is unsafe: `onSlide === 0` means only that THIS RUN has
not drawn there, which is equally true of the user's own slide full of content.

**What Stage 0 found instead, and it is better news.** The ownership guarantee is
not needed for the short-read case, because **the matcher already proves
ownership by our own ids**. It matches `created[k].id` against the re-read's
collection; the 20 of 24 that matched on chart 4 are provably ours, on any slide,
without any new flag. Nothing has to be promised.

So the two failures separate:

| | chart 4 — short read | chart 5 — empty read |
| --- | --- | --- |
| what came back | 20 of 24, matched BY OUR IDS | nothing at all |
| ownership | proven, no flag needed | nothing to prove it with |
| what blocks it | a TRADE-OFF, not a signal | genuinely needs the guarantee |

#### 1. ~~Group a partial match rather than declining~~ — CALLED AND SHIPPED 2026-08-19

**The owner took it: group the 20.** The comparison it was decided on:

    group the 20    chart re-editable, 4 shapes stranded in its box
    group nothing   chart whole, and loses its config about 7 times in 10

**What shipped**, so a round is read against the right thing:

- The matcher groups a partial match instead of discarding it, and the
  short-read trace carries `grouping:` saying which — a round archive spans the
  change, and the same line meant "so it was not grouped" before it.
- **A bound the call did not name.** The subset is taken only when it is a strict
  MAJORITY of the chart. 20 of 24 is what every round produces; 1 of 24 would be
  a group holding one label with twenty-three shapes around it — the config saved
  and the object destroyed. Majority rather than a tuned share, because "more of
  the chart is inside the group than outside it" is a statement about the object.
  **Worth the owner's eye**: it is the one part of the change he did not specify.
- **`wholeMatch`**, because `freshMembers` now carries subsets and two readers
  need whole lists. `ungroupedFallback` builds the parts tag off that map, and a
  SHORT parts tag is worse than none — the next update deletes what it can name,
  redraws everything and leaves the rest, so the chart grows on every edit. That
  case is reachable: grouping can throw after a subset was chosen. The
  single-member tag-target swap is the other reader.
- **The stranded remainder is NOT recorded for the update to delete.** Tried
  first, and it is how this trade would become data loss: the only ids we hold
  for those shapes are the ones creation returned, and creation ids are what this
  host has been seen not to answer to (`withOwnId 7 of 7 … matched 0`, rounds
  068/069). The parts tag is a list the update path deletes BY.
- `chooseGroupMembers` unchanged, deliberately: it sees the re-read's members and
  not `created.length`, so it cannot tell two of two from two of twenty-four and
  cannot judge whether the group would still be the chart.

Three guards, each mutation-proven against the line it protects: the subset
groups and says what it left behind, a minority does not group, and a partial
group that then throws still writes a whole parts tag.

**WHAT A ROUND SHOULD SAY.** Staked before it runs, per the method below.
`same scale across the deck` fails at chart 4 in every round on record, and
chart 4 is the 20-of-24 one — so it should now group, keep its config, and the
scenario should reach chart 5 (the EMPTY read, item 2, which this does not
touch). `grouped the chart's shapes` should carry `partial=1 left=N:4` for it,
which is the first time that field reports an intended outcome. If chart 4 still
loses its config after grouping, the tag write is refused THROUGH THE GROUP —
which is what rounds 064/065 found on a freshly-added slide, and a different
problem from this one.

#### 2. The empty re-read (chart 5) — BLOCKED, and honestly so

Nothing came back, so nothing is proven ours and the positional rule would be a
guess on a user's slide. The only remaining route is `chooseGroupMembers`'
`use: "created"`, which hands `addGroup` the drawing proxies — and this host
throws on those, which takes the whole batch's tagging with it. So it needs
per-chart isolation of the grouping loop FIRST (today one try/catch wraps every
chart), and even then it is a gamble on a host that refuses those handles.

**And the paths that matter cannot make the promise anyway.** Stage 0 checked the
callers: only the self-test and the demo path add their own blank slides. The
pane's Insert draws on the slide the user is looking at, and Same Scale
(`updateChartsInSlides`) redraws onto the user's existing slides. **A fix gated on
"we added this slide" would improve the self-test and not the product** — which is
optimising the measurement rather than the thing measured, and this project has
made that mistake before.

#### 3. Retry a short or empty re-read once, after a delay — BUILT 2026-08-16, AWAITING ITS ROUND

**This did not come out of the archive. It came out of the office-js tracker on
2026-08-16, and it reframes items 1 and 2 above.**

**IT IS BUILT AND MERGED. What it has NOT had is a real host** — every claim
below is a prediction until a round tests it, and the prediction is staked here
on purpose so the round can refute it.

PowerPoint Online has a **known settling delay on a slide that has just been
materialised**: the shapes collection is not populated immediately after
`slides.add()`, and the community workaround — in
[#2903](https://github.com/OfficeDev/office-js/issues/2903), the very issue this
project already cites, and echoed in
[#5022](https://github.com/OfficeDev/office-js/issues/5022) — is to **wait one to
two seconds before reading it.**

This repo had read #2903 and recorded "upstream has nothing, `sleep(2000)` only".
That dismissal was reasonable when the failure looked like a tag problem. It is
wrong now: the failure has been isolated to *exactly* the state the workaround
addresses — a slide this run has just added. **We did not know that was our state
when we read the issue.**

**The shape of it: retry, do not delay.** Not a blanket wait before every
re-read — that taxes the 99% of charts whose read is already complete. Re-read
once more, after ~1.5s, only when the first read came back short or empty. Then:

    complete first read   costs nothing at all, which is charts 1-3
    short or empty        costs ~1.5s, on a chart that is otherwise losing its config
    still short           falls through to today's exact behaviour, nothing lost

**Why it beats both items above.** It needs no ownership guarantee (item 2's
blocker) and strands no shapes (item 1's trade). It is additive: every existing
path is reachable and unchanged, and the worst case is the current behaviour plus
a second and a half.

**What would prove it:** one same-build pair. `WHICH SLIDE THE CHART LANDED ON`
should move `freshly added, empty` off 1% grouped. If it does not move, the
settling delay is not our mechanism and the tracker lead is spent for the price
of one round.

> **ANSWERED, 2026-08-25 — it moved to 100%.** `freshly added, empty` reads 36 of
> 36 across eight consecutive rounds. The prediction staked above was correct and
> the mechanism is ours.
>
> **Read the delay before celebrating it.** The metric read `0/0` for the twenty
> rounds before that, so this had been true for an unknown number of rounds with
> nothing able to say so — see the entry in the journal. The prediction was
> right; the loop simply could not hear the answer.

**Stake the prediction before the round** (the method that earned itself
overnight): charts 4 and 5 of `same scale across the deck` group, and the
scenario stops failing at 34 of 34.

### ROUNDS 064 + 065 ANSWERED IT — half right, and the half that failed is the useful half

**Run as a PAIR on `bcd5773`, no merge between them, and the two are structurally
identical chart for chart.** `the settled retry repaired 2` in both; charts 4 and
5 take the retry, group, and are refused through `from: group×1` in both; the
slide ids differ between rounds, so what repeats is the position and the
freshness rather than an id. The noise floor is about counts and none of this is
a count — no third round is needed. What DID move is the downstream score, 4 of 8
then 3 of 8, which was always the noisy half.

**Charts 4 and 5 grouped**, both after the retry, both for the first time in five
rounds of failing identically. `the settled retry repaired 2`. No chart in the
round went ungrouped at all.

**The scenario still failed**, because the tag write is refused *through the
group*:

    4/8   tagging failed   from: group×1   slide 258#4111159134   5010
    5/8   tagging failed   from: group×1   slide 259#3844610554   5010

**This refutes "grouping is what saves a config."** That 2%-vs-79% split was
measured on a population that excluded freshly-added slides by construction —
they could not group, so they could never be in the grouped column. Round 064's
own split is `grouped 5, 2 lost = 40%`.

**The next target is the SLIDE handle, and it is named rather than guessed.** A
shape proxy carries its parent's object path; the group is made fresh in the
grouping batch but hangs off a slide handle Office has rewritten to
`slides.getItem(id)`, and a freshly-added slide's id does not round-trip here
(`shape-add-held-slide-proxy: threw`). The round says it outright: *the host
grouped through a slide handle two syncs old.* The members were never too old —
the parent was.

Keep the retry regardless: it is what turned an invisible failure into a named
one, and it costs nothing on a healthy read.

**What shipped**, so the round is read against the right thing:

- `REREAD_RETRY_MS = 1500`, `REREAD_ATTEMPTS = 1` in `powerpoint.ts`. The
  pre-grouping re-read runs in an attempt loop over a `pending` list; a chart
  whose answer was complete leaves the loop after the first pass and pays
  nothing. A fresh slide handle is taken per attempt, because Office.js rewrites
  the previous one to `slides.getItem(id)` and this host refuses shapes hanging
  off that.
- Two new trace readings. `re-reading the slide's shapes again after a settle
  delay` says the retry ran, with the slides it ran for; and both failure lines
  now carry `afterRetry: true`, so a round log distinguishes "short" from "short
  even after a pause" — the second is a much stronger claim about the host and
  the logs could not make it before.
- `hostFriction.emptyReReads` counts only reads the retry failed to repair, so a
  round is not told the host failed on a chart that came out fine.
- `faults.readsMissingFirst` in the fake — short once, then honest. Neither
  existing fault could express a slide that settles (`readsMissing` is permanent,
  `hollowReads` heals but only from empty), so without it the suite could show
  the retry costing a delay and never show it saving anything.

Guarded by two tests that assert **the group**, not the absence of a complaint,
and mutation-proven: `REREAD_ATTEMPTS = 0` turns both red.

**If the round refutes it**, the thing to check first is whether 1500ms is simply
too short — the tracker's reports say "one to two seconds" and this sits at the
bottom of that range deliberately, because the cost is paid on the live insert
path too.

#### DONE — cleared 2026-08-16

- **Revert `tagAnchorIndex`.** No measured effect across five rounds and four
  builds. The anchor is `created[0]` again, the draw loop loads every id, the
  matcher's contiguity deduction is gone, and `parts()` is `.slice(1)`. Kept, as
  planned: the `origin tag lost` trace line, the fake's `handleResolved` split
  (the host confirmed it), and the `taggedShape` test helper. The two assertions
  that flip back are annotated with what they asserted in between and why, and
  the reasoning behind the anchor move is preserved in full beside
  `CHART_ORIGIN_TAG` rather than deleted — it was good reasoning that this host
  does not reward.
- **`Grouping.refreshShapes` says one thing and is set from another.** Fixed the
  comment rather than introducing the guarantee, because the guarantee has no
  honest caller. Two sibling comments in the same file carried the same stale
  "last N on the slide" claim and were swept with it.

### Grouping is what saves a config, not the tag handle

**Measured 2026-08-15 over the whole archive, and it had been sitting there for
eleven rounds unqueried:**

    grouped      64 chart(s),  1 lost the tag =  2%
    NOT grouped  62 chart(s), 41 lost the tag = 66%

Per round it is almost mechanical — three grouped and none lost, two or three
ungrouped and two lost, round after round after round. `npm run rounds` prints it
now, under **DOES GROUPING SAVE THE CONFIG**.

**Why it is that stark.** When grouping succeeds the tag target is the GROUP —
a handle made in the grouping batch, never loaded, never resolved — and the write
lands. When grouping is skipped the target falls back to a `created` handle, and
that is the path that loses two charts in three.

**What it means for everything below.** The three sections that follow are about
WHICH handle the fallback should use, and four rounds plus a merged renderer
change (`tagAnchorIndex`) went into them. They are a question about the losing
path. A chart that never has to take that path does not care what the answer is.

`not grouping: no member handle this host will accept` carries `refreshed: 0`
every time, so what actually decides a chart's config is **whether the
pre-grouping re-read returned anything**. That is one level up from where this
work has been aimed, and it is where the next attempt belongs — the re-read, not
the tag.

**AND THE RE-READ FAILS THE SAME TWO WAYS EVERY ROUND.** Per-chart, identical in
rounds 042 through 046 without exception:

    1/8, 2/8, 3/8   GROUPED by=ids                          config kept
    4/8             re-read PARTIAL, matched 20 of 24       no group → config lost
    5/8             re-read EMPTY, drew 24                  no group → config lost

Two different faults wearing one outcome, and both are deterministic rather than
moody — the same chart, the same numbers, five rounds running.

- **Chart 4 matches 20 of 24, every time.** A partial match is thrown away on
  purpose (grouping a subset strands the rest), so this chart is ungrouped by our
  own rule and then loses its config to the fallback. Twenty of twenty-four is
  not noise; four specific shapes are not being named. Find those four and the
  chart groups.
- **Chart 5's re-read comes back EMPTY** with 24 drawn — the host listing nothing
  for a slide it has just been drawn on.

**The comment in the partial branch has been corrected** (`powerpoint.ts`): it
claimed an ungrouped chart "is still tagged, still re-editable", which the 66%
above refutes.

**BOTH BULLETS ABOVE ARE HISTORY AS OF 2026-08-19, AND THIS SECTION ASSERTED THEM
AS CURRENT UNTIL THEN.** They describe rounds 042-046. Two things have happened
since, and the file said the opposite of each in another section while still
saying this here:

1. **The partial branch is no longer left alone.** #586 groups the majority the
   host names. "The branch itself is left alone deliberately" was the rule for
   eight days and is not the rule now.
2. **The re-read WAS fixed, which is why neither harm is being chosen.** The
   settled retry (`REREAD_RETRY_MS`) repairs it: rounds 079-087 report
   `repaired=5, re-editable=8` every time, `same scale across the deck` passes,
   and no round records a short re-read since the retry shipped (42 in the whole
   archive, the `afterRetry` field ABSENT on every one — it postdates them all).
   "Fix the re-read and neither harm has to be chosen" was the right instruction
   and it has been carried out.

   **CORRECTED 2026-08-20 — that zero is a floor, not a count.** Until then the
   COLD read's outcome was never traced: attempt 0 pushed the entry onto the
   retry list and returned in silence. So "0 since the retry" compared a
   post-settle read against 42 COLD ones — different units, and the archive
   could not distinguish "the fault stopped happening" from "the retry hides
   it". The cold read is traced now; a few rounds will say which.

   **ANSWERED 2026-08-20, by rounds 111 and 112 as a pair: THE RETRY HIDES IT.**
   The cold read fails 8-11 times a round, every round, and that replicates.
   111: 11 failures (3 short, 4 empty, 4 zero), 11 retries, 0 survivors.
   112: 8 failures (2 short, 3 empty, 3 zero), 8 retries, 2 survivors (1 empty,
   1 zero). So the retry does not repair everything — but it repairs `short`
   FIVE TIMES OUT OF FIVE across the pair, and `short` is precisely what the
   subset branch below needs. The branch is starved by the one case the retry
   is best at, while the two cases that do survive are ones it cannot use.
   That is a sharper answer than either round gave alone.

So chart 4 does NOT match 20 of 24 any more — 084-087 group every chart with
`partial:0`. #586's subset branch is presently unreachable on this host, which is
not an argument against it: it is the guard for a regime this host has left, and
the round that would exercise it has not happened.

**THE `NNN#0` SLIDE LEAD IS DEAD — refuted 2026-08-15 from the archive, before
the instrument built for it had even deployed.** The draw trace already carried
both `chart` and `onSlideKey`, so the join could be done on rounds already
archived:

    1/8  slide 256#0            GROUPED   config kept    <- a #0 slide
    2/8  slide 256#0            GROUPED   config kept    <- a #0 slide
    3/8  slide 256#0            GROUPED   config kept    <- a #0 slide
    4/8  slide 257#0            no group  CONFIG LOST    <- a #0 slide
    5/8  slide 288#1168146411   no group  CONFIG LOST    <- an ordinary id

Identical in rounds 043-046. A `#0` slide carries the three best-behaved charts
AND one failure, and an ordinary slide fails too. **The id shape predicts
nothing.**

**WHAT DOES PREDICT IT: POSITION IN THE DECK-WIDE UPDATE.** Those five charts are
`same scale across the deck`, and its own summary says `the host flipped at chart
4 of 8, so the last 3 were not attempted` — in every round it has run (34 of 34
on 2026-08-16; `npm run rounds` has the live count).
Read the re-read outcomes in order and it is a decay curve, not a coin:

    charts 1-3   re-read matches all 24     grouped, config kept
    chart 4      re-read matches 20 of 24   partial, thrown away, config lost
    chart 5      re-read returns NOTHING    config lost
    charts 6-8   never attempted

**The hypothesis, and it points at our own perf work.** `updateChartsInSlides`
was deliberately made **one context, four syncs, flat in N** — the fix for
`doSameScale` spending 80 syncs across 20 contexts. That is the right shape for a
host that can hold a context, and this one degrades as a context is used: the
re-read is the first casualty, at chart four, every time. `#112` already made the
opposite call for the demo deck, one `PowerPoint.run` per slide, because a
context that has done too much stops answering.

**The experiment: chunk the deck-wide update into a fresh context every ~3
charts** — not per chart, which is what the perf work correctly removed, but at
the boundary the data actually shows. Cost is two extra contexts for an 8-chart
deck against 20 for the old shape. **Prediction: `same scale` moves off 17-of-17
failures, and charts 4-8 start keeping their configs.** If it does not move, the
context is not the limit and the re-read fails for a reason that survives a fresh
one — which is worth knowing for the price of one round.

Two further leads on the re-read, both cheap and neither yet followed:

- **The `NNN#0` slide ids.** Two of every round's added slides come back with an
  id whose second half is `0` — not the shape this host gives a slide it has
  finished adding — and the one line that names a slide in a tag failure named
  `257#0` in rounds 043, 044 and 045. The failure lines now carry the slide
  (`not grouping`, `a chart's tag could not even be queued`, `tagging failed`),
  so the next round says outright whether the charts that lose their config are
  the ones sitting on those slides. If they are, this is one fault and not
  three, and it is a SLIDE fault.
- **The re-read comes back empty rather than short.** `the re-read before
  grouping came back empty` appears beside the losses, and
  `shapes-items-count-honest` answers `unreadable` with `items` undefined. A
  slide whose shape collection will not enumerate cannot be grouped and cannot
  be tagged; both symptoms may be the same slide problem.

### The settle pass cannot repair what this host loses

MEASURED, 2026-08-15, rounds 029-034. `same scale across the deck` has failed six
rounds running at 4-5 of 8, and every loss ends the same way:

    settle pass: could not repair any config tag the drawing context lost
      charts=1  settled=0  lost=1  withId=0

`withId: 0`, every time. The settle pass is the second chance that exists to put
a config tag back on a chart the drawing context could not tag — and it works by
resolving the shape by id. This host does not give ids back for shapes it has
just made (`tag-through-refetched-shape: no-id`, three rounds, nine samples), so
the repair has nothing to try. It is not failing; it is unreachable.

The host's own error names the same wall from the other side:

    errorLocation: ShapeCollection.getItem
    statement: var shape = shapes.getItem(...) /* originally addTextBox(...) */;

Office.js rewrites the CREATION proxy's object path into `getItem(id)`, so even
the code path that never mentions an id has become an id lookup by the time the
host sees it.

What makes this actionable rather than another dead end: round 033 and 034 both
answered `tag-the-creation-proxy-a-sync-later: yes`. **A write through the
creating handle goes through on this host, a sync later, when neither the id nor
a read does.** So the settle pass has a route it is not using.

**The settle pass cannot be where this is fixed.** It opens a FRESH
`PowerPoint.run`, so the creation handle — the one thing this host still accepts
writes through — does not exist by the time it runs. Its two routes are an id and
a collection read, and this host refuses both (`no-id`; `shapes-items-count-honest:
unreadable`). `withId: 0` is that fact showing up as a number.

So the retry belongs in the DRAWING context, before the handle is thrown away:
when the tag write is refused there, write it again through the proxy that
created the shape rather than handing the chart to a repair pass that has nothing
to work with.

This is the same goal `binding-names-shape-later` was written for — "whether the
repair pass can be given a handle that does not go through
`ShapeCollection.getItem(id)`" — and that question answers `unreadable` on this
host, four rounds running. `tag-the-creation-proxy-a-sync-later` is a second
route to it that does answer `yes`, and it needs no 1.8 surface.

**ATTEMPTED 2026-08-15, and the one-line version does not work.** Worth recording
so nobody spends the hour again.

The rule was measured first: it is RESOLUTION, not age. `faults.strictTags`
models "refuse anything older than one sync", and this host does not do that —
`tag-the-creation-proxy-a-sync-later` answers `yes` four rounds running, so a
handle that made a shape keeps taking writes however old it is. What it refuses
is a handle a `load()` has resolved, which Office.js rewrites into
`shapes.getItem(id)`. `faults.refuseTagWritesOnResolvedProxy` now models exactly
that, and it is the first thing in the fake that can.

The obvious fix follows and is wrong: `finishCharts` replaces the tag target with
the pre-grouping re-read handle (`tagTargets[i] = fresh[0]`), so keep the
creation handle instead. Tried, and the reproduction still failed — because the
pass loads the CREATED shapes too before the tag is written, to read their ids
for the parts tag. By the time the write happens there is no unresolved handle
left to use.

So the fix is not a swap, it is an ordering change: one handle has to stay
unresolved all the way to the tag write. That is a change to the shape of the
whole pass and wants a session of its own.

**THE BUDGET IS MEASURED NOW, and it is not a constraint (round 039, 2026-08-15,
replicated from the recovered round-42 log).**
`how-many-syncs-a-creation-handle-survives` answers **`survives-8`** — five
samples across two builds, `stable: true`, and every sample taken while the host
was in the `collection-refused` regime. An unresolved creation handle accepted a
tag write at each of eight successive syncs.

That removes the gamble the probe was written to remove. The ordering change does
not have to move the tag write EARLIER in time; it only has to keep one handle
unresolved. Eight syncs is past anything a chart does, so the constraint is
purely `load()` — and the resolver has a name and a line number:
`powerpoint.ts:6956`, `created[k].load("id")` on each batch's own sync, which
resolves every created shape including `created[0]`, the tag target. Round 039's
host dump shows the consequence with the ids in it: `shapes.getItem("27")
/* originally addTextBox(...) */`, where 27 is the first batch's title box.

`ungroupedFallback` is NOT that resolver, so swapping it changes nothing:
`parts()` slices index 0 off deliberately, and the anchor is never in its
`load("id")` list.

**AND THE LAST ALTERNATIVE IS CLOSED (round 040, 2026-08-15).**
`tag-through-refetched-shape` — "can a fresh shape be tagged through a handle
re-fetched by its own id?" — answers **`no-id`**, stable over three passes: *the
fresh shape would not report an id*. Not "the id is refused"; there is no id to
re-fetch it by. Route by route on this host:

    created    works, and does not age out (survives-8) — the pass throws it away
    refreshed  resolved by a load, rewritten to getItem(id), refused
    by-id      cannot be constructed at all — no id
    group      never loaded, but only exists where grouping worked, and here it does not

So the ordering change is not the best remaining option, it is the only one.

One caveat the probe cannot cover, stated so the fix is not built on it: it asks
through `ctx.scratch()`, a slide resolved in the same batch, while production's
parent is a slide handle Office has rewritten to `slides.getItem(id)`. A refusal
originating in the PARENT path would be invisible to this question. The host's
`errorLocation` is `ShapeCollection.getItem` rather than the slide, which is
evidence against that reading but not proof.

### Why it is a redesign and not a reordering — the three-way knot

Attempted properly on 2026-08-15 and stopped at a constraint rather than at
effort. Recorded because every route below LOOKS like the fix for about twenty
minutes, and the thing that kills it is somewhere else in the file.

Three requirements, each independently non-negotiable, and no two-line change
satisfies all three:

1. **The tag target must be unresolved when the tag is written.** Measured:
   `survives-8` and `tag-through-refetched-shape: no-id`. A `load()` is what
   makes Office rewrite the handle into `shapes.getItem(id)`, which this host
   refuses.
2. **The pre-grouping re-read needs every created id.** `powerpoint.ts:7347`
   matches created shapes to the re-read BY ID, and a PARTIAL match is thrown
   away on purpose — the comment above it is explicit that keeping one strands
   shapes inside the chart's box, with `grouped … partial=1 left=0:4` from a real
   slide as the evidence. Skip loading the anchor and every chart becomes an N-1
   of N partial match, so **grouping stops working everywhere, desktop included**.
   The positional fallback is not available either: it is deliberately reachable
   only when NOTHING matched, because a slide holding the user's own shapes can
   satisfy `items.length >= created.length` and "the last N" would then group the
   user's content in, to be deleted with the chart on the next update.
3. **The tag must not land before the chart is complete.** Tagging `created[0]`
   in the first draw batch is the obvious way to get requirement 1, and it makes
   a stalled or stopped draw leave a tagged partial chart behind.

   **DO NOT CITE `stop a run part-way` AS EVIDENCE FOR THIS — corrected
   2026-08-19.** This entry rested on that scenario's pass ("nothing left
   claiming to be a chart"), and the scenario has never drawn a shape: it calls
   `requestStop()` BEFORE the insert, `throwIfStopped()` sits at the top of the
   batch loop, and across all 69 archived rounds it has committed **zero**
   batches. It tests that a stop asked for before a draw prevents it — a real
   promise, but not this one. **Requirement 3 is currently untested**, and this
   design constraint rests on reasoning alone.

   The seam for a real test is `onPhase("commit", …)`, which fires once per
   batch: requesting the stop from the first commit aborts a genuinely
   half-drawn chart. It is not a change to slip in — a mid-draw abort leaves a
   PARTIALLY DRAWN SLIDE, and `same scale across the deck` discovers its chart
   population from that same deck, so it would contaminate every scenario after
   it in the round. It wants its own pair.

The two routes the earlier note proposed both die on requirement 3 or on cost:

- **Tag early, drop the anchor's tag once the group's lands.** `TagCollection.delete(key)`
  does exist, so the mechanism is real, and on this host grouping never succeeds
  so the delete would never need to run. It still leaves requirement 3 broken.
- **A second key only the recovery path reads.** For the chart to be re-editable
  the READERS have to accept that key, and `chartTagsOf` does not read the scene
  tag today — so deduping "the group and its anchor are one chart" costs another
  per-shape tag lookup deck-wide, on a host already observed reading collections
  short. That is a real cost against the scan that Same Scale depends on.

**The shape of the actual fix, now that the knot is stated:** the tag target has
to be a shape created in the LAST draw batch and never loaded — then it is
unresolved (1), every OTHER created shape keeps its id for the re-read (2), and
nothing is tagged until the chart is finished (3). That means the anchor stops
being `created[0]`, which `ungroupedFallback`'s "everything after index 0 is a
part" and `CHART_ORIGIN_TAG`'s frame both assume.

### It was BUILT, and the fake refused it — 2026-08-15

Written in full and reverted the same hour, because the thing that refuses it is
worth more than the change was.

What was built: `tagAnchorIndex` moving the anchor to the last shape drawn, the
draw loop's `load("id")` running one shape behind so that shape is never
resolved, `ungroupedFallback`'s parts list taking everything but the anchor, and
the pre-grouping matcher learning that exactly one shape is deliberately
unmatchable — deducing it as the contiguous neighbour of the last matched
sibling, and falling through to today's don't-group behaviour when that does not
land cleanly. `CHART_ORIGIN_TAG` needs no change: it records the TAGGED shape's
own corner and shifts by how far that shape has moved, so it is self-consistent
whichever shape carries it. 211 office-render tests green, including the eleven
grouping ones.

**Then the reproduction still failed, and the trace said why.** The config tag
write was refused, `from: created`, in the drawing context — with the anchor's
own `load` held back. Instrumenting the fake's `load()` with a stack per call
found the resolver: not the draw loop, but `groupAndTagAll`'s
`shapes.load("items/id")`, the pre-grouping re-read. It resolved eight shapes;
the draw loop had loaded seven.

**The fake makes that read poison every handle onto the shape, and that is the
open question.** `freshHandle` is a `Proxy` over the same object: it deliberately
gives the fresh handle its own `syncCreated` and its own tag writer, and shares
everything else "because it is one shape". `loadedProps` is on the shared side —
so resolving a collection item resolves the creation handle too. Office.js gives
each proxy its own object path, which says they should be independent; the fake's
own treatment of `syncCreated` says the same. Either `loadedProps` is handle
state modelled as shape state, or it is the one place the fake is right.

**No round can be argued into answering it, so it is now asked directly:**
`collection-read-poisons-the-creation-handle` ships with this note.

- `yes` → the read is innocent, the fake is wrong, and the fix above works as
  written. Rebuild it from this description; it is not large.
- `refused` → no arrangement of loads saves the drawing context's write while
  grouping needs a re-read, and the second-key route is the only one left.

### ROUND 042 SAID `yes`, AND THE FIX IS IN — 2026-08-15

Three passes, `stable: true`, every one taken while the host was refusing
collection reads. The pre-grouping re-read does not touch the creation handle.
The fake was modelling handle state as shape state — `loadedProps` shared across
every handle onto a shape, while `freshHandle` had always given each handle its
own `syncCreated` and its own tag writer. Split, and the collection read now
loads through a fresh handle, which is what it hands back anyway.

Shipped exactly as described above: `tagAnchorIndex` (the last shape drawn), the
draw loop's `load("id")` one shape behind, the parts list taking everything but
the anchor, and the matcher deducing the one deliberately-unmatchable shape as
the contiguous neighbour of the last matched sibling. The reproduction under
`refuseTagWritesOnResolvedProxy` now shows the drawing context's own write
landing with no repair, and asserts the absence of a `tagging failed` line rather
than the presence of one.

**Two things it changed that were not planned, both kept:**

- **`origin tag lost` is its own trace line.** The config tag commits one sync
  before the origin tag, so the good case on this host is config-lands-origin-
  fails, and the old shared catch called that "charts are not re-editable until
  repaired" — false, and about to be the common case.
- **A short deck scan can no longer see a tail-anchored tag, and that is the
  price.** `faults.readsMissing` drops from the tail, so a scan blinded that way
  misses the shape carrying the config; `web-host.test.ts` asserts the new
  behaviour and, in the next line, that an unblinded scan finds the chart — so it
  is a visibility cost, not a lost tag. Taken deliberately: the write failed
  CERTAINLY, on every chart big enough to span batches, four rounds running,
  while this scan fails intermittently.

  **AND THE COST IS MOOT ON THIS HOST — rounds 043 and 044.**
  `which-end-a-short-read-drops` answers `unreadable` both times: *the collection
  would not list its items*. `shapes-items-count-honest` says the same on the
  same sheet, with `items` undefined rather than short. **A read that returns NO
  list cannot drop one end rather than the other**, so where the anchor sits in
  the collection cannot decide whether a scan sees it. The trade was real to
  take and turns out to cost nothing here. The question stays for a host that
  lists items — it costs one scratch slide — but nothing should wait on it.

**And the trace now says which handle was used**, which is what the six rounds
before it could not. `tagging failed` carries `from: created×N, refreshed×N`, one
of four routes:

    created    the proxy that drew the shape, never loaded — writes go through
    refreshed  the pre-grouping re-read, RESOLVED by a load and therefore
               rewritten to shapes.getItem(id)
    group      the group made in the grouping batch, also never loaded
    by-id      an explicit getItemOrNullObject(id)

The field earned itself on its first run. The suspect going in was `refreshed`;
the trace answered `created×1` — the target WAS the drawing handle and the write
was refused anyway, because the pass loads the created shapes before it writes.
That is the evidence for "ordering, not swap", and it took one test run rather
than a round to get.

Also measured while trying: arming `refuseShapeIdLoads` alongside — to starve
the settle the way the real host does — models something HARSHER than the host.
It makes the insert throw outright, where a real round carries on and merely
loses the tag. The reproduction in `test/office-render.test.ts` therefore pins
the first half (the drawing context's write is refused) and names the second (the
settle repairs it here and cannot there).

NOT attempted further, deliberately. Two things have to be true first and only
one of them is: the write-through-the-handle answer needs a third round, and the update
path needs the same treatment `insertSceneIntoSlide` already has — a test in
`test/office-render.test.ts` shows the INSERT path survives a refused
`load("id,left,top")` unharmed, so whatever costs the tag in the update path is
not the shared batch, and guessing which line it is has already produced one
wrong theory tonight.

### Stop the probe carrying seventy scratch slides through a round

**MOSTLY WRONG — THE DECK NEVER GREW. Corrected 2026-08-14 from round 29.**

The premise below was that the deck fills with abandoned scratch slides and that
this is what makes `listing the deck's slides` slow enough to time out. Round 29
measured it directly and says otherwise:

    gave the scratch slides back  returned=0 swept=0 left=73
      deckBefore=1 deckAfter=1 shrankBy=0 stillListed=0
      heldIds=["4123571114#123571113", …]  deckIds=["287#62081387"]

Seventy-three slides taken, and the deck is **one slide before and one after**.
None of the held ids is in the deck; `stillListed` is zero. The deck evidence in
the same round agrees — 7 slides scanned, 6 added, all by the self-test. The
scratch slides never landed at all.

Look at the two id namespaces. The scratch slides come back as
`4123571114#123571113` and the deck lists `287#62081387` — that is the
`getItemAt` → `getItem(id)` rewrite from `RESEARCH.md` seen from the other end: a
freshly added slide reports an id the deck will not answer to, so the probe can
neither find nor delete it, and its counter climbs while the deck stands still.

So the cost is NOT paid in an O(deck) slide listing, because the deck stays at
one; and the wedge these numbers were used to explain has since been identified
as the host's editing session dying (`docs/ROUNDS.md`, "The wedge"), which the
probe does not cause. **The counter is the bug, not the deck.** What is worth
building is not a cheaper clean-up but an honest count: `left=73` should say
`never landed=73`, and the round after that can ask why the host accepts an add
whose slide never appears.

The rest of this entry is kept because the two reverted fixes below are still
reverted for the reasons given, and both would still break the same guards.

Two fixes were built and both reverted on 2026-08-14:

- **Delete the abandoned slide, then add its replacement.** Safe-looking, because
  the abandoned slide is the deck's last at that moment so a positional delete
  can take it. It leaves the run with NO slide when the add then fails, and
  `asks every question even when the host keeps losing the scratch slide` caught
  it — `shape-add-fresh-slide-proxy` stopped being put at all.
- **Add the replacement, then delete the old one by its index.**
  `deleteTrailingSlides` deletes any range despite its name, so this works
  mechanically. It broke four guards, including
  `left its replacement slides in the deck` (11 where 2 was expected) and
  `shape-resolve-held-slide-proxy` changing its answer. Deleting mid-question
  does not merely tidy the deck — it changes what the probe measures.

So this is not a tidy-up, it is a change to the instrument, and it needs its own
round with a control rather than a green suite.

**The pass-boundary variant was then tried too, and it is blocked by something
else — the accounting, not the timing.** Sweeping between passes does leave the
questions alone, exactly as predicted: none of the four guards that killed the
per-replacement attempts fired. What fires instead is `leaves the deck exactly as
it found it`, and it keeps firing through every reconciliation:

- the end-of-run clean-up derives what it owes from `scratchIds.length` and from
  the deck's growth since the run began (`positionalSweepPlan`), and BOTH are
  invalidated by a slide going back early;
- scoping the clean-up to what is still outstanding fixes the double-count and
  then under-sweeps, because the positional plan is capped by `deckNow -
  deckAtStart`, which the early handbacks have already shrunk;
- feeding the handbacks back into the plan as `alreadyDeleted` changes nothing —
  the deck still ends with slides the run cannot account for.

Three spellings, one cause: **the clean-up assumes every scratch slide is still
there when it runs.** That assumption is load-bearing in `slidesActuallyReturned`,
`positionalSweepPlan` and the `scratch-slides-returned` answer, and no amount of
arithmetic at the call sites removes it.

The clean-up rewrite was then done — `outstandingScratch`, the ledger, tests that
exercise an early return — and **the pass-boundary sweep still fails, so the
accounting was not the blocker after all.** The instrumented numbers say why, and
they are worth more than the three theories that preceded them:

    taken 10 · handed back early 5 · deckAtStart 2 · deckBefore 12

`deckBefore` is read when the clean-up starts, after the handbacks. 2 + 10 = 12.
**The deck at clean-up time is as if nothing went back at all** — even though
`deleteTrailingSlides` reported five successful deletions and the deck visibly
shrank at the moment each ran.

So the handback deletes A slide and not necessarily THE slide it named. Which
makes sense of everything else: this host does not honour delete-by-id, and the
only reason the end-of-run sweep works is that by then every remaining scratch
slide is a **contiguous trailing block** — a positional RANGE needs no slide to
be identified individually. During a run that block does not exist, because the
live scratch slide sits at the end of it.

That looked like the real prerequisite — "there is no way to delete one known
slide mid-run on this host" — and **it is wrong.** Asked directly rather than
inferred from a misbehaving run, the answer is yes: `deleteTrailingSlides(i, 1)`
removes exactly the slide at index `i`, confirmed by id either side, and it keeps
doing so under `renumbersOnAdd` — the very behaviour that makes delete-by-id
useless here. `test/delete-one-slide.test.ts` pins it, and a mutation that
deletes a fixed index instead turns it red.

The run's bookkeeping is not at fault either: on clean main the ids it records
match the deck's growth exactly, 13 for 13.

So both things the last four attempts blamed are sound.

**The pass boundary was then instrumented — deck length either side of every
handback, and again at the next one — and the deletes are innocent too:**

    boundary 1   deck 7 -> 3   reported 4   actual shrink 4   MATCH
    boundary 2   deck 8 -> 7   reported 1   actual shrink 1   MATCH

Every slide the handback claimed to remove did leave, immediately, verified by a
count either side. The leak is in the SEGMENTS BETWEEN, and it is an ADD problem
rather than a delete problem:

    segment 1 -> 2         deck 3 -> 8  (+5)   run noted +4
    segment 2 -> clean-up  deck 7 -> 12 (+5)   run noted +2

The deck grows faster than the run records, and only while the handback is
active. Four candidates for the un-noted adds have been read and eliminated:
`addScratchSlide` reports a slide it could not remove (`onAdded` fires on exactly
that branch), `deleteSlideById` verifies with `slideIsGone` before returning
true, `deleteSlideByPosition` re-reads the deck and returns false rather than
guessing when the id is absent, and on clean main the run's own count matches the
deck's growth exactly.

So the question is now precise and small: **which call adds a slide between one
pass boundary and the next without the run recording it, and why only once slides
have been handed back?** Everything upstream and downstream of it has been
measured and is sound.

The next attempt logs the deck length either side of every single
`addScratchSlide` in one run — not the probe's behaviour, not the accounting,
just which add fails to conserve the count. That is one number and it ends this.

The ledger from the clean-up rewrite stays: it is right, it is tested, and it
removes a genuine assumption. It simply was not the blocker.

Four implementations have been reverted here, and every theory each was abandoned
on has since been disproved by a one-line experiment: the primitive works, the
bookkeeping is accurate, the deletes land where aimed. The pattern is worth more
than any of them — reasoning from a full probe run was wrong every time, and
asking a single component directly gave a clean answer in under a minute, every
time.

### Improvements the 2026-08-15/16 overnight run was owed — ALL CLEARED 2026-08-16

All cleared 2026-08-16. See git.

### Method that earned itself overnight, and should be kept

- **Run the same build TWICE.** Nineteen archived rounds and almost every one on
  its own build, so no two could be compared. The one exception, `cabb357` run
  twice, is the only noise measurement this project owns — 1 and 5 for the same
  fault with nothing changed. Pairs are what let `tagAnchorIndex` be called
  no-effect rather than unproven.
- **Judge a change on a count that did NOT move, or on a line that appears where
  none did.** Those are the two readings that survived; every count that moved
  moved with the host's mood.
- **Ask the archive before spending a round.** Grouping-vs-config, the `#0` ids
  and the fresh-slide split were all answered from rounds already filed. Three
  hypotheses died for the price of a query.
- **A question that cannot be asked is an answer about the harness.** Three
  probes died on the scratch slide refusing to enumerate a collection; the
  measurement moved into production and worked first time.
- **Stake a prediction before the round that tests it**, especially against your
  own work. The one staked on 2026-08-15 failed, and the failure is what
  redirected the effort.

## 2. Rejected or already covered (do not re-propose)

- **A PowerPointApi 1.8 binding as a durable handle to a drawn chart** — the
  last untried route out of `same scale across the deck`, and it cannot be
  ASKED on this host. Not refuted: unanswerable, which is a different and more
  annoying thing.

  The idea is sound and stays written down in case the host changes.
  `ShapeCollection.getItem(id)` is where every 5010 lands, and it is what leaves
  a chart drawn and nameless — no group, no config tag, nothing to settle one
  onto. `bindings.add` takes the live Shape proxy inside the batch that CREATED
  it, so it needs neither an id round trip nor a collection read, and the
  document persists it. If it worked, the repair pass would have a handle it
  does not have today.

  The probe `binding-names-shape-later` asked eight times across eight rounds
  and never once reached its own question:

      13  no-scratch-slide   15  no-scratch-slide   18  no-scratch-slide
      14  no-scratch-shape   16  no-scratch-slide   19  no-scratch-slide
      17  no-scratch-shape   21  no-scratch-shape

  Six distinct causes were found and five of them fixed — the end-of-run second
  pass (#353/#356), the poisoned slot after `shape-add-held-slide-proxy` (#360),
  an unclaimable add when the host renumbers a slide (#364), a slide-acquisition
  budget of ninety seconds against a question budget of eight (#364), and the
  probe's position in the order (#359). Each fix got it one step further. The
  sixth is not a defect and is not going away.

  **A shape question needs the scratch slide resolved once per batch, and this
  host hands out about one usable resolution per slide** — round 21's trace
  shows nearly every question taking a replacement before it can answer. The
  binding question needs more resolutions than any other shape probe, because
  an honest answer needs a CONTROL: the same batch without the binding, proving
  the host would have taken a plain shape. Without that control the probe claims
  `commit-threw` on a host refusing every shape add, which
  `test/host-probe.test.ts` catches. Running the control second does not help —
  a failed sync poisons its own context, so the control cannot resolve the slide
  afterwards. Both were tried and reverted, in that order.

  So: honest and unanswerable, or answerable and dishonest. Recorded rather than
  re-litigated. `same scale across the deck` therefore has **no route found**,
  and the failure under it (a chart that draws and is then unreachable) stands
  as a host limit. Anything that revisits this needs a host that resolves a
  scratch slide more than once, and the probe will say so the day one appears.

- **A golden-image gate on the generated deck** — rejected as a hash comparison,
  and largely covered as a structural one. The cheap half of this DID ship:
  `validate-ooxml.mjs` checks the deck against the OOXML grammar and
  `verify-deck.mjs` gained a duplicate-`cNvPr`-id check that neither tool had,
  both gating CI on `examples/showcase.pptx`.

  The measurements that made the visual half look feasible are real and still
  true (2026-08-02): LibreOffice headless renders the deck to PDF in 7s and to
  122 PNGs in 17.6s, byte-identically across fresh profiles, and its usual PPTX
  weak spots do not apply — this generator emits no gradients, no pattern fills,
  no `normAutofit`, no effects. What kills it is the **pinned container**. The
  deck asks for Segoe UI and Calibri, neither present in CI, so every render is a
  substituted render: self-consistent, different from a laptop with the real
  fonts, and different again after any LibreOffice minor bump. The baseline
  churns, and a visual gate that cries wolf is one that gets switched off — the
  failure this entry predicted about itself.

  **The structural half is thin rather than missing.** All three renderers
  consume the SAME scene graph, and `scripts/visible-charts.mjs` already proves
  that scene rasterises to something visible (ink present, more than one colour,
  inside the frame, at least half its usual coverage). A pptx-only invisibility
  bug would have to originate in `pptx-paint.mjs`'s own mapping, which is
  asserted per node at 100% statements / 88% branches, on top of `verify-deck`
  auditing the bytes and `validate-ooxml` checking the grammar.

  **Do not read this as "the pptx output is pixel-checked" — it is not.** The
  headless pptx renderer is the one of the three with no look at its pixels, and
  the residual gap is a composition bug that survives every node-level assertion.
  That is a narrow target, not an empty one. If it is ever revisited, copy the
  shape of `visible-charts.mjs` (assert properties that hold on any renderer, no
  committed images) and reuse its `judge()` wholesale; the part that wants
  pinning is LibreOffice, not the method.

  Two traps worth keeping whatever is built. Frame any such gate as "did our own
  output change", never "does this match PowerPoint" — no FOSS renderer is close
  enough for the second claim, and a gate that overclaims gets ignored. And
  **never** use a LibreOffice round-trip as a PowerPoint proxy: converting the
  showcase deck back to `.pptx` silently deleted all 122 `ppt/tags/*.xml` parts,
  leaving 121 charts non-re-editable. Anything learned that way is a fact about
  LibreOffice.

- **An image / icon node** — not reachable in the live add-in, so nothing can
  be built on it. PowerPoint's `ShapeCollection` exposes exactly
  `addGeometricShape`, `addGroup`, `addLine`, `addTable`, `addTextBox`;
  `addImage` exists only in the Excel namespace. The one route to pixels is
  `ShapeFill.setImage`, which is **PowerPointApi 1.8** against a manifest
  pinned at **1.4** — it would vanish silently on any older host. It also
  bloats twice over: `ChartConfig` round-trips through a shape tag verbatim
  (`tagData: JSON.stringify(cfg)`), and PptxgenJS does not dedupe media, so an
  identical icon is re-embedded once per point. The only thing presets cannot
  do is a **logo** (a competitor-positioning map). If that is ever wanted, the
  shape is: user-supplied data URI (no asset library, so no licensing call),
  gated on `supports("1.8")` like grouping already is, and drawn *additively
  over* a marker so an old host degrades to a plain point rather than a
  missing one.
- **Scatter/bubble point icons** — the primitive above is unavailable, and the
  real itch was not bitmaps: shape as a categorical channel is what Excel and
  Highcharts mean by point markers, and PowerPoint's presets give it filled and
  native at 1.4. Shipped as `scatter.markers`.
- **Heatmap per-cell icon overlays** — a heatmap's rows ARE its data series, so
  a "Glyph" row yields one row of glyphs, not a per-cell matrix; carrying a
  genuine second dimension would mean inventing a datasheet convention for
  every heatmap, for a want nobody has stated. The demonstrable gap was
  narrower and is fixed: a diverging scale states direction in hue alone, and
  its strongest + and − sit at 1.12:1 in greyscale — the same tone — with no
  label to fall back on under `sizeEncode`. Shipped as `heatmap.symbols`.
- **An X / cross marker, and a five-point star** — a marker shape has to
  reproduce its OOXML preset *exactly*, because the SVG renderer draws the
  points while the PowerPoint renderers name the preset, and `markerScale`
  measures the area off those same points. A shape that only approximates its
  preset therefore breaks the bubble's "area ∝ size" claim in the deck while
  keeping it in the preview — the worst kind of divergence, since the preview
  looks right. `mathMultiply` fails on angled arms (and is redundant with
  `plus` at 3-4pt). `star5` fails on its `hf`/`vf` stretch: the preset widens
  itself 1.05146x and heightens itself 1.10557x so the star fills its box
  (a 5-point star spans 1.902R by 1.809R, and those factors are exactly
  2/1.902 and 2/1.809), which makes PowerPoint's star **16.2% larger in area**
  than an inscribed SVG one. Reproducing that faithfully is possible on paper
  but not verifiable here without a PowerPoint rasteriser, so the set stops at
  the shapes whose geometry is exact: circle, square, diamond, triangle, plus.

- **Waterfall connector re-routing** (drag to skip columns) — the drag is out of
  Office.js reach, and the rendering feature underneath is ill-posed: the
  connector is not an object with authored endpoints, it is a derived assertion
  that "this level carries into the next bar". Re-routing it is truthful only
  where the skipped columns leave the level unchanged — and there it already
  works, via `spacerIndices` and `totalIndices`. Anywhere else the line would
  end in mid-air pointing at a bar it does not touch. What was actually wanted
  shipped instead as `waterfall.detailGroups`.
- **Scatter-on-combo** — a continuous-x scatter over categorical columns needs a
  second x scale that means nothing beside category slots. The coherent reading
  — unconnected marks at category positions — shipped as `type: "marker"`
  (overlay) and already existed as `decorations.barStyle: "dot"` (clustered).
- **Free 2D bubble repulsion** — both axes carry data, so moving a marker in 2D
  corrupts two readings at once with nothing to bound them. The honest version
  shipped as `scatter.spread`: one named axis, hard-capped, cap printed in the
  footnote.
- **Gantt resource capacity vs load** — a per-resource per-week histogram against
  a capacity line. That is a stacked column chart with a `Target` row, both
  shipped; it needs a value axis, which a timeline does not have. Recipe in
  docs/MANUAL.md. (Lane grouping itself shipped as `gantt.lanes`.)
- **Dial / needle gauge** — bullet charts replace it deliberately (Few); low
  deck demand. Note the *semi-circle scorecard* gauge did ship (`doughnut` +
  `pie.semi`); what stays rejected is the dial-with-needle and its threshold
  bands.
- **Sankey / chord / arc** — need curved ribbons; infographic genre.
- **Ridgeline** — stacked density curves; academic register. (The single-column
  `violin` kind shipped at owner request; ridgeline stacking not pursued.)
- **Stream graph** — feasible but editorial aesthetic; no deck demand.
- **Pictogram with icon libraries** — needs the image node above, which is not
  reachable in the live add-in at all (see the first entry), plus an asset
  library. Waffle is the deliberate substitute: it covers the part-to-whole
  genre with square cells, it does not render icons.
- **Histogram as a kind** — the look is `clustered` + `gapWidth: 0`, both
  shipped. If this is ever revisited, auto-binning raw samples into categories
  is the only real gap; the bar geometry is not. (`histogramBins` in
  src/core/format.ts bins over a fixed domain, but nothing derives categories.)
- **Choropleth maps, 3D, drill-down interactivity** — out of scope by design
  (see CLAUDE.md). Tilemap proportional-area cartograms and tilemap drill-down
  fall here too: hard/infeasible.
- **Population pyramid, plain dot chart** — already covered by `butterfly`
  (+ `butterfly.split`) and `decorations.barStyle: "dot"`.
- **Radar vertex markers** — already there: radar emits `marker-*` ellipse
  nodes, which the Office.js renderer draws, so they appear in the live add-in
  too.
- ~~Candlestick / OHLC~~ — shipped as the `candlestick` kind at owner request,
  despite the thin consulting-demand signal.
- ~~Alt text in the headless pptx renderer~~ — **shipped 2026-08-17.** The
  rejection was right about pptxgenjs and stopped one step short: this project
  already hand-patches the generated OOXML (`injectGroupsAndTags`), and it
  writes the very element alt text lives on — the group's `<p:cNvPr>`. One
  attribute, and `scene.desc` now reaches the deck. Image-mode charts take
  pptxgenjs's own `altText`, which it does expose on a picture. Kept here as the
  shape of the mistake rather than as a candidate: a limit of a DEPENDENCY is
  not automatically a limit of the product when the product already
  post-processes that dependency's output.

## 3. Findings from the round archive

What ~290 rounds against the live host have established. Kept because the
finding outlives the fix: each one says what was measured and how, so nobody
re-derives it. Open questions among them are marked as such.

### The probe has been blind on GROUPS for the whole archive — found 2026-08-16

Found and fixed 2026-08-16; the probe reads groups now. See git.

### A SINGLE-BATCH CHART CANNOT GROUP ON THIS HOST, AND THAT IS OUR BUG — 2026-08-16

`npm run rounds` prints it under **DID THE CHART SPAN BATCHES**. Pooled over 41
rounds and 537 draws:

    spanned batches   452 draw(s),  353 grouped = 78%
    one batch only    214 draw(s),   49 grouped = 23%

**353 of 452 against 49 of 214.** (First reported here as 333 of 333 — 100% —
which was a bug in the pooling function and not a fact: it searched a fixed few
entries for the verdict and did not count `not grouping` at all, so every honest
decline was dropped from both arms. The separation survived the correction; the
absolutes did not. On builds carrying the settled retry the multi arm really is
10 of 10.) The sharpest separation in the archive, and
unlike the fresh-slide split it is not a fact about the host — it is a
consequence of one line of ours.

**The mechanism, end to end, every link measured:**

1. `refreshShapes` is set from `spansBatches(created, opts)`, so **only a
   multi-batch chart gets the pre-grouping re-read.**
2. A single-batch chart therefore hands `addGroup` the raw `created` proxies.
3. This host refuses those: `InvalidParam passed to GetItem(id)`, 5010, at
   `grouping the chart's shapes`.
4. The failed group takes the tag with it — `target.tags` comes back
   **undefined**, and it is *always* this: **155 of 155 across the whole archive,
   every one immediately after a 5010 group, zero exceptions.**

So `tags-undefined` is not an independent fault and never was. It is the second
half of a failed group, and the TAG FAULTS table's columns are not independent —
one refused group produces `group-5010`, `tags-undefined`, `no-queue` and
`tagging-failed`, four counts for one event.

**THE RASTERISE WAS A RED HERRING.** It first looked like draws after a rasterise
group 22% of the time against 93% for everything else. Splitting the arms by
batch count collapses it to **22% against 27% — nothing**. The scenarios that
rasterise simply draw small charts. This is the same confound `poolEveryDraw`
already warns about, met from the other direction, and it is written down because
it took a deliberate control to kill and would otherwise be re-found.

**THE FIX, and why it is not in this commit.** Refresh before grouping for every
groupable chart rather than only multi-batch ones. One line — `refreshShapes`
stops being gated on `spansBatches()`.

The original reason for the gate is in the code at the `needsRefresh` call site:
asking for a re-read a chart does not need was "a way to LOSE a group, not gain
one", because this host answers a re-read short or empty. **That objection is
much weaker now the settled retry exists** — a short or empty answer is asked
again after 1.5s before anything is decided.

**Stake the prediction before the round that tests it:** single-batch grouping
moves off 24% toward the multi-batch 100%, `tags-undefined` falls with it because
it has no other cause, and `the chart is actually visible` starts reporting a
chart that carries its config. **Run it as a pair**, and watch for the cost: an
extra sync per single-batch chart on the live insert path, which is the path a
user waits on.

**Owner's call, because it changes every insert.** The evidence is as strong as
this project gets, and the risk is a round-trip on the interactive path.


### THE RE-READ NEVER MATCHES OUR IDS — named precisely by rounds 068/069

**The clearest statement of the grouping defect this project has, and it took one
instrumented field to get.**

    withOwnId 7 of 7, 9 of 9   our handles are fine — the ids we hold are real
    listed    9, 10, 16, 17    the host named plenty of shapes...
              1, 1, 1, 1, 31   ...and in the next round, ONE
    matched   0                none of them ours, in either round

So the host lists a freshly drawn slide's shapes **under ids that are not the
ones it returned at creation**, and how many it lists at all is moody. Every one
of those traces carries `afterRetry: true`: it read that way *after* a
1.5-second settle, so waiting longer is not the fix.

**Consequence.** Grouping for these charts now rests entirely on the positional
fallback — "the last N on the slide" — which is a guess, is only legal when
nothing matched, and needs `listed >= drawn`. That is why the same build grouped
4 of 5 in one round and 1 of 5 in the next.

**The internal parallel says what kind of bug this is.** Scratch slides came back
as `4123571114#123571113` while the deck listed `287#62081387` — a freshly added
SLIDE reporting an id in a namespace the deck will not answer to. This is that,
one level down, on shapes.

#### The route worth reopening: a BINDING, measured in production

`bindings.add` takes the live Shape proxy inside the batch that created it — **no
id round trip and no collection read**, which are precisely the two things
failing above. Microsoft's own shape-binding documentation surfaced unprompted in
two of three web searches on this symptom.

This repo parked bindings as **unanswerable, not refuted**: `binding-names-shape-
later` was asked eight times in eight rounds and never once reached its own
question, because a shape probe needs the scratch slide resolved more times than
this host allows. **That is a fact about the probe, not about bindings** — and it
is the same shape as the six questions found 0-for-41.

The settled retry proved the alternative: ask in PRODUCTION. `repaired N` on a
verdict line settled in one evening what twelve rounds of probing could not.

**Plan it before building it** (the four points): the defect is that no durable
handle survives from creation to grouping; the seam is `groupAndTagAll`'s member
choice, beside `chooseGroupMembers`; what proves it is a `by: "binding"` reading
on the grouping trace against today's `ids`/`created`/positional; and it must not
touch the id path that already works for multi-batch charts, which group at 100%
on current builds and must stay that way.

**Owner-gated on cost, not on doubt**: bindings are PowerPointApi 1.8 against a
manifest pinned at 1.4, so it needs a `supports("1.8")` gate like grouping
already has, and it must degrade to today's behaviour on an older host.

### `same scale` PASSES — and the next work is where the failures moved to

**Rounds 070/071/072 on `01f3607`: 12 of 12 scenarios, `same scale` 8 of 8, three
times.** After 35 rounds of failing. Full numbers in the journal.

**The cause is not the code that was written for it.** The binding was added to
give the SETTLE a durable handle; the settle route has never executed (`bind 0/0`
in all three). What it did instead was upstream: `bindings.add` on the live proxy
in the drawing batch stabilises the shape's identity, so the pre-grouping re-read
stops listing shapes under ids we cannot match, and the config tag write lands.
`cfg5010` is 0 in all three rounds against 8 before.

**Reframe the item accordingly.** This is not "the settle got a handle", it is
**"binding a shape makes its identity resolvable"** — and if that is the real
mechanism, the settle-by-binding route is dead weight riding along with a
one-line side effect. Worth deciding deliberately rather than leaving both.

#### 1. ~~THE ORIGIN TAG IS NOW THE TOP FAILURE, and nothing can check it~~ — SHIPPED

Both halves are done and this entry outlived them by two days. The scenario is
`an update follows a moved chart` (`dragThenUpdate` in `src/taskpane/selftest.ts`,
in the battery's list): it moves a chart PROGRAMMATICALLY, confirms the move
landed, updates it, and asserts the redraw follows the delta rather than snapping
back to where the chart was inserted. No selection call anywhere in it.

And the defect it was written for is closed — see **THE ORIGIN TAG IS FIXED**
below: rounds 075/076, 13 of 13 scenarios twice, `origin5010` 0 both times, after
routing that one tag write through the chart's binding.

What is still owed is the part a scenario cannot do: a REAL mouse drag, by a
human. That is in the owner-only list at the end of this file and nowhere else.

#### 2. ~~NOTHING ENFORCES A ROUND RESULT~~ — SHIPPED

`scripts/rounds-gate.mjs` (`npm run rounds:gate`) is that gate: it reads the
archive, and exits 1 when a scenario that passed the previous three rounds of the
SAME profile fails in the newest one. `scenarioRegressions` in `scripts/triage.mjs`
is the decision, kept pure; a skipped scenario counts as "did not measure" rather
than as a fall, and exit 2 means the gate could not do its job at all.

Two limits worth knowing rather than re-proposing. It cannot be a CI check — CI
has no rounds — so it runs after archiving, and `cycle.mjs` acts on its exit
code. And it reads the pass/fail FLAG, so a scenario that goes on passing while
its numbers drift is not what it watches; `npm run rounds` is still where a
person reads the pooled counts.

This entry stood for a day after the gate landed, claiming the result had no
guard. Read the scripts before quoting a gap here.

#### 3. THE LOOP STILL NEEDS SIX MANUAL BROWSER STEPS A ROUND

Download the log, clean the deck, reopen the pane after a reload, verify the
toggles. This is throughput, but it is also correctness: hand-driven cleanup is
what took a deck to **0 slides** on 2026-08-16 and produced `slide 1 REFUSED`.
`cleanDeckScript` already exists — readiness REFUSES on a dirty deck instead of
using it.

#### 4. OWNER-ONLY, and neither has ever happened

- **Desktop PowerPoint.** Every round in this archive is the web host.
  `PUBLISHING.md` asks for desktop at least once.
- **A 4:3 deck**, same.
- **A real mouse drag.** The native-editability premise the product rests on has
  never been exercised by a human hand — and it is the thing item 1 can only
  partially substitute for.

### HALF THE CHARTS CANNOT FOLLOW A DRAG, and a passing scenario was hiding it

`npm run rounds` prints it now, under **CHARTS THAT CANNOT FOLLOW A DRAG**.

    rounds 073/074:  origin tag lost on 9 of 19 and 8 of 17 charts
    the scenario:    `an update follows a moved chart` PASSED in both

The scenario tests ONE chart and picks the least loaded. Roughly half the
population loses its origin tag in the same round, and every one of those would
snap back to where it was inserted rather than following the user's drag.

**The same shape this project keeps finding.** `does a rasterise poison the next
draw` counted only its own four draws while the round held 195; the fresh-slide
split sat unqueried for eleven rounds. **A scenario samples; a pooled count does
not**, and a green verdict over a sample is the most expensive kind of quiet.

**What it costs the user, stated precisely.** The chart is re-editable and its
config is intact — `cfg5010` is 0 across every post-binding round. What is lost
is only the drag delta. So this is a real defect and NOT a data-loss one, which
is why it can sit behind an item that is.

**Where the failure came from.** It moved here. Before the binding change the
config tag took the 5010s (8 in round 069) and the origin tag none; now the
config tag takes none and the origin tag takes 8-9 a round. The second write is
the one that fails, and it fails on the handle the first write just succeeded
through — which is worth a probe question of its own, since nothing currently
asks whether a shape accepts a SECOND tag in a later sync.

**No upstream twin.** Searched 2026-08-16: no office-js issue reports a second
`tags.add` on the same shape failing where the first succeeded. Two adjacent
ones are worth knowing and neither bites us — #6079 (tags forced uppercase, and
case-sensitive on web only; every tag here is already uppercase) and #3784
(shape tags lost on cut/paste on web). Recorded so the absence is not
rediscovered.

**No rate is printed, deliberately.** Only failures are traced — a successful
origin write says nothing — so there is no honest denominator, and inventing one
would be the kind of number this file has already had to correct once.

### THE ORIGIN TAG IS FIXED — rounds 075/076, and the binding is proven

Fixed and proven by rounds 075/076. See git.

### PUTTING 4:3 INTO THE NIGHTLY RUNS — four blockers — ALL CLEARED 2026-08-16

All four blockers cleared 2026-08-16. The 4:3 leg is `cyclePlan`'s third; see docs/ROUNDS.md.

### THE INSTRUMENTS WERE THE PROBLEM — rounds 082-087, 2026-08-16/17

Six consecutive 13/13 rounds, zero `UnexpectedError`. **Two product defects were
found across the whole stretch. Nine were in the reporting**, five of them
introduced the same day they were found.

That ratio is the finding. The product has been stable for six rounds; nearly
everything that went wrong was a number about the product rather than the product
itself. **A report that lies is worse than a crash**, because a crash stops the
night and a wrong number redirects a week.

#### The four ways one instrument was wrong

The orphan reading — "does an update strand the rest of an ungrouped chart" —
produced four confident wrong answers in two days:

| | what it reported | what was true |
| --- | --- | --- |
| mismatched units | 283 shapes stranded | its own line said `before: 3, after: 3` |
| an identity | 0, always | the two terms summed to zero by construction |
| one stale read | 92 shapes grew | the deck showed one grouped chart per slide |
| two agreeing stale reads | growth 23, `settled` | both reads fell inside one lag |

**Every one was caught by `deck.inventory`** — a second source already in the
round file, taken at end of round, long after any host lag. It has never been
wrong. It is now cross-checked on every reading and disagreements are COUNTED,
because an instrument's own error rate belongs in its own report.

**The rule that comes out of this**: never quote a number until a second,
independent measurement in the same data agrees. Not a re-run of the same
instrument — two reads of one lagging source agree perfectly.

#### What is left open here

- **The stranding question is at three observations, all zero growth.** Three is
  not five. It needs rounds that fail to group something; 082-086 grouped
  essentially everything and could not answer either way.
- **`reading back an ungrouped chart's shape ids`** is still the one live
  `GetItem(id)` refusal site, in 56 of 63 rounds. What it costs is now measured
  rather than assumed: it denies a chart its parts list, and only a chart in that
  state can strand anything.

#### Two product defects, for the record

- **A blank slide of ours shipped in the finished deck.** One `slides.add()` can
  land two; the branch reporting un-nameable landings counted the event, not the
  slides. Nine archived rounds affected; five of the last thirteen began on a
  deck already dirty by one slide, invisible because the scenarios measure
  GROWTH rather than absolute size.
- **An update died on a refused id.** The resolve sync was unguarded, so one
  by-id lookup this host would not honour took the whole update down. Guarded,
  then re-asked through a collection read. **Still unexercised on a real host** —
  it has not fired once in six rounds.

#### Mechanics that earned themselves

- Every `playwright-cli` call is bounded; the round's own 30-minute deadline
  could never fire because it is checked between calls.
- An unexpected throw is a `threw` reason rather than a dead process.
- The cycle stops when a leg finished but archived nothing — the state that
  otherwise lets the next leg overwrite the only copy of the evidence.
- `sweepDeck` reads its own answer instead of always claiming success.
- The gate's novelty bucket uses a five-round window and names the build a
  signature FIRST appeared in. It had spent fifteen rounds blaming nine innocent
  builds for one 064-era signature.

#### Owner-only, unchanged

Sign-in, a real mouse drag, and desktop PowerPoint. Re-sideloading an add-in into
a document is drivable and now automated, but a **browser death does not always
lose the sideload** — once it did, once it did not — so that path is written and
tested and has not yet met the state it exists for.

The browser deaths themselves are **connected standby**, not crashes: Windows
Event 507 four times in a morning, `ERR_NETWORK_IO_SUSPENDED` in the console
tail. Power settings on AC now prevent it; on battery the display still sleeps at
four minutes, which on a Modern Standby machine is what triggers it.

### Retire the positional group-member mapping with `Shape.creationId` — CLOSED 2026-08-29, REFUSED BY THE HOST

**The remedy does not exist here.** This host reports PowerPointApi **1.10** in
`requirementSets` and does not populate `Shape.creationId`. All three probe
questions say so, unanimously, over the 25 rounds to 2026-08-28:

    creationid-on-fresh-shape       absent          25 of 25
    creationid-survives-a-sync      no-creation-id  25 of 25
    creationid-survives-grouping    no-creation-id  25 of 25

So the blocker recorded below — "the paths that reach this code are gated at
1.8, so it needs a 1.10 branch" — was never the blocker. A 1.10 branch would
compile, ship, and find nothing to read. The positional mapping stays, and the
test asserting the TARGET of a node-0 write is its permanent guard rather than
a stopgap.

Worth noting how this stayed open: the three answers had been on every sheet
for months, and the questions sat in `PENDING_QUESTIONS` marked "the committed
sheet predates this question" because the fixture beside the gate was a
2026-08-12 capture. The probe was even spending a scarce scratch slide
re-asking each of them. The analysis below is kept because the reasoning about
what the mapping rests on is still true.

The in-place update maps scene nodes to a grouped chart's shapes by POSITION:
member 0 is the anchor, the rest line up in drawing order. Round 145 showed what
a mistake in that mapping costs — node 0's properties were written onto the
group itself, and only the host caught it.

**The order is not documented.** Checked 2026-08-21: the `PowerPoint.ShapeGroup`
reference says `shapes` is "the collection of Shape objects in the group" and
nothing more. The assumption rests on the classic Office object model, where a
shape's index in a Shapes collection is its z-order position and the anchor —
drawn first — is at the back. Consistent with everything observed, and inferred
rather than promised.

**`Shape.creationId` (PowerPointApi 1.10) is a durable per-shape identifier.**
Recording creation ids at draw time would make the node-to-shape mapping
explicit and ordering irrelevant, which is the only real fix for the assumption
above. Blocked on the requirement set: the paths that reach this code are gated
at 1.8, so it needs a 1.10 branch with the positional mapping kept as fallback.

Until then the guard is the test that asserts the TARGET of a node-0 write
(`writes node 0 to the group's ANCHOR, not to the group itself`) — a count check
cannot catch a mapping error that preserves the count.

### Two questions production answered before the probe sheet could — RETIRED 2026-08-21

#### What actually answered them

**`tag-on-group-survives` — "Does a tag written on a GROUP read back?" YES, from
production, over 149 rounds.** It is the mechanism the whole product runs on:
every chart is discovered by reading `POWERCHART_CONFIG` off its group, and
since #645 the in-place update reads `POWERCHART_SCENE` off the same group
thirteen times a round. Rounds 146, 147 and 149 each completed three in-place
updates, which is not possible unless a tag on a group reads back. A probe
cannot beat that evidence; it can only agree with it more slowly.

**`grouped-child-by-id-from-slide` — "Can a shape INSIDE a group still be
resolved by id off the slide?" MOOT.** It was described as the question that
decides whether the in-place update can ever work here. It is not, because the
update no longer takes that route: #643 reads a grouped chart's members from
`shape.group.shapes` (a `ShapeScopedCollection` at PowerPointApi 1.8), and #646
writes through those member proxies. The by-id-off-the-slide path is not used
and does not need an answer.

Note what that means for the older entry above, "The probe has been blind on
GROUPS for the whole archive": the blindness was real, and the product routed
around it without the probe ever seeing. **A question worth asking is not the
same as a question worth waiting for.**

#### The branch that tried the other remedy

`claude/ask-the-decisive-probes-early` (`f09041c`, 2026-08-13) proposed moving
both to positions 5 and 6. It is pushed and preserved, and it was NOT merged,
for three independent reasons: it ships a deliberately failing test ("KNOWN RED,
AND THE REASON IS THE POINT"); 23 commits have touched `host-probe.ts` since,
and it no longer rebases cleanly; and its premise — that these two questions
gate the in-place update — is now false.

Its diagnosis was still worth keeping, and is recorded above: these questions
never get put. The conclusion inverted once production answered them.

#### What is left open

The `no-scratch-*` split in the older entry still stands for the four remaining
starved questions. Retiring two of the six is not a fix for the harness; it
removes two questions that could never pay, and gives their slide back to the
ones that might.

### The slide-size ladder's first rung hangs about twice as often as it answers — RE-MEASURED, AND IT IS WORSE THAN THAT

**RE-MEASURED 2026-08-29 over 270 rounds, and the headline below is wrong in
the direction that matters. It is not "about twice as often as it answers".
Every round that reads a slide size times out on rung 1 first:**

    rounds that hang on rung 1, then record pageSetup        82
    rounds that hang on rung 1, then record exportedSlide    75
    rounds that record a size WITHOUT hanging on rung 1       0
    rounds that hang and never get a size                     0

**157 of 157, always the full 4000ms — one hang per round, never two.** The four
seconds is universal, not occasional. And rung 1 is not the loser it reads as:
it supplies the final answer in 82 of those 157, more often than rung 2 does.

The 2026-08-23 figures below (22 · 22 · 44) were the same questions asked with
a pattern that matched a fraction of the archive. Both counts were low and the
ratio between them survived, which is what made them look reliable.

**So neither obvious action follows.** "Lower the bound" now saves four seconds
in EVERY round rather than a third of them — a much better prize — but rung 1
is also the rung that wins half the time, and a lower bound might buy the time
by giving those 82 answers away to an export whose cost is still unmeasured.

**What decides it is one number nobody has: how long a SUCCESSFUL rung-1 read
takes.** Under 4000ms and the timeout is being paid by a different call than the
one that answers, so the bound comes down for free. Over it, and the bound is
exactly what makes those 82 answers possible.

`slideSize()` now traces `ms` on all four rungs (2026-08-29), so the next rounds
answer it. Four tests, four mutants — three killed by the runner and the fourth
confirmed by hand after the runner failed to apply it.

**ANSWERED THE SAME NIGHT, rounds 297-299 — and it is the cheap answer.**

    297   STALL afterMs=4000   READ pageSetup ms=270
    298   STALL afterMs=4000   READ pageSetup ms=246
    299   STALL afterMs=4000   READ pageSetup ms=263

A successful rung-1 read costs **246-270ms**, against a stall of a flat 4000.
`ms` is counted from `slideSize()`'s own entry, so a read that traced 270 cannot
be the call that waited 4000 — **the stall and the answer are different calls**,
and the second one succeeds cheaply without the first's bound helping it at all.
The 4000ms is not buying those 82 answers. It is being paid by a first call the
host is not yet ready for, and a later call gets the same answer in a quarter of
a second.

So the bound can come down, and by a lot: 1000ms leaves a **4x margin** over
every successful read yet observed and returns ~3 seconds per round. What it
must not do is come down to where the export on rung 2 starts winning races it
used to lose, since nothing has measured what THAT costs — the reason this was
not simply lowered in the first place, and still the reason to measure rung 2
before choosing the number.

**Three samples.** Enough to answer the question that was asked (is a successful
read cheaper than the bound? yes, by 15x) and not enough to pick the new bound
from. Every round from 47db58a on carries the field, so the sample grows on its
own; re-read before choosing.

**AND ROUND 300 CONTRADICTS THE "157 OF 157" ABOVE, which is the most useful
thing here.** It carries no stall at all and a rung-1 read of **138ms** — the
first round in the archive to read a slide size without paying the four seconds
first.

What was different is not the code, which is the same build as 298 and 299. It
is that round 300 ran immediately after a crash recovery: PowerPoint's own
Refresh, a full tab reload, and a pane reopened seconds before. **So the stall
looks like a property of a host that has been sitting, not of the call.** A
freshly loaded document answers `pageSetup` first time, in 138ms.

That is one round and it is offered as one round. But it points the remedy
somewhere different from "lower the bound": if the four seconds is the first
call meeting a cold host, then a shorter bound simply reaches rung 2 sooner on
exactly the runs that are already unwell, and the thing worth measuring is
whether a cheap warm-up call before the ladder removes it entirely.

**ROUND 301 SUPPLIES THE MISSING NUMBER, and it unblocks the decision.** It is
the first round in the archive to show both halves in ONE call:

    301   STALL afterMs=4000   READ source=exportedSlide ms=4280

`ms` counts from `slideSize()`'s entry and the bound is 4000, so **rung 2's
export cost about 280ms** — the same as a successful rung-1 read (246-270ms).
That was the one thing this entry said it was blocked on: "nothing has measured
what the export costs".

It is measured now, and it is cheap. **So the bound can come down and the
fallback is nearly free.** At a 500ms bound a cold host would spend 500 on rung
1 and ~280 on rung 2 — under 800ms against today's 4,280, a saving of about
three and a half seconds, with the same answer at the end of it. The worry that
made this wait ("a lower bound might trade those 82 answers for an export whose
cost nobody knows") is retired: the export costs what the read costs.

301 also supports the cold-host reading. Its tab had been reloaded 35 minutes
earlier and then sat idle, and the stall came back — where round 300, running
seconds after a reload, had none.

**Round 302 then narrowed it, which is more useful than a third confirmation.**
Its PANE was closed and reopened three minutes before the round, and it stalled
anyway (`READ pageSetup ms=504`, after the usual 4000). So "recently touched"
is not the variable. What round 300 had that 301 and 302 did not is a full TAB
RELOAD after PowerPoint's own crash-Refresh — a new editing session, not a new
pane. Reopening the pane does not buy it.

Three samples, one stall-free, and the thing that distinguishes it is the
heaviest possible reset. That is a narrower and more testable claim than the one
it replaces, and it is still three samples.

**What remains before changing the number:** decide between the two remedies
rather than doing both blindly. Lowering the bound saves ~3.5s on every stalled
run; a warm-up call might remove the stall altogether and save the rung-2 hop as
well. They are not exclusive, and neither is now blocked on a measurement.

One defect found writing those tests, worth recording because it is this
project's recurring one: the rung-3 assertion used `traceLog().entries.find`,
and that log accumulates across the file, so it read a line an EARLIER test had
written and passed against a deliberately broken rung. Measuring something
adjacent to what you meant to measure — the same shape as the gauge total
measured unbold and drawn bold. It searches from a captured index now.

**Original entry, 2026-08-23, kept because the reasoning about WHY not to act
blindly is still right:**

`slideSize()` rung 1 reads `presentation.pageSetup` directly, bounded by
`SELECTION_TIMEOUT_MS` (4000ms). Across the archive:

    slide size read, by source:  pageSetup 22 · exportedSlide 22
    gave up waiting on "reading the slide size":  44 rounds

So rung 1 **is not a dead rung** — it answers 22 times — and it hangs roughly
twice as often, first seen in round 117 and in 26 of the last 30 rounds. Every
hang costs a flat four seconds before rung 2 exports a slide and answers.

The trace already names the shape precisely, and it is worth quoting because it
rules out "the host was busy":

    "THIS CALL ALONE is stuck — the host answered something else 3845ms AFTER
     this call"

**Why this is recorded rather than fixed.** The obvious change is to lower the
bound, and the obvious change is not obviously right:

- rung 2 exports a slide, and nothing has measured what THAT costs. Trading a
  4-second wait for an unmeasured export is not known to be a saving.
- 4000ms is `SELECTION_TIMEOUT_MS`, shared with the selection path. Lowering it
  here means either a separate constant or a change to a budget something else
  depends on — see the 0.5→0.8 `UPDATE_SHARE_LIMIT` episode for what happens
  when a shared budget is raised on a good measurement of the wrong thing.
- the cost is ~4s in a ~490s round, under one percent. It is worth knowing and
  it is not worth a blind change.

**What would settle it:** time rung 2 on the rounds where rung 1 hangs. Both
already trace; neither carries a duration. One field on each, and the next round
answers whether a lower bound saves anything at all.

Found by the archive sweep of 2026-08-23. The sweep reported it as "every round
eats a 4-second timeout — 38 of 38, never zero", which is the recent era and not
the archive: it is 44 of 147 overall, and universal only since round 117.

### The four dependabot advisories — ALL FOUR NOW FIXED, 2026-08-27

All four fixed 2026-08-27. See `package.json`'s `_overrides_why` and git.

### The original analysis — 2026-08-26

Every push prints "GitHub found 3 vulnerabilities on the default branch (2 high,
1 moderate)". Checked once so it does not have to be re-checked every push.

    image-size  *        HIGH x2   ICNS and JXL/HEIF parsers loop forever on a
                                   malformed image (GHSA-w3rx-r6r6-pgpr,
                                   GHSA-5p2g-fcmc-qvqq)
    qs  6.11.1-6.15.1    MODERATE  qs.stringify crashes on null entries in
                                   comma-format arrays (GHSA-q8mj-m7cp-5q26)

**`qs` is dev-only, and is now FIXED** — `@stryker-mutator/core` ->
`typed-rest-client` -> `qs`, never shipped, never runs outside mutation testing.

Two things I wrote here first were wrong, and both were measured rather than
assumed the second time:

`npm audit fix` does NOT clear it. npm reports `fixAvailable: true` and then
changes nothing — zero lockfile delta, four vulnerabilities before and after —
because it cannot move a transitive that `typed-rest-client` pins. A fix that is
"available" and inert is worse than one that is absent, because the report reads
as actionable.

What does clear it is an `overrides` pin to `qs@^6.15.3`. The advisory range is
`6.11.1 - 6.15.1` and 6.15.3 is outside it, so this is a PATCH bump on the same
minor rather than a version gamble. `npm audit` goes from 4 vulnerabilities (2
moderate, 2 high) to 2 high. The `_overrides_why` key in package.json says when
to remove it.

**`image-size` is NOT REACHABLE.** It arrives under `pptxgenjs`, which IS a
runtime dependency — `src/render/pptx-deck.ts:74` imports it dynamically — so the
tree alone makes it look shipped and live. It is not: pptxgenjs calls
`image-size` to size images handed to `addImage`, and **SSF Charts never calls
`addImage`**. There is no image path in `pptx-deck.ts`, no PNG or JPEG or data
URI anywhere in it. The deck is shapes and text. Both parsers named in the
advisories need an image to parse and are handed none.

**There is also nothing to upgrade TO.** Both advisories give their vulnerable
range as `<=2.0.2`, and 2.0.2 is the latest published `image-size`. Every version
that exists is flagged, so an `overrides` pin — which is what fixed `qs` above —
has no target here.

And npm's proposed remedy is not a bump but a THREE-MAJOR DOWNGRADE: `fixAvailable`
names `pptxgenjs@1.1.5` against the 4.0.1 in use. `npm audit fix --force` would
gut the deck writer to patch a code path this product never executes. Not taken,
and the direction is worth stating — I first recorded this as "a major bump",
which sounds like progress. It is the opposite.

**What would change this.** Any feature that puts a bitmap into a generated deck:
a logo, a screenshot, a rasterised chart fallback, an image placeholder. If
`addImage` ever appears in this repo, `image-size` becomes live and this entry is
wrong — upgrade pptxgenjs before shipping that feature, not after.

### The arm that would settle the first-chart penalty has been run ONCE — 2026-08-26

The largest single cost in a round is the first chart of the deck-wide rescale:
~44.1s against ~17.6s for the same chart size, same changed count, and the same
slide occupancy (n=26 each — see `WHAT THE SLIDE ACTUALLY HELD` in triage). Every
one of the three WRITE syncs carries the penalty at 2.1-2.35x while the tag sync
carries only 1.16x, and `our-idle-is-negligible` puts our own waiting at 1ms
median. So the time is going to the host, during writes, on the first chart.

The leading explanation is contention with the DECK SCAN that runs immediately
before the rescale. `scanSettleMs` exists precisely to test it — a pause inserted
between the scan and the first update, "the only gap where a pause can tell the
two apart", traced at zero as well so every archived round records its arm.

**It has one nonzero sample in the whole archive.**

    settleMs = 0        58 round(s)
    settleMs = 10000     2 round(s)   -- one with a usable first-chart timing
    absent              60 round(s)   -- predates the trace

n=1 against a first-chart distribution spanning 13.3s to 58.3s says nothing at
all, and the single sample (38552ms) sits in the middle of it.

**Why it is never run: it cannot be set from the driver.** `scanSettleMs()` reads
the pane's `localStorage`, so arming it means typing

    localStorage.setItem("powerchart-scan-settle-ms", "3000")

into the pane console by hand, before a round, every time. `scripts/round.mjs`
has no way to do it — it drives entirely through playwright-cli refs and never
evaluates JavaScript. So the arm is available in principle and unreachable in
practice, which is exactly the shape of a control that quietly never runs.

**What would fix it.** `playwright-cli` DOES have `eval <func> [target]`. The
pane is a cross-origin iframe, so a page-level eval cannot touch its
`localStorage` — but Playwright evaluates in the frame that OWNS the target
element, so passing a ref from inside the pane should reach it:

    eval "() => localStorage.setItem('powerchart-scan-settle-ms','3000')" <pane-ref>

That is untested. It needs a `--scan-settle <ms>` flag on the driver, defaulting
to off so no existing round changes what it measures, and then enough rounds in
each arm to clear a noise floor the archive puts at IQR 14%.

**THE DECK CANNOT SEPARATE POSITION FROM OCCUPANCY, and this is the concrete
blocker.** The rescale's charts 1, 2 and 3 all sit on the same busy slide, but 1
is a 24-node chart and 2 and 3 are 16-node. Charts 4-8 are 24-node and sit on
1-shape slides. So the archive offers `24/18 on a 3-shape slide, FIRST` and
`24/18 on a 1-shape slide, LATER` — and no `24/18 on a 3-shape slide, LATER` cell
at all. Position and occupancy move together by construction, in every round.

One round makes the point without any pooling. From the crashed run of
2026-08-26 (`crashes/2026-08-26T19-56-29-crashed-run.json`), one run, one deck:

    chart 1/8   24 nodes, 18 changed   38174ms   syncMs [12635,11803,13078,...]
    chart 4/8   24 nodes, 18 changed   17999ms   syncMs [ 5975, 5692, 5736,...]

Same size, same changed count, same run — 2.1x, and EVERY write sync carries it
rather than one slow step. Charts 2 and 3, on the SAME slide as chart 1, cost
13.6s and 13.3s, so the slide is not what makes chart 1 expensive. What the deck
cannot tell us is whether a 24-node chart LATER in the run on that same busy
slide would also be cheap. Building one such slide would answer it in one round.

**THE SCAN'S OWN DURATION DOES NOT PREDICT THE PENALTY — checked 2026-08-26.**
`scanned the deck for charts` carries an `ms` that no tool reads. Split 69 rounds
at the median scan time and compare each round's first chart against ITS OWN
later charts, so a generally slow host cannot manufacture the result:

    short scans   n=34   median scan 1139ms   median first/later ratio 2.22
    long  scans   n=35   median scan 1383ms   median first/later ratio 2.15

No effect, and what there is runs the wrong way.

**This does NOT clear contention**, and the limit is worth stating rather than
overreading a null. The natural range is narrow — 1.1s against 1.4s, 21% apart —
and a weak contrast cannot detect a real effect. A scan may also leave the host
busy for a period unrelated to how long the scan itself took. What it does rule
out is a simple dose-response, where a longer scan costs a slower first chart.

That is an argument FOR running the settle arm rather than against it: a 3-10s
pause is a far stronger manipulation than the 250ms of natural variation above,
and it is the only one that separates the two.

**What NOT to do.** Do not make the pane counterbalance this on its own. Half of
every future round would then carry a pause, which changes what a round measures
for every claim already resting on the archive. The flag is opt-in for that
reason.

### Text drawn over text — 2,148 to 317, and what is left

**The measurement.** Every kind under 24 option variants and 10 data shapes, at
eight frame sizes, two fonts and both orientations — about 24,000 charts — with
overlap measured by the frame gate's own ink rule. It ran for the first time on
2026-08-27 at **2,148 overlapping pairs** and stands at **317**.

**AND THE GATE HAD THE SAME BLIND SPOT IT WAS BUILT TO CLOSE — 2026-08-29.**

`frame-fit.test.ts` swept kinds, frames and fonts and missed the option and
data-shape variants; that is why this family lived nine days behind a green
gate. `overlap-budget.test.ts` was written to close that hole, and it swept
every option and every data shape **and never one of each** — the two tables
were concatenated, not crossed. A chart with a secondary axis AND ten series
was not among its 24,000. Neither was a footnote on twenty-four categories.

Crossing a slice of them (six layout-changing options against four
label-stressing data shapes) took the count from **537 to 4,010** and produced
**twenty-two shapes this engine has never been seen to draw**. Nothing about the
engine changed that day. The number moved because the sweep did, and a figure
from before it is not comparable with one after — the same trap as the
"75 pairs" of 2026-08-19, sprung a second time in ten days.

    2,683   `p#-*`, small-multiples panels — ALL of it new
    1,058   `value-axis-title` — the open decision below, unchanged in kind
      269   everything else, including three new shapes

**THE FIND — AND THE FIRST ANSWER WAS WRONG, so read the correction first.**
This entry said, from the count alone, that "small multiples do not thin their
labels when the data gets dense". Tested on 2026-08-29 and refuted. The category
axis thins perfectly well — `catFs` shrinks until every name fits its slot and
`clipToWidth` clips what no floor can fit — and **zero** of the 910 overlaps are
between two names in the SAME panel:

    panel category-name overlaps:  WITHIN a panel 0 · ACROSS panels 910
    by frame:                      300x60  910   (every other frame: none)

All of it is one frame, and these labels are not merely touching. On a 300x60
chart with ten series in two columns, `p0-category-0` is drawn at **y = 307 on a
chart 60 points tall**, on top of the next row's names at 316.

**A DEFENSIVE CLAMP TURNS AN IMPOSSIBLE GRID INTO A PLAUSIBLE WRONG ANSWER**,
and that is the whole finding. It took two wrong explanations to get here, both
written from arithmetic rather than measurement; the third is measured end to
end.

Ten series in two columns is five rows, and

    panelH = (cfg.height − titleH − footH − gap × (rows − 1)) / rows
           = (60 − 22 − 0 − 40) / 5
           = −0.4

`buildMultiples` builds the grid anyway and hands each panel `height: −0.4`.
Then `normalizeConfig` runs, and `clampDim` says:

    if (typeof v !== "number" || !Number.isFinite(v) || v <= 0) return fallback;

**−0.4 is `<= 0`, so the panel is given `DEFAULT_SIZE.height`, 300.** Every
panel is therefore laid out as a full 300-point chart, and the composition
places those at a 9.6-point pitch — ten 300-point charts stacked 9.6 points
apart inside a box 60 points tall. 129 of the scene's 139 text nodes end up
below the bottom of the chart.

Measured rather than reasoned, and this is the number that settles it:

    a standalone 155x300 chart puts category-0 at   y = 285.0
    the panel puts category-0 at, in its own frame  y = 285.0

Identical. The panel is not a squeezed panel; it is a full-size chart wearing a
panel's offset.

`clampDim` is not at fault and should not be changed. It exists because
`width: NaN` arrives from pasted configs, saved templates and shape tags written
in other decks, and repairing that beats a stack trace in a headless render. The
fault is that an INTERNAL arithmetic error reached it. A repair meant for
malformed input from outside absorbed a computation that should never have been
attempted, and converted "this grid cannot exist" into "here are ten charts on
top of each other" — a loud failure made quiet.

So the fix belongs in `buildMultiples`: check that a panel has a usable height
BEFORE building the grid, and if it has not, do what this engine does everywhere
else with a reservation it cannot pay for — decline it, and render the chart
whole. This is the house pattern named in the memory *a fit and a clamp fight,
and the clamp wins*, in its most expensive form yet: here the clamp does not
merely undo the fit, it manufactures a chart nobody asked for.

The counts below stand as measured; only the explanation changed. Crossed with
twenty-four categories or ten series:

    910  p#-category# / p#-category#
    408  p#-title / p#-title
    378  p#-value-axis / p#-value-axis
    292  p#-label# / p#-label#
    261  p#-label## / p#-label##
    198  p#-total# / p#-total#

**Every family traced, 2026-08-29 — it is one defect, not six.** The
inference above ("the same grid seen through different labels") was checked
rather than left standing. Over all twenty-three `p#-` shapes:

    shape                          within  across  off-chart  frames
    p#-category# / p#-category#         0     910        910  300x60
    p#-value-axis / p#-value-axis       0     420        420  300x60, 80x60
    p#-title / p#-title                 0     408        312  300x60, 80x60
    p#-label# / p#-label#               0     252        252  300x60, 80x60
    p#-total# / p#-total#               0     140        140  80x60, 300x60
    …and eighteen more, all the same shape

Three things hold across every one of them. **`within` is zero** — no two
labels in the same panel ever collide. **`off-chart` equals the overlap count**,
family by family: essentially every pair involves a label drawn outside the
chart's own bounds. And **only two frames appear at all**, 300x60 and 80x60,
the two shortest in the sweep. That is the signature of the grid, seen twenty-
three ways.

(Those counts are from a `multiples 2 columns` × four-data-shape probe, so they
do not match the ratchet's totals line for line — it crosses six options. The
structure is what is being claimed, not the arithmetic.)

**~~One shape is NOT the grid, and it is the interesting leftover.~~ WRONG, and
the ratchet said so within the hour.** This claimed `p#-total# / p#-cagr-label`
was "a real label collision at a size somebody would present" — 9 of its 26
within-panel, and the only family appearing at 480x300 and 960x540 — and that
it "will still be there after the grid is fixed".

It is not there. After the fix the budget table holds **no `p#-` entry at all**,
and the ratchet's "draws no shape it has never drawn before" test passes, which
it could not if any `p#-` pair still occurred. The grid took the whole family,
that one included.

**Why the probe and the ratchet disagreed, because that is the reusable part.**
They were not measuring the same charts. The probe built its `10 series` shape
from scratch — ten fresh series with generated values — where `DATA_SHAPES`
spreads the sample's existing series and keeps their colours and roles. Same
name, different data, different decorations, different overlaps. A diagnostic
written beside a gate has to reuse the gate's own variant table or its numbers
answer a different question, and the two agreeing everywhere else is exactly
what made this one look trustworthy.

**The lesson, and it is the same one twice in a night.** The first explanation
here was written from a count and read plausibly — this engine really does thin
labels everywhere else, so "not inside a panel" fitted. It took one diagnostic,
which cost minutes, to find that the true answer was in a different file
entirely and an order of magnitude worse than the story. A count tells you where
to look and never why.

**FIXED the same night, once it was diagnosed rather than guessed.** The whole
2,683 goes to zero — every one of the twenty-three `p#-` shapes — and the
sweep's total falls from 4,010 to **1,327**. The change is one line in
`buildMultiples`:

    if (!(panelW > 0) || !(panelH > 0)) return null;

`> 0` and no more, deliberately. Every positive size passes `clampDim`
untouched, so this cannot alter a chart that renders today; it refuses only the
case that was being silently rewritten. Declining the grid renders the chart
whole, which is what this engine does with every other reservation it cannot
pay for — and the test asserts BOTH halves, because declining must not cost the
user their chart.

Two mutants, both dead, and the second is the point: removing the guard fails
the test, and so does weakening `> 0` to `>= 0`. That second one survived the
first version of the test, because the case it used has `panelH = −0.4` and
fails both operators. `clampDim`'s own test is `v <= 0`, so **exactly zero** is
rewritten too — there is now a case pinning it (four series, two columns, height
10, no title → `panelH = (10 − 10) / 2`).

This was written up as "not fixed tonight, it wants the sweep, the ratchet and a
round". It got the sweep and the ratchet; the round is what it still owes, and
it is a RENDERING change, so it must not be called done until one has run.

**Still not crossed**, so nobody has to re-derive it: every option outside
`CROSS_OPTIONS` against every data shape, and any product of three or more.
Widen that list before calling a family closed.

**It is a gate now, not a script somebody remembers.**
`test/overlap-budget.test.ts` runs the sweep in eight seconds on every build and
holds a budget PER SHAPE. Three of its four tests catch a regression; the fourth
catches a budget left ABOVE the real figure, so improving the engine is supposed
to fail the file and the numbers get edited down. A shape absent from its table
is a regression by definition — it is text this engine has never been seen to
draw over text. **The per-shape table there is the live list; this section does
not duplicate it.**

The fixes themselves are in git, with a commit each explaining what the defect
was and what the fix cost. What is worth keeping here is what remains, and what
the exercise taught.

---

**STILL OPEN, AND IT IS THE OWNER'S CALL: `valueAxisTitle`, 916 of the 1,014** —
**smaller since the clip shipped on 2026-08-29 (it was 1,208 of 1,327), and the
same decision.** The fifth remedy, at the end of this entry, is the one that
worked: the option is a SHORT UNIT and the sweep was testing it with a sentence.

The count is bigger than the "287 of 317" this said before and the problem is
the same size — the sweep widened underneath it, and the panels that briefly
outnumbered it have been fixed. It is now the whole of what remains bar a
119-pair tail, which makes it the only overlap decision left worth the name.

**It said 1,058 and a 269-pair tail for half a day, and both were arithmetic
rather than measurement.** Those were the 4,010 sweep's split, carried forward
by subtracting the 2,683 that the small-multiples fix removed. Re-summing the
table gives 1,208 and 119 — the family GREW by about 150 while the tail shrank
by the same, because a chart that used to render as an impossible grid now
renders as an ordinary one, and an ordinary chart with a unit label contributes
`value-axis-title` overlaps of its own. **A subtraction is not a
re-measurement**, and this file has now made that mistake with two different
figures in two days.

Four collision shapes, and they are not one problem. Two were this label's own
`y` clamped into the title and onto the topmost tick — fixed. The other two are
the WIDTH, and they are the question:

    value-axis-title / legend#    ~100
    value-axis-title / total#      ~43

Its width is `Math.max(frame.x - 4, textWidth(…))` — a floor that RAISES a width
— so a long unit grows right across the totals row and the legend. Fitting it to
the axis gutter was tried in 2026-08-19 and reverted: `frame.x` is the value
axis's own column and a chart drawn WITHOUT a value axis has none, so the fit
dropped the unit from ordinary charts.

**Three remedies were implemented and measured on 2026-08-28. All three were
reverted, and the numbers are why:**

    drop on box overlap     100 of 352 units survive — seven charts in ten lose
                            the unit, most to collisions only the empty part of
                            an over-wide box was having
    drop on ink overlap     249 of 352 — both clamp shapes go to zero, but a
                            clustered chart at 480x300 in 18pt loses its unit,
                            and that is a size people present at
    move below the title    keeps more, and lands it on the tick numbers — the
                            collision it was moved away from

**Why it cannot be settled the way everything else was.** Every other remedy in
this engine drops or shrinks a label the layout generated. This one is TEXT THE
AUTHOR TYPED. Dropping a tick number costs a reading the gridline still carries;
dropping `valueAxisTitle` deletes something a person wrote and cannot see is
gone. The band above the plot holds the chart title, the unit and the topmost
tick number, and there is not always room for three.

**A FOURTH WAY WAS TRIED ON 2026-08-29 AND DOES NOT WORK, which is worth
knowing because it is the one this engine would reach for by reflex.**

The unit is the only label here drawn at a FIXED size and fitted to nothing —
every other label shrinks to its room and drops past `MIN_LABEL_FS`. Shrinking
is not dropping, so it does not raise the author's-text objection below, and
none of the three remedies above tried it. Measured over every kind, frame, font
and orientation, shrinking the unit until it fits the chart's own width:

    charts where the unit overlaps something    174  ->  153

    what it collides with       before   after
      legend#                      100     100
      total#                        68      39
      title                         41      53
      value-axis                    33      18

It TRADES. The totals row and the tick numbers get most of their space back and
the title loses some of its own, because a smaller unit sits lower in its box
and moves up into the title's ink. Twenty-one charts of 352 is not worth a
change to how every unit is drawn, and `legend#` — the largest partner — does
not move at all, because the legend is nowhere near the unit's width.

So the fit is not the missing piece, and the entry's judgement stands: there is
no bound available here that is not a coupling to something drawn later. It is a
decision about WHERE A UNIT BELONGS, and the three ways out below are still the
three ways out.

**A FIFTH WAY SHIPPED ON 2026-08-29, and it is the one that worked: clip the
unit, because the option is a unit and not a subtitle.**

The four remedies above all argue about WHERE to put a long unit. None of them
asked whether a long unit is a thing this option supports. It is not.
`ChartConfig.valueAxisTitle` is documented as "units label shown at the top of
the value axis (e.g. `€m`)", the pane's input is 56px wide with the placeholder
`e.g. €m`, and the two uses in the 123-chart shipped deck are `€m` and
`$m (log)`. The sweep tests it with **"Revenue in millions of euro, restated"** —
twenty-seven characters. Most of the largest overlap family in this engine was
the cost of a string the feature does not offer.

So: `clipToWidth` at 40% of the chart's width. Measured on the ratchet, not on a
scratch harness:

    total across the sweep      1,327  ->  1,014      and no new shapes
    value-axis-title family     1,208  ->    916
    the tail                      119  ->     98      (unchanged in kind)

It keeps author text where the three 2026-08-28 remedies deleted it — `€m` is
untouched, a sentence is truncated with the start kept — and truncating what
does not fit is what this engine already does to gantt and category names.

**AND THE OTHER HALF OF THE SAME RECOMMENDATION WAS REFUSED BY THE
MEASUREMENT.** Flooring the unit's `y` at the title's ink — `title /
value-axis-title` is 205 and length-INDEPENDENT, which is exactly the clamp
signature this file has recorded five times — was written and swept in the same
hour:

    clip alone                1,014, no new shapes
    clip + floor at the ink   1,156, and `value-axis-title / category#` at 310,
                              a family that did not exist before

The floor buys the title collisions by moving the unit into the category names
on short charts. **A clamp moves a label whether or not the destination is
free** — which the CAGR caption's own note already said, about the same move.
Pinned by `test/value-axis-title.test.ts`, which says re-measure before
"fixing" it.

**What is still open, and it is unchanged in kind:** where a unit belongs when
the band above the plot cannot hold the title, the unit and the topmost tick all
three. Three ways out, none of them engineering work:

- Put the unit at the END of the axis rather than above it — Excel's rotated
  axis title, without the rotation.
- Fold it into the tick numbers themselves: `€m` on the top tick only.
- Accept it as the one label allowed to displace the plot, which is the gutter
  idea the 2026-08-19 attempt reverted.

---

**DIAGNOSED AND DELIBERATELY NOT PATCHED: the dual-axis gutter.**

**MEASURED PROPERLY 2026-08-29, and the answer is: leave it alone.** The "30
pairs" below is not what this costs, and the trade for fixing it is far worse
than when the decision was taken. Both numbers, over every kind, frame, font and
orientation, with `secondaryAxis` and with `pareto`, across four data shapes:

                              scoping kept    merge applied to secondary too
    gutter overlaps                     14                                 6
    secondary ticks DRAWN            4,864                             3,470

**Eight overlaps fixed for 1,394 tick numbers deleted.** That is 174 readings
lost per pair gained. The entry below records the same trade at 34 fixed for 335
deleted — about ten — so the ratio has got seventeen times worse as the rest of
the engine improved, and the original decision to scope the merge off is not
merely still right, it is far more clearly right than when it was made.

The remaining family is 14 pairs of **1,014** (it was 1,327 before the unit-label
clip, which touches nothing here), or 1.4 percent, and **not one of
them is a tick number** — they are `title / series-label#` (8),
`combo-series-label#` on itself (5) and `series-label#` on itself (1). Zero
overlapping pairs anywhere in the sweep involve a `secondary-axis` label, out of
4,864 of them drawn across 2,228 charts.

So this is no longer a design waiting to be built. It is eight overlapping pairs
whose obvious remedy costs 1,394 axis readings, and the design below stays
written for the day somebody wants those eight badly enough to pay for a real
allocator. Nobody should, on these numbers.

A combo's column names, its line names and its secondary-axis tick numbers all
occupy the same two points of x. Laying the first two out in one pass fixed the
ordinary combo (see `seriesLabelNodes`' `extra` parameter), and doing the same
where the secondary axis is present is worse, not better: measured, it fixed 34
overlaps and DELETED 335 tick numbers, leaving a comfortable 480x300 pareto with
two of them. So that case keeps its 30 pairs.

Three families in one strip wants a design — a gutter that is allocated once,
with each family's claim on it stated — not a third pass over the same nodes.

**THE DESIGN, written 2026-08-29. Not implemented: it changes how labels are
placed on every combo chart, so it wants the sweep, the ratchet and a round,
not a late-night patch.**

The failed attempt shared the band by YIELD — merge everything, then let the
secondary strip drop whatever lands on a name. That is why it deleted 335 tick
numbers: `seriesLabelNodes` spreads names over the WHOLE band from the title's
bottom to the frame's bottom whenever they collide, so after the merge there is
almost nowhere a tick can land that is not on a name.

The asymmetry the fix should turn on is not importance, it is **whether a label
may move at all**:

- A secondary tick's `y` is DETERMINED — it is that value's position on the
  secondary scale. Move it and it is a lie. There are typically four to six.
- A series name's `y` is a PREFERENCE — the last segment's midpoint — and
  `seriesLabelNodes` already overrides it freely: it pushes names apart by
  `lineH`, clamps the overflow, spreads them evenly over the band when the gap
  cannot be honoured, shrinks to fit, and drops them all past `MIN_LABEL_FS *
  1.25`.

So the anchored family should be placed FIRST and become geometry, and the
movable family should be fitted around it. Today it is exactly backwards: the
ticks are emitted first and then yield to whatever was already drawn, while the
names spread as though the strip were theirs alone.

**The change, concretely.** Give `seriesLabelNodes` a fourth input — the y
intervals the secondary strip has claimed — and have the spread place names in
the gaps between them rather than across the band. Everything else it does
already applies: too little room and it shrinks, too little for that and it
drops, which is the answer this engine gives everywhere else a reservation
cannot be paid for. No new pass, no new yield rule, and the secondary strip
stops needing one at all.

**What it costs, stated in advance so the measurement is honest.** Names lose
band, so charts with many series in a dual-axis combo will shrink or drop names
that are drawn today. That is the trade this file has already made twice, and
it is the right way round only if the count moves: the 30 pairs should go to
roughly zero without the 335 tick numbers being spent. Anything else and the
answer is that a dual-axis combo with many series has no room for both, which
is a legitimate finding and should be recorded rather than forced.

**Order of work:** intervals in first with the parameter unused and the sweep
re-run to prove it is inert; then the spread; then re-measure. Two numbers
decide it — pairs, and tick numbers drawn.

---

**FIVE INSTANCES OF ONE DEFECT, worth naming because it will recur.** A strip is
fitted to the room it has, and a CLAMP a few lines below — there to keep a label
on the canvas — moves it back toward its neighbour and takes the room away
again. Found in the secondary value axis, the scatter's x tick numbers, the
gantt's date strip, the value-axis title, and the gauge's slice labels. The
symptom is labels that are SIZED correctly and still touch, beside a
`Math.max(0, …)` or `Math.min(width - w, …)` on the label's own x or y.

Three remedies, in order of preference: one shift for the WHOLE strip so every
gap is preserved; pay in SIZE, shrinking until the clamped layout clears; or
test the fit on the CLAMPED position rather than the raw one.

**And its sibling: a fit that measures something ADJACENT to what is drawn.**
The gauge's total measured unbold and drawn bold. The gantt's nudge measured at
`fs * 0.9` and drawn at `headFs`. A test measuring node boxes where the gate
measures ink. A label's `y` read as its ink when it is the box top and the ink
sits lower inside it — that one was mine, and the post-condition I wrote against
it never fired once, which is how it was caught.

**THE BLIND SPOT IS THE LESSON ABOVE ALL OF THEM.**
`test/frame-fit.test.ts` sweeps every kind at eight frames and seven fonts and
allows no overlap at all. It was green for the nine days that this entire family
sat in the option and data-shape variants it does not sweep. A gate is only as
wide as its sweep, and a green one is a statement about its sweep rather than
about the product.

A related caution, learned the same way: the figure that stood in this file from
2026-08-19 — "75 overlapping pairs" — was measured with a cruder ink rule and
was never comparable to anything. Re-measuring gave 2,148. A number in prose
decays; the sweep is four lines of script and the archive is on disk, so re-run
it rather than quote it.

### A reconcile window sized from a slide count the host reports late — 2026-08-29

**Round 297 stopped a whole cycle on a verdict that was arithmetically correct
about the wrong population.** `two slides claiming one slot` FAILED with
`4 slides inserted, 4 kept, 4 of 2 still re-editable; 0 queued as duplicates`,
and the fatal check did its job: a scenario that was passing had stopped.

It is not the flake this scenario already has. Five earlier failures in 273
rounds: four of them — 060, 148, 253, 273 — read `2 queued as duplicates`, so
the dedup RAN and left slides behind, and the fifth (287) is the stale-chunk
`Failed to fetch dynamically imported module` that `src/render/lazy.ts` was
written for. None of the five found zero duplicates. This one did, which is a
different thing.

**The trace names the cause in three lines:**

    handed the host a generated deck   expectedSlides 2, landed 0
    handed the host a generated deck   expectedSlides 2, landed 2
    read the deck back                 range [3,5], read 2, unread 0

The first insert reported **nothing landed**. Its slides did land — the deck
went 3 to 7, all four of them — but not by the time anyone counted. So
`afterInsert = await slideCount()` answered 5, `reconcileDeck` was handed the
window `[3,5]`, and it examined two slides out of the four that existed. Two
distinct titles in the window, no duplicates, verdict `0`. Every step correct;
the window was wrong.

`unread` is **0** on that read, which matters: the host was not refusing. It was
BEHIND. Those are different failures and only one of them the guards already
cover — `duplicateSlot` checks `blind` on its readback and there was nothing
blind to catch.

**This is `addSlides`' problem seen from the other side.** That function exists
because "PowerPoint on the web silently drops some `slides.add()` calls", and it
verifies growth from a fresh context rather than trusting the batch. Here the
host does the mirror image: it reports slides as NOT landed that did land. And
`slideCount()` is already the fresh-context read, so a fresh context is not the
remedy — the host is simply late, the same lateness `slideShapeCounts` pays
`COUNT_SETTLE_MS` for on the shape counts.

**The fix, and why it is not in tonight.** `duplicateSlot` knows it asked for
four slides. When `afterInsert - before` is not four, the reconcile is about to
be given a window that cannot contain the duplicates, and the honest answer is
the one `blindSkip` already gives for the other kind of unreadable host: say the
run could not be judged, rather than judge it. That is a few lines.

It is not in tonight because it changes what a round MEASURES, and this
scenario's pass/fail series is 273 rounds long and is itself evidence. A change
to a verdict rule belongs in daylight, with the series re-read afterwards, not
at 2am between cycles. Recorded here so it is not re-derived.

**Read alongside** the entry above about the ratchet's own blind spot: three
times this week a gate has been correct about the thing it was measuring and
wrong about the thing it was believed to measure.

### The crash, counted across the whole archive — 2026-08-28

Twenty-four crash records over fourteen distinct builds, taking the LAST step
each one wrote and counting each phase once per build so one bad night cannot
vote three times:

    4 build(s)   draw  — parts list outcome
    4 build(s)   pane  — collecting deck evidence — scanning
    2 build(s)   probe — asking
    2 build(s)   draw  — batch issued
    2 build(s)   probe — re-asked what the empty deck could not answer
    1 build each update / group / probe answered / probe second pass

**Counting RECORDS rather than builds would have said `parts list outcome` five
times to the scan's four** — and three of those five are the same build on the
same day. One build that crashed three times is one piece of evidence, not
three, and the difference decides which phase looks dominant.

**2026-08-29 adds one build, and it breaks the tie toward the scan.** Build
`4275306` crashed five times in thirty-six minutes on the 4:3 leg, and four of
the five ended at the SAME step:

    01-49   513.4s   pane  — collecting deck evidence — scanning
    02-00   533.5s   pane  — collecting deck evidence — scanning
    02-10   523.6s   pane  — collecting deck evidence — scanning
    02-21   538.9s   pane  — collecting deck evidence — scanning
    02-25   160.8s   probe — did not re-acquire, the slide resolved

By the rule above that is **one** build, taking the scan to 5 against draw's 4 —
a lead, not a verdict, on fifteen builds. What is new is not the count but the
**consistency**: four crashes on one build, one phase, and all four between 513
and 539 seconds. Round 290's 524s crash sits inside that band too. So this
failure has a time signature — roughly nine minutes into a round, during deck
evidence collection — where the archive as a whole looked scattered.

Worth one more build's worth of records before anyone acts on it. Four of five
is a pattern; it is not yet the answer to "why nine minutes".

### FOUR MORE BUILDS ARRIVED, AND THE SCAN WINS — 2026-08-29

The entry above asked for one more build. That evening supplied four, and they
are unanimous. Counted the same way — one vote per build, so a bad night cannot
vote six times:

    e9222a8   4 of 4 crashes   pane — collecting deck evidence — scanning
    3495fb8   5 of 5           same
    557b10e   1 of 1           same
    4275306   4 of 6           same (the other two: one probe, one selftest)

**That is 9 builds on the scan against 4 on `draw`** — 14 of the day's 17
records, in a band of 441-572s. The sentence below ("no phase dominates") was
true of fifteen builds and is not true of nineteen. `collectDeckEvidence`'s own
comment blamed `slideShots`, the screenshot loop, as "the prime suspect"; the
per-phase tracing it added in the same breath has now acquitted it. Every one of
these died in `listChartsInDeck({ withInventory: true })`, the FIRST phase,
before a single screenshot was attempted.

**And what it costs is not the round — it is the FILING of the round.** Every
verdict is in before that scan. 33 of 49 crash records hold a complete
fourteen-scenario result, over 13 builds; on 2026-08-29 one 4:3 leg produced a
full result five times, 14/14 in four of them, and archived none of the five.
The pane now assembles and banks its run log BEFORE the scan, and
`scripts/salvage-crashed.mjs` recovered 22 of the historical ones into
`rounds-salvaged/` — all 22 of them 4:3, taking that arm from 26 rounds to 48.

**A rate nobody had tracked**, and it is a reason to watch rather than a finding.
Crashes per archived round, by day:

    08-24  0.14      08-27  0.75
    08-25  0.12      08-28  1.33
    08-26  0.23      08-29  1.24

Roughly tenfold, breaking on 08-27. The cause is unknown and the denominator is
sensitive: a leg that burns seven attempts contributes seven crashes and zero
rounds, so a single bad leg moves the ratio a long way. Round duration is not it
— 12 scenarios/364s then, 14/390s now. What else changed that day is the domain
move and v0.4.0. Nobody should conclude anything from this yet; it is here so the
next reading has something to be compared against.

So — of the sentence that follows — no phase dominated across FIFTEEN builds.
The per-phase traces did their job — before them, six of these would have been
filed under the probe re-ask — and what they showed then was a host that falls
over in several places late in a long round, not one broken call.

**The host says the channel was healthy.** `crashes/2026-08-28T01-51-11.md` is
PowerPoint's own account of the 524s crash in round 290: the document channel
answered `200` on every one of its last twelve `GetUpdates` calls, and every
console error in the window belongs to Microsoft's own infrastructure. Nothing
in it points at us. One number worth keeping: **87,331 document-channel requests**
by the time it fell over. That is what a ten-minute round costs a live
PowerPoint session, and it is the scale at which this happens.

**Severity, stated plainly:** the scenarios pass and the round archives 14/14.
This costs the EVIDENCE COLLECTION after a run, not the run. That is why it has
been survivable for a month and why it has never been urgent — and it is also
what stopped the 4:3 leg on 2026-08-28.
