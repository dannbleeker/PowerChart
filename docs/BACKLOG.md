# Backlog

Curated candidate work, from a research sweep (July 2026) comparing every
existing chart kind against think-cell, Excel, Highcharts, Datawrapper, and
Mekko Graphics, plus a chart-type survey across the Zelazny / FT Visual
Vocabulary taxonomies and competitor add-ins (Zebra BI, Vizzlo, UpSlide).

**This is the only backlog document.** Items graduate from here into PRs and are
deleted when they ship — what has shipped is recorded by the README feature
table and by git, not here. Rejected ideas stay in §2 so they aren't
re-proposed.

Feasibility is judged against the live-add-in constraint: rects, lines, text,
ellipses, and any of PowerPoint's 177 preset geometries (all of them
PowerPointApi 1.4), plus polygon *outlines* — no freeform curves, and no
images. The SVG and skill-pptx renderers additionally have filled polygons and
patterns.

## 1. Open

### Two of every three redraws are a grouped chart the update cannot map — 2026-08-26

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
that path draws shape by shape and pays the curve. **PowerChart is fast exactly
when it CREATES slides and slow exactly when it ADDS to one.**

**Why this is a product decision and not a fix.** The side-by-side layout — two
or three small charts on one slide — is the everyday think-cell workflow, and
it is the slow case. Nothing in the pane tells the user that, and nothing
should be changed unilaterally: the cost is the host's, the mitigation
(`leastLoadedChart`) is correctly scoped to the self-test, which picks a chart
arbitrarily and so should pick wisely. A user picks the chart they mean, so
there is no choice for the product to make on their behalf.

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
### Text that overlaps text on data the samples do not carry — MEASURED 2026-08-19, not fixed

The frame gate's overlap half sweeps `sampleConfig(kind)`, and a sweep of the
same kinds under **option** and **data-shape** variants finds 75 overlapping text
pairs it cannot see. The overflow half of both sweeps is closed and gated; this
is what is left, measured, so nobody has to re-derive it.

    46  24 categories        pie/doughnut adjacent outside labels, radar category names
     8  10 series            legend rows against each other and against the plot
     8  long category names  the shared category axis, already fitted, at its floor
    11  valueAxisTitle       against the column totals and the topmost tick number
     4  pareto               the secondary axis against the primary's numbers
     2  secondaryAxis        same
     2  every value negative
     1  pie.semi

**Why each is left rather than clamped.**

- **24 categories** is the honest one: a 24-slice pie has more labels than it has
  ring. The outside labels already shrink to the gap between NEIGHBOURS (the
  sunburst rule), and at 24 slices that floor is reached and the labels are
  dropped or collide. What is missing is a decision — drop every label past N, or
  draw a legend instead — and that is a product call, not a bound.
- **`valueAxisTitle` was attempted and REVERTED**, which is the useful record.
  Its width is `Math.max(frame.x - 4, textWidth(…))`, a floor that raises a
  width, so a long unit grows right over the totals. Fitting it to the axis
  gutter drops it from every chart drawn WITHOUT a value axis — the stacked
  sample among them, where `frame.x` is zero — and `keeps a numeric axis title`
  caught that immediately. The room it really has is "whatever else is in the
  band above the plot", i.e. the totals row, which is a coupling rather than a
  bound. It wants a decision about where a unit belongs on a chart with no axis
  column.
- The rest are one shape: two independent numeric strips sharing a band.

**Priority:** low. Every one is inside the frame, so nothing lands on the slide;
they are legibility, not damage. Worth doing when someone is next in `frame.ts`
with the totals row open in front of them.

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

**Researched:** 2026-08-02.

The standing test run is down to four items because the in-host battery
absorbed the rest. The remaining ones need a human because nothing can click
the pane — but something can. Playwright reaches an add-in task pane, and there
is a public working reference implementation to copy from
(`kzarzycki/powerpoint-mcp`, its `e2e/` directory).

**What makes it tractable here:** `manifest.xml` already points at
`https://localhost:3000`, which is the constraint Microsoft's own (undocumented)
sideload seam enforces — `office-addin-dev-settings` appends
`wdaddindevserverport` / `wdaddinmanifestfile` / `wdaddintest` to an
Office-on-the-web document URL, gated behind a `WEB_SIDELOAD_TEST` env var that
exists for Microsoft's internal CI. `playwright-core` is already a
devDependency.

**Blockers, all defeated in that reference repo:** headless user agents are
sniffed and sideloading silently skipped (spoof UA + `sec-ch-ua`); Chrome's
Private Network Access blocks the WAC origin from reaching loopback (intercept
with `context.route` and answer `Access-Control-Allow-Private-Network`); the
`wdaddintest` flag does *not* in fact suppress the developer-mode dialogs, so
they must be clicked. Reach the pane with `page.frames()` — it is a nested
cross-origin iframe, so `frameLocator()` chaining cannot get there, and Selenium
cannot reach it at all (office-js#5264).

**The real wall is login**, not the add-in machinery. That reference repo primes
a browser profile by hand in a headed window and runs headless afterwards; its
own CI does not run the suite. Headless would need Entra certificate-based auth
(Playwright `clientCertificates`, ≥1.46 — Microsoft's Power Platform samples do
exactly this), which needs a test user with CBA that the owner probably cannot
self-serve on a corporate tenant.

**It cannot be developed from an agent container** (checked 2026-08-06, not
assumed). The environment's network policy answers 403 to `CONNECT` for
`www.office.com` and `powerpoint.officeapps.live.com`; only
`login.microsoftonline.com` resolves, which is useless on its own. So the local
half is *local* in a stronger sense than this entry first implied — it needs the
owner's own machine and his own signed-in browser, and no session working on
this repo can build or run it end to end.

**Priority:** medium — but only the local half. Run locally it turns the
remaining manual items into `npm run test:e2e` and pays for itself without
solving login. CI is a separate, owner-gated question and may never be worth it.
Covers PowerPoint **on the web** only; desktop stays human either way.

**What is left for it to absorb, now that the battery has taken the rest.** The
selection round trip is a scenario again (*edit the chart YOU click*), but it
waits for a **human** click and records the verdict — so what still needs a
person is the click and the drag, not the checking.

Do not have such a driver call `setSelectedShapes` to supply that click. On the
web it wedges the host's selection subsystem (`docs/RESEARCH.md` §4b), and the
wedge is in PowerPoint rather than in how it is called. A Playwright click on
the **canvas** is a real click and might dodge it entirely — the one thing worth
trying first if this is ever built.

### ~~Take more than two draws per arm~~ — MOOT, the question it served is closed

**Removed as an action 2026-08-16.** This existed to reach 60 draws an arm
faster, because `does a rasterise poison the next draw` could not be answered at
two per round. The rounds got there by accumulation: **60 per arm, 0 stalls in
both**, and the question is closed. Building a faster instrument for a question
already answered is optimising the measurement rather than the thing measured,
which is the mistake this file records elsewhere.

The rest is kept because the TRAP in it is general and still live — see below —
and because a future question sampled per-round will meet it again.

The original entry follows.

---

The rasterise question cannot be answered at the rate it is being sampled.
`does a rasterise poison the next draw` collects FOUR draws a round — two per
arm — and the report is explicit about what that buys:

    after a rasterise     0 stalled /  24 drawn = 0.0%
    after a cheap read    0 stalled /  24 drawn = 0.0%
    NOT an answer yet: 24 draws in the smaller arm. Telling rates this close
    apart needs nearer 60-100 an arm.

Twelve rounds produced 24. Sixty needs thirty rounds, and each round costs about
twelve minutes of a real PowerPoint plus a recovery when it wedges. At the rate
of 2026-08-14/15 — six landed rounds in a night — that is a week of nights for
one question.

Six draws an arm would get there in ten rounds instead of thirty, and costs
about a minute of extra host time per round against the twelve it already takes.

WHAT MAKES IT DELICATE, and why it is written down rather than done at the end of
a long session: `rasteriseArmVerdict` is a pure function whose FIRST version
manufactured a finding. With the cheap arm first and the rasterise arm second the
rasterise arm always ran later, and round 11 reported "the draw after a RASTERISE
did not land" for a difference that position explains just as well. The fix was
counterbalancing — each call type once early and once late — and the verdict
reads position out of the pairing by INDEX: `raster[0]` against `raster[1]`,
`[raster[0], cheap[0]]` as the early pair.

So raising the count is not a constant. The function has to keep the
early/late reading with N per arm — interleave the order, split each arm at the
midpoint, and decide "both late arms failed" from halves rather than from a pair.
The existing verdicts must keep their exact wording for N=2, because twelve
rounds of history are read through them.

Worth doing, and worth doing with the same care the counterbalancing got.

### ~~READ THIS FIRST: a chart on a FRESHLY ADDED SLIDE cannot be grouped~~ — FIXED, measured 2026-08-25

**36 of 36 charts on freshly added slides now group, eight consecutive rounds at
four of four.** Against the 1% below. The chain this item describes is broken at
its first link: the pre-grouping re-read no longer comes back short on a slide the
run has just added, so the chart groups, so its tag goes through the group handle
instead of a `created` one, so it keeps its config.

**Keep the 1% figure and keep its date.** It is a true measurement of a host that
no longer behaves this way, and it is quoted in several places as though it were
current.

**And read how it was found.** The metric read `0/0` for twenty rounds before
2026-08-25 — not fixed, not broken, nothing — because `poolFreshVsEstablished`
keys on a per-chart label the draw path had stopped emitting once in-place
updates started succeeding. The problem had been closed for an unknown number of
rounds while the instrument that would have said so reported nothing at all. The
fix is not the interesting part; the silence is.

---

**Measured 2026-08-15 over the whole archive, and it WAS a switch rather than a
tendency:**

    slide already had shapes   82 chart(s), 81 grouped = 99%
    freshly added, empty       74 chart(s),  1 grouped =  1%

`npm run rounds` prints it under **WHICH SLIDE THE CHART LANDED ON**.

**The chain, each link measured:**

1. A chart drawn onto a slide this run has just added gets a pre-grouping re-read
   that comes back **short or empty** — 20 of 24, or nothing at all.
2. A short match is thrown away and an empty one has nothing to match, so the
   chart is **not grouped**.
3. An ungrouped chart's tag falls back to a `created` handle, which this host
   refuses about seven times in ten.
4. So it loses its config, and `same scale across the deck` fails — as it has for
   every round it has run — 34 of 34 as of 2026-08-16, and `npm run rounds`
   carries the current count rather than a number that rots here.

**This is not a new problem. It is THE problem, and this repo has been circling
it since #108.** `shape-add-held-slide-proxy` answers `threw`, a web-new-slide id
does not round-trip, and the #108-#111 saga was four attempts at drawing on a
freshly added slide. The tag work of the last four rounds — which handle, which
anchor, which context — was aimed one level above this.

**Ruled OUT along the way, each cheaply:**

- **Context wear.** `contextSyncs` says the failing re-read is the FIRST sync of
  its context, not the thirtieth. Chunking `updateChartsInSlides` would change
  nothing, and that 390-line restructure is ruled out before anyone starts it.
- **The `NNN#0` slide ids.** `256#0` carries the three best-behaved charts in
  every round; the id shape predicts nothing.
- **A failed `addGroup` poisoning the context.** The charts that lose their tag
  never attempt a group.

**Where a fix would go, and why it is not in this commit.** The code already
knows the honest rule — *"the positional rule is still right for a slide this run
added blank"* — but that branch is reachable only when NOTHING matched, so a
chart matching 20 of 24 falls past it and declines to group. On a slide this run
added blank we KNOW our shapes are the only ones there, so a short read is a host
lie we can detect rather than obey. That is a contained change to the matcher
rather than a restructure — but it is still surgery on the grouping path, which
carries three shipped-broken fixes on record, and it wants a person awake.

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

Nineteen rounds, ten pairs. Kept as a record of what the run left and what was
done about it, because two of the four turned out to be worth more than the
rounds that produced them.

- ~~**`signedOut` cannot tell a popup from a signed-out browser.**~~ **Done.**
  `signInIsPopup` distinguishes them by whether a document tab is still open, and
  the two messages now describe what the reader can actually see. Both still stop
  the round and both still hand the job back — if Office is asking for
  credentials, nothing measured past it can be trusted, and it needs a password
  either way. Deliberately loose about which document: the driver's recovery
  looks for `Presentation63` by name, and hard-coding one deck's name into a
  diagnostic is how it goes quietly wrong for the next deck.
- ~~**Two of the run's own fixes shipped without guards.**~~ **Done in the run
  itself**, and **the check is the lesson**: a green suite proves nothing about a
  path nothing tests, and "I watched it work in production" leaves no guard
  behind. Both were found by grepping each new symbol in the test tree — a pass
  now worth running at the end of any fixing session.
- ~~**`same scale`'s scenario summary counts what it carried, not why.**~~
  **Done.** `reReadNote` adds the cause to the verdict line: how many re-reads
  read short, how many read empty, and how many the settled retry repaired. It is
  silent on a clean round and it names `repaired 0` explicitly — that clause is
  the one that would refute the retry, and a clause that vanished at zero would
  read identically to a round where the retry never ran.
- **Rounds cost ~12 minutes and a pair costs a build.** Every pair after the
  mechanism was settled produced an identical result, and the audit found more
  than the rounds did. When a question is closed, say so and stop — the brief now
  opens with what is settled for that reason. **Not a task**; kept because the
  next long run should read it before spending its first hour.

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

### The probe has been blind on GROUPS for the whole archive — found 2026-08-16

`npm run rounds` prints it now, under **QUESTIONS THAT NEVER ANSWERED**. Six
questions have produced nothing in **41 of 41 rounds**:

    never asked — the harness could not set the question up (OURS to fix)
      grouped-child-by-id-from-slide     41 round(s)  no-scratch-shape
      shape-resolve-held-slide-proxy     41 round(s)  no-scratch-shape
      tag-on-group-survives              41 round(s)  no-scratch-slide
    asked, and the host would not answer (a fact ABOUT the host)
      addgroup-returns-usable            41 round(s)  unreadable
      group-children-via-getcount        41 round(s)  unreadable
      shape-proxy-survives-one-sync      41 round(s)  unreadable

> **PARTLY RESOLVED 2026-08-21.** Two of these six — `grouped-child-by-id-from-slide`
> and `tag-on-group-survives` — were RETIRED rather than fixed: production
> answered one and the product routed around the other. See "Two questions
> production answered before the probe sheet could" at the end of this file.
> The entry below is left as it was written.

**Four of the six are the group cluster**, and that is the cluster rounds 064 and
065 just made urgent. `tag-on-group-survives` asks precisely what those two
rounds discovered: whether a tag written on a group is honoured. It has never
once been put — and **production answered it in one evening, twice**: no, refused
5010, when the group sits on a freshly-added slide.

**Why it stayed invisible.** The per-round report names what was "never put in
this round", so a permanently dead question read as bad luck forty-one times
running. Nothing pooled it. Every round paid a scratch slide and host time for
each of these regardless.

**The split is the actionable part.** `no-scratch-*` means the harness could not
set the question up — ours, and the fix is to move the measurement into
production or retire the question. `unreadable` means the question was put and
the host declined, which is a finding and should be left alone. Pooling the two
would hide exactly the distinction that decides what to do.

**What this is evidence for, beyond the six.** The archive already records "a
question that cannot be asked is an answer about the harness", and this is its
largest instance: three questions the probe could never put, one of which was the
most important open question in the repo, answered by a trace line in production
the first evening anyone instrumented for it. **Prefer production instrumentation
to a scratch-slide question** whenever the thing being asked about happens in the
real path anyway.

**Not done here:** retiring or relocating the three `no-scratch-*` questions.
That is a real change to the instrument and wants its own session with a control
— the archive records four reverted attempts at changing the probe's slide
handling, every one of which changed what the probe measured.

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

    round   scenarios  origin5010  charts that cannot follow a drag
    074        12/13        8                    8
    075        13/13        0                    0
    076        13/13        0                    0

**Thirteen of thirteen, twice.** The item above — half the charts unable to
follow a drag — is CLOSED. The pooled count is history now: 34 charts across 5
rounds, all of them 072-074, and it will not climb again unless something
regresses, which is what it is there to notice.

**The fix, and why it is not what it looks like.** The origin tag was written
after `load("id,left,top")` had resolved the target into `shapes.getItem(id)`.
Routing that ONE write through the chart's binding — a handle taken from the live
proxy in the drawing batch, which never becomes a resolved one — took it from
eight or nine failures a round to zero, twice.

**This also settles a question the probe could never reach.**
`binding.getShape()` is NOT `shapes.getItem(id)` renamed. That was the staked
alternative and it is refuted: the same write lands through the binding and is
refused through the resolved proxy. `binding-names-shape-later` asked eight times
across eight rounds and never once got to its own question — a fact about the
PROBE, and the third time production has answered something the probe could not.

#### What is left on this line

- **`settledByBinding` is still 0 across seven rounds.** The settle-by-binding
  route has never executed, because the settle only runs when the config tag
  fails and the config tag stopped failing. Kept deliberately — it costs nothing
  at rest and is the only route left if the config tag regresses — but it is
  **untested in production** and must not be mistaken for exercised code.
- **A real mouse drag is still owed.** `an update follows a moved chart` proves
  the ARITHMETIC by moving a shape programmatically; test 4 of the standing run
  is what proves a mouse. Nothing here substitutes for it.
- **4:3 HAS NOW BEEN RUN ONCE, and the result is confounded.** Round 077 scored
  10 of 13 against 13 of 13 twice on the same build, with a failure class this
  archive has never held: **52 `UnexpectedError`**, 36 on the config tag write
  and 16 on the settle's binding route, mostly through the GROUP handle. Every
  5010 stayed fixed — `cfg5010`, `orig5010` and `origLost` are all zero.

  **It also exercised the settle-by-binding route for the first time in eight
  rounds**, and that route was refused too. The fallback is no longer untested;
  its first test says it does not save this case.

  **But TWO variables moved**: the aspect ratio, and the DECK. Round 077 ran on
  `Presentation65`, not the `Presentation64` carrying all 52 previous rounds —
  and that deck was created during a browser crash the same day, with its whole
  ribbon greyed out until a reload. `UnexpectedError` is Office.js's generic
  failure rather than the specific `InvalidParam / 5010` that names a refused
  handle, which fits a sick document as well as it fits an aspect ratio.

  **THE CONTROL RAN, AND 4:3 IS EXONERATED.** Round 078 — same deck, same build,
  16:9 instead of 4:3 — scored **10 of 13 with 52 `UnexpectedError` and 18
  tagging failures, identical to 077 on every number**. The failure follows the
  DECK. `Presentation65` was created during a browser crash, has greyed out
  twice, and began round 078 holding 100 slides.

  **Nothing has been learned about 4:3 yet** — round 077 measured a sick
  document. The question needs a clean 4:3 deck, which is item 4 below.
- **Desktop remains untested against a host.** Every archived round is
  PowerPoint on the web.

### PUTTING 4:3 INTO THE NIGHTLY RUNS — four blockers — ALL CLEARED 2026-08-16

**All four shipped.** The run log carries `slideSize`; `scenarioRegressions` and
`profileDivergence` are profile-scoped; readiness verifies the size and refuses
`wrong-size`; and `npm run cycle` runs the agreed schedule. Rounds 079-081 were
the first full cycle and 4:3 was exonerated — see the journal.

**One blocker they did not anticipate**, found on 2026-08-17: `PW_DECK` reached
only `recover`, so the 4:3 leg measured whichever deck the previous leg left open
and refused with `wrong-size` every night. `selectDeck` now fronts the named deck
before anything is measured, and refuses with `deck-missing` when no tab carries
it. A deck the add-in is not registered for refuses with `addin-missing` instead
of retrying seven times.

The reasoning below is kept because it is the record of what each blocker cost.

Round 077 was driven by hand. Everything below is what stopped it being a
command, found by doing it.

#### 1. A ROUND DOES NOT RECORD ITS SLIDE SIZE — fix this first

`077-357632b.json` carries `build`, `host`, `hostAnswers`, `selftest`, `deck`,
`trace`, and **nothing that says it ran at 4:3**. So nothing downstream can tell
the two apart, and `npm run rounds` pools them into one number: `WHICH SLIDE THE
CHART LANDED ON`, `DID THE CHART SPAN BATCHES`, `CHARTS THAT CANNOT FOLLOW A
DRAG` would each silently average two different experiments.

That is the rounds 24-and-25 mistake — "differed only in this, and were compared
as though they did not" — except automated and running every night.

`slideSize()` already resolves it at runtime through three fallback rungs and is
unit-tested at 720x540. Put `{ width, height, source }` in the run log beside
`build` and `host`. **Everything else here depends on this one field.**

#### 2. THE REGRESSION GATE WOULD CRY WOLF EVERY NIGHT

`scenarioRegressions` establishes on the previous three rounds and judges the
newest. Alternate profiles nightly and a 4:3 round is judged against three 16:9
rounds — and round 077 scored 10 of 13 where 16:9 scored 13 of 13 twice.

The gate fires, the fire is spurious, and someone switches it off. This file
already records that happening to a gate that cried wolf, and the fix for the
LAST false alarm shipped hours ago. **Segment establishment and comparison by
profile**, so a 4:3 round is only ever measured against 4:3 rounds.

#### 3. THE SLIDE SIZE IS NEVER VERIFIED, AND A SILENT MISS IS EASY

Readiness checks the build stamp, the deck, both toggles and the host's
liveness. It does not check the size.

**This is not theoretical.** Setting Widescreen during the control run SILENTLY
DID NOT TAKE — the click landed while the document was in its greyed "Loading"
state, the menu accepted it, and nothing changed. It was caught only by
reopening the menu and reading which box was checked. A round that believes it
is 4:3 and is not proves exactly nothing, which is the same class of harm as a
round on a stale pane — and that one is already a hard stop.

Add a `wrong-size` refusal with its own code, against an expected profile
(`PW_EXPECT_SIZE`), and let `recover` NOT try to fix it: changing a deck's slide
size mid-run would change what the round measures.

#### 4. THE DECK IS SET UP BY HAND, AND THIS ONE HAS A HISTORY

Slide size is a ribbon click the driver cannot make, and a sideload is
per-document.

**HALF OF THAT IS NO LONGER TRUE — corrected 2026-08-20.** `PageSetup.slideWidth`
and `slideHeight` are writable at PowerPointApi 1.10, so the driver CAN set a
deck's size, behind `PW_SET_SIZE=1` and only for a deck that exists to be that
profile. It was used for the first time on 2026-08-20 to take `Presentation67`
from 960x540 to 720x540, verified by a read-back in a separate call. Only the
sideload is still owner-only. That is a one-time cost per deck and is fine — but the deck used
for round 077 is not.

`Presentation65` was created during a browser crash, and has now gone into a
greyed, unusable ribbon state **twice**, each time needing a reload before it
would accept input. A nightly series must not start on a document with that
history. **Create a clean 4:3 deck for it**, sideload once, and point rounds at
it with `PW_DECK` — which already exists.

#### The order, and why

**1 before everything.** Segmenting (2), verifying (3) and reporting all need a
round to say what it was. Building any of them first means building on a guess
about which rounds were which, and the archive has 53 rounds with no size on any
of them — so the field also needs a documented default of 16:9 for everything
already filed, stated once, rather than inferred per reader.

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

## Retire the positional group-member mapping with `Shape.creationId`

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

`grouped-child-by-id-from-slide` and `tag-on-group-survives` are gone from
`PROBES`, from `FAKE_BASELINE`, from `KNOWN_DIVERGENCES`, from
`PENDING_QUESTIONS` and from `WHAT_IT_MEANS`. The probe sheet is 31 questions,
down from 33.

**They never answered once, in the entire archive.**

    rounds carrying a probe sheet: 125

    grouped-child-by-id-from-slide   no-scratch-shape 117, no-scratch-slide  8   = 125/125 starved
    tag-on-group-survives            no-scratch-slide 122, no-scratch-shape  3   = 125/125 starved

Every one of those 250 attempts spent a scratch slide and host time on a
question that could not be put, in a run where scratch slides are contended —
so they were not merely useless, they were starving questions that could still
pay.

**Why they starved is settled, and it is not lateness.** They sat directly
under `group-reports-its-children`, which carries `burnsTheSlide: true`. A
question placed beneath a slide-burner finds the slide gone. The rival theory —
that they starved for sitting at positions 22 and 23 — is refuted in that
probe's own comment: round 26 answered #31, the last question of all, while #8,
#16, #22 and #23 starved, and the positional split came out 55% for questions
1-8 against 62% for 9-and-later. **Starvation here is a property of the SLOT.**
Anything moved in under that probe inherits the same fate, which is why the
answer was never "promote them".

## What actually answered them

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

## The branch that tried the other remedy

`claude/ask-the-decisive-probes-early` (`f09041c`, 2026-08-13) proposed moving
both to positions 5 and 6. It is pushed and preserved, and it was NOT merged,
for three independent reasons: it ships a deliberately failing test ("KNOWN RED,
AND THE REASON IS THE POINT"); 23 commits have touched `host-probe.ts` since,
and it no longer rebases cleanly; and its premise — that these two questions
gate the in-place update — is now false.

Its diagnosis was still worth keeping, and is recorded above: these questions
never get put. The conclusion inverted once production answered them.

## What is left open

The `no-scratch-*` split in the older entry still stands for the four remaining
starved questions. Retiring two of the six is not a fix for the harness; it
removes two questions that could never pay, and gives their slide back to the
ones that might.

## The slide-size ladder's first rung hangs about twice as often as it answers

**Measured 2026-08-23, over all 147 archived rounds. Not acted on, because the
numbers do not yet say which way to act.**

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

### The four dependabot advisories, and why three of them are not ours — 2026-08-26

Every push prints "GitHub found 3 vulnerabilities on the default branch (2 high,
1 moderate)". Checked once so it does not have to be re-checked every push.

    image-size  *        HIGH x2   ICNS and JXL/HEIF parsers loop forever on a
                                   malformed image (GHSA-w3rx-r6r6-pgpr,
                                   GHSA-5p2g-fcmc-qvqq)
    qs  6.11.1-6.15.1    MODERATE  qs.stringify crashes on null entries in
                                   comma-format arrays (GHSA-q8mj-m7cp-5q26)

**`qs` is dev-only.** `@stryker-mutator/core` -> `typed-rest-client` -> `qs`. It
is never shipped and never runs outside mutation testing. `npm audit fix` clears
it without a major bump; it is queued behind anything that matters.

**`image-size` is NOT REACHABLE.** It arrives under `pptxgenjs`, which IS a
runtime dependency — `src/render/pptx-deck.ts:74` imports it dynamically — so the
tree alone makes it look shipped and live. It is not: pptxgenjs calls
`image-size` to size images handed to `addImage`, and **PowerChart never calls
`addImage`**. There is no image path in `pptx-deck.ts`, no PNG or JPEG or data
URI anywhere in it. The deck is shapes and text. Both parsers named in the
advisories need an image to parse and are handed none.

The fix would be `npm audit fix --force`, which is a MAJOR pptxgenjs bump — real
breakage against an unreachable path. Not taken.

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
