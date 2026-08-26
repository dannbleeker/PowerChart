# What we know now

`ROUND-LOOP-JOURNAL.md` is the reasoning record: eight thousand lines of how each
thing was learned, including the wrong turns, in the order they happened. It is
the right file to read when you want to know *why*.

It is the wrong file to read when you want to know *what is currently true*, and
that has cost this project real work — findings have been rediscovered because
nobody could locate them, and a fixed problem went on being quoted as broken for
an unknown number of rounds.

**This file is the short answer.** Lines marked ✓ carry a claim id and are
re-checked against the archive on every `npm run rounds` — see
`scripts/claims.mjs`. If one prints STALE, the world moved and this file is
wrong. Unmarked lines are believed but unchecked, which is a weaker thing.

---

## The host

- **✓ A chart on a freshly added slide groups and keeps its config** — 36 of 36.
  `fresh-slides-group`
  It was 1 of 74 on 2026-08-15. That figure is history and must not be quoted as
  current; it still appears, marked, in six places.
- **✓ The first chart of a multi-chart update costs ~2.2x a later chart of the
  same size.** `first-chart-costs-more`
  Position is now demonstrated with load measured (round 239). Load's independent
  effect is still open.
- **✓ The write syncs carry that 2.2x; the tag sync does not.**
  `tag-sync-is-not-the-writes`
  A within-call asymmetry, so it does not pay the between-round floor.
- **✓ The slowdown is entirely host-side** — our idle between calls is 1ms in the
  fastest round and the slowest alike. `our-idle-is-negligible`
  If this ever climbs, every timing conclusion in the journal needs re-reading.
- **✓ An id this host named still resolves through a slide handle a sync old.**
  `id-through-aged-slide-handle-reads`
  `shapes.getItemOrNullObject(<id>)` on a slide handle acquired a sync earlier
  reads the shape back — which is what `deleteShapesById`,
  `setShapeSelection` and the selection path all depend on, and what nothing had
  ever measured. The question answered `no-scratch-shape` in 216 of 216 rounds
  because it minted a scratch shape and this host will not name one; it becomes
  askable with an id from a chart the self-test has drawn and tagged.
  **The five `unreadable` answers in rounds 241-247 say nothing about
  PowerPoint.** Each was our own stale id, and round 247 proved it by listing the
  slide: `the id is NOT among the slide's 11 listed shapes`. Do not quote them.
  The claim counts only rounds whose detail says the shape WAS in the listing.
- **✓ Dropping the selection is what stops the stall** — 10 stalls in 457 draws
  made with a shape still selected, **0 in 516** with it dropped.
  `dropping-the-selection-stops-the-stall`
  So `dropShapeSelection` is not a precaution, it is the fix, and there is now a
  number behind it. The scenario's own note concluded "over the whole history
  they do not separate" from seventeen ROUNDS; counting DRAWS over 235 rounds
  they separate cleanly, and the zero is the load-bearing half.
- **✓ When the probe DOES buy a replacement slide, the question usually answers
  on it** — 15 of 18. `buying-a-replacement-slide-rescues-the-question`
  Including 9 of 12 where a SHAPE refusal prompted the buy, which the code's own
  comment called "a weaker reason to suspect the slide". It is a weaker reason
  and it pays. Recorded because 18 slide adds a round looks like pure waste from
  the outside, and that is how a correct behaviour gets optimised away.
- **✓ The probe gets its scratch slide back by asking the deck, not by adding
  another.** `scratch-slides-are-re-acquired-not-rebought`
  The id `addScratchSlide` captures is real when captured and names nothing
  later — the run held `4123571114#123571113` while the deck listed the same
  slide as `256#2587447327`. Two id spaces, not a renumbered neighbour. Round
  254 against 252/253: slides bought **63 to 16**, deck peak **110 to 38**, probe
  phase **133s to 73s**, and delete-by-id returned **21 where the whole archive
  before it returned 0**. A STALE here means the run is silently buying a slide
  per question again — silent because the replacement path still works, which is
  how the cost hid for 250 rounds.
- **✓ A by-id lookup that refuses is always rescued by re-reading the slide** —
  105 of 105, and not one re-read threw.
  `the-re-read-always-rescues-a-refused-lookup`
  This host refuses `shapes.getItemOrNullObject(<id>)` often enough to matter,
  and the refusal poisons the whole SYNC — unguarded it took every chart in the
  batch down and the caller saw a null target, which is a silent no-op on a
  chart the user is looking at. The report has always counted the refusals
  (`idRefusals`, 380 in `explode a degraded picture` alone) and never the
  recoveries, so the rate could not be formed and the recovery's worth was
  assumed. It is now watched: a STALE here means updates have started dying
  wholesale again.
- **A slide tag lands, and a second write overwrites** — measured directly on a
  REAL slide 2026-08-26, three runs, `overwrites` every time. `DEMO_SLOT_TAG` is
  sound. The probe's `tags-add-same-key-twice` has answered the uninformative
  `other` in 224 rounds because it asks on the SCRATCH slide, which this host
  treats differently — the same slide whose shape collection reads `unreadable`
  in 92% of rounds. **Do not quote that probe answer as a fact about tags.**
- **A grouped chart's children are not addressable by id off the slide** —
  measured directly 2026-08-26, three runs: `no-such-shape`, the child read back
  as `undefined`. Grouping takes the shapes OUT of the slide's collection, and
  the only route back is `shape.group.shapes`, which this host will not
  enumerate. So a grouped chart's update has no way to map nodes to shapes and
  MUST redraw — that is not a missed optimisation, it is the only correct
  behaviour available, and it is the price of the grouping that keeps the config
  tag. See `docs/BACKLOG.md`.
- **Group children are unreachable.** 548 throws, 0 answers. office-js#3014 is
  closed upstream; the fix has not reached this host.
- **The host is unversioned** (`0.0.0.0`) — but probe answers are a behavioural
  fingerprint, and it has been flat across the whole archive.

## Running the loop

- **✓ Rest 45+ minutes between rounds.** `rested-rounds-rarely-skip`
  Rest buys COMPLETENESS, not speed. The driver warns from session position 5.
  - Rounds 230-238: nine rested, zero skips. Rounds 216-225: ten back-to-back,
    ten skips, starting at the fifth.
  - **Not "never skips".** Round 226 was rested, did not crash, and skipped the
    rescale anyway — 1 in 12. The first version of this claim said *zero* and its
    own checker refuted it within minutes, because the window had been
    hand-picked. It is a large difference in rate, not an absolute.
  - The 10-in-10 half cannot be machine-checked: those rounds predate
    `sessionIndex`, so the claim reads `?` until enough deeper rounds carry it.
- **The noise floor:** two identical rounds differ by **14% (IQR)** typically,
  **73%** at worst. Under 14% is noise. A pair sees a 2x effect and cannot see a
  30% one.
- **Never move HEAD while a round runs** — the driver refuses as `site-behind`.
- **Never run the gate while a round is starting** — spawn starvation killed
  round 197.

## Open

- **Why an insert onto an occupied slide loses draws.** (Narrowed 2026-08-25 —
  the selection half of this is now answered, see the claim above.) Pooled over the archive:
  `a selected shape survives an insert` stalls 10 of 455 draws (2.2%),
  `insert onto a slide that already has content` 8 of 920, while
  `edit a chart on the visible slide` and `explode a degraded picture` stall on
  none of 366 and 171. A stall costs the scenario its verdict — it SKIPS and
  checks nothing — so this is the largest source of unchecked rounds in the
  battery, and it reads as "14 of 14" every time.
  What REMAINS open is insert-versus-update and how full the target slide is.
  I wrote here that "the experiment that would settle it does not exist yet" and
  that was wrong twice over: a dedicated arm was built and deliberately removed
  (every slot allocated, widening the band broke the no-overlap invariant), and
  the comparison it was for has been running all along as
  `edit the chart the user selected`. Read the scenario's own note before
  proposing it again. `npm run rounds` carries the rates.

- **✓ The first chart always lands on the deck's busiest slide**, by construction.
  `first-chart-is-on-the-busiest-slide`
  This is the CONFOUND behind four corrections. A STALE reading here would be
  good news: it would mean the harness stopped confounding position with load.
- **✓ No chart has ever reached an update carrying a parts list** — 0 of 1091.
  `parts-list-never-consumed`
  **Corrected 2026-08-26: this is not a fix that is late, it is impossible by the
  route everyone assumed.** A grouped chart's children are not addressable by id
  off the slide, so a parts tag written for one would name ids that resolve to
  nothing — and essentially every chart groups now. The redraw it causes is the
  only correct behaviour available, and the price of the grouping that keeps the
  config tag. What would move this line is a host that enumerates group children,
  or charts starting to FAIL to group — the second would be a warning, not a
  victory.
- **✓ Tag faults are zero** across the last 20 rounds. `tag-faults-are-zero`
  Thread 1 therefore cannot be tested — both arms of any experiment score zero.
  It needs a trigger that reproduces tag loss on demand.
- **✓ The `no-queue` trace is dead** (last round 065). `no-queue-trace-is-dead`
  Thread 3 as written is closed; its live successor is a sparse single-chart
  `tagging failed`.
- **Whether a STORED id survives, as opposed to a freshly observed one.**
  Corrected 2026-08-25: an earlier version of this line said the tested charts
  were "observed seconds later" and called the aged case untested. That
  conflated two different ages. The chart's age was measured from the archive
  and is **9-10 minutes** — drawn by ms 251-275k of the round, asked at ms
  794-845k (rounds 248 and 251). It is the OBSERVATION that is seconds old,
  because the re-ask scans the deck immediately before asking.
  So what remains open is narrower than it looked, and it already has a claim:
  an id written down and reused later WITHOUT re-enumeration is the parts-list
  case, tracked by `parts-list-never-consumed`. **Closed 2026-08-26 for the case
  that matters**: a grouped chart's children are not addressable by id at all, so
  a stored id would name nothing. It is not blocked by 5010 waiting for a fix —
  there is no route. A persistent-deck harness was considered for this and is not
  needed.
- **Load's independent effect on update cost.** The experiment is named: the lone
  arm on a LOADED slide against the lone arm on a clear one, holding run-length
  at one. Occupancy is now recorded, so both arms are readable.

## Habits that earned their place

- **Every difference gets its sample size and its spread before it gets a
  sentence.** Two claims died on their error bars on 2026-08-25; one IQR column
  would have caught both.
- **When a question survives two attempts at analysis, instrument it.** The
  position-vs-load question took four corrections and ~200 rounds of argument,
  and one `getCount()` in the right place.
- **Read the code that writes a field before grouping by it.** A proxy renamed
  for being a bad proxy was then used as one, twice in a day.
- **Mutate every new test.** Five tests passed against the code they were written
  to catch, in one day, and no gate caught any of them.
- **Before believing a host finding, prove the input was good.** The probe read
  `unreadable` five times and it looked like office-js#2903 confirmed. It was a
  stale id every time. What settled it was making the instrument report the
  things that would explain the answer WITHOUT the host being at fault — did the
  slide resolve, does it hold shapes, is the id in its listing — each of which
  was added only after an answer had already been misread once.
- **The noise floor is RANGE 73%, IQR 14%** — build `eba1c4d`, nine
  first-of-session rounds (230-238), in-place update 18-of-24. The IQR is the
  honest headline: the range only ever grows with n, so it is a lower bound that
  looks like an estimate. **A difference smaller than 14% is not evidence of a
  change.**
- **A report that opens with "too few to call it" when an answer exists is
  wrong, not merely unhelpful.** The floor section prints a block per build and
  most builds have one qualifying round, so it opened with five refusals and
  buried the n=9 answer in the middle. It was read as "the floor is
  unmeasurable" and a four-hour plan to re-measure it was proposed against a
  figure the archive already held. It now leads with the best available floor.
- **Every headline probe answer in the archive is a COLD answer.** `record`
  keeps the first real answer and lets nothing real displace it — deliberately,
  so a sheet means today what it meant yesterday. Measured 2026-08-25, the cost
  of that rule is real: `shape-add-positional-slide-proxy` says `yes` 85% of the
  time on pass 1 and 67% later; `binding-names-shape-later` is `silent` 15% cold
  and ~0% warm; `shape-add-held-slide-proxy` says `yes` 3% cold and 11% warm.
  The directions differ — **collection reads are worse cold, positional slide
  reads are worse warm** — so this is the shape of the HARNESS, not of
  PowerPoint. `npm run rounds` now prints the shifts. A single archived answer
  should be read as "what this host said first", not "what this host does".
- **A renamed trace message is indistinguishable from a fixed fault.** Both read
  as a zero. `a slide's shape count would not settle — not counting it` sat in
  the quiet-instruments list for 140 rounds while the instrument fired every
  round under a rewritten tail. `npm run rounds` now classifies each quiet line
  as STOPPED, RENAMED or REMOVED by searching the source's string literals —
  comments excluded, because this codebase names retired lines in prose and a
  comment saying "this is dead" otherwise vouches for it being alive.
- **When a record keeps going stale, stop recording and start observing.** The
  named id was fixed twice for freshness and was still wrong; a deck scan taken
  seconds before the question is right by construction rather than by
  maintenance.
- **A correction sweeps its own section.** Five stale statements survived four
  hours directly above the correction that replaced them.
- **State findings over the population, not the window you measured.**
  "Rested rounds skip nothing" was true of the nine rounds it came from and false
  of the twelve that qualified.
