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

- **✓ The first chart always lands on the deck's busiest slide**, by construction.
  `first-chart-is-on-the-busiest-slide`
  This is the CONFOUND behind four corrections. A STALE reading here would be
  good news: it would mean the harness stopped confounding position with load.
- **✓ No chart has ever reached an update carrying a parts list** — 0 of 1032.
  `parts-list-never-consumed`
  Blocked by the 5010 stale-proxy bug, office-js#2903, closed as not-planned. It
  causes 1197 redraws. If this goes stale, that is GOOD NEWS.
- **✓ Tag faults are zero** across the last 20 rounds. `tag-faults-are-zero`
  Thread 1 therefore cannot be tested — both arms of any experiment score zero.
  It needs a trigger that reproduces tag loss on demand.
- **✓ The `no-queue` trace is dead** (last round 065). `no-queue-trace-is-dead`
  Thread 3 as written is closed; its live successor is a sparse single-chart
  `tagging failed`.
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
- **A correction sweeps its own section.** Five stale statements survived four
  hours directly above the correction that replaced them.
- **State findings over the population, not the window you measured.**
  "Rested rounds skip nothing" was true of the nine rounds it came from and false
  of the twelve that qualified.
