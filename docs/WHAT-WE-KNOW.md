# What we know now

`ROUND-LOOP-JOURNAL.md` is the reasoning record: eight thousand lines of how each
thing was learned, including the wrong turns, in the order they happened. It is
the right file to read when you want to know *why*.

It is the wrong file to read when you want to know *what is currently true*, and
that has cost this project real work — findings have been rediscovered because
nobody could locate them, and a fixed problem went on being quoted as broken for
an unknown number of rounds.

**This file is the short answer. It says what is believed today and nothing about
how it was reached.** Every claim marked ✓ is checked against the archive on
every `npm run rounds` — see `scripts/claims.mjs`. If one goes STALE, the world
moved and this file is wrong.

---

## The host

- **✓ A chart on a freshly added slide groups and keeps its config.** 36 of 36.
  It was 1 of 74 on 2026-08-15; that figure is history and must not be quoted as
  current.
- **✓ The first chart of a multi-chart update costs ~2.2x a later chart of the
  same size.** Cause unresolved — position and slide-load are confounded in every
  archived sample, and four attempts to separate them failed.
- **The slowdown is entirely host-side.** Our idle time between calls is 1ms in
  every round from fastest to slowest.
- **Group children are unreachable.** 548 throws, 0 answers. office-js#3014 is
  closed upstream; the fix has not reached this host.
- **The host is unversioned** (`0.0.0.0`) — but probe answers are a behavioural
  fingerprint and it has been flat across the whole archive.

## Running the loop

- **Rest 45+ minutes between rounds.** Nine rested rounds skipped zero scenarios;
  ten back-to-back skipped ten, starting at the fifth. Rest buys COMPLETENESS,
  not speed. The driver warns from session position 5.
- **The noise floor:** two identical rounds differ by **14% (IQR)** typically and
  **73%** at worst. Under 14% is noise. A pair sees a 2x effect and cannot see a
  30% one.
- **Never move HEAD while a round runs** — the driver refuses as `site-behind`.
- **Never run the gate while a round is starting** — spawn starvation killed
  round 197.

## Open

- **✓ No chart has ever reached an update carrying a parts list** (0 of 1029).
  Blocked by the 5010 stale-proxy bug, office-js#2903, closed as not-planned. It
  causes 1197 redraws. If this claim ever goes stale, that is GOOD NEWS.
- **✓ Tag faults are zero** across the last 20 rounds. Thread 1 therefore cannot
  be tested — both arms of any experiment would score zero. It needs a trigger
  that reproduces tag loss on demand.
- **✓ The `no-queue` trace is dead** (last round 065). Thread 3 as written is
  closed; its live successor is a sparse single-chart `tagging failed`.
- **Slide occupancy is now read** before the deck-wide rescale. Unsampled — the
  next pair settles the position-vs-load question.

## Habits that earned their place

- **Every difference gets its sample size and its spread before it gets a
  sentence.** Two claims died on their error bars on 2026-08-25 and both would
  have been caught by one IQR column.
- **Read the code that writes a field before grouping by it.** A proxy renamed
  for being a bad proxy was then used as if it were not, twice in one day.
- **Mutate every new test.** Five tests passed against the code they were written
  to catch, in one day, and no gate caught any of them.
- **A correction sweeps its own section.** Five stale statements survived four
  hours directly above the correction that replaced them.
