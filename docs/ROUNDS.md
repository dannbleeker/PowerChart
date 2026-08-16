# Real-host rounds: what to do with one

A round is the only instrument this project has for the thing it cannot test — a real PowerPoint,
refusing real calls. Everything in `test/` asserts against a fake whose `sync()` returns instantly
and never fails, so a round is where facts come from. This is what to do with one.

## After every round, all five. It is a gate, not a menu.

**Write the outcomes down before reporting anything.** All five get an answer, including "nothing
found" — that is a result and it is worth recording. Skipping one is not visible unless it is
written, which is exactly how the two below went unnoticed for several rounds.

### 1. Mine the whole trace, not the headline

A round carries 380-440 entries and the pane summary names a fraction. The headline is never all of
it, and cross-referencing against earlier rounds is where the finding usually is.

Round 28 is the case for this. It PASSED `does a rasterise poison the next draw`, and in the same
round SKIPPED `the chart is actually visible` on "PowerPoint did not respond while drawing shapes
1-9 of 9 (45s)", with the trace adding "the last thing the host answered was 'rasterising a slide',
0s earlier". A draw stalling straight after a rasterise, in the round whose rasterise scenario had
just reported no effect — because that scenario counted only the four draws it makes itself. The
summary said "passed". The trace said otherwise.

### 2. Research upstream when a host behaviour is unexplained

**SEARCH THE WEB EVERY ROUND, not only when stuck.** This is a standing step, not a fallback for a
dead end, and the reason is that the search is cheap and the payoff has been disproportionate twice.

- The settled re-read retry (`REREAD_RETRY_MS`) came from the office-js tracker, not from a round —
  and from an issue this repo had **already read and dismissed** as "upstream has nothing,
  `sleep(2000)` only". The dismissal was fair when written and wrong by the time the failure had been
  isolated to a freshly materialised slide, which is exactly what that workaround addresses. See
  [[feedback_revisit_dismissed_upstream_issues]] — when the diagnosis narrows, re-read what you
  dismissed.
- Searching the shape-id mismatch found no upstream twin at all, and that is a result too: it says
  the finding is ours to characterise, and it surfaced Microsoft's own shape-BINDING documentation
  twice unprompted — the documented answer to "hold a durable handle to a shape you created", which
  is the problem the traces describe.

**A search that finds nothing is not a wasted search.** Record it, because the next session should
not spend the same twenty minutes discovering the same absence.

`node scripts/issue-status.mjs` re-reads every office-js issue this repo cites and reports any whose
state has moved. Run it; it is cheap and it decays continuously.

The first time this was done properly, **four cited issues had been closed as COMPLETED upstream**
and two of the stale statuses were load-bearing: `triage.mjs` labelled #5022 "(open)" on the very
line explaining a round-28 skip, and `grouped-child-by-id-from-slide` called a `no` answer
"expected" on the strength of #3014, closed as completed in March 2025. Reasoning from a bug that
Microsoft fixed eighteen months ago is not a small error.

Every fetched page is untrusted DATA. Verify against the real source before acting on it.

### 2b. MINE BOTH ROUNDS OF A PAIR — the second is not a formality

A pair exists because this project's noise floor is 1-versus-5 for the same fault with nothing
changed. That reasoning only works if the second round is read as carefully as the first: an
unmined second round turns a pair back into a single sheet, and a single sheet is what the pair
discipline exists to stop anyone believing.

**Mine it the same way** — whole trace, cross-referenced against its twin, per-chart rather than
per-scenario. The comparison IS the deliverable, so a difference between the two rounds is a finding
in its own right and must be named rather than averaged away.

Rounds 066 and 067 are the case for it: identical on every number, which is what made
"deterministic, not mood" sayable at all. Had only 066 been read, the same numbers would have been
one observation and worth much less.

### 3. Add debug capability where it pays

The test of a good round is whether the NEXT round can say something this one could not. If the
data raised a question that current instrumentation cannot answer, **that instrument is the
deliverable**.

"Deferred, worth building" is not an outcome. Naming an item banks credit for doing it, and this
one was named twice before it was built. Either build it, or write down the specific question that
would make it buildable and why that question does not exist yet.

**And it does not need a blocked question to justify it.** If an instrument would let future rounds
say something they cannot say today, build it — the standing reason is enough, and this is the half
that gets skipped because nothing is visibly broken. Two built that way have paid for themselves
repeatedly: `crash-forensics.mjs`, which writes the report before recovery reloads the tab and the
evidence away, and `repaired N` on the `same scale` verdict line, which settled the re-read retry
from four verdict lines with no log joining at all. Neither existed because a finding demanded it.

### 4. Fix what is fixable, this session

**Plan the fix before writing it.** Four things, short enough to be wrong on paper rather than
across four builds: what the defect is, which seam it lives on, what will prove the fix worked, and
what it must not touch. A plan that names no proof is a guess wearing a fix's clothes.

Worked example, and it is why this is written down. `tagAnchorIndex` moved the tag onto a different
handle, merged, and produced **no measured effect across five rounds and four builds** before it was
reverted. The prediction staked alongside it says why in its own words — the change was *"aimed one
level too low"*. What sits between the draw and the tag is a grouping attempt this host refuses, and
no choice of handle addresses that. Naming the seam would have caught it on paper; five rounds
caught it instead.

**A defect the current evidence cannot reach is NOT deferred, and does not get a speculative fix
either — it becomes an instrument.** That is its fix for this session: build the output that would
let the NEXT round decide it, and write in the journal which reading settles it which way. Guessing
at a fix nothing can judge is how a change survives four builds without anyone being able to say it
did nothing. See §3 — for a blocked item the instrument is the *whole* deliverable.

### 5. Correct the DOCTRINE, not just the code

The one that is easiest to skip and costs the most. If a round contradicts something written in
`CLAUDE.md`, in `docs/`, or in a source comment, **the written claim is a defect** and it gets
fixed in the same session as the code.

Worked example, because it cost a whole branch. `host-probe.ts` carried "needs a shape, positions
1-8: 77%; positions 9+: 47%" as doctrine, and three separate arguments were built on it — that two
probes starve because they sit at 22 and 23. Recomputed per attempt over two rounds it is flat, and
in round 26 it INVERTS (55% against 62%), because round 26 answered #31, the last question on the
sheet, while #8, #16, #22 and #23 all starved. Position was the correlate. The cause was the slot
above each one burning the scratch slide — see `Probe.burnsTheSlide`. A branch moving those probes
to positions 5 and 6 was built, run, and reverted.

## Keep the round

`rounds/` — see the README there. `poolRasteriseArms` and `poolEveryDraw` need many rounds and had
none for the entire life of the project, because round files lived in a temp directory and died
with the session.

    npm run rounds        # every round, pooled

## Predictions

A change made because of a round should say what the NEXT round will show, and the next round
should judge it. #468 predicted three questions would answer for the first time in eleven rounds;
round 27 said no; #469 recorded the failure rather than leaving a claim the data contradicts.
**A prediction that cannot fail is not worth staking.**

`rounds/predictions.json` is the ledger and `npm run rounds` judges the open ones. An entry names
the build it was made on, the claim in a form a machine can check, and — the half worth keeping —
`because`, the reasoning. Four claim kinds: `probe-answers`, `probe-starves`,
`scenario-passes`, `probe-detail-matches`.

Two rules the judge enforces, both learned by getting them wrong:

- **A prediction is never judged against the round that prompted it.** The change it predicts about
  is not in that round's build, so judging there fails every prediction the moment it is written.
  Rounds are ordered oldest-first and only rounds AFTER the prompting one count.
- **Never-put is `undetermined`, not `FAILED`.** A question the host declined to be asked has not
  refuted anything, and blaming the prediction for the refusal is the same class of error as
  reading `no-scratch-slide` as an answer.

Predicting that something will NOT change is worth staking too: it is what shows a change was
scoped, and #470's "the other two stay blocked" is the reason its result reads as a controlled
one rather than a coincidence.

## Verdicts that oscillate are noise, not regressions

`npm run rounds` prints every scenario's verdict per round and flags the ones that have said both
pass and fail about code that did not change between them. `explode a degraded picture` reads
`FAIL pass FAIL FAIL` — read one of those as a regression and you hunt a bug that is not there.

**A skip is not a flip.** `the chart is actually visible` reads `pass pass pass skip`: it has never
disagreed with itself, the host simply stopped answering during the fourth. Sometimes-unmeasured
and genuinely-contradictory are different facts and the report keeps them apart.

## Starting from nothing

    . scripts/pw.sh                                   # pw / paneref / pwclean
    pw open "https://onedrive.live.com/"              # sign in, open the deck,
                                                      # sideload from Home > Add-ins
    node scripts/round.mjs --check --dir .pw-session
    node scripts/round.mjs --dir .pw-session --retry 6

`scripts/pw.sh` exists because those helpers lived in `/tmp` for a day and a
half, rebuilt by hand every session, carrying paths keyed to whichever agent
session created them. Archive that session and the browser, its profile and every
helper went with it.

**The session directory has to be stable, and that is not a detail.** The CLI
daemon keys a browser by the working-directory STRING, so a browser is only
findable from the exact path that opened it. Parking it in an agent scratchpad
meant the next session looked in a new directory, got `(no browsers)`, and could
not reach a browser sitting on screen. `.pw-session/` in the repo root,
gitignored, is reachable from any shell on the machine — and it is what `--dir`
should be given.

Sign-in is the owner's, always: `--check` names that state on its own and says
so.

## Driving a round

The owner drives PowerPoint and the agent fixes fallout — but a round CAN be driven end to end from
here now; see `CLAUDE.md`, "Looking at the task pane", for the browser mechanics and the four traps
(Add-ins lives on the HOME ribbon, the file chooser is swallowed, the pane is two iframes deep so a
plain click does nothing, refs go stale on every DOM change).

Before starting, check the pane's build stamp is the commit you mean to test. PowerPoint caches the
pane HTML for ten minutes and a whole round can otherwise test code the host never fetched.

    npm run round -- --check    # preconditions only, nothing driven
    npm run round               # check, run, poll, archive, triage

`--check` refuses on a stale pane, an unpublished build, a dirty deck, either toggle off, a host that
is not answering, a host that answers but will not resolve slide 1 — and, ahead of all of them, on a
PowerPoint that has crashed or on there being no browser at all.

**Most of that list is now the driver's to fix rather than yours.** With `--retry N` it recovers a
crash, a silent host, a refused slide, a closed or stale pane, a dirty deck, a round that wedged
mid-flight, and a browser process that died — reopening the last from the persistent profile, which
still holds the sign-in. What it will NOT retry is a stop a reload cannot clear: a build Pages has
not published yet, or a pane toggle someone deliberately turned off. Every refusal carries a code,
and `RECOVERABLE_STOPS` in `scripts/round.mjs` is the list — derived from what `recover` actually
does, not from a judgement about which refusals feel transient.

## A browser death takes the sideload with it

**This is the one failure a night cannot recover from, and it looks like an
ordinary closed pane.** `recover` handles a dead browser well — the persistent
profile keeps the sign-in, so it reopens the window, finds the deck by name and
fronts it, all without a password. What it cannot restore is the ADD-IN. A web
sideload does not survive the browser process, so the deck comes back with its
slides intact, its ribbon showing `Add-ins`, and no PowerChart command anywhere.

Observed on 2026-08-16: the browser died mid-round, recovery brought back
`Presentation64` perfectly, and every subsequent attempt refused with "could not
read the pane's build stamp — is the add-in open?" The add-in was not closed. It
was gone.

The driver now names it `addin-missing` and **stops on the first attempt**,
because retrying is what the previous behaviour did seven times for nothing. It
is deliberately outside `RECOVERABLE_STOPS`.

Putting it back is a person's job: **Add-ins ▸ Upload My Add-in ▸ Browse** and
the repo's manifest. Nothing in this driver will do it, and nothing should try
unattended — the flow ends in a modal that a failed attempt would leave sitting
over the document for the rest of the night.

**The check requires a readable slide list before it fires.** A tab that is
merely mid-reload answers nothing to every read, so its ribbon looks exactly as
bare as a document with no add-in — and since this refusal is not retried,
firing it on a loading tab would end a night on a state that clears itself in
twenty seconds.

## What stops the driver hanging

Three bounds, and they are not interchangeable. The round's 30-minute deadline is
checked at the TOP of each poll, so it can only fire **between** calls — it
bounds a slow round, never a stuck one. Each `playwright-cli` call therefore
carries its own `timeout` (`PW_CLI_TIMEOUT_MS`, three minutes by default),
because a CLI that never returns leaves that deadline permanently out of reach.
And the page-side `budgetMs` inside an `eval` bounds only the *page*: it does
nothing if the browser connection itself is gone.

A timed-out call is `unreachable`, which is the truth — nothing was measured.

**An unexpected exception is a reason like any other.** `attempt` used to throw
straight out of the process: no receipt, no retry, and a night with six attempts
left ending on its first surprise. It now comes back as `threw`, gets the same
`--retry` a crash does, and lands in the receipt with its message. Retrying is
bounded by `--retry`, so a deterministic bug fails that many times and stops.

## The wedge

**The host's session is gone, and a reload brings it back.** Rounds 24, 25, 29 and 30 each died the
same way — `listing the deck's slides`, 90s timeout — and each cost most of an hour to reach that
point. The cause was never in this add-in, and it was never worth waiting out. It comes in two forms,
seen an hour apart on 2026-08-14, and both clear the same way.

**Form 1 — PowerPoint has crashed.** It raises its own error and puts up a modal:

> **Microsoft PowerPoint** — Sorry, we ran into a problem. Please try again. \[Refresh]

Its own ULS log, which the browser ships to `RemoteUls.ashx` and `playwright-cli request-body` can
read straight out of the tab, names it exactly:

    OnServerFindSucceeded could not find target slide, time elapsed: 449 ms
    GlobalErrorHandler:DisplayErrorDialog: 5341289
    ErrorDialog::ShowErrorDialog BSQMErrorCode: 5341289; ErrorName: errorLocalChangeLostSingleUser

Every crash on record shows the same three lines, with `ActionName=ExecuteAddinBatchOperation` — an
add-in `context.sync()` — as the action in flight and `MergeChanges` merging inbound revisions
alongside it. PowerPoint held a local change whose target slide the server could not find, decided
the change was lost, and gave up on the session. `errorLocalChangeLostSingleUser` is not documented
anywhere public; the log is the only source.

**The calls in flight are READS.** Three crashes logged the same `PptApi Call` sequence immediately
before the failure, and there is no `addGeometricShape` anywhere in it:

    PptApi Call - Presentation.GetSlides
    PptApi Call - Slide.GetId  (×2)
    PptApi Call - SlideCollection.GetItemOrNullObject
    …the block repeats, then…
    Failed to restore selection after load content.
    ReplicateOutbound → FindCommentRequest → UpdateNextPopulatedContextDetails
    OnServerFindSucceeded could not find target slide

That is the probe RESOLVING its scratch slide, not writing to it. An earlier reading of this file
blamed the round's first write, because that is where the trace stops; the host's own log says the
write never happened. `Failed to restore selection after load content` appears in all three, which
points at the slide the view sits on going away underneath it.

Four rounds have now died at the same question, `shape-add-fresh-slide-proxy`. In three of them the
recorded answer is `no-scratch-slide` after a 90s timeout — the probe never got as far as asking. The
question is where the trace stops, not what stopped it.

**Form 2 — the network moved under it.** No dialog, a document that looks perfectly normal in a
screenshot, and a host that will not answer. The console carries `net::ERR_NETWORK_CHANGED` (five of
them, plus `ERR_CONNECTION_REFUSED`), and the tab has fetched `RemoteSessionTermination.ashx`: the
editing session was severed and nothing reconnected it. The first crash had the same fingerprint just
before it — `ERR_CONNECTION_REFUSED` and `ERR_QUIC_PROTOCOL_ERROR.QUIC_NETWORK_IDLE_TIMEOUT` across
several hosts at once — which reads as form 2 turning into form 1 when the add-in next tries to
commit something.

**Either way, every Office.js call hangs forever and none of them throws** — including a
`context.sync()` with nothing queued. `PowerPoint.run` still ENTERS its callback, which is exactly
why it reads as a host thinking rather than a host that is gone, and why the round dies at whatever
call happens to be next rather than where the damage was. Both forms measured through the pane, the
same answer both times:

    Office:object | PowerPoint:object | ctx:yes | host:PowerPoint | api1.5:true
    entered:true | synced:false | threw: | emptySync:false

In form 1 the document UI is frozen behind the modal, so the deck reads back as `?` in the same
breath — that pairing is a tell. In form 2 the UI renders normally and tells you nothing.

**Reloading the document clears both, in seconds.** Form 1: click Refresh in PowerPoint's own dialog
— no answer in 8002ms before, 13–66ms after, on a session that had been dead for hours. Form 2: an
ordinary tab reload — 8011ms silent before, **7ms** after. The pane closes with the document either
way, so reopen it from Home ▸ Add-ins ▸ Insert chart.

Earlier revisions of this file said to wait it out and specifically not to reload, "because reloading
has never cleared it". That was written the same morning, before anyone had looked at the document,
and both halves of it were wrong.

### What has been ruled out

The add-in's most suspicious calls were replayed by hand against a healthy host, each on a loop, with
a `getCount()` ping between iterations. Every one left the host answering in tens of milliseconds:

| replayed                                                              |     | result                             |
| --------------------------------------------------------------------- | --- | ---------------------------------- |
| `shape-add-held-slide-proxy` — writing through a proxy resolved a sync earlier | ×8  | `GeneralException` every time, host fine |
| add a slide, add a shape through a same-sync proxy, list the deck      | ×6  | all `yes`, host fine               |
| write to a fresh slide then delete it in the same batch                | ×6  | host fine                          |
| the same in separate batches, deck oscillating 1↔2                     | ×10 | host fine                          |

So the trigger is not a single illegal call, and specifically **not** the `burnsTheSlide` question the
timing first suggested. Nor is it anything about the round: deck size at start, build, browser-session
age and time of day were each checked across the healthy and wedged rounds and none of them separates
the two. The variable that does track it is the machine's network, which is not something a round
controls and not something the add-in can be blamed for.

That leaves one thing genuinely open — what makes form 1 fire rather than form 2, i.e. why one severed
session raises `errorLocalChangeLostSingleUser` and another just goes quiet. What is settled: what the
wedge IS, that it is not the add-in, that it costs seconds rather than an hour to spot, and that a
reload clears it.

**The method mattered more than any of it.** Five earlier attempts reasoned from probe traces and were
wrong every time. What answered it in one session was looking at the thing itself: a screenshot showed
the dialog, `playwright-cli console` named the failing request, and `request-body` pulled PowerPoint's
own ULS log — which says, in Microsoft's words, what went wrong. That log ships from the tab on every
round; it had been there the whole time.

### The evidence keeps itself now

`scripts/crash-forensics.mjs` runs the moment the driver sees a wedge — all three
exits, the dialog, the silent pane and the thirty-minute timeout — and writes
`crashes/<timestamp>.md` before recovery reloads the tab.

That ordering is the whole point. The console log, the request list and the ULS
batches live in the tab, and reloading it is the first thing recovery does, so
until now the only copies of PowerPoint's own account of a crash were three
hand-typed passes in a chat log. Each cost about a quarter of an hour to reach
the same three lines.

The report holds the console errors, where the document's data channel stopped
(`GetPopWacUpdates` going quiet is what separates a dead session from a slow
one), and the thirty ULS lines before the fatal entry — which is where the
`PptApi Call` sequence in flight shows up.

Every read is guarded and named in the report rather than thrown: it runs on a
host that has just died, half the reads are expected to fail, and a forensics
pass that took the driver down with it would be worse than none. An absent fatal
entry is itself a finding — that is what the quiet form looks like.

`crashes/` is deliberately not `rounds/`. Everything downstream pools that
directory, and a crash report is not a round.

### Asking cheaply

`--check` pings `slides.getCount()` against an 8s budget before anything else, and looks for the
dialog directly, so a crashed host is named rather than waited on. The probe asks from the inside
too: `opened: { ms, answered }` on the sheet times the same call before question 1. It is a NUMBER,
not a verdict, deliberately — three healthy and three wedged rounds are not enough to set a
threshold, and a made-up one would turn a real measurement into a guess wearing a verdict's clothes.

## The nightly cycle: 16:9 twice, 4:3 once

**Two 16:9 rounds on one build, then one 4:3 round.** The pair is the
measurement — this project's noise floor is 1-versus-5 for the same fault with
nothing changed, so a claim needs two rounds that agree. **4:3 is validation, not
measurement**: one round, asking only whether anything behaves differently at a
slide size nothing else covers.

    npm run cycle

runs exactly that — two 16:9 legs then one 4:3, the gate after each round, and
it stops the night at the round that caused a problem rather than three rounds
later against a changed deck. By hand it is:

    PW_EXPECT_SIZE=16:9 node scripts/round.mjs --dir .pw-session --retry 6   # ×2, same build
    PW_DECK=<4:3 deck> PW_EXPECT_SIZE=4:3 node scripts/round.mjs --dir .pw-session --retry 6

**What the cycle will not do.** It does not judge whether a round found
something, it does not decide whether a divergence is real, and it never sets a
slide size — that would change what the round measures rather than restore it,
which is the one thing recovery is forbidden to do. It does not retry: by the
time it sees a refusal the driver has already exhausted `--retry`, and a second
implementation of "is this worth another attempt" is one too many. And it goes
nowhere near sign-in, which needs a password.

It stops on four things, and they want different responses: a **regression**
(read it now — that is the one fatal check), a **refusal recovery does not
address** (a hand on the machine), **no receipt at all** (the driver did not
reach the end of its own run), and **a gate that could not judge** (an archive
it cannot read — nothing was checked, and that is not a regression). A round
whose scenarios merely FAILED is not one of them — that is the measurement
working, and stopping there would throw away the second half of the pair.

**The gate's exit code carries that distinction: 1 is a regression, 2 is the
gate saying it could not do its job.** They must not be folded together. `archive`
writes straight to the final path rather than writing-then-renaming, so an
interrupted write leaves a truncated round behind — and before the gate had its
own code, meeting one threw a `SyntaxError`, node exited 1, and the night
stopped blaming the build for a fall that never happened. A round that will not
parse is now named and read past, loudly, because a round left out of the
comparison is a round whose regression cannot be seen.

**`PW_EXPECT_SIZE` is not optional in a cycle.** Without it the round does not
check the deck's size, and a deck can be the wrong one silently: setting
Widescreen on 2026-08-16 was accepted while the document was loading and did
nothing at all, caught only by reopening the menu. A round filed under the wrong
profile is worse than a round not run.

### When 4:3 disagrees

`npm run rounds:gate` answers several different questions and they must not be
confused. **Exactly one of them is fatal.**

- **A REGRESSION** — a scenario fell against its OWN profile's history. Fatal,
  exit 1. Judged only against rounds at the same slide size, because a 4:3 round
  measured against three 16:9 rounds would be flagged for scoring differently,
  which it does by design.
- **A DIVERGENCE** — a scenario passed at one slide size and failed at another,
  on the same build. **Reported, never fatal.** The response is to run 4:3
  again, or as a pair, before treating the difference as a property of the slide
  size at all.
- **UNSTABLE WITHIN** — a scenario passed AND failed at the same slide size on
  the same build. Not divergence: a profile that disagrees with itself has said
  nothing about its aspect ratio. Sending someone to investigate 4:3 for a
  scenario that is merely flaky is how a useful report teaches people to ignore
  it.

That escalation is the whole point of 4:3 being a single round: it is cheap
enough to run every night, and a disagreement buys a second round rather than a
conclusion.

### What the trace said that the archive has not

The three above are about the thirteen named scenarios. The gate also reads the
TRACE, which is ~95K characters per round and which nobody can count by eye —
round 081 was 512 entries collapsing to 44 distinct shapes, **none of them new**
against 57 prior rounds, and the archive's entire vocabulary is 81 signatures.
Reading all of it to rediscover that is the cost this exists to cut.

**It does not replace reading the trace.** It says where to start. Three buckets,
none fatal:

- **NOVEL** — a shape the archive has never produced. Read the trace; this is the
  reason to.
- **NEW BEHAVIOUR** — seen rarely before, common now. Usually a mechanism that
  just started working: rounds 079 and 081 both report the settle-delay re-read
  11 times against a median of 0, because `needsPreGroupRefresh` widened which
  charts reach a re-read that already existed.
- **A SPIKE** — a signature that had a baseline and left it.

The split between the last two is load-bearing. Collapsing them would mean every
fix this project lands announces itself as a fault on the night it works, which
is precisely how a report gets ignored and then switched off.

### Every round before 2026-08-16 was 16:9

Fifty-three of them, and none carries a slide size because the field did not
exist. **Anything reading a round must default to 16:9 rather than guess**;
`roundProfile` does, and this sentence is why it is allowed to.
