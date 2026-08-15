/**
 * What the fake claims, what a real host said, and the diff between them —
 * with no Node in it.
 *
 * Split out of `host-diff.mjs` so three callers can share one copy: the CLI,
 * the CI gate, and the TASK PANE. The pane is why the split exists — it can now
 * tell the owner whether a probe run found anything before they send the file,
 * and it could not import a module that opens with `import ... from "fs"`.
 *
 * Nothing here may reach for the filesystem, the network, or `process`. The
 * moment it does, the pane stops building and the reason will not be obvious.
 */
/** The fake's frozen answers, as asserted in `test/host-probe.test.ts`. */
export const FAKE_BASELINE = {
  "load-isnullobject-populates": "unreadable",
  "load-id-populates-isnullobject": "yes",
  "getitemornullobject-missing": "null-object",
  "shape-add-fresh-slide-proxy": "yes",
  "shape-add-held-slide-proxy": "threw",
  // Its partner does the identical thing a moment later on another fresh slide.
  // The fake refuses held proxies consistently, so the pair AGREES here — which
  // is the "the host has a definite behaviour" outcome the question exists to
  // tell apart from a coin. A real host that answered these two differently in
  // one run would be saying the opposite, and that is the finding.
  "shape-add-held-slide-proxy-again": "threw",
  "shape-resolve-held-slide-proxy": "yes",
  "shape-add-fresh-getitem-slide": "yes",
  "shape-add-positional-slide-proxy": "yes",
  "shape-proxy-survives-one-sync": "yes",
  "shapes-items-count-honest": "at-least-5",
  "shapes-items-via-positional-slide": "at-least-5",
  "getcount-populates-same-sync": "yes",
  "tags-add-same-key-twice": "overwrites",
  "tags-on-fresh-shape": "yes",
  "tag-through-refetched-shape": "yes",
  "how-many-syncs-a-creation-handle-survives": "survives-8",
  // `yes` here is the fake saying its collection read did NOT poison the
  // creation handle — on the scratch slide, where the shape has no other
  // handles onto it. Production's does, which is the whole question; see
  // PENDING_QUESTIONS.
  "collection-read-poisons-the-creation-handle": "yes",
  "delete-then-lookup": "reports-gone",
  "addgroup-returns-usable": "yes",
  "group-children-via-getcount": "two",
  "group-reports-its-children": "two",
  "grouped-child-by-id-from-slide": "yes",
  "tag-on-group-survives": "yes",
  "binding-names-shape-later": "yes",
  "getitemat-past-end": "threw",
  "picture-then-shape-read": "yes",
  "group-of-existing-shape-readable": "2",
  "slide-layout-readable": "yes",
  "layouts-readable": "yes",
  "untrack-available": "no",
  // The fake models `untrack` on a SHAPE and not on a null-object slide, which
  // is why this pair is worth asking as a pair: the fake already behaves the
  // way the confound predicts, and only the real host can say whether it does
  // too. Written from what the fake actually answers, after a guess at `no` was
  // corrected by the gate.
  "untrack-available-on-shape": "yes",
  "scratch-slides-returned": "all",
};

/**
 * Divergences between the fake and the committed real-host sheet that are
 * EXPECTED, each with the reason it is allowed to stand.
 *
 * The point of the table is what it makes impossible. `test/host-contract.test.ts`
 * diffs the fake's baseline against `test/fixtures/host-answers-web.json` and
 * fails on any divergence that is not declared here — so a change to the fake
 * that contradicts a real PowerPoint stops being something a human discovers by
 * running PowerPoint, which is how every host bug in this project has been
 * found so far, and starts being a red test in two seconds.
 *
 * A divergence belongs here for exactly one of two reasons, and the note has to
 * say which:
 *
 * - the fake models a DIFFERENT host on purpose (its happy path is a host that
 *   behaves, which is what makes most tests readable), or
 * - the real answer is known to be about the probe rather than the host, and
 *   the question has been re-asked but not yet re-run.
 *
 * "We have not looked into it" is not a reason. An entry with a note like that
 * is a to-do wearing a passing test's clothes.
 */
export const KNOWN_DIVERGENCES = {
  "untrack-available-on-shape":
    "The fake models the host `untrack()` was designed for — one where a shape proxy carries the method — and PowerPoint on the " +
    "web is not that host. ANSWERED 2026-08-12 (`89675b6`, reproduced on `1789749`): `no`, asked of a proxy " +
    "`addGeometricShape` had just returned, which is the kind `renderShapesChunked` holds hundreds of. That removes the " +
    "confound its trigger carried — `untrack-available` asks a NULL-OBJECT slide proxy, the one kind most likely to lack any " +
    "method — so the `no` is about the platform rather than about the probe. " +
    "Microsoft's performance guidance names untracking as the remedy for our exact symptom (\"large batch operations may " +
    'generate a lot of proxy objects... a noticeable performance benefit when using large numbers of proxy objects"), so this ' +
    "closes that idea on evidence rather than leaving it as an omission in the draw path. Do not re-propose it. The fake keeps " +
    "the method because `untrack` is best-effort everywhere this repo calls it and a host that has it is a real host.",
  "binding-names-shape-later":
    "RETIRED AS AN IDEA, 2026-08-12 (`957aca0`), and the fake is left saying `yes` deliberately. The question was " +
    "whether `settleAndTagChart` could be handed a shape handle that never goes through `ShapeCollection.getItem(id)` — " +
    "every 5010 this host throws is at that call. A binding is made from the live Shape proxy inside the batch that " +
    "created it, so it needs neither an id round trip nor a collection read. The host answers `commit-threw`: the batch " +
    "carrying the binding is REJECTED (`ErrorPointer`), and it counts as an answer rather than as background noise " +
    "because the probe's control arm committed the same batch WITHOUT a binding seconds earlier and it landed. Twelve " +
    "attempts across nine rounds never reached the commit; this one did, twice in three passes. " +
    "The fake keeps `yes` because there is NO CALLER to protect — nothing in this repo makes a binding, and modelling " +
    "a batch-poisoning API would be " +
    "fiction with no caller to protect. What the divergence is FOR is the direction it points: the fake is the " +
    "optimistic one here, so anybody who reaches for bindings as the way out of the id refusals will find this entry " +
    "before they find out from a deck.",
  "load-isnullobject-populates":
    "The fake models the host `queueNullCheck` was written for, where loading the flag by name populates nothing. PowerPoint on the web does populate it. Both hosts are real; the workaround is harmless on this one rather than necessary.",
  "shape-proxy-survives-one-sync":
    "office-js#2903. The fake keeps proxies alive on its happy path so that ordinary tests read as tests rather than as stale-proxy exercises; `applyWebProfile` is where the refusal lives. The web host refuses them, and says so here.",
  "shapes-items-count-honest":
    "The fake's happy path answers a shape collection honestly, and `faults.hollowReads` is where the refusal lives. RE-ASKED AND ANSWERED 2026-08-08: `short-0`, items=0, with `getcount-populates-same-sync` answering `yes, value=8` in the same run. Same host, same minute: the count is right and the list is empty. That is the sharpest form this bug has taken, and it is what `slideShapeNames`' corroboration check exists to catch. The earlier WITHDRAWN note (the `items`-undefined answer being about the handle, not the collection) is settled by the partner below.",
  "tags-add-same-key-twice":
    "WITHDRAWN, awaiting a re-run. The real `other — value=undefined` was one shape proxy held across four syncs, not an opinion about tag keys.",
  "addgroup-returns-usable":
    "The fake's happy path hands back a group usable in the sync that made it. RE-ASKED 2026-08-08 and still `unreadable`, but read the detail before calling it settled: `members via same-batch`. The strict id route could not supply members — the collection answers empty, see above — so this asked with same-batch proxies, which is the weaker form of the question. What it establishes is that the weak form fails; the strong form has still never been put on this host.",
  "group-reports-its-children":
    "The fake's happy path lists a group's children. RE-ASKED 2026-08-08 with the load queued in the sync that MADE the group — the friendliest form there is — and this host still answered `threw`, \"The property 'items' is not available\": office-js#6363's signature. Its sibling `group-children-via-getcount` was added to ask the other way and answered `unreadable`. Both routes refused, so `contentShapes` returning UNKNOWN_CONTENT for a grouped slide is the permanent answer rather than a gap.",
  "tag-on-group-survives":
    "WITHDRAWN, awaiting a re-run. Taken at face value the real `no` says no chart in any deck is re-editable, which the same run disproves — its repair pass landed 23 retags on grouped charts. Still owed: the 2026-08-08 round could not put this question either (`no-scratch-slide`).",
  "group-children-via-getcount":
    "The fake's happy path counts a group's children; the web host's refusal lives in a named fault rather than the default. ANSWERED 2026-08-08: `unreadable`. Read with `group-reports-its-children` above — both ways into a group's children are refused on this host. Also carried in UNSTABLE_ANSWERS, because it has been asked once and once is a sample.",
  "group-of-existing-shape-readable":
    "The fake's happy path names a group it has just made, so the later-batch question can be put at all. This host would not: `no-group-id`. That is an answer and not a setup failure — a host that will not name a fresh group cannot be asked about resolving one from the deck afterwards, and the fact belongs in the sheet. It also means `countGroupChildrenPage`, which swallows failures per shape, produces no error and no measurement here.",
  "shape-add-fresh-getitem-slide":
    "The fake models the host `getTargetSlide` was written for, where `slides.getItem(id)` resolves any slide. On the web it answered `threw` (GeneralException) for a slide added moments earlier — and its follow-up partner `getitem-durable-slide` answered `yes` in the same run. So the two readings are separated: it is not the by-id form that fails, it is the by-id form applied to a NEW slide. A pre-existing slide's id round-trips fine, which is why editing a chart in place has always worked.",
  "shapes-items-via-positional-slide":
    "The fake's happy path answers a shape collection honestly whichever handle names the slide. ANSWERED 2026-08-08: `short-0`, the same as its by-id partner in the same run. That is what the partner was added to decide, and it decides it — the parent handle was never the problem, the COLLECTION is. Every readback in `powerpoint.ts` is therefore no better for being renamed positionally.",
};

/**
 * Questions this host has answered DIFFERENTLY on different runs of one build.
 *
 * Deliberately not a `KNOWN_DIVERGENCES` entry, and the contract gate rejecting
 * the first attempt at that is the reason this list exists. A divergence is
 * "the fake and the real host disagree", which is a fact about two systems and
 * is either true or it is not. This is a fact about ONE system: the real host
 * gave answer A, then gave answer B, minutes apart, same build. The gate has
 * nothing to say about that and should not be made to.
 *
 * What it is for is stopping the next reader — human or agent — from building
 * on whichever answer a sheet happens to carry. `shape-add-positional-slide-
 * proxy: yes` is exactly the answer that makes a positional slide handle look
 * like the safe route out of the by-id refusals, and it is one of the two that
 * flipped.
 *
 * The mechanism is in the run log rather than inferred: three `scratch slide
 * landed but its id will not resolve` lines mid-run, two replacement scratch
 * slides taken, and every question asked inside that window answering
 * `no-scratch-slide` before the host came back. The host's ability to resolve a
 * freshly added slide's id comes and goes within a single 37-second run — the
 * same reversible bimodality the draw times show.
 *
 * A question in here has been SAMPLED, not answered. Removing an entry needs
 * several runs agreeing, not one.
 */
export const UNSTABLE_ANSWERS = {
  "picture-then-shape-read":
    "office-js#5022's question — can a shape collection be read after a picture insert — and a COIN, seen flipping inside one " +
    "round on 2026-08-12 (`1789749`): `yes` on pass 1 while the host was in slide-trouble, then `unreadable` on passes 2 and 3 " +
    "once the collection was refusing. The committed fixture carries `unreadable`. " +
    'It is in this table because of what the other value would licence: a `yes` reads as "#5022 does not affect this host", ' +
    "and `drawDemoItem` performs exactly that sequence — insert a picture, then read shapes — so somebody would reasonably " +
    "stop guarding it. One pass in three is not permission. Note the regimes: both readings came from the same round minutes " +
    "apart, so this is the host moving rather than two builds disagreeing.",
  "shape-add-held-slide-proxy-again":
    "RETIRED FROM STABLE ON 2026-08-12 (`89675b6`), by the mechanism it exists to be. It came off `PENDING_QUESTIONS` on " +
    "`756682e` as stable across three passes; this round flipped it inside ONE round — `threw` on pass 1 with the host " +
    "healthy, `yes` on pass 2 in slide-trouble — while its TRIGGER answered `threw` both times and reported stable. So the " +
    "pair has been seen the other way round from the way this project describes it: the partner is the coin here and the " +
    "trigger held. Do not build on either value. Worth keeping as a pair rather than collapsing: a round where the two " +
    "disagree is exactly the evidence that this host's refusal of a held proxy is a state it moves through, not a rule.",
  "shapes-items-count-honest":
    "MEASURED AGAIN 2026-08-10: `unreadable` on the last SEVEN consecutive rounds (1fa0509 through 3d17165), not the three this " +
    "entry claimed. `short-0` has not appeared since 619d24b. The ANSWER never moved — this host will not tell a caller what is on " +
    "a slide — and the FORM now looks settled too, though `slideShapeList` still handles both and should, because nothing explains " +
    "why the form ever changed. Original note follows. " +
    "`short-0` again on 2026-08-09 (`619d24b`), against `unreadable` in the committed fixture — nine samples now, still alternating between the two forms with no trend. " +
    "The ANSWER is stable; its FORM is not. `unreadable` (2026-08-05), `short-0` (2026-08-08 on 2f1e8c4), `unreadable` again " +
    "(2026-08-08 on a546897), `short-0` again (2026-08-09 on 8bb9e8f), then `unreadable` on the last three (448ffc6, cfa1f50 and the round after it, which answered identically). Seven samples, two forms, no trend — though the last three held still, which is the longest it has. Every one of those says the same thing — this host will not tell a caller what is on a slide — and they " +
    "differ only in how the refusal arrives: a collection that throws, versus one that answers with zero items. Worth keeping apart " +
    "because the two want different code (a catch versus a corroborated count), and `slideShapeList` handles both for exactly this reason.",
  "shapes-items-via-positional-slide":
    "As its by-id partner above, and moving in step with it: `short-0` (2f1e8c4), then `not-listed` on every run since (a546897, d812d0c, 448ffc6, cfa1f50). Every run agrees with the " +
    "by-id form in the same run, which is the finding — the parent handle is not the variable. What varies is the host, run to run.",
  "group-children-via-getcount":
    "NO LONGER ONCE: eight of eight `unreadable` across the ten rounds to 2026-08-10 (the other two never put the question). " +
    'The hedge below — "two consistent answers from two routes is a strong hint, not a finding" — was right when it was written and ' +
    "is now overtaken: both routes into a group's children are refused, consistently, on every round that could ask. `contentShapes` " +
    "returning UNKNOWN_CONTENT for a grouped slide is the permanent answer. Original note follows. " +
    "ASKED AND ANSWERED ONCE, on a degraded host: `unreadable`, 2026-08-08. Its sibling `group-reports-its-children` answered `threw` " +
    "(\"The property 'items' is not available\") on a healthy round the same day, so BOTH routes into a group's children have now been " +
    "refused and `contentShapes` returning UNKNOWN_CONTENT for a grouped slide looks permanent rather than a gap. " +
    "Listed here rather than treated as settled because that round put only 17 of 27 questions — `getcount-populates-same-sync` itself " +
    "came back `no-scratch-slide` in it, having answered `yes, value=9` the round before. One sample from a host in that state is a " +
    "sample. Two consistent answers from two routes is a strong hint, not a finding.",
  "shape-add-held-slide-proxy":
    "THE `yes` PAIR HAS BEEN CAUGHT, AND THE REGIME IS NOT THE STATE — 2026-08-13 (`cd3b60c`). The note below " +
    "names its own honest limit: four agreeing pairs, all of them `threw` pairs, so what was shown was consistency " +
    "in the `threw` state rather than consistency in general, and a pair taken while the host answers `yes` 'cannot " +
    "be scheduled — it has to be caught'. This round caught it, and two more besides:\n" +
    "  pass 1   trigger `threw` 16.3s / partner `threw` 17.0s   (healthy)\n" +
    "  pass 2   trigger `yes`   33.9s / partner `yes`   34.4s   (collection-refused)   <- the missing one\n" +
    "  pass 3   trigger `threw` 55.6s / partner `threw` 56.8s   (collection-refused)\n" +
    "Seven pairs now, seven agreements, zero disagreements, and the `yes` state is paired half a second apart. A " +
    "fifty-fifty coin agrees half the time, so seven for seven is p=0.0078 against one. The reading the entry " +
    "already reached — at any instant this host has a DEFINITE answer, and the variation across a run is a state " +
    "changing rather than a coin landing — is now supported in BOTH states rather than one. " +
    "AND THE NEW HALF: `scripts/host-regimes.mjs` calls this question a COIN, correctly, because " +
    "`collection-refused` produced BOTH `yes` (pass 2) and `threw` (pass 3). Those two facts are not in tension — " +
    "together they say something neither says alone. There IS a definite state, it changes during a run, and the " +
    "probe's `regime` stamp is NOT that state. So the next question is not 'is it a coin' (it is not) but 'what is " +
    "the state', and `regime` has been eliminated as the answer. " +
    "The pair is also LOCKSTEP by that tool's other reading — trigger and partner move at the same pass boundary — " +
    "so they are one mechanism sampled twice, not two questions. Treat a flip in either as a flip in both. " +
    "Original notes follow. THE PARTNER HAS ANSWERED, AND IT IS NOT A COIN — 2026-08-11 (`756682e`). Three paired asks in one round, " +
    "TWO of them on later passes, and all three agreed:\n" +
    "  pass 1   trigger `threw` 37.8s  / partner `threw` 39.2s\n" +
    "  pass 2   trigger `threw` 89.0s  / partner `threw` 90.1s\n" +
    "  pass 3   trigger `threw` 127.5s / partner `threw` 129.1s\n" +
    "With the agreeing pair from `96461eb` that is four pairs, four agreements, zero disagreements. A fifty-fifty " +
    "coin agrees half the time, so four for four is p=0.0625 against it — not proof, and much stronger than anything " +
    "fifteen one-sample rounds could say. Read it as: at any given instant this host has a DEFINITE answer, and the " +
    "variation across a run is a state changing, not a coin landing. " +
    "The honest limit: all four pairs are `threw` pairs. Nothing yet pairs a `yes`, so what is shown is consistency " +
    "in the `threw` state rather than consistency in general. A pair taken while the host is answering `yes` is the " +
    "one still missing, and it cannot be scheduled — it has to be caught. " +
    "The later-pass tally moved too, and against the old story: this round answered `threw` on ALL THREE passes, so " +
    "later passes now stand at 3 x `yes` and 3 x `threw`. Pass 1 remains near-deterministic (18 of 19 `threw`). " +
    "Original notes follow. THE CLEAN SPLIT IS GONE — 2026-08-11 (`96461eb`), the round that first carried the partner. A LATER pass " +
    "answered `threw` for the first time (pass 3, 46.7s), so the tally below is no longer 3-of-3 either side:\n" +
    "  pass 1 : 18 x `threw`, 1 x `yes`   (19 observations)\n" +
    "  later  :  3 x `yes`,   1 x `threw` ( 4 observations)\n" +
    "Read that as: pass 1 is near-deterministic and the LATER passes are where the variability lives — which is a " +
    "different claim from two clean populations, and a better fit to every round on file. " +
    "THE PARTNER ANSWERED ONCE AND AGREED: trigger `threw` at 7.7s, `shape-add-held-slide-proxy-again` `threw` at " +
    "8.4s, seven tenths of a second apart on two different fresh slides. One agreeing pair is weak evidence against " +
    "a fast coin (a coin agrees half the time) and it is the LEAST informative pair available, because pass 1 is the " +
    "condition that barely varies. The pair worth having is on a LATER pass, and that round did not get one — the " +
    "pass-3 partner came back `no-scratch-slide`. The pair costs two scratch slides on the host least willing to " +
    "give them, which is the thing to fix if this is still open after another round or two. " +
    "Original notes follow. THE POPULATIONS WERE NEVER SEPARATED — 2026-08-11 (`7027f96`), and this supersedes both readings below. " +
    "Every observation this entry has ever recorded, sorted by WHICH PASS asked it:\n" +
    "  pass 1 (asked ~1-2s into the probe):  17 x `threw`, 1 x `yes`   (18 observations)\n" +
    "  any later pass (34s-80s in):           3 x `yes`,   0 x `threw` ( 3 observations)\n" +
    "The second row did not exist until 2026-08-11, because `PROBE_PASSES` shipped that day (780cf02) and every " +
    "round before it asked each question exactly ONCE. So the sixteen pre-3x observations are not sixteen samples " +
    "of a coin — they are sixteen samples of ONE condition, and the ten-in-a-row `threw` that the note below calls " +
    "a trend is ten rounds asking at the same moment and getting the same answer. Both 3x rounds agree: `threw` " +
    "first, `yes` on every later pass (R12 threw@2.1s then yes@79.4s; R13 threw@1.1s then yes@34.4s and yes@47.7s). " +
    "DO NOT read that as settled. Three readings still fit and this round cannot choose between them: the age of " +
    "the RUN, the pass number itself, or the state of the SCRATCH SLIDE — and the third is live, because the run " +
    "log shows this question WRECKS its own scratch slide every time it is asked (`giving up the scratch slide " +
    "this question wrecked`, three times in R13), so pass 1 meets a deck with no scratch history and later passes " +
    "do not. What is settled is that the variable is not the round, which is what fifteen one-sample rounds were " +
    "implicitly testing. The partner question that separates run-age from slide-state is the `Probe.follow` case " +
    "to write next: ask it twice in quick succession LATE in a run, once on a fresh scratch slide and once on one " +
    "the run has already used. One answer to each, one round, no reasoning. " +
    "The regime stamps ARE readable now (`regimeFrom` shipped in #390 and R13 is the first round on a build " +
    "carrying it — 24/19/6 across three regimes, moving back and forth, against 55-of-65-identical before). They " +
    "say `healthy` for the `threw` and `slide-trouble` for both `yes`. That is one round; it is consistent with " +
    "the scratch-slide reading and does not establish it. Original notes follow. SETTLED 2026-08-11 (`3223293`), and it is a coin after all. The first round to ask every question three times " +
    "got BOTH answers out of one host, one tab, 77 seconds apart: `threw` at 2.1s and `yes` at 79.4s, run never " +
    "restarted, build never changed. That is the observation this entry has wanted since the single `yes` of " +
    "2026-08-08, and it retires the reading below: a mechanism with one early outlier cannot reproduce its outlier " +
    "on demand inside one run. It is not a claim of fifty-fifty — two samples are two samples — but the variation is " +
    "WITHIN a session, so re-running whole rounds was never going to settle it, and the fifteen observations below " +
    "were fifteen one-sample rounds rather than a trend. " +
    "The regime stamped on those two samples (`healthy` for the `threw`, `collection-refused` for the `yes`) is NOT " +
    "yet evidence: the flag it came from latched 8.9s into a 110s probe, so `collection-refused` may mean no more " +
    "than `later`. Fixed the same day (`regimeFrom`), so the next round's stamps can be read. Original note follows. " +
    "MEASURED AGAIN 2026-08-10 over ten consecutive rounds (`scripts/host-history.mjs`): `threw` TEN times out of ten, " +
    "551ad42 through 3d17165. That makes fifteen of sixteen observations, with the single `yes` now eleven rounds and two " +
    "days behind — this entry called it a coin, and a coin does not do that. Read it as a mechanism with one early outlier " +
    "rather than a fifty-fifty, and keep the entry only because the outlier has never been explained. Original note follows. " +
    "ALTERNATES. Six observations: `threw` (2026-08-05), `threw` (2026-08-07), `yes` (2026-08-08 run a), `threw` (2026-08-08 run b), " +
    "`threw` (2026-08-08 run c, the sheet now committed), `threw` (2026-08-09 on d812d0c). Five of the six are `threw`. Earlier wording here said it flipped once, which " +
    "reads as though the newer value were the true one and the old one a mistake — it is not a sequence of corrections, it is a coin, and " +
    "five of six landings do not make the seventh a mechanism. " +
    "The fake keeps refusing held proxies, which is the safe direction: code that never holds one across a sync is correct whichever way the coin lands.",
  "shape-add-positional-slide-proxy":
    "A REGIME MAPPING THAT DID NOT REPRODUCE — 2026-08-13 (`cd3b60c`), and the failure is the useful part. Run " +
    "against the COMMITTED fixture (`1789749`), `scripts/host-regimes.mjs` read this question as EXPLAINED by " +
    "regime: `yes` in `slide-trouble` twice and `not-listed` in `collection-refused`, a mapping with a way to come " +
    "out wrong that did not. One round later it is a COIN — `collection-refused` produced BOTH `not-listed` " +
    "(pass 1, 19.3s) and `yes` (pass 2, 38.7s) in the same regime.\n" +
    "So the mapping was one round's shape, not a mechanism, which is exactly what that tool's footer warns and why " +
    "an `explained` verdict is a reason to RE-READ an entry rather than to rewrite it. Had this entry been rewritten " +
    "as 'a degradation you can test for' on the strength of the first run, it would have been wrong inside a day. " +
    "Two rounds of `explained` is the bar; one is a reason to look again. " +
    "Original notes follow. MEASURED AGAIN 2026-08-10 over ten rounds: `yes` on the last five consecutively, `not-listed` on three before that, and `threw` " +
    "NOT ONCE. So the face that made this dangerous — `threw`, the one that says the question was put and refused — has not appeared " +
    "in ten rounds, while `yes` is on its longest run. That does NOT retire the warning: `yes` is still exactly the answer that makes " +
    "a positional slide handle look like a way around the by-id refusals, and `shapes-items-via-positional-slide` (genuinely " +
    "unstable, still alternating) is why it would not help anyway. Original note follows. " +
    "A ninth observation on 2026-08-09 (`619d24b`): `not-listed` again, the third face, on a run that answered 21 of 28 — so it is not a symptom of a badly degraded round either. " +
    "THREE-SIDED, which is worth saying because this entry called it a coin. Eight observations: `yes`, `yes`, `threw`, `yes`, `yes`, " +
    "`threw`, `not-listed` (2026-08-08, `1fd6aa3`), `yes` (2026-08-09, `d812d0c`). The third face is not a variant of the other two — `not-listed` means the DECK'S " +
    "OWN SLIDE LIST did not contain the scratch slide's id, so the question was never really put, while `threw` means it was put and " +
    "refused. A host that will not list a slide it has just added is the same story as everything else here, arriving at a different door. " +
    "For the first six observations it did alternate in lockstep with its partner above and always opposite to it: `yes`, `yes`, `threw`, `yes`, `yes`, `threw`. " +
    "The more dangerous of the two, because `yes` is exactly the answer that makes a positional slide handle look like a way around the " +
    "by-id refusals — and four of the five samples say `yes`. A majority is not a mechanism. Whatever decides these two flips within a " +
    "single run (see the `no-scratch-slide` windows in any probe log), and until that is understood neither answer may be built on. " +
    "`shapes-items-via-positional-slide` is the reason it would not help anyway: a positional handle reads a shape collection exactly as " +
    "short-0 as a by-id one does.",
};

/**
 * Questions the committed real-host sheet cannot answer, because they were
 * added after it was taken.
 *
 * Declared for the same reason divergences are: a hole in the comparison is
 * worth seeing. `host-contract.test.ts` fails on an UNdeclared hole, so a probe
 * added without a re-run cannot quietly reduce what the gate covers, and fails
 * on a stale entry too, so the list shrinks the moment a newer sheet arrives.
 *
 * Every entry here is a reason to ask the owner for a probe run.
 */
export const PENDING_QUESTIONS = {
  "collection-read-poisons-the-creation-handle":
    "Added 2026-08-15, and it is the question the ordering fix is now blocked on — the fix was BUILT that day and the " +
    "Office.js fake refused it. `survives-8` says a creation handle keeps taking tag writes for at least eight syncs, and " +
    "`tag-through-refetched-shape: no-id` says there is no id to re-fetch one by, so the fix looked settled: make the tag " +
    "anchor a shape the draw loop never `load()`s and the write goes through. Production does one thing neither probe " +
    "does — `groupAndTagAll` re-reads the whole slide's shape collection before grouping, because grouping needs fresh " +
    "handles — and in the fake that read marks the shape resolved for EVERY handle onto it, creation handle included, so " +
    "holding the anchor's own load back changes nothing and the write is refused anyway. Whether the real host works " +
    "that way is unknown. Office.js gives each proxy its own object path, and the fake itself takes that view everywhere " +
    "else: a fresh handle gets its own `syncCreated` and its own tag writer, sharing only the shape's state. " +
    "`loadedProps` is handle state modelled as shape state, which is either a bug in the fake or the one place it is " +
    "right. `yes` means the collection read is innocent, the fake is wrong, and the ordering fix works — build it. " +
    "`refused` means no arrangement of loads saves the drawing context's write while grouping needs a re-read, and the " +
    "fix has to be a second tag key or nothing.",
  "how-many-syncs-a-creation-handle-survives":
    "Added 2026-08-15 after round 037 carried the new `from` field and answered the question it was built for: EVERY tag " +
    "failure in that round went through a `created` handle — seven batch-level and four per-chart, not one `refreshed` " +
    "or `by-id`. So swapping the tag target changes nothing and the ordering is what has to change. Four of those " +
    "failures were also a different fault from the 5010: `Cannot read properties of undefined (reading 'add')`, i.e. " +
    "`.tags` GONE rather than refused. `tags-on-fresh-shape` asks in the creating batch and answers yes every round; " +
    "`tag-the-creation-proxy-a-sync-later` asks one sync later and also answers yes. Production's handles are older " +
    "than either — the renderer chunks a chart across batches — so this asks HOW LONG one lasts. The answer is the " +
    "budget an ordering fix gets built against, and guessing it is what would make that change a gamble. The fake " +
    "answers `survives-8` because nothing in it ages a handle this way; a real number below 8 is the finding.",
  "tag-through-refetched-shape":
    "Added 2026-08-14 after rounds 29 and 30 failed `same scale across the deck` identically: `InvalidParam passed to " +
    "GetItem(id)` (5010) at `writing the chart's config tag`, charts left with no config, the scenario stopping on the " +
    "second consecutive loss. That is the exact path `finishCharts` takes — write the tag through " +
    "`shapes.getItemOrNullObject(id)` with an id read off a shape created a sync earlier. `tags-on-fresh-shape` already " +
    "says the fresh shape's own `.tags` works, every round, so the id round trip is the untested half. A question rather " +
    "than a rewrite ON PURPOSE: this path has a history of changes reverted on a theory. `threw` means stop re-fetching " +
    "and tag the creation proxy; `yes` means the 5010 comes from somewhere else and a rewrite would have been wasted.",
  "grouped-child-by-id-from-slide":
    "Added 2026-08-11, and it decides whether the in-place update has a future on this host. `tryInPlaceUpdate` needs a " +
    "node-to-shape mapping and gets one from CHART_PARTS_TAG, which is written only for UNGROUPED charts — so the " +
    "`53ec985` round declined ELEVEN times out of eleven with `the chart has no parts list`, the grouped charts having no " +
    "tag and the ungrouped ones having had their id readback refused. But 'the parts tag does not list them' is a fact " +
    "about our code, not about the host. If the slide's own collection still resolves a child by id, a grouped chart can " +
    "carry a parts list written from ids the grouping pass already holds, and the fast path applies to the fourteen " +
    "grouped charts a round produces instead of none. office-js#3014 SAID sub-shapes cannot be reached, and a no used to " +
    "be called expected on that basis — but #3014 was closed as COMPLETED on 2025-03-03 (GitHub API, checked 2026-08-14), " +
    "so neither answer is expected any more: a yes means the upstream fix reached this host, a no that it did not. Which " +
    "is why it has to be a measured answer rather than an assumed one, because the whole feature turns on it.",
  "shape-resolve-held-slide-proxy":
    "Added after the fixture's build. It decides whether `deleteShapesById`, `setShapeSelection` and the selection path are bugs or merely untidy: all three resolve a slide, sync, then reach through that same handle for a shape. The write form of this is known to fail; the read form has never been asked, and the fake's windowed handle does not gate it either way.",
};

/**
 * What each divergence would MEAN, so the report says something actionable
 * rather than just "these differ".
 *
 * Keyed by probe id. Absent is fine — the diff still reports the divergence,
 * it just cannot say what rests on it.
 */
const WHAT_IT_MEANS = {
  "addgroup-returns-usable":
    "Whether a group can be used in the sync that made it. WITHDRAWN: the 2026-08-04 'unreadable' came from asking for the id one sync after the group was made, and from grouping members that were themselves a sync old — so it measured proxy age, which three other questions already establish. Members are resolved in the grouping batch now and the id is asked for in it.",
  "shapes-items-via-positional-slide":
    "The partner to the question above, and the only one of the four contaminated answers that could not be cleaned up by re-resolving: a collection load is queued in one batch and read in the next by definition. If this reads back and the by-id form does not, the collection was never the problem and the parent handle was — which decides how every readback in `powerpoint.ts` should name its slide.",
  "load-isnullobject-populates":
    "`queueNullCheck` loads 'id' instead of 'isNullObject' precisely because the flag cannot be loaded by name. If a real host populates it, that whole comment is wrong for this host — and the workaround is merely harmless rather than necessary. ANSWERED: PowerPoint on the web (2026-08-04) said yes, and read the flag back as false. The negative is host-specific; the workaround stays because the host it was written for is real too.",
  "load-id-populates-isnullobject":
    "If a real host does NOT populate the flag from a real property load, `queueNullCheck` does not work and every `isLive` check is answering 'not live' for live objects. `isLive` treats unreadable as NOT live, so the failure mode is refusing to act on slides that are fine.",
  "shape-add-fresh-slide-proxy":
    "Whether this host will take a shape at all on a slide added moments ago, asked through a slide proxy resolved in the same sync as the add. ANSWERED: PowerPoint on the web (build a609c9c) said YES. A new slide takes shapes perfectly well — it is never the slide's newness that fails.",
  "shape-add-held-slide-proxy":
    "The same add through a slide proxy resolved one sync earlier — what Office.js has by then rewritten to `slides.getItem(id)`. ANSWERED: THREW, GeneralException, while the fresh and positional forms of the same add both worked. So it is the HOLDING that fails, not the id and not the slide. The fake models this unconditionally now (`expiringSlideHandle`), which is why the baseline beside it reads `threw` — the one question here whose expected answer is a refusal.",
  "shape-add-positional-slide-proxy":
    "The third way to name the same slide. ANSWERED: YES — so by-index is not the fix for anything; by-id was never the problem. Kept as the control that makes the pair above readable.",
  "shape-proxy-survives-one-sync":
    "office-js#2903. The fake keeps proxies alive by default, which is the kindness that hid a whole class of stale-proxy bug until a human found it in a real host. If a real host refuses a one-sync-old proxy, `applyWebProfile` should be the default rather than a named profile. ANSWERED, sideways: the 2026-08-04 self-test run threw `InvalidParam passed to GetItem(id)` at `ShapeCollection.getItem` while grouping a chart's shapes, five charts in a row — so on that host the answer is no. The probe's own attempt never reached the question.",
  "shapes-items-count-honest":
    "`faults.hollowReads` models a host answering SHORT without throwing — a readback asked about 19 shapes and was told 3. WITHDRAWN: the 2026-08-04 answer (`items` undefined) was about the handle the collection hangs off, not the collection. A collection read cannot avoid crossing a sync — the load is queued in one batch and read in the next — so this one could not be cleaned up by re-resolving, and got a partner instead: `shapes-items-via-positional-slide` asks the same thing through a positional parent. Read the two together or neither.",
  "tags-add-same-key-twice":
    "Re-editing a chart rewrites POWERCHART_CONFIG on the same shape every time. If a host appends rather than overwrites, a chart edited ten times carries ten configs and the reader picks one arbitrarily. WITHDRAWN: the 2026-08-04 'other — value=undefined' was the probe holding one shape proxy across four syncs, not an opinion about tag keys. Every write now goes through a shape resolved in its own batch, so the next sheet's answer is the first real one.",
  "tags-on-fresh-shape":
    "`faults.tagsUndefinedOn` models `.tags` coming back undefined, where reading `.add` throws SYNCHRONOUSLY and escapes the tagging loop — losing the config for every chart after it in the batch, not just the one.",
  "tag-through-refetched-shape":
    "The production path `finishCharts` actually uses: write POWERCHART_CONFIG through `shapes.getItemOrNullObject(id)`, where the id was read off a shape created a sync earlier. Rounds 29 and 30 both failed `same scale across the deck` with `InvalidParam passed to GetItem(id)` (5010) at `writing the chart's config tag`. `tags-on-fresh-shape` says the fresh shape's `.tags` is usable; this asks whether the id round trip is what the host refuses. `threw` means stop re-fetching; `yes` means the 5010 is about something else.",
  "delete-then-lookup":
    "`deleteSlideById` re-checks from a FRESH context because the same-context answer was not trusted. If a host answers honestly here, that second round trip is removable.",
  "group-children-via-getcount":
    "Added 2026-08-08, in response to the sheet taken that day. Its sibling `group-reports-its-children` asked through `group/shapes/items/id` and this host answered `threw` — \"The property 'items' is not available\", office-js#6363's exact signature — with the load queued in the sync that MADE the group, which is the friendliest form the question has. But the same sheet says `getcount-populates-same-sync: yes, value=9`: this host COUNTS a shape collection it will not LIST. Nobody has asked whether that holds for a GROUP's collection, and the answer decides something concrete — `contentShapes` returns UNKNOWN_CONTENT for every grouped slide, which is what makes the reconcile report a slide complete without counting it. A count is all it needs.",
  "group-reports-its-children":
    "The single most load-bearing answer here. A chart IS a group, and the readback measures whether a chart survived by counting what is inside it — so a host that groups successfully and then reports no children makes every chart read back as wreckage, and the repair pass 'fixes' charts that were never broken. WITHDRAWN: the 2026-08-04 PropertyNotLoaded was a nested load queued a sync after the group was made. It is queued in the grouping batch now.",
  "tag-on-group-survives":
    "Where a chart's config actually lives. WITHDRAWN: the 2026-08-04 NO was the probe writing through a group proxy a sync old. Taken at face value it says no chart in any deck is re-editable, which the same run disproves — its repair pass landed 23 retags on grouped charts. The group's id is what crosses the sync now, and every use resolves its own handle.",
  "binding-names-shape-later":
    "Whether the repair pass can be given a handle that does not go through `ShapeCollection.getItem(id)`. Every 5010 this host throws is at that call, and it is what leaves a chart drawn and nameless — no group, no tag, nothing to settle one onto. A binding is made from the live Shape proxy inside the batch that created it, so it needs neither an id round trip nor a collection read, and the document persists it. A real `yes` means `settleAndTagChart` has a route it does not have today; a real `no` retires the idea. Watch the answer WORD, because five of them are different facts: `no-binding-api` is a missing 1.8 surface and says nothing about the idea; `add-threw` is `bindings.add` objecting on the spot; `commit-threw` is the host rejecting the batch that carried it, which counts as an answer ONLY because the probe commits the same batch without a binding first — see below; `unreadable` means the binding was made and then would not name its shape, the same refusal wearing a new coat; `yes` is the one that changes what can be built. FIRST REAL SIGNAL, 2026-08-09 evening (`2a44f64`): the commit came back `UnexpectedError` in 1.3 seconds while `shape-add-fresh-slide-proxy` answered `yes` in the same sheet. That points at the binding, but it is cross-question inference on a host that flaps between minutes, so the probe was given its own control arm rather than the reading being written down as fact. One sample, and the question stays open until a sheet answers it with the control in place. WHAT MICROSOFT'S OWN DOCS ADD (searched 2026-08-13, `bind-shapes-in-presentation`): the API is real and current — `bindings.add(shape, BindingType.shape, id)` and `Binding.getShape()` are both PowerPointApi 1.8, and this host reports 1.10, which is consistent with `commit-threw` rather than `no-binding-api` and means the surface is not the problem. The documented FLOW differs from the probe's in one respect worth a variant: Microsoft creates the shape, sets its fill, and retrieves through `bindings.getItem(id).getShape()` in a SEPARATE `PowerPoint.run`, whereas the probe binds inside the batch that created the shape. So `commit-threw` may be the host refusing to bind a shape the document does not yet have, rather than refusing bindings at all. A variant that binds in a LATER batch than the one that drew the shape is untried, and is the partner question this entry wants next. Do not read that as a prediction — it is the one difference between a flow Microsoft publishes and the flow that failed here.",
  "getitemat-past-end":
    "Nothing in this repo currently depends on the answer — it is here to find out before something does.",
  "untrack-available":
    "The fake does not implement `untrack`, so every `untrack()` call in `powerpoint.ts` is a no-op under test and the proxy-release path is entirely unexercised. A real host saying 'yes' does not fix that; it means the path is real and still untested.",
};

/**
 * Find the answer sheet inside whatever arrived — the sheet itself, or a whole
 * round's file with the sheet nested in it.
 *
 * Split out of `answersOf` because the ANSWERS were being unwrapped and the
 * HEADER was not: `host-diff.mjs` read `source` and `requirementSets` off the
 * outer object, so a round file — the shape the pane actually writes, and the
 * one `answersOf` exists to accept — printed `REAL HOST ?` and
 * `requirement sets: unknown` above a page of real answers. Which host, and
 * which API versions it offers, is not decoration here: the whole binding lead
 * turns on whether PowerPointApi 1.8 is present, and the file said 1.1 through
 * 1.10 while the report said it did not know.
 *
 * "Run the whole round" writes one file for both halves precisely so there is
 * one thing to send; every reader of it has to unwrap the same way.
 */
export function sheetOf(file) {
  if (file?.kind === "powerchart-host-answers") return file;
  if (file?.hostAnswers) return sheetOf(file.hostAnswers);
  return file ?? null;
}

/** Read a sheet, whichever shape it arrived in. */
export function answersOf(file) {
  const sheet = sheetOf(file);
  if (sheet?.kind === "powerchart-host-answers" && Array.isArray(sheet.answers)) {
    return Object.fromEntries(sheet.answers.map((a) => [a.id, a.answer]));
  }
  // A bare map, e.g. the committed baseline.
  if (sheet && typeof sheet === "object" && !Array.isArray(sheet)) return sheet;
  return null;
}

/**
 * Answers that mean the question was never put, so they are never divergences.
 *
 * The probe's own vocabulary, kept in step with `NOT_ASKED` in
 * `src/render/host-probe.ts`. No probe can produce any of these words as an
 * answer, which is what makes them safe to read this way.
 *
 * "Kept in step" was a hope, and it drifted: `not-asked` — what the probe's mute
 * breaker records when it abandons the rest of a sheet — was added on the probe
 * side and never here. Every question that breaker gives up on was therefore
 * compared against the fake and reported as a real-host DIVERGENCE, which is the
 * exact failure the third bullet of `diffAnswers` was written for ("PowerPoint
 * refused eight setups on 2026-08-04 and this tool reported eight host
 * divergences from questions nobody had asked"). The same bug, wearing a word
 * that did not exist yet. No round has tripped the breaker so far, so nothing
 * published has been wrong — it would have fired the first time a host degraded
 * badly enough to matter. `test/host-probe.test.ts` now asserts the two sets are
 * equal, so the next word cannot be added to one side alone.
 */
export const NEVER_ASKED = new Set(["no-scratch-slide", "no-scratch-shape", "not-asked"]);

/**
 * Compare two answer sheets.
 *
 * Three ways a question can fail to be a match, and all three used to be one:
 *
 * - `onlyReal` / `onlyFake` — one side was never asked. A gap, not agreement.
 * - `notAsked` — the real host never got far enough to answer. The probe run
 *   says so in a word no probe can produce, because it learned the hard way
 *   what happens otherwise: PowerPoint on the web refused eight setups on
 *   2026-08-04 and this tool reported eight host divergences from questions
 *   nobody had asked.
 * - `differ` — the only one that is actually a finding.
 */
export function diffAnswers(real, fake) {
  const ids = [...new Set([...Object.keys(real), ...Object.keys(fake)])].sort();
  const agree = [];
  const differ = [];
  const notAsked = [];
  const onlyReal = [];
  const onlyFake = [];
  for (const id of ids) {
    const r = real[id];
    const f = fake[id];
    if (r === undefined) onlyFake.push(id);
    else if (f === undefined) onlyReal.push(id);
    else if (NEVER_ASKED.has(r)) notAsked.push({ id, why: r });
    else if (r === f) agree.push(id);
    else differ.push({ id, real: r, fake: f, means: WHAT_IT_MEANS[id] });
  }
  return { agree, differ, notAsked, onlyReal, onlyFake };
}
