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
  "shape-resolve-held-slide-proxy": "yes",
  "shape-add-fresh-getitem-slide": "yes",
  "shape-add-positional-slide-proxy": "yes",
  "shape-proxy-survives-one-sync": "yes",
  "shapes-items-count-honest": "at-least-5",
  "shapes-items-via-positional-slide": "at-least-5",
  "getcount-populates-same-sync": "yes",
  "tags-add-same-key-twice": "overwrites",
  "tags-on-fresh-shape": "yes",
  "delete-then-lookup": "reports-gone",
  "addgroup-returns-usable": "yes",
  "group-children-via-getcount": "two",
  "group-reports-its-children": "two",
  "tag-on-group-survives": "yes",
  "binding-names-shape-later": "yes",
  "getitemat-past-end": "threw",
  "picture-then-shape-read": "yes",
  "group-of-existing-shape-readable": "2",
  "slide-layout-readable": "yes",
  "layouts-readable": "yes",
  "untrack-available": "no",
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
  "picture-then-shape-read":
    "The fake's happy path re-reads a shape collection after an image insert. office-js#5022 (open, Microsoft-assigned) reports `context.sync()` running indefinitely on exactly that sequence, and this host answered `unreadable` on 2026-08-08. `drawDemoItem` does this shape whenever a chart degrades to a picture, since `needsRefresh` is true for any item carrying `pictureBase64`.",
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
  "shapes-items-count-honest":
    "`short-0` again on 2026-08-09 (`619d24b`), against `unreadable` in the committed fixture — nine samples now, still alternating between the two forms with no trend. " +
    "The ANSWER is stable; its FORM is not. `unreadable` (2026-08-05), `short-0` (2026-08-08 on 2f1e8c4), `unreadable` again " +
    "(2026-08-08 on a546897), `short-0` again (2026-08-09 on 8bb9e8f), then `unreadable` on the last three (448ffc6, cfa1f50 and the round after it, which answered identically). Seven samples, two forms, no trend — though the last three held still, which is the longest it has. Every one of those says the same thing — this host will not tell a caller what is on a slide — and they " +
    "differ only in how the refusal arrives: a collection that throws, versus one that answers with zero items. Worth keeping apart " +
    "because the two want different code (a catch versus a corroborated count), and `slideShapeList` handles both for exactly this reason.",
  "shapes-items-via-positional-slide":
    "As its by-id partner above, and moving in step with it: `short-0` (2f1e8c4), then `not-listed` on every run since (a546897, d812d0c, 448ffc6, cfa1f50). Every run agrees with the " +
    "by-id form in the same run, which is the finding — the parent handle is not the variable. What varies is the host, run to run.",
  "group-children-via-getcount":
    "ASKED AND ANSWERED ONCE, on a degraded host: `unreadable`, 2026-08-08. Its sibling `group-reports-its-children` answered `threw` " +
    "(\"The property 'items' is not available\") on a healthy round the same day, so BOTH routes into a group's children have now been " +
    "refused and `contentShapes` returning UNKNOWN_CONTENT for a grouped slide looks permanent rather than a gap. " +
    "Listed here rather than treated as settled because that round put only 17 of 27 questions — `getcount-populates-same-sync` itself " +
    "came back `no-scratch-slide` in it, having answered `yes, value=9` the round before. One sample from a host in that state is a " +
    "sample. Two consistent answers from two routes is a strong hint, not a finding.",
  "shape-add-held-slide-proxy":
    "ALTERNATES. Six observations: `threw` (2026-08-05), `threw` (2026-08-07), `yes` (2026-08-08 run a), `threw` (2026-08-08 run b), " +
    "`threw` (2026-08-08 run c, the sheet now committed), `threw` (2026-08-09 on d812d0c). Five of the six are `threw`. Earlier wording here said it flipped once, which " +
    "reads as though the newer value were the true one and the old one a mistake — it is not a sequence of corrections, it is a coin, and " +
    "five of six landings do not make the seventh a mechanism. " +
    "The fake keeps refusing held proxies, which is the safe direction: code that never holds one across a sync is correct whichever way the coin lands.",
  "shape-add-positional-slide-proxy":
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
  "shape-resolve-held-slide-proxy":
    "Added after the fixture's build. It decides whether `deleteShapesById`, `setShapeSelection` and the selection path are bugs or merely untidy: all three resolve a slide, sync, then reach through that same handle for a shape. The write form of this is known to fail; the read form has never been asked, and the fake's windowed handle does not gate it either way.",
  "binding-names-shape-later":
    "Added 2026-08-09, after the round on `551ad42` failed `same scale across the deck` for the fifth time in the same shape: five charts of eight drew all 24 shapes and were then unreachable — `InvalidParam passed to GetItem(id)`, 5010, at `ShapeCollection.getItem`, three times each (ids, config tag, positions) — so each left 24 shapes on a slide that is no longer a chart, and the settle pass repaired none of them (`{charts:1, settled:0, lost:1}`). Both handles that pass has are already known-refused here: `shapes-items-count-honest` says the collection reads back empty and `shapes-items-via-positional-slide` says a positional parent reads no better. A PowerPointApi 1.8 binding is the only reference that goes through neither — made from the live proxy in the shape's CREATING batch, persisted by the document, asked for later by our own key. If it survives, the repair pass gets the handle it lacks and a lost config tag becomes repairable instead of a chart the user cannot edit; if it does not, that closes the last cheap idea and the answer is worth as much. Nothing in `src/` uses bindings today — this is a question, not a half-built feature. ASKED TWICE, ANSWERED NEITHER TIME: `no-scratch-slide` on `2a44f64` morning (the run never got a slide for it), then `no-scratch-shape` on `2a44f64` evening — that one reached its own commit and got `UnexpectedError` in 1.3 seconds, which is the first real signal and is written up in `WHAT_IT_MEANS`. The control arm added after that round is what will make the third attempt an answer either way.",
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
  "delete-then-lookup":
    "`deleteSlideById` re-checks from a FRESH context because the same-context answer was not trusted. If a host answers honestly here, that second round trip is removable.",
  "group-children-via-getcount":
    "Added 2026-08-08, in response to the sheet taken that day. Its sibling `group-reports-its-children` asked through `group/shapes/items/id` and this host answered `threw` — \"The property 'items' is not available\", office-js#6363's exact signature — with the load queued in the sync that MADE the group, which is the friendliest form the question has. But the same sheet says `getcount-populates-same-sync: yes, value=9`: this host COUNTS a shape collection it will not LIST. Nobody has asked whether that holds for a GROUP's collection, and the answer decides something concrete — `contentShapes` returns UNKNOWN_CONTENT for every grouped slide, which is what makes the reconcile report a slide complete without counting it. A count is all it needs.",
  "group-reports-its-children":
    "The single most load-bearing answer here. A chart IS a group, and the readback measures whether a chart survived by counting what is inside it — so a host that groups successfully and then reports no children makes every chart read back as wreckage, and the repair pass 'fixes' charts that were never broken. WITHDRAWN: the 2026-08-04 PropertyNotLoaded was a nested load queued a sync after the group was made. It is queued in the grouping batch now.",
  "tag-on-group-survives":
    "Where a chart's config actually lives. WITHDRAWN: the 2026-08-04 NO was the probe writing through a group proxy a sync old. Taken at face value it says no chart in any deck is re-editable, which the same run disproves — its repair pass landed 23 retags on grouped charts. The group's id is what crosses the sync now, and every use resolves its own handle.",
  "binding-names-shape-later":
    "Whether the repair pass can be given a handle that does not go through `ShapeCollection.getItem(id)`. Every 5010 this host throws is at that call, and it is what leaves a chart drawn and nameless — no group, no tag, nothing to settle one onto. A binding is made from the live Shape proxy inside the batch that created it, so it needs neither an id round trip nor a collection read, and the document persists it. A real `yes` means `settleAndTagChart` has a route it does not have today; a real `no` retires the idea. Watch the answer WORD, because five of them are different facts: `no-binding-api` is a missing 1.8 surface and says nothing about the idea; `add-threw` is `bindings.add` objecting on the spot; `commit-threw` is the host rejecting the batch that carried it, which counts as an answer ONLY because the probe commits the same batch without a binding first — see below; `unreadable` means the binding was made and then would not name its shape, the same refusal wearing a new coat; `yes` is the one that changes what can be built. FIRST REAL SIGNAL, 2026-08-09 evening (`2a44f64`): the commit came back `UnexpectedError` in 1.3 seconds while `shape-add-fresh-slide-proxy` answered `yes` in the same sheet. That points at the binding, but it is cross-question inference on a host that flaps between minutes, so the probe was given its own control arm rather than the reading being written down as fact. One sample, and the question stays open until a sheet answers it with the control in place.",
  "getitemat-past-end":
    "Nothing in this repo currently depends on the answer — it is here to find out before something does.",
  "untrack-available":
    "The fake does not implement `untrack`, so every `untrack()` call in `powerpoint.ts` is a no-op under test and the proxy-release path is entirely unexercised. A real host saying 'yes' does not fix that; it means the path is real and still untested.",
};

/** Read a sheet, whichever shape it arrived in. */
export function answersOf(sheet) {
  if (sheet?.kind === "powerchart-host-answers" && Array.isArray(sheet.answers)) {
    return Object.fromEntries(sheet.answers.map((a) => [a.id, a.answer]));
  }
  // A whole round's file, with the sheet inside it. "Run the whole round"
  // writes one file for both halves precisely so there is one thing to send;
  // this tool refusing to read it would put the second upload straight back.
  if (sheet?.hostAnswers) return answersOf(sheet.hostAnswers);
  // A bare map, e.g. the committed baseline.
  if (sheet && typeof sheet === "object" && !Array.isArray(sheet)) return sheet;
  return null;
}

/**
 * Answers that mean the question was never put, so they are never divergences.
 *
 * The probe's own vocabulary, kept in step with `NOT_ASKED` in
 * `src/render/host-probe.ts`. No probe can produce either word as an answer,
 * which is what makes them safe to read this way.
 */
const NEVER_ASKED = new Set(["no-scratch-slide", "no-scratch-shape"]);

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
