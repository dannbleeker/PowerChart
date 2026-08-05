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
  "shape-add-positional-slide-proxy": "yes",
  "shape-proxy-survives-one-sync": "yes",
  "shapes-items-count-honest": "at-least-5",
  "shapes-items-via-positional-slide": "at-least-5",
  "getcount-populates-same-sync": "yes",
  "tags-add-same-key-twice": "overwrites",
  "tags-on-fresh-shape": "yes",
  "delete-then-lookup": "reports-gone",
  "addgroup-returns-usable": "yes",
  "group-reports-its-children": "two",
  "tag-on-group-survives": "yes",
  "getitemat-past-end": "threw",
  "untrack-available": "no",
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
    "WITHDRAWN, awaiting a re-run. The real answer (`items` undefined) came from a build before the probe stopped reading a collection through a handle the same sync resolved. Its partner `shapes-items-via-positional-slide` was added to separate the two; neither has been asked of a host since.",
  "tags-add-same-key-twice":
    "WITHDRAWN, awaiting a re-run. The real `other — value=undefined` was one shape proxy held across four syncs, not an opinion about tag keys.",
  "addgroup-returns-usable":
    "WITHDRAWN, awaiting a re-run. The real `unreadable` came from asking a group for its id one sync after making it, out of members that were themselves a sync old.",
  "group-reports-its-children":
    "WITHDRAWN, awaiting a re-run. The real PropertyNotLoaded was a nested load queued a sync after the group was made.",
  "tag-on-group-survives":
    "WITHDRAWN, awaiting a re-run. Taken at face value the real `no` says no chart in any deck is re-editable, which the same run disproves — its repair pass landed 23 retags on grouped charts.",
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
  "shapes-items-via-positional-slide":
    "Added with the probe rewrite, after the fixture's build. It is the partner that decides whether `shapes-items-count-honest` was ever about collections or only about the handle they hang off, so the pair is worth a run on its own.",
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
  "group-reports-its-children":
    "The single most load-bearing answer here. A chart IS a group, and the readback measures whether a chart survived by counting what is inside it — so a host that groups successfully and then reports no children makes every chart read back as wreckage, and the repair pass 'fixes' charts that were never broken. WITHDRAWN: the 2026-08-04 PropertyNotLoaded was a nested load queued a sync after the group was made. It is queued in the grouping batch now.",
  "tag-on-group-survives":
    "Where a chart's config actually lives. WITHDRAWN: the 2026-08-04 NO was the probe writing through a group proxy a sync old. Taken at face value it says no chart in any deck is re-editable, which the same run disproves — its repair pass landed 23 retags on grouped charts. The group's id is what crosses the sync now, and every use resolves its own handle.",
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
