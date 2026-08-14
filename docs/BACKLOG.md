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

### Deck-level style, so a shared deck keeps its branding

**Researched:** 2026-08-01 (`docs/RESEARCH.md` §4b).

The imported style file and saved templates live in `localStorage`
(`powerchart-style`, `powerchart-templates`), so they follow the **browser**,
not the deck. Send a branded deck to a colleague and their charts do not match
yours — and every chart they add drifts further.

**Shape of the fix:** a presentation-scoped custom XML part
(`Presentation.customXmlParts`), which is the one thing it is genuinely the
right tool for. Chart config stays in shape tags — it has to travel with the
shape, through copy/paste into another deck and through PowerPoint's own
Duplicate Slide, which a presentation-scoped part cannot do.

**Known trap:** Office.js enumerates only parts related from
`ppt/presentation.xml`. A part written at the package root
(`/customXml/itemN.xml`) is invisible to it. If the generated deck writes one,
it must be related from the presentation, not dropped at the root.

**Priority:** low. Nothing is broken today; this is a sharing gap, and it needs
a decision about precedence (deck style vs the user's own imported style) before
any code.

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

NOT attempted yet, deliberately. Two things have to be true first and only one of
them is: the write-through-the-handle answer needs a third round, and the update
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
- **Alt text in the headless pptx renderer** — the hunt flagged the missing
  `descr=` as a break of the SVG/Office.js/pptx parity contract, but pptxgenjs
  exposes alt text only on pictures and native charts, and every PowerChart
  shape is an `addShape`/`addText` autoshape. There is no seam to write it
  through short of hand-patching the generated OOXML. Recorded as a documented
  limit in `skill/reference.md` (#197) alongside the fixed deck font.
