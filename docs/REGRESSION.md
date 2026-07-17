# Real-host regression: the demo deck

The **Insert demo deck** action (Automation tab) is PowerChart's real-host
regression harness. It drops one slide per chart kind + feature + element onto
fresh slides, so a human — or a diff — can see what PowerPoint *actually* draws.
This catches host-only bugs the vitest fakes cannot (ellipse repaint, wedge
tessellation, slide-timing races), which is the whole point: the fake asserts our
intent, the deck asserts the host's behaviour.

The deck is **self-identifying and self-summarising**: it opens with a **title
slide** stamping the running build *and* the host (`Office.context.diagnostics` →
`PowerPoint · OfficeOnline · 16.0.x`), a **contents** slide indexing every chart
with its office-shape count, and closes with a **results slide** — a summary line
plus a table of only the skipped/failed items. So an exported PDF is a complete,
comparable record of one run without opening the console: which build, which host,
what failed, how long it took.

Two conveniences: tick **Smoke test (10 slides)** in the pane for a fast pass over
one representative chart per family (`demoItems({ smoke: true })`) instead of the
full deck; and a slide that stalls is **retried once** automatically in a fresh
context (the recovered item is marked `retried`, and — because a mis-identified
last-slide delete is not worth the risk — attempt 1's stray slide is left in place;
`addsIssued`, below, accounts for it).

## 1. The cheap pass — self-check (every run)

Insert the deck. When it finishes, the pane reports and the **console** (F12)
prints a per-chart table plus the run's integrity numbers:

```
chart        shapes  status      retried   ms
Bubble         44    rendered      false   180
Combo          22    failed        false 45012   ← host stalled mid-draw (near the 45s timeout)
Doughnut       15    rendered      false   240
Area            0    skipped       false     2   ← too dense, stamped
deck grew by 33, issued 35 adds — 2 LOST; blank slots 24, 34 · total 78.4s
```

Read `insertDemoDeck`'s `DemoReport` (`src/render/powerpoint.ts`): `results[i]` is
`{created, status, ms, retried}`; `slidesAdded` is the deck's ACTUAL growth (settled
`getCount`, after − before); `totalMs` is the whole run's wall-clock.

**Lost slides — measure against `addsIssued`, not `items.length`.** The retry and a
double-failed item each leave a **stray** slide, so `slidesAdded` can equal
`items.length` even when the host lost a real slide — the stray cancels it. So loss
is `addsIssued − slidesAdded`, where `addsIssued` = one add per item plus one more
for each retried/failed item (both make a second attempt). A stray that LANDED
cancels; a swallowed/lost add does not — so this reads through the coincidence.
(A real run once lost 2 slides and reported 0 under the old `items.length` formula.)

**`blankSlides`** is the list of **1-based deck positions** of added slides that read
back with **zero shapes** — the host kept the slide but its content detached. It is
reported **by position, never by item name**: a blank slide has no content and no
config tag, so it cannot be attributed to an item, and under load the host reorders/
merges/loses slides, which breaks any positional item mapping anyway. Each `0` is
re-read once (a struggling host reports transient `0`s) before it counts, and
`blanksRead` is `false` if the readback faulted — so an empty list is never mistaken
for "no blanks" when it means "not fully measured".

Honest limits of `blankSlides`: it cannot see a **merge** (two items on one slide —
that slide isn't blank), loss *inside* a group, or a **paint-only** blank
(office-js#2699 — the shapes exist, so `getCount > 0`). Naming the missing/merged
charts by their config tag (`CHART_TAG`, deck-wide, order-independent) is a
**documented follow-up**, not yet built.

This pass needs no PDF and catches structural regressions: skipped, failed, lost
(via `addsIssued`), and empty slots (via `blankSlides`).

It does NOT catch *paint* bugs — a shape created but not rendered (office-js#2699)
still counts. That's what the visual pass is for.

## 2. The visual pass — PDF diff (when paint correctness matters)

Export the deck to PDF, then compare each page to its SVG reference (the SVG
renderer is the source of truth — see the byte-identical snapshot invariant).

```
# render the PDF pages
python -c "import fitz; d=fitz.open('deck.pdf'); [p.get_pixmap(matrix=fitz.Matrix(2,2)).save(f'p{i+1:02}.png') for i,p in enumerate(d)]"
# map pages -> chart by title
python -c "import fitz; [print(i+1, p.get_text().split(chr(10))[0]) for i,p in enumerate(fitz.open('deck.pdf'))]"
# render references: sceneToSvg(buildChart(sampleConfig(kind))) -> Chrome --screenshot
```

**Pair by TITLE, not page number** — a lost or misaligned slide shifts the page
order (map the title text on each page to the chart).

**Two artifacts that WILL bite you** (both flagged bogus "defects" on a real run):

- *Framing.* The rendered chart sits in a *sub-region* of a 960×540 slide (placed
  at ~60,90 with margins), while a raw SVG reference fills its own frame. So
  "rendered is smaller / shifted down-right" is not a defect — it flagged 15
  charts once. Either place the reference on a 960×540 canvas at the same offset,
  or tell the reviewer to ignore absolute scale/position and judge only data,
  proportions, colour, and completeness.
- *Title text.* `demoItems()` overrides each sample's title with a short label
  (`{...sampleConfig(kind), title: label}`), so a reference built from
  `sampleConfig(kind)` shows the sample's LONG title — a "title differs / colour
  lost" mismatch that isn't real. Build references with the SAME title override
  the demo uses.

Net from the run that produced this doc: of ~20 "findings", exactly ONE was a real
render bug (doughnut arc gaps); the rest were these two artifacts, known host
limits, or the stamp/stall. Distrust the diff; verify each finding against the
scene before filing.

## Known host limitations (not regressions)

- **Filled polygons render as outline only** (radar/violin fills missing) — Office.js
  has no freeform fill; the SVG reference fills them, the host can't.
- **Dense charts (>~90 shapes) are skipped + stamped** on web (area, tile map,
  waffle, smoothed line) — they exceed the per-slide shape budget.
- **Freshly-inserted shapes may not repaint** until zoom/navigate (office-js#2699).

Relevant OfficeDev/office-js issues: #2699 (repaint), #2903/#2474 (slide-add
timing / non-round-trip ids), #4272 (>50-item load hang). See `docs/repro/` for a
Script Lab repro of the ellipse case.
