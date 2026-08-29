# Changelog

Notable changes, newest first. Dates are the release date.

This file starts at 0.4.0, the release that renamed the project. Earlier
releases are in the git history and on the [releases page][releases]; there was
no changelog before this one, and inventing entries for 451 commits after the
fact would produce a document nobody could trust.

## Unreleased

### Charts stop drawing text over text

A sweep of every chart kind under twenty-four options and ten data shapes — at
eight frame sizes, two fonts and both orientations — counted **2,148 places
where this engine drew one piece of text through another**. What is left is one
open question about where a unit label belongs, plus a small tail, and the sweep
runs on every build so the number can only go down.

**Three figures appeared while this was in progress and only one of them was a
change to the engine**, so none of them should be read against another without
saying which sweep produced it: 2,148 → 467 was the fixes below; 467 → 4,010 was
the sweep itself widening to cross options with data shapes, which it had never
done; and 4,010 → 1,327 was the small-multiples bug in the next entry.

What that looked like on a real chart:

- **A pie or doughnut with many slices** drew its outside labels through each
  other. They now shrink to the room the neighbour leaves and are dropped when
  even that is too small — the wedge is still drawn, so a dropped label loses a
  name, not a number.
- **A radar** did the same with its spoke names. On a web this matters more than
  anywhere else: the chart *is* the mapping from name to axis, so a name nudged
  onto its neighbour's spoke does not look untidy, it lies. Nothing is moved.
- **A combo with a label on every point** ran the numbers together at twenty-odd
  categories, and a combo with several line series stacked their labels on one
  category. Both now yield to their neighbours.
- **A scatter or bubble** printed point labels across the axis numbers. The
  engine had already decided that a point's label is data and an axis number is
  chrome; it just never acted on it. The chrome yields now, and the placer tries
  to dodge a number before overwriting it, so the axis keeps most of its scale.
- **A dual-axis or pareto chart's second axis** had no fit at all: five tick
  numbers, each placed at its own tick and never measured against the next. On a
  short plot they were simply drawn on top of one another.
- **A gantt's date row** thinned by a fixed gap that could not know how wide a
  date is, so "December 2024" ran into the next one.
- **A gauge** crowded its slice labels and printed them through the big total in
  the middle — and drew that total wider than the arc it sits in.
- **A bubble chart's size key** centred each reference number in a box the width
  of its own circle, so the smallest number spilled out over its neighbours.

### Small multiples no longer scatter their labels off the slide

Asking for a grid of panels that could not fit produced a chart whose text ran
hundreds of points below the bottom of it — on the slide, under whatever came
next, and nowhere near the chart it belonged to. Ten series in two columns on a
short chart is the case: five rows, and after the title and the gaps there was
nothing left per panel.

The engine was computing a negative panel height and handing it on, where a
guard meant for malformed input from outside — a width of `NaN` pasted from
somewhere — quietly replaced it with a default. Each panel was then laid out as
a full-size chart, and ten full-size charts were stacked nine points apart
inside a box sixty points tall.

A grid whose panels have no room is no longer drawn: the chart renders whole
instead. Nothing that fits today changes, because only a size of zero or less
was ever being rewritten.

### Editing a chart is two to four times faster

An edit that changes one thing used to send the host everything about every
shape it touched — twenty separate instructions to change a single word of a
title. It now sends only what actually differs.

Measured on the real thing, not in a simulator: retitling sends two instructions
where it sent twenty, recolouring a series 44 where it sent 152, and a deck-wide
rescale 180 where it sent 272. On PowerPoint on the web that is a rescale across
eight charts falling from around 150 seconds to 103, and the smaller edits
falling further in proportion.

### Inserting a chart no longer pays a four-second pause first

Every insert asked PowerPoint for the slide's size, and on a document that had
been sitting a moment that question went unanswered for a flat four seconds
before the pane gave up and got the answer another way. The wait was not buying
anything: when the question is answered at all it is answered in about a quarter
of a second, and the fallback costs about the same again.

The pane now waits a second and a half rather than four, so a cold document
costs roughly two and a half seconds less per insert. Nothing else changed about
how the size is found.

### The pane says when an insert will be slow

Adding a chart to a slide that already holds content costs several times what
the same chart costs on an empty one, and the pane used to say nothing — a
loaded slide simply looked like it had hung. It now estimates the wait from
2,917 timed inserts, and where a slide of its own would at least halve it,
offers one. Both real waits are quoted; neither is called instant.

### It tells you when a release lands mid-session

If a new version is published while your pane is open, the pane is holding an
older page and asks the server for files that have been replaced. That used to
surface as a browser error naming a URL. It now says SSF Charts has been
updated, that closing and reopening the pane fixes it, and that your slides are
untouched.

## 0.4.0 — 2026-08-27

### Renamed to SSF Charts — **this breaks existing installs**

The add-in is now **SSF Charts**, hosted at
`ssf-chart.struktureretsundfornuft.dk`.

**Everyone must re-install it.** An Office add-in manifest pins the pane to a
host, GitHub Pages serves one custom domain per repository, and the old address
no longer answers. An add-in sideloaded from a previous release will open a pane
that cannot load. There is no way to migrate this from our side — the manifest
lives in your PowerPoint, not in ours.

**What to do:** remove the old add-in, then sideload the new
[`manifest-prod.xml`][manifest]. Nothing in your decks changes: every chart you
have already inserted stays editable, because the tags and shape names written
into your slides were deliberately left on their old values. See "Not renamed"
below.

Also renamed: the repository (`dannbleeker/SSF-Charts`, old links redirect), the
npm package, and the Claude Agent Skill — which uploads as a **new** skill
rather than replacing the old one, so delete the `powerchart-charts` entry after
installing `ssf-charts.zip`.

### New visual identity

The pane wears the SSF design system: navy header and headings, blue for chrome,
and a single orange accent — the tick above the content, which is the system's
signature. Ribbon icons are redrawn to match. Both light and dark themes were
checked for contrast; every foreground/background pair clears WCAG AA and most
clear AAA.

### Fixes

- **Fewer full redraws when updating a chart in place.** A single refused group
  read used to disable in-place updates for the rest of a run, so every
  remaining chart was redrawn whole — the expensive path. The refusal is now
  scoped to the batch that saw it. Measured across ten rounds: redraw rate
  **21.4% → 14.3%**, and in-place updates after a refusal went from 0 to 1 per
  round.
- **Deck style is stored in the deck**, so a shared deck keeps its branding for
  whoever opens it.
- **Chart output**: fourteen fixes across options, data shapes and decorations
  where a chart could be left wrong; the text-overlap gate now runs against
  every font, and the value axis no longer climbs into the title.
- **Stability**: the per-sync shape budget is measured rather than guessed —
  earlier values crashed PowerPoint on the web.

### Security

`npm audit` reports **0 vulnerabilities**, down from four. `qs` was pinned past
its advisory range. `image-size` has no patched release — every published
version is covered by two HIGH advisories — but nothing imports it, so it is
replaced by a stub that throws if anything ever does. See
[`vendor/image-size-stub/README.md`][stub].

### Not renamed, deliberately

These are written **into your decks** and read back to recognise a chart as
ours. Renaming them would orphan every chart inserted by an earlier build:

    POWERCHART_CONFIG / _PARTS / _ORIGIN / _SCENE / _DEMO_SLOT   shape tags
    PowerChart                                                   group/shape name
    PowerChart:not-complete                                      banner name
    <powerchartStyle>                                            deck style element

The manifest's `<Id>` is also unchanged: a new GUID would make this a different
add-in rather than an update. Two archive format markers
(`powerchart-host-answers`, `powerchart-crash-log`) are read under both
spellings so 257 archived test rounds stay readable.

[releases]: https://github.com/dannbleeker/SSF-Charts/releases
[manifest]: https://github.com/dannbleeker/SSF-Charts/blob/main/manifest-prod.xml
[stub]: https://github.com/dannbleeker/SSF-Charts/blob/main/vendor/image-size-stub/README.md
