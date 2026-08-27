# Changelog

Notable changes, newest first. Dates are the release date.

This file starts at 0.4.0, the release that renamed the project. Earlier
releases are in the git history and on the [releases page][releases]; there was
no changelog before this one, and inventing entries for 451 commits after the
fact would produce a document nobody could trust.

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
