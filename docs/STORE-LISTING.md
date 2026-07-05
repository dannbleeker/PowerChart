# AppSource store listing — copy & submission checklist

Everything needed to submit PowerChart to **Microsoft AppSource** (via Partner
Center). This is the *public store* path — heavier than sideloading or an
org-wide admin deploy (see `PUBLISHING.md` "Distribution beyond sideloading").

> **Trademark rule for anything store-facing:** the internal docs describe
> PowerChart as a "think-cell clone / think-cell-style" tool, but the public
> listing, name, description, and screenshots must **not** use the "think-cell"
> mark as branding. Describe the features generically. A nominative disclaimer
> ("not affiliated with … think-cell") is fine and lives on `/terms.html`.

## Listing copy (paste into Partner Center)

**Name:** PowerChart

**Subtitle / summary (≤ ~100 chars):**
> Consulting-grade charts as native, editable PowerPoint shapes.

**Short description (≤ ~1,000 chars):**
> PowerChart turns a simple data table into the charts that consultants and
> analysts actually use — waterfall/bridge, Mekko/Marimekko, stacked and
> clustered columns, 100% charts, lines, areas, pie/doughnut, scatter/bubble,
> Gantt plans, and more — inserted onto your slide as **native, fully editable
> PowerPoint shapes**, never flat pictures or opaque objects. Add the
> annotations that tell the story: CAGR arrows, difference arrows, value lines,
> automatic column totals, and collision-avoiding labels. Every chart stays
> re-editable: reopen the pane, change the data, and it updates in place.

**Long description — feature bullets:**
- 18 chart kinds incl. waterfall bridges, Mekko, stacked/clustered/100%,
  Gantt, combo, scatter/bubble, radar, heatmap, treemap, and more.
- Native PowerPoint shapes — recolour, move, or restyle any element by hand.
- Signature annotations: CAGR & difference arrows, value lines, totals, smart
  labels with a global de-overlap pass.
- Re-editable charts, saved templates, and an import/export style file for a
  consistent corporate look.
- Runs entirely in your client — your data never leaves your device.

**Categories:** Productivity; Data visualization
**Search keywords:** waterfall chart, bridge chart, Mekko, Marimekko, Gantt,
consulting charts, CAGR, editable charts, data visualization
**Support URL:** https://github.com/dannbleeker/PowerChart
**Privacy URL:** https://powerchart.struktureretsundfornuft.dk/privacy.html
**Terms URL:** https://powerchart.struktureretsundfornuft.dk/terms.html

## Assets to produce
- **Store logo** 300×300 PNG (Partner Center listing image — separate from the
  16/32/80 ribbon icons already in the manifest).
- **1–5 screenshots** 1366×768 of the pane + an inserted chart (use the live
  gallery / a real deck; no "think-cell" text in-frame).
- Optional short demo video.

## Submission checklist
- [ ] Partner Center account created (free for Office Store apps).
- [ ] `manifest-prod.xml` validated: `npx office-addin-manifest validate manifest-prod.xml`.
- [ ] Add-in works on **every** platform the manifest claims (web + Windows +
      Mac — testers check all of them). Do Phase 2 validation first.
- [ ] Privacy + Terms pages live (they build to `/privacy.html`, `/terms.html`).
- [ ] Listing copy above is trademark-clean; screenshots contain no competitor marks.
- [ ] Value is demonstrable **without a login** (PowerChart needs none — good).
- [ ] Submit → respond to Microsoft validation feedback (days–weeks).

## Faster alternative (recommended first)
For BESTSELLER-internal use you don't need the store at all: a Microsoft 365
admin uploads `manifest-prod.xml` under **Admin Center → Settings → Integrated
apps → Upload custom app** and deploys it to chosen users/groups. No review,
appears automatically.
