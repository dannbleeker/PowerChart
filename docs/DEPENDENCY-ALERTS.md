# Dependency alerts, and what was decided about each

Dependabot has been posting a banner on every push for weeks — "9 vulnerabilities
(3 high, 6 moderate)" — and nobody had read it. A standing red count that nobody
triages is worse than no scanner at all: it trains everyone to scroll past the
one that matters.

This file is the same idea as `KNOWN_ISSUES` in `scripts/office-js-watch.mjs`.
An entry says what was checked and what was decided, **including "no exposure"**,
which records that somebody looked. An alert not listed here has not been read.

Re-read it with:

```bash
npm audit
gh api repos/dannbleeker/PowerChart/dependabot/alerts --jq '.[] | select(.state=="open") | "\(.security_advisory.severity)\t\(.dependency.package.name)\t\(.dependency.scope)"'
```

## Cleared 2026-08-10

`npm audit fix` (no `--force`, so nothing semver-major moved) took the count from
8 to 4: **brace-expansion**, **nanoid**, **postcss** and **undici**, all
development-scope transitives. `pptxgenjs` stayed at 4.0.1, which was the thing
to check — see below.

## Open, with a reason

### `image-size` — 2 × high — NO EXPOSURE

> ICNS parser allows denial of service through an infinite loop
> JXL and HEIF parsers allow denial of service through infinite loops

The only alert scoped **runtime**, which is what makes it worth writing down
rather than shrugging at.

- It arrives one level down: `pptxgenjs@4.0.1 → image-size@1.2.1`.
- **There is no forward fix.** 1.2.1 is the newest 1.x, the patch is in 2.x, and
  pptxgenjs — already on its newest release — pins `^1.2.1`. `npm audit`'s
  proposed remedy is to move `pptxgenjs` to **1.1.5**, three majors backwards,
  which would take the whole pptx renderer with it. That is not a fix.
- **It is not reachable.** `image-size` runs when pptxgenjs embeds a picture.
  This repo calls `addImage` in exactly one place — `skill/scripts/render-pptx.mjs`,
  on the `render: "image"` path — and what it hands over is a **PNG this code
  rasterised itself** from its own scene graph through resvg. The vulnerable
  parsers are ICNS, JXL and HEIF. No user file, of any format, reaches it.
- **It is not shipped to the browser either.** pptxgenjs is loaded by dynamic
  `import()` so it stays out of the pane's first-load bundle, and the built
  `dist/assets/pptxgen.es-*.js` carries Vite's `__vite-browser-external` stub
  where the Node-only imports were — none of the parsers are in the pane at all.

Revisit when pptxgenjs takes `image-size@2`. Do not downgrade pptxgenjs.

### `qs`, `typed-rest-client` — 2 × moderate — NO EXPOSURE

Both arrive under `@stryker-mutator/core`, the mutation-testing runner. It is a
devDependency, it runs in one weekly scheduled job on a GitHub runner, and it is
in neither the add-in nor the `test` job that gates a PR. Nothing here parses a
query string or makes an HTTP request at runtime.

Revisit when Stryker updates. Nothing to do meanwhile.
