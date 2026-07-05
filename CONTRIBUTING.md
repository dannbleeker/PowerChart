# Contributing to PowerChart

## Setup & daily loop

```bash
npm install
npm run dev          # demo gallery at http://localhost:3000, pane at /src/taskpane/taskpane.html
npm test             # 300+ tests incl. SVG snapshots, fuzzing, and mocked Office.js hosts
npm run coverage     # same suite + report; CI enforces the thresholds in vitest.config.ts
npm run typecheck
```

The layout engine (`src/core`) is pure TypeScript — most changes are testable
without PowerPoint. The Office.js renderer (`src/render/powerpoint.ts`) is a
thin mapping layer; changes there need a sideloaded host to verify.

## The lockstep rule

Two artifacts must always match the feature set, and CI enforces both:

1. **The Claude Agent Skill** (`skill/`) — `test/skill-docs.test.ts` fails if a
   chart kind, datasheet row, or decoration key is missing from the docs. The
   zip is rebuilt by CI from these sources on every push.
2. **The showcase deck** (`examples/showcase.pptx`) — extend
   `scripts/build-showcase.mjs` for new features, run `npm run showcase`, and
   commit the regenerated json+pptx. `test/showcase.test.ts` checks coverage
   and CI diffs the slide XML against a fresh build.

## Snapshots

`test/snapshots.test.ts` freezes every sample chart's SVG. If your change
intentionally alters layout, review the gallery visually first, then
`npx vitest run -u` and mention the change in the PR.

## Releasing

Merges to `main` refresh the rolling `skill-latest` release. For a versioned
release, push a tag: `git tag v0.x.0 && git push origin v0.x.0` — the Release
workflow attaches the skill zip, both manifests, and the showcase deck.
