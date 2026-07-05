## What & why

<!-- One paragraph: the feature/fix and the think-cell behavior it maps to, if any. -->

## Feature-set lockstep

- [ ] New chart kinds/decorations/datasheet rows documented in `skill/SKILL.md` + `skill/reference.md` (enforced by `test/skill-docs.test.ts`)
- [ ] Showcased in the demo deck: `scripts/build-showcase.mjs` extended + `npm run showcase` committed (enforced by `test/showcase.test.ts` + CI staleness gate)
- [ ] User manual updated: `docs/MANUAL.md` (enforced by `test/manual.test.ts`)
- [ ] README feature table updated

## Verification

- [ ] `npm run typecheck` and `npm test` pass (snapshots updated intentionally, if at all)
- [ ] Visual check: demo gallery reviewed / screenshots for layout changes
- [ ] For Office.js renderer changes: exercised on a real PowerPoint host, or the graceful-degradation path is noted
