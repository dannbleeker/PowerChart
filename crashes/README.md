Crash reports written by `scripts/round.mjs` when a round wedges.

NOT rounds. Everything that pools `rounds/` — verdict histories, the rasterise
arms, the flip detector — must never see these, which is why they live here.

Each file is PowerPoint's own account of why it died: the console errors, where
the document's data channel stopped, and the ULS window around the fatal entry.
The browser discards all of it when the tab reloads, and recovery reloads the
tab, so this is the only copy.

## Some of these ARE complete rounds

"NOT rounds" is about the directory, not about every record in it. The host dies
in `collectDeckEvidence`, which runs after every verdict is in — 9 builds against
4 on the per-phase traces — so a `-crashed-run.json` written there holds a full
scenario result that was never filed. **33 of 49 do.**

`scripts/salvage-crashed.mjs` turns those into rounds, and refuses the rest by
name. It writes to `rounds-salvaged/`, never here and never to `rounds/`; see
that directory's README for why they cannot join the numbered archive.

The rule above still holds for the pooling functions: nothing should read this
directory as evidence. Read the salvages instead.
