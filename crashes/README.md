Crash reports written by `scripts/round.mjs` when a round wedges.

NOT rounds. Everything that pools `rounds/` — verdict histories, the rasterise
arms, the flip detector — must never see these, which is why they live here.

Each file is PowerPoint's own account of why it died: the console errors, where
the document's data channel stopped, and the ULS window around the fatal entry.
The browser discards all of it when the tab reloads, and recovery reloads the
tab, so this is the only copy.
