# The round archive

Every real-host round, kept. `NNN-<build>.json`, oldest first.

## Why this exists

`scripts/triage.mjs` was built to pool evidence across rounds — `poolRasteriseArms` and
`poolEveryDraw` both take a LIST of logs, and the CLI accepts many paths. It has never had
anything to pool. Round files landed in a temp directory, were read once, and died with the
session, so every report in practice said `pooled over 1 round(s)`.

That is not a small loss. The rasterise question needs 60-100 draws a side to separate the rates it
is looking at; a single round contributes four through the counterbalanced arms and about forty
through every-draw counting. Kept, that is a handful of rounds. Discarded, it is never.

triage's own comment said it first — "the evidence was accumulating in files nobody was adding up.
This adds them up." The files still have to be there.

## What is in a file, and what is not

The run log exactly as the pane saves it — `hostAnswers`, `selftest`, `deck`, `trace` — with one
change: the base64 slide images are replaced by a marker. They are 48% of the bytes and they are
evidence only at triage time, for deciding whether a slide that read back empty really was blank.
Once that verdict is in `selftest` and `deck`, the pixels are not worth carrying forever.

Keep the unstripped original if a round's blank-slide verdict is ever disputed. Nothing here
depends on it.

## Using it

    node scripts/triage.mjs rounds/*.json        # pooled across every round
    node scripts/triage.mjs rounds/028-*.json    # one round in detail
    npm run rounds                               # the pooled view, shortest path

## Adding a round

Strip the images, name the file for the build the pane reported (NOT the build you merged — round
27 ran on `fef1c2a` while `162f80a` was the last thing merged, and naming it wrong is how an
archive starts lying), and commit it in the same change as whatever the round taught you.

`test/rounds.test.ts` checks the naming and that every file still parses as a round.
