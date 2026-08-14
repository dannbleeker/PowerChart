import { describe, expect, it } from "vitest";
import { installHost, makeSlide, faults } from "./helpers/office-host";
import { deleteTrailingSlides, deckSlideIds } from "../src/render/powerpoint";

/**
 * Can ONE known slide be removed, or only a trailing range?
 *
 * Four attempts to stop the probe carrying seventy scratch slides through a
 * round were reverted, and the last of them concluded — in `docs/BACKLOG.md` and
 * in a merged commit message — that this host offers no way to delete a single
 * identified slide mid-run, so the whole idea was blocked. That conclusion was
 * reached by watching a full probe run misbehave, which is three thousand lines
 * of interleaved state away from the question actually being asked.
 *
 * Asked directly, the answer is yes. `deleteTrailingSlides` is named for its
 * first use and takes a range; given a range of one it removes exactly the slide
 * at that index, confirmed by id either side, and it keeps doing so on a host
 * that renumbers its slides on every add.
 *
 * So the primitive is sound and the blocker recorded against it was wrong. What
 * remains unexplained is narrower and belongs to the run rather than the tool:
 * with the mid-run handback wired in, the deck at clean-up time was the same
 * size as if nothing had gone back, while each individual delete reported
 * success. That is worth its own investigation, and it starts from here rather
 * than from "the host cannot do it".
 *
 * These two cases are kept because the contract they pin was ASSUMED by four
 * implementations and checked by none of them.
 */
describe("removing one known slide", () => {
  it("removes exactly the slide at the index it was given", async () => {
    installHost(["s1", "s2", "s3", "s4", "s5", "s6"].map(makeSlide));
    expect(await deckSlideIds()).toEqual(["s1", "s2", "s3", "s4", "s5", "s6"]);

    const gone = await deleteTrailingSlides(3, 1); // index 3 is "s4"

    expect(gone, "it did not report removing anything").toBe(1);
    expect(await deckSlideIds(), "a slide went, but not the one named").toEqual(["s1", "s2", "s3", "s5", "s6"]);
  });

  it("still takes the named slide on a host that renumbers on add", async () => {
    // `faults.renumbersOnAdd` is the behaviour that makes delete-by-id useless
    // here — the ids the deck lists are not the ids the run wrote down. Position
    // survives it, which is the whole reason the end-of-run sweep is positional.
    installHost(["s1", "s2", "s3", "s4"].map(makeSlide));
    faults.renumbersOnAdd = true;
    try {
      const before = await deckSlideIds();
      const target = before![2];

      await deleteTrailingSlides(2, 1);

      const after = await deckSlideIds();
      expect(after, `expected ${target} to be the one that left`).not.toContain(target);
      expect(after).toHaveLength(3);
    } finally {
      faults.renumbersOnAdd = false;
    }
  });

  it("takes nothing from a range that is not there", async () => {
    // Behaviour, not the guard. Removing `from < 0` from the source leaves this
    // green — `getItemAt(-1)` no-ops here — so this pins what callers can rely
    // on and does NOT prove the early return that makes it explicit. Said out
    // loud because a test quoted as proving something it does not is worse than
    // no test: the next person deletes the guard and the suite agrees.
    installHost(["s1", "s2"].map(makeSlide));
    expect(await deleteTrailingSlides(-1, 1), "a negative index is not a slide").toBe(0);
    expect(await deleteTrailingSlides(0, 0), "an empty range removes nothing").toBe(0);
    expect(await deckSlideIds()).toEqual(["s1", "s2"]);
  });
});
