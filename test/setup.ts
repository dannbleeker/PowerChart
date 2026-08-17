import { beforeEach } from "vitest";
import { _setReReadRetryDelayForTest } from "../src/render/powerpoint";

/**
 * A fake host has no lag to settle, so no suite should pay for one.
 *
 * `REREAD_RETRY_MS` buys two things against the real host: the pre-grouping
 * re-read's second attempt, and the orphan instrument's second shape count. Both
 * exist because PowerPoint on the web answers a question it has already
 * committed the answer to — round 084 reported four slides growing by 23 shapes
 * apiece, measured 1.3 seconds after an `addGroup` that had synced, on slides
 * the deck inventory then showed holding one grouped chart each.
 *
 * In a suite that lag is fiction. The fake answers correctly the first time, so
 * every one of those waits is dead time — and it is per UPDATE, which took four
 * selftest scenarios past their five-second budget the moment the second count
 * landed. Set once here rather than in each test that happens to enable tracing,
 * because the next such test would hit the same wall and the wall has nothing to
 * do with what it is testing.
 *
 * The production default is deliberately NOT changed: 1.5s is sized to a real
 * host that was measured lagging 1.3s.
 */
beforeEach(() => {
  _setReReadRetryDelayForTest(1);
});
