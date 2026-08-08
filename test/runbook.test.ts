import { describe, expect, it } from "vitest";
import { readFileSync } from "node:fs";

/**
 * The runbook must name buttons that exist.
 *
 * `docs/PUBLISHING.md` is the one document written for someone who is not
 * reading the code: it tells the owner which control to click, in a PowerPoint,
 * on a build he did not make. A label that has drifted out of the pane costs him
 * the round — he goes looking for it, does not find it, and stops.
 *
 * That happened. The starred row said **Run the whole round**, which has never
 * been the text on any control; the button reads *Probe, then self-test*. Two
 * more had drifted the same way — the demo path options were written up as
 * *File (fast)* and *Shapes (everyday)* while the pane offered *One file insert*
 * and *Shape by shape*.
 *
 * So each instruction is pinned to the element it means, by id. Not a blanket
 * scan of every bold phrase in the file: most of those are prose, the noise
 * would swamp the signal, and a gate people learn to ignore is not a gate. This
 * list is the controls the standing test run actually tells the owner to press.
 * Add a row when the runbook starts naming another one.
 */
describe("the standing test run names controls that exist", () => {
  const html = readFileSync("src/taskpane/taskpane.html", "utf8");
  const runbook = readFileSync("docs/PUBLISHING.md", "utf8");

  /** The visible text of the button or option carrying this id. */
  function labelOf(id: string): string {
    const el = new RegExp(`<(button|option)[^>]*\\bid="${id}"[^>]*>([\\s\\S]*?)</\\1>`).exec(html);
    if (el) return el[2].replace(/<[^>]*>/g, "").trim();
    // Options are commonly written without an id — match on the value instead,
    // which is what the pane's own code selects them by.
    const byValue = new RegExp(`<option[^>]*\\bvalue="${id}"[^>]*>([\\s\\S]*?)</option>`).exec(html);
    if (byValue) return byValue[1].replace(/<[^>]*>/g, "").trim();
    throw new Error(`no button or option in taskpane.html for "${id}"`);
  }

  const CONTROLS = [
    { id: "demo-round", what: "the one-click round the starred row sends the owner to" },
    { id: "demo-probe", what: "test 0" },
    { id: "demo-selftest", what: "test 1" },
    { id: "demo-insert", what: "tests 2a and 2b" },
    { id: "file", what: "the demo path for test 2a" },
    { id: "shapes", what: "the demo path for test 2b" },
    { id: "demo-log", what: "what to send after a run" },
    { id: "demo-crashlog", what: "what to send after a run that died" },
  ];

  for (const { id, what } of CONTROLS) {
    it(`spells the control for ${what} the way the pane does`, () => {
      const label = labelOf(id);
      // Non-vacuity: a control whose label went empty would otherwise pass by
      // matching the empty string everywhere.
      expect(label.length, `#${id} has no visible label to match against`).toBeGreaterThan(2);
      expect(
        runbook,
        `the runbook never says "${label}" — the standing test run points at #${id}, so it has to name it`,
      ).toContain(label);
    });
  }
});
