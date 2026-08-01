import { describe, expect, it } from "vitest";
// @ts-expect-error — plain .mjs tools, no types; both are deliberately
// independent of src/ so they cannot inherit a bug from the code they audit.
import { readDeckBytes } from "../scripts/verify-deck.mjs";
// @ts-expect-error — as above.
import { triage, runsIn, selfTestIn } from "../scripts/triage.mjs";
import { buildDeckBase64 } from "../src/render/pptx-deck";
import { buildChart } from "../src/core/chart";
import { sampleConfig } from "../src/core/samples";

/**
 * The run/deck joiner.
 *
 * Its whole job is to disagree — to say where the add-in's account of a run
 * and the file that run produced do not match. So the cases that matter are
 * the ones where they differ, and every one of them below is paired with the
 * same slot in the clean deck to prove the verdict came from the difference
 * and not from the tool's default mood.
 */

const cfg = (title: string) => ({ ...sampleConfig("line"), title });

interface Piece {
  title: string;
  slot: number;
  run?: string;
  tagged?: boolean;
}

/** Build a real .pptx from slot descriptions and read it back as the auditor sees it. */
async function deckOf(pieces: Piece[]) {
  const built = await buildDeckBase64(
    pieces.map((p) => ({
      scene: buildChart(cfg(p.title)),
      title: p.title,
      configJson: p.tagged === false ? undefined : JSON.stringify(cfg(p.title)),
      slot: p.slot,
      run: p.run ?? "run-x",
    })),
  );
  return readDeckBytes(Uint8Array.from(atob(built.base64), (c) => c.charCodeAt(0)));
}

/** A run log claiming every named item rendered and was tagged. */
function logOf(titles: string[], over: Record<string, unknown> = {}) {
  return {
    build: "test-build",
    run: "run-x",
    host: "test host",
    totalMs: 1234,
    path: "file",
    items: titles.map((title) => ({
      title,
      status: "rendered",
      shapes: 3,
      ms: 10,
      grouped: true,
      tagged: true,
      chart: true,
      lateOutcome: "",
    })),
    deck: { slidesAdded: titles.length, addsIssued: titles.length, lost: 0, blank: [] },
    ...over,
  };
}

const verdicts = (t: { slots: { title: string; verdict: string }[] }) =>
  Object.fromEntries(t.slots.map((s) => [s.title, s.verdict]));

describe("triage — joining a run log to the deck it produced", () => {
  it("finds no disagreement when the deck holds exactly what the run claimed", async () => {
    const t = triage(
      await deckOf([
        { title: "A", slot: 0 },
        { title: "B", slot: 1 },
      ]),
      logOf(["A", "B"]),
    );
    expect(verdicts(t)).toEqual({ A: "ok", B: "ok" });
    expect(t.disagreements).toBe(0);
    expect(t.run).toBe("run-x");
    expect(t.inferredRun).toBe(false);
  });

  it("calls a slot the deck does not hold lost", async () => {
    // Negative control: B is `ok` above, from the same log, with its slide.
    const t = triage(await deckOf([{ title: "A", slot: 0 }]), logOf(["A", "B"]));
    expect(verdicts(t)).toEqual({ A: "ok", B: "lost" });
    expect(t.disagreements).toBe(1);
  });

  it("does not call a skipped item lost — the run never claimed to draw it", async () => {
    const log = logOf(["A", "B"]);
    log.items[1].status = "skipped";
    const t = triage(await deckOf([{ title: "A", slot: 0 }]), log);
    expect(verdicts(t)).toEqual({ A: "ok", B: "skipped" });
    expect(t.disagreements).toBe(0);
  });

  it("calls a tag the run believed it wrote and the file does not carry tag-lost", async () => {
    // The file path cannot produce this by construction, which is exactly why
    // it is worth naming when it appears: it means the deck lost a tag the
    // add-in had already committed to the bytes.
    const t = triage(await deckOf([{ title: "A", slot: 0, tagged: false }]), logOf(["A"]));
    expect(verdicts(t)).toEqual({ A: "tag-lost" });
    expect(t.disagreements).toBe(1);
  });

  it("calls a tag the run gave up on and the file carries repaired, and not a fault", async () => {
    const log = logOf(["A"]);
    log.items[0].tagged = false;
    const t = triage(await deckOf([{ title: "A", slot: 0 }]), log);
    expect(verdicts(t)).toEqual({ A: "repaired" });
    // The repair pass working is not a disagreement to chase.
    expect(t.disagreements).toBe(0);
  });

  it("calls a chart object with no config not-editable when the run never claimed the tag", async () => {
    const log = logOf(["A"]);
    log.items[0].tagged = false;
    const t = triage(await deckOf([{ title: "A", slot: 0, tagged: false }]), log);
    expect(verdicts(t)).toEqual({ A: "not-editable" });
    expect(t.disagreements).toBe(1);
  });

  it("calls two slides claiming one slot duplicated", async () => {
    const t = triage(
      await deckOf([
        { title: "A", slot: 0 },
        { title: "A again", slot: 0 },
      ]),
      logOf(["A"]),
    );
    expect(verdicts(t)).toEqual({ A: "duplicated" });
    expect(t.slots[0].rows).toBe(2);
    expect(t.disagreements).toBe(1);
  });

  it("counts another run's slides as foreign, never as this run's losses", async () => {
    // The case the run token exists to survive. Before it existed, a second
    // insert into the same deck made every item match two slides and the
    // repair pass deleted a healthy run as the other one's duplicate.
    const t = triage(
      await deckOf([
        { title: "A", slot: 0 },
        { title: "A", slot: 0, run: "run-earlier" },
      ]),
      logOf(["A"]),
    );
    expect(verdicts(t)).toEqual({ A: "ok" });
    expect(t.foreign).toHaveLength(1);
    expect(t.foreign[0].run).toBe("run-earlier");
    expect(t.disagreements).toBe(0);
  });

  it("calls a slide carrying a slot this run never issued an orphan", async () => {
    const t = triage(
      await deckOf([
        { title: "A", slot: 0 },
        { title: "stray", slot: 7 },
      ]),
      logOf(["A"]),
    );
    expect(t.orphans).toEqual([{ slot: 7, verdict: "orphan", indexes: [2] }]);
    expect(t.disagreements).toBe(1);
  });

  it("infers the run from the deck when the log predates run ids, and says it guessed", async () => {
    const log = logOf(["A", "B"]) as Record<string, unknown>;
    delete log.run;
    const t = triage(
      await deckOf([
        { title: "A", slot: 0 },
        { title: "B", slot: 1 },
      ]),
      log,
    );
    expect(t.run).toBe("run-x");
    expect(t.inferredRun).toBe(true);
    expect(t.disagreements).toBe(0);
  });

  it("leaves an item that was never a chart alone, config tag or not", async () => {
    // The bug that produced this test: a clean 38-of-38 run triaged as seven
    // broken slides. The title page, both contents pages and four elements are
    // drawn as `PowerChart` objects with NO config by design, and the rule
    // "a chart object without a config is not re-editable" flagged every one.
    const log = logOf(["Title"]);
    log.items[0].tagged = false;
    log.items[0].chart = false;
    const t = triage(await deckOf([{ title: "Title", slot: 0, tagged: false }]), log);
    expect(verdicts(t)).toEqual({ Title: "ok" });
    expect(t.disagreements).toBe(0);
  });

  it("will not guess at intent a log is too old to state", async () => {
    // Same deck and the same missing config, with `chart` absent: the tool
    // cannot tell a title page from a chart that lost its tag, so it says so
    // instead of picking one. `tagged` cannot stand in — on the shape path
    // false means both "never had a config" and "the write did not land".
    const log = logOf(["Mystery"]) as { items: Record<string, unknown>[] };
    log.items[0].tagged = false;
    delete log.items[0].chart;
    const t = triage(await deckOf([{ title: "Mystery", slot: 0, tagged: false }]), log);
    expect(verdicts(t)).toEqual({ Mystery: "no-config" });
    expect(t.disagreements).toBe(1);
  });

  it("prefers the log's own run id over the deck's commonest", async () => {
    // A deck where the run under investigation is the MINORITY. Inferring from
    // the deck would pick the other one and report every slot lost.
    const t = triage(
      await deckOf([
        { title: "A", slot: 0, run: "run-mine" },
        { title: "old", slot: 0, run: "run-older" },
        { title: "old", slot: 1, run: "run-older" },
      ]),
      logOf(["A"], { run: "run-mine" }),
    );
    expect(t.run).toBe("run-mine");
    expect(verdicts(t)).toEqual({ A: "ok" });
    expect(t.foreign).toHaveLength(2);
  });
});

describe("triage — a log holding more than one run", () => {
  it("reads each run in a both-paths log separately", async () => {
    // One click can take both insert paths, and they fail in completely
    // different ways. Merging them would produce a report that is true of
    // neither: the file run's slides are present while the shape run's are
    // still missing, in the same deck, at the same moment.
    const deck = await deckOf([
      { title: "A", slot: 0, run: "run-file" },
      { title: "B", slot: 1, run: "run-file" },
      { title: "A", slot: 0, run: "run-shapes" },
    ]);
    const file = { ...logOf(["A", "B"], { run: "run-file", path: "file" }) };
    const shapes = { ...logOf(["A", "B"], { run: "run-shapes", path: "shapes" }) };
    const runs = runsIn({ build: "b", host: "h", runs: [file, shapes] });
    expect(runs).toHaveLength(2);
    expect(verdicts(triage(deck, runs[0]))).toEqual({ A: "ok", B: "ok" });
    // The shape run only landed its first slide, and says so — while the file
    // run beside it in the same deck is clean.
    expect(verdicts(triage(deck, runs[1]))).toEqual({ A: "ok", B: "lost" });
  });

  it("reads a single-run log as a list of one", async () => {
    // Every artifact captured so far is shaped this way. Refusing them would
    // make the tool useless on exactly the runs it was written to explain.
    const flat = logOf(["A"]);
    expect(runsIn(flat)).toEqual([flat]);
  });
});

describe("triage — logs that are not inserts", () => {
  it("finds the host self-test's verdicts in a log that holds no runs", async () => {
    // The self-test writes a log with an empty `runs` list: its scenarios are
    // not inserts and own no slots. Reading only `runs` made the whole report
    // come out blank and exit 0 — indistinguishable from a clean deck, on a
    // file someone handed this tool expecting an answer.
    const log = {
      build: "b",
      host: "h",
      runs: [],
      selftest: [
        { name: "insert twice", ok: true, detail: "deck grew by 4", ms: 10 },
        { name: "explode a picture", ok: false, detail: "config did not survive", ms: 20 },
      ],
    };
    expect(runsIn(log)).toEqual([]);
    expect(selfTestIn(log).map((s: { name: string }) => s.name)).toEqual(["insert twice", "explode a picture"]);
  });

  it("finds nothing to report in a log that is neither", async () => {
    expect(selfTestIn({ build: "b", host: "h", runs: [] })).toEqual([]);
    expect(selfTestIn(undefined)).toEqual([]);
  });

  it("calls a slide tagged with a slot before the run's range an orphan", async () => {
    // Only slots PAST the end were checked, so a negative slot — a tag written
    // wrong, or carried in from something that was not this run — matched no
    // item, fell outside the orphan test, and was reported by nothing.
    const deck = await deckOf([
      { title: "A", slot: 0 },
      { title: "stray", slot: -3 },
    ]);
    const t = triage(deck, logOf(["A"]));
    expect(t.orphans).toEqual([{ slot: -3, verdict: "orphan", indexes: [2] }]);
    expect(t.disagreements).toBe(1);
  });
});
