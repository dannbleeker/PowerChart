/**
 * WHAT THIS PROJECT BELIEVES, written so the archive can contradict it.
 *
 * `docs/ROUND-LOOP-JOURNAL.md` is eight thousand lines of prose containing
 * statements of fact about the archive — "1% of charts on a freshly added slide
 * group", "the first chart of a run costs 2.2x a later one". Every one of them
 * was true when written. None of them can go stale LOUDLY.
 *
 * On 2026-08-25 that cost the most consequential finding of the day: the
 * fresh-slide failure had been fixed for an unknown number of rounds while the
 * metric that would have shown it read `0/0`, and the 1% figure went on being
 * quoted as current. Nothing was broken. Nothing was watching.
 *
 * So each claim here carries the query that checks it. `npm run rounds` runs
 * them over the whole archive and prints HOLDS or STALE. A claim that goes stale
 * is not a failure — it is the single most interesting line in the report, and
 * it is how a fix announces itself.
 *
 * NEVER FAILS THE GATE. A stale claim means the world moved, which is
 * information; making it a build break would teach people to delete claims
 * rather than write them.
 *
 * Adding one: state it the way you would say it out loud, give it the narrowest
 * query that could refute it, and say WHEN it was measured. A claim nobody can
 * refute is a slogan.
 */

/** Every round's `updated only the shapes that changed` rows, flattened. */
function updateRows(logs) {
  const rows = [];
  for (const log of logs ?? []) {
    for (const e of log?.trace?.entries ?? []) {
      if (String(e.message ?? "") !== "updated only the shapes that changed") continue;
      const d = e.data ?? {};
      if (!d.chart || typeof d.ms !== "number") continue;
      const [i, n] = String(d.chart).split("/").map(Number);
      rows.push({ round: String(log.roundName ?? "").slice(0, 3), i, n, changed: d.changed, of: d.of, ms: d.ms });
    }
  }
  return rows;
}

const median = (a) => (a.length ? [...a].sort((x, y) => x - y)[Math.floor(a.length / 2)] : null);

/** The last N rounds, so a claim about "now" is not answered by 2026-08-12. */
function recent(logs, n = 20) {
  return (logs ?? []).slice(-n);
}

export const CLAIMS = [
  {
    id: "fresh-slides-group",
    says: "A chart on a freshly added slide groups, so it keeps its config.",
    measured: "2026-08-25, 36 of 36 over eight rounds — it was 1 of 74 on 2026-08-15",
    check(logs) {
      let fresh = 0;
      let grouped = 0;
      for (const log of recent(logs)) {
        const onSlide = new Map();
        const isGrouped = new Set();
        const decided = new Set();
        for (const e of log?.trace?.entries ?? []) {
          const d = e.data ?? {};
          if (!d.chart) continue;
          if (typeof d.onSlide === "number" && !onSlide.has(d.chart)) onSlide.set(d.chart, d.onSlide);
          if (/^grouped the chart/.test(String(e.message))) {
            isGrouped.add(d.chart);
            decided.add(d.chart);
          }
          if (/^not grouping/.test(String(e.message))) decided.add(d.chart);
        }
        for (const [chart, n] of onSlide) {
          if (!decided.has(chart) || n > 0) continue;
          fresh++;
          if (isGrouped.has(chart)) grouped++;
        }
      }
      // Under twenty it cannot say either way — the same bar the rate report uses.
      if (fresh < 20) return { ok: null, actual: `only ${fresh} fresh-slide chart(s) in the window` };
      const pct = Math.round((100 * grouped) / fresh);
      return { ok: pct >= 90, actual: `${grouped}/${fresh} = ${pct}%` };
    },
  },
  {
    id: "first-chart-costs-more",
    says: "The first chart of a multi-chart update costs about twice a later chart of the same size.",
    measured: "2026-08-24, n=14 first against n=70 later, no overlap in range",
    check(logs) {
      const rows = updateRows(logs).filter((r) => r.changed === 18 && r.of === 24 && r.n > 1);
      const first = rows.filter((r) => r.i === 1).map((r) => r.ms);
      const later = rows.filter((r) => r.i > 1).map((r) => r.ms);
      if (first.length < 5 || later.length < 5) return { ok: null, actual: `n=${first.length}/${later.length}` };
      const ratio = median(first) / median(later);
      return { ok: ratio >= 1.8, actual: `${ratio.toFixed(2)}x (${median(first)}ms vs ${median(later)}ms)` };
    },
  },
  {
    id: "parts-list-never-consumed",
    says: "No chart has ever reached an in-place update carrying a parts list.",
    measured: "2026-08-25, 0 of 1005 charts",
    check(logs) {
      let charts = 0;
      let withParts = 0;
      for (const log of logs ?? []) {
        for (const e of log?.trace?.entries ?? []) {
          if (!/shapes left on the slide after an in-place/.test(String(e.message))) continue;
          const d = e.data ?? {};
          charts += d.charts ?? 0;
          withParts += d.withParts ?? 0;
        }
      }
      // STALE HERE MEANS GOOD NEWS. If this ever fails, the parts list started
      // working and the 1197 redraws it causes are on their way out.
      return { ok: withParts === 0, actual: `${withParts} of ${charts} charts`, staleIsGood: true };
    },
  },
  {
    id: "tag-faults-are-zero",
    says: "No chart loses its config tag any more.",
    measured: "2026-08-25, zero across the last 20 rounds; last fault was round 206",
    check(logs) {
      let faults = 0;
      for (const log of recent(logs)) {
        for (const e of log?.trace?.entries ?? []) {
          if (/tagging failed/.test(String(e.message))) faults++;
        }
      }
      return { ok: faults === 0, actual: `${faults} in the last 20 rounds` };
    },
  },
  {
    id: "no-queue-trace-is-dead",
    says: "The trace thread 3 waits on has stopped firing entirely.",
    measured: "2026-08-25, last seen round 065",
    check(logs) {
      let last = null;
      for (const log of logs ?? []) {
        for (const e of log?.trace?.entries ?? []) {
          if (/could not even be queued/.test(String(e.message))) last = String(log.roundName ?? "").slice(0, 3);
        }
      }
      const newest = Number(String((logs ?? [])[logs.length - 1]?.roundName ?? "").slice(0, 3));
      const gap = last == null ? Infinity : newest - Number(last);
      return { ok: gap > 40, actual: last == null ? "never seen" : `last round ${last}, ${gap} rounds ago` };
    },
  },
];

/** Run every claim and report what the archive says about it. */
export function checkClaims(logs) {
  return CLAIMS.map((c) => {
    let r;
    try {
      r = c.check(logs);
    } catch (err) {
      // A claim whose query throws is a broken claim, not a broken archive, and
      // it must say so rather than reading as a refutation.
      r = { ok: null, actual: `the check threw: ${err?.message ?? err}` };
    }
    return { id: c.id, says: c.says, measured: c.measured, ...r };
  });
}
