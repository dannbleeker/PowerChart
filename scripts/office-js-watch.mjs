#!/usr/bin/env node
/**
 * Watch the office-js tracker for defects that land on code this add-in ships.
 *
 * Five of the bugs this project now guards against — a text box deleting the
 * selected shape, a sync that never returns after an image insert, `Shape.group`
 * throwing, layouts unreadable on custom-template decks — were found in a SINGLE
 * manual sweep of the tracker, on one afternoon, after months of not looking.
 * That is not a process, it is luck with a good afternoon attached. Every one of
 * them sat under code that ships, open upstream, for a year or more.
 *
 * So: look every week, automatically, and say only what is NEW.
 *
 * **`KNOWN_ISSUES` is the load-bearing half.** Without it the sweep reports the
 * same twenty issues forever and is ignored within a month. With it, the report
 * is exactly "here is something nobody here has looked at", and the table
 * doubles as the record of what this project did about each one — which is the
 * question a reader actually has when they meet an issue number in a comment.
 *
 * Usage:
 *   node scripts/office-js-watch.mjs                 # fetch and report
 *   node scripts/office-js-watch.mjs --from a.json   # report from a saved page
 *   node scripts/office-js-watch.mjs --json          # machine-readable output
 *
 * The fetch half cannot run in every environment — an agent session is bound to
 * its own repositories and the office-js API is out of reach from one. The
 * matching half is pure and is what the tests drive, through `--from`.
 */

/**
 * The Office.js surface this add-in actually calls.
 *
 * The bar is the same as the host probe's: not "this is interesting about
 * Office.js" but "if this broke, code in this repo would be wrong". A term that
 * matches nothing we call produces noise, and noise is how a weekly report stops
 * being read.
 */
export const WATCHED_APIS = [
  { term: "addGeometricShape", why: "every native shape a chart is made of" },
  { term: "addTextBox", why: "every label, title and axis tick" },
  { term: "addLine", why: "connectors, axes, difference arrows" },
  { term: "addGroup", why: "a chart IS a group; ungrouped charts are the failure mode" },
  { term: "shape.group", why: "how the repair pass counts a chart's children" },
  { term: "setSelectedShapes", why: "the selection round trip behind Edit it" },
  { term: "setSelectedSlides", why: "showSlide, and every off-screen redraw" },
  { term: "getSelectedShapes", why: "reading the chart the user clicked" },
  { term: "getImageAsBase64", why: "the only way the add-in can see its own output" },
  { term: "insertSlidesFromBase64", why: "the fast path for a whole generated deck" },
  { term: "slideMasters", why: "blankLayoutId — every slide the add-in creates" },
  { term: "SlideLayout", why: "same, and reported broken on custom templates" },
  { term: "getItemOrNullObject", why: "how every slide and shape is resolved" },
  { term: "shape.tags", why: "config tags — what makes a chart re-editable" },
  { term: "pageSetup", why: "the slide size every placement depends on" },
  { term: "fill.setImage", why: "the degrade-to-picture path" },
  { term: "context.sync", why: "the one call every failure this project has met goes through" },
];

/**
 * Issues this project has already read and responded to.
 *
 * The value is not a status, it is WHAT WE DID — so an issue number met in a
 * code comment can be traced to a decision without re-reading the thread. "No
 * exposure" is a legitimate entry and an important one: it records that someone
 * checked, which is otherwise indistinguishable from nobody having looked.
 */
export const KNOWN_ISSUES = {
  1650: "A slide add whose sync never resolves though the slide lands. Every slide-add in powerpoint.ts is bounded and verified by re-reading the deck rather than by trusting the promise.",
  2328: "SlideMaster.shapes throws GeneralException on the web. Asked by the `layouts-readable` probe; blankLayoutId already degrades to the inherited layout.",
  2699: "A blank slide that is only blank to the eye. Handled by the demo path's blank re-read.",
  2775: "Adding a text box deletes the SELECTED shape, web only. Guarded by `dropShapeSelection` on the insert path, and asked by the self-test's `a selected shape survives an insert`.",
  2172:
    "addGeometricShape refused on a slide that is completely blank, web only; the reporter's workaround was to add a text element first. " +
    "Closed `Status: fixed`. NO EXPOSURE, and worth saying why rather than leaving it to the label: every slide this repo draws onto gets " +
    "its shapes through `addAndRenderItem`, which stamps a title text box before the chart's shapes on the demo path and otherwise draws " +
    "onto a slide the user already has. Found on 2026-08-07 while searching for the ShapeCollection.getItem refusal; it is a different " +
    "failure (the ADD, not the lookup) and it is not this one.",
  2780: "Carried as a caveat in the docs; no code depends on the behaviour.",
  2881: "Complex SVG renders wrong through the picture path. Why charts are native shapes rather than an image by default.",
  2903:
    "A stale shape proxy answers InvalidParam passed to GetItem(id). The reason `refreshShapes` and `settleAndTagChart` exist. " +
    "Re-read 2026-08-07 and it says more than we had taken from it: the report is about a freshly ADDED slide, not a stale shape — on Online " +
    "a slide that has been added and synced is not usable yet (text does not render, images land on the first slide instead), and the " +
    "reporter's only workaround is to wait a couple of seconds. Microsoft closed it `not planned`, so the wait is the fix that exists. " +
    "TRIED AND REVERTED: `addScratchSlide` settled for 2s on 2026-08-07 and the next real-host round answered 1 of 25 questions against 19 of 26 before it — the add landed, the wait ran, and the liveness check after it found nothing. This host resolves a fresh slide id ONCE and refuses it ever after, so waiting spends that one resolution later instead of buying time. Do not re-apply the workaround from this issue.",
  3014:
    "PowerPoint's API has no grouping story: creating and reading groups is a known parity gap, grouped shapes come back from getItem() " +
    "as type `unknown`, and sub-shapes cannot be reached. Open since 2022, `Status: in backlog`. No exposure to fix, but it is why " +
    "`ungroupedFallback` and CHART_PARTS_TAG exist at all — a chart that cannot be grouped has to carry its parts some other way — and " +
    "why `chooseGroupMembers` prefers an ungrouped, tagged chart over a failed addGroup that takes the tagging down with it.",
  3083: "setSelectedShapes([]) does not clear the selection on the web. clearShapeSelection re-selects the slide instead.",
  3269: "Office.js cannot read speaker notes at all. Recorded as a limitation; nothing here reads them.",
  3309: "SVG cannot be read back out of a shape. Same reason as #2881.",
  3698: "A picture cannot be inserted while a shape is selected, and setSelectedShapes([]) may never resolve. Both covered by dropShapeSelection and by the scenario ordering.",
  3826: "A freshly-added slide's layout shapes throw GeneralException. Asked by the `slide-layout-readable` probe.",
  4272: "A load of more than ~50 items answers short. Why the deck scan is paged at READBACK_PAGE.",
  4906: "SlideLayout.shapes throws on decks built from a custom template — the owner's case. Asked by the `layouts-readable` probe.",
  5022: "context.sync() runs indefinitely when shapes are re-read after an image insert. Asked by the `picture-then-shape-read` probe; drawDemoItem does exactly that sequence.",
  5101: "A placeholder keeps type: Placeholder when reused. NO EXPOSURE — this repo never reads a shape's type. Checked, not assumed.",
  5264: "A part of the object model Office.js cannot reach. Recorded as a limitation.",
  5849: "Shape.group throws GeneralException. Asked by the `group-of-existing-shape-readable` probe; countGroupChildrenPage reads groups exactly that way.",
  5896: "Reported alongside another SVG defect; same handling.",
  6363: "Carried in the docs as a known platform limitation.",
  2714: "setSelectedDataAsync converts points to pixels. NO EXPOSURE — this repo never calls it. Checked, not assumed.",
};

/** Lower-cased haystack for one issue. */
const haystack = (issue) => `${issue.title ?? ""}\n${issue.body ?? ""}`.toLowerCase();

/**
 * Which of these issues are worth a human's attention.
 *
 * Pure, and that is deliberate: the network half cannot run everywhere, and a
 * filter nobody can test is a filter nobody should trust. Pull requests are
 * dropped — the GitHub issues endpoint returns them too, and a PR against
 * office-js is not a defect report about it.
 */
export function freshIssues(issues, known = KNOWN_ISSUES, apis = WATCHED_APIS) {
  const out = [];
  for (const issue of issues ?? []) {
    if (issue.pull_request) continue;
    if (issue.number in known) continue;
    const text = haystack(issue);
    const hits = apis.filter((a) => text.includes(a.term.toLowerCase()));
    if (!hits.length) continue;
    out.push({
      number: issue.number,
      title: issue.title,
      state: issue.state,
      url: issue.html_url ?? `https://github.com/OfficeDev/office-js/issues/${issue.number}`,
      updated: issue.updated_at,
      hits: hits.map((h) => h.term),
      why: hits.map((h) => h.why),
    });
  }
  // Newest activity first — a reader works down the list and stops when they
  // recognise everything below.
  return out.sort((a, b) => String(b.updated ?? "").localeCompare(String(a.updated ?? "")));
}

/** The issue body a run posts. Markdown, and short enough to read on a phone. */
export function reportBody(fresh, checked) {
  if (!fresh.length) {
    return (
      `No office-js issue mentioning an API this add-in calls has appeared since the last sweep.\n\n` +
      `Checked ${checked} recently-updated issue(s) against ${WATCHED_APIS.length} watched calls ` +
      `and ${Object.keys(KNOWN_ISSUES).length} already answered.\n`
    );
  }
  const rows = fresh
    .map(
      (f) =>
        `### [#${f.number}](${f.url}) — ${f.title}\n\n` +
        `\`${f.hits.join("`, `")}\` · ${f.state} · updated ${f.updated}\n\n` +
        f.why.map((w) => `- ${w}`).join("\n"),
    )
    .join("\n\n");
  return (
    `${fresh.length} office-js issue(s) mention an API this add-in calls and are not yet in ` +
    `\`KNOWN_ISSUES\`.\n\n` +
    `For each: decide whether this repo is exposed, then add it to \`KNOWN_ISSUES\` in ` +
    `\`scripts/office-js-watch.mjs\` **with what was done about it** — including "no exposure", ` +
    `which records that somebody checked. An issue left out of that table comes back next week.\n\n` +
    `${rows}\n`
  );
}

/** Recently-updated issues from the tracker. Only reachable where the API is. */
async function fetchIssues(pages = 3, perPage = 100) {
  const all = [];
  for (let page = 1; page <= pages; page++) {
    const url =
      `https://api.github.com/repos/OfficeDev/office-js/issues` +
      `?state=all&sort=updated&direction=desc&per_page=${perPage}&page=${page}`;
    const res = await fetch(url, {
      headers: {
        Accept: "application/vnd.github+json",
        ...(process.env.GITHUB_TOKEN ? { Authorization: `Bearer ${process.env.GITHUB_TOKEN}` } : {}),
      },
    });
    if (!res.ok) throw new Error(`office-js issues page ${page}: HTTP ${res.status}`);
    const batch = await res.json();
    if (!Array.isArray(batch) || !batch.length) break;
    all.push(...batch);
  }
  return all;
}

async function main() {
  const args = process.argv.slice(2);
  const fromAt = args.indexOf("--from");
  const issues =
    fromAt >= 0
      ? JSON.parse(await (await import("fs/promises")).readFile(args[fromAt + 1], "utf8"))
      : await fetchIssues();
  const fresh = freshIssues(issues);
  if (args.includes("--json")) {
    process.stdout.write(JSON.stringify({ checked: issues.length, fresh }, null, 2));
    return;
  }
  process.stdout.write(reportBody(fresh, issues.length));
  // Non-zero when there is something to look at, so a workflow can branch on it
  // without parsing the body.
  if (fresh.length) process.exitCode = 3;
}

// Only when run directly — the tests import the pure half.
if (process.argv[1] && import.meta.url.endsWith(process.argv[1].split("/").pop())) {
  await main();
}
