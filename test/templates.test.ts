// @vitest-environment jsdom
import { beforeEach, afterEach, describe, expect, it, vi } from "vitest";
import { readFileSync } from "fs";
import { BUILTIN_TEMPLATES } from "../src/taskpane/templates";
import type { ChartConfig } from "../src/core/types";

/**
 * Saved chart templates — a whole user-facing feature with no tests at all.
 *
 * Save the chart you have set up, pick it again later, delete it when you are
 * done. It rides on `localStorage` and on a plain object keyed by whatever the
 * user typed, which is the combination this repo has been bitten by before:
 * "object lookups keyed by a config string must use
 * `Object.prototype.hasOwnProperty.call`" is in the project's own notes, with a
 * list of the tables it was applied to. This table was not on the list.
 */

const $ = <T extends HTMLElement>(id: string) => document.getElementById(id) as T;
const TEMPLATES_KEY = "powerchart-templates";

async function bootPane() {
  window.localStorage.clear();
  window.history.replaceState({}, "", "/taskpane.html");
  const parsed = new DOMParser().parseFromString(readFileSync("src/taskpane/taskpane.html", "utf8"), "text/html");
  parsed.querySelectorAll("script").forEach((s) => s.remove());
  document.body.innerHTML = parsed.body.innerHTML;
  vi.resetModules();
  await import("../src/taskpane/app");
}

/** Re-open the pane WITHOUT clearing storage — what a reload really is. */
async function reopenPane() {
  const parsed = new DOMParser().parseFromString(readFileSync("src/taskpane/taskpane.html", "utf8"), "text/html");
  parsed.querySelectorAll("script").forEach((s) => s.remove());
  document.body.innerHTML = parsed.body.innerHTML;
  vi.resetModules();
  await import("../src/taskpane/app");
}

/** Save the pane's current chart under `name`, as the button does. */
function saveAs(name: string) {
  vi.spyOn(window, "prompt").mockReturnValue(name);
  $("template-save").click();
}

/** The user-template names the picker is offering. */
const offered = () =>
  [...$<HTMLSelectElement>("template-list").querySelectorAll("option")]
    .map((o) => o.value)
    .filter((v) => v.startsWith("user:"))
    .map((v) => v.slice("user:".length));

const pick = (value: string) => {
  const sel = $<HTMLSelectElement>("template-list");
  sel.value = value;
  sel.dispatchEvent(new Event("change"));
};

const stored = (): Record<string, unknown> => JSON.parse(window.localStorage.getItem(TEMPLATES_KEY) ?? "{}");

beforeEach(async () => {
  await bootPane();
});
afterEach(() => {
  vi.restoreAllMocks();
  window.localStorage.clear();
});

describe("saved chart templates", () => {
  it("saves the chart you have, offers it back, and loads it", async () => {
    ($("chart-title") as HTMLInputElement).value = "Q3 revenue";
    ($("chart-title") as HTMLInputElement).dispatchEvent(new Event("input"));
    ($("chart-w") as HTMLInputElement).value = "640";
    ($("chart-w") as HTMLInputElement).dispatchEvent(new Event("input"));
    saveAs("my layout");

    expect(offered(), "the saved template was not offered back").toContain("my layout");
    expect(Object.keys(stored())).toContain("my layout");

    // Change the pane, then pick the template — it must come back.
    ($("chart-w") as HTMLInputElement).value = "300";
    ($("chart-w") as HTMLInputElement).dispatchEvent(new Event("input"));
    pick("user:my layout");
    expect(($("chart-w") as HTMLInputElement).value, "picking a template did not restore its size").toBe("640");
  });

  it("survives a reload — the point of saving one", async () => {
    saveAs("keeps");
    await reopenPane();
    expect(offered(), "the template did not survive reopening the pane").toContain("keeps");
  });

  it("deletes a user template, and refuses to delete a starter", async () => {
    saveAs("throwaway");
    pick("user:throwaway");
    $("template-delete").click();
    expect(offered()).not.toContain("throwaway");
    expect(Object.keys(stored())).not.toContain("throwaway");

    // A starter is not the user's to delete, and the guard is a string prefix
    // — the kind that stops working the moment someone renames the option
    // values, so it is pinned.
    const starter = BUILTIN_TEMPLATES[0];
    pick(`builtin:${starter.name}`);
    $("template-delete").click();
    const still = [...$<HTMLSelectElement>("template-list").querySelectorAll("option")].map((o) => o.value);
    expect(still, "a built-in starter was deleted").toContain(`builtin:${starter.name}`);
  });

  it("loads a built-in starter", () => {
    const starter = BUILTIN_TEMPLATES.find((t) => (t.config as ChartConfig).kind);
    expect(starter, "no starter carries a kind to check against").toBeTruthy();
    pick(`builtin:${starter!.name}`);
    // The pane's own type summary is the visible proof it took.
    expect($("type-sub").textContent, "picking a starter changed nothing").toBeTruthy();
  });

  it("keeps a template whose name collides with a JavaScript builtin", async () => {
    // `all[name] = config` on an object from `JSON.parse` is a plain assignment
    // for every name but one. For `__proto__` it hits the inherited SETTER and
    // re-parents the object instead of storing anything — so the template reads
    // back correctly for the rest of the session, `JSON.stringify` writes it
    // out as `{}`, and it is gone the next time the pane opens. Saved,
    // apparently fine, silently lost.
    //
    // Nobody names a template `__proto__` on purpose. That is not the point:
    // the point is that this is the third table in this repo to be keyed by a
    // user-supplied string, and the project's own notes say to guard every one.
    saveAs("__proto__");
    expect(offered(), "the odd name was not even offered in the same session").toContain("__proto__");

    await reopenPane();
    expect(offered(), "a saved template vanished on reload, with nothing said").toContain("__proto__");
  });

  it("cannot be tricked into loading Object.prototype's members as a chart", () => {
    // The read side of the same table. `loadTemplates()[name]` for a name that
    // is not stored — `constructor`, `toString` — reaches Object.prototype and
    // hands back a FUNCTION, which is truthy, so the pane would apply it as a
    // config. Nothing offers those names today because the picker is built
    // from `Object.keys`, which is exactly the sort of accident that stops
    // being true when someone changes how the list is built.
    window.localStorage.setItem(TEMPLATES_KEY, JSON.stringify({ real: { kind: "clustered" } }));
    // A width that is NOT the default, so "applied the wrong thing" and
    // "applied nothing" cannot look the same. They did in the first version of
    // this test, which read the default back and called it a pass.
    ($("chart-w") as HTMLInputElement).value = "640";
    ($("chart-w") as HTMLInputElement).dispatchEvent(new Event("input"));
    const before = ($("chart-w") as HTMLInputElement).value;
    expect(before).toBe("640");
    // Reach past the picker and ask for the dangerous name directly.
    const sel = $<HTMLSelectElement>("template-list");
    const opt = document.createElement("option");
    opt.value = "user:constructor";
    sel.appendChild(opt);
    pick("user:constructor");
    expect(($("chart-w") as HTMLInputElement).value, "applied something that was not a template").toBe(before);
  });
});
