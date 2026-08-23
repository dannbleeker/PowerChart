// @vitest-environment node
import { describe, expect, it } from "vitest";
import { readFileSync } from "fs";
import { isGroupReadRefusal } from "../src/render/powerpoint";

/**
 * WHAT THIS COST: an exception and a failed resolve in every round for 62 rounds.
 *
 * Build 408d00b (#642) taught the in-place update to read a group's members
 * through `shape.group.shapes`. The archive says what happened next:
 *
 *     rounds 023-142   0.00 GeneralException/round —  0 of 118 rounds
 *     rounds 143-204   1.00 GeneralException/round — 62 of 62 rounds
 *
 * Every one is `errorLocation: Shape.group` on `var group =
 * itemOrNullObject1.group;`. The probe sheet had settled the same fact from the
 * other side for just as long: `group-of-existing-shape-readable = no-group-id`,
 * `group-reports-its-children = threw`.
 *
 * `queueGroupMembers` wraps the access in try/catch and its comment calls the
 * attempt silent, "because on a host that has not implemented it the property
 * access itself can throw synchronously". It does not. The access only QUEUES a
 * proxy read; the throw arrives at the sync, under another label, outside that
 * try, inside a batch it can poison.
 */
describe("the Shape.group refusal", () => {
  it("recognises the refusal this host has produced every round since 408d00b", () => {
    // The real shape of it, from rounds/204-d2236ec.json.
    const real =
      'GeneralException | at=resolving the charts\' shapes | code=GeneralException | debugInfo={"code":"GeneralException",' +
      '"message":"GeneralException","errorLocation":"Shape.group","statement":"var group = itemOrNullObject1.group;"}';
    expect(isGroupReadRefusal(real)).toBe(true);
  });

  it("matches the errorLocation FIELD, not the phrase wherever it turns up", () => {
    // MY FIRST VERSION OF THIS TEST WAS VACUOUS, and mutation is what said so.
    //
    // It fed an error whose `surroundingStatements` held
    // `var group = itemOrNullObject1.group;` and claimed a loose matcher would
    // trip on it. Loosening the predicate to /Shape\.group/ left the test GREEN,
    // because a real Office.js statement echo never contains the TYPE name
    // "Shape.group" at all — only `<proxy>.group`. The fixture demonstrated
    // nothing, and the comment above it claimed a protection it had not shown.
    //
    // What is demonstrable, and is the real property: the phrase can appear in a
    // host MESSAGE while the error comes from somewhere else. Latching a
    // capability off on that would disable a working feature for the rest of the
    // session on the strength of a sentence.
    const elsewhere =
      'GeneralException | debugInfo={"code":"GeneralException","errorLocation":"ShapeCollection.addGroup",' +
      '"message":"Shape.group is not supported here"}';
    expect(isGroupReadRefusal(elsewhere), "the phrase outside errorLocation latched the refusal").toBe(false);

    // And the real one still matches, so narrowing to the field lost nothing.
    expect(isGroupReadRefusal('debugInfo={"errorLocation":"Shape.group"}')).toBe(true);
  });

  it("does not fire on unrelated errors, or on nothing", () => {
    expect(isGroupReadRefusal('GeneralException | debugInfo={"errorLocation":"ShapeCollection.addGroup"}')).toBe(false);
    expect(isGroupReadRefusal("")).toBe(false);
    expect(isGroupReadRefusal(undefined as unknown as string)).toBe(false);
  });

  it("is actually consulted before queueing the read, not merely defined", () => {
    // A predicate nothing calls is decoration. Asserted on the source because
    // the latch is module state set from an error path a unit test cannot reach.
    const src = readFileSync("src/render/powerpoint.ts", "utf8");
    expect(src, "the classifier stopped setting the latch").toContain(
      "if (isGroupReadRefusal(text)) groupReadRefused = true;",
    );
    expect(src, "queueGroupMembers stopped checking the latch").toContain("if (groupReadRefused) return undefined;");
  });
});
