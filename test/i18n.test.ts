// @vitest-environment jsdom
import { describe, expect, it } from "vitest";
import { localizePane, t } from "../src/taskpane/i18n";

describe("runtime string translation (t)", () => {
  it("translates keyed status strings for a supported language", () => {
    localizePane("de-DE");
    expect(t("Done.")).toBe("Fertig.");
    expect(t("Working…")).toBe("Arbeite…");
    expect(t("Chart loaded — edits will update it in place.")).toBe(
      "Diagramm geladen — Änderungen aktualisieren es direkt.",
    );
  });

  it("passes an unkeyed (interpolated) message through unchanged", () => {
    localizePane("de");
    expect(t("Inserted 3 of 5 shapes")).toBe("Inserted 3 of 5 shapes");
  });

  it("is a no-op for an unsupported language or none", () => {
    localizePane("fr");
    expect(t("Done.")).toBe("Done.");
    localizePane(undefined);
    expect(t("Done.")).toBe("Done.");
  });
});
