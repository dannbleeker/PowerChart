/**
 * Minimal pane localization: translate static UI strings when Office reports
 * a matching display language. Chart output localization (separators) is
 * handled separately via NumberFormat.locale.
 */

const DE: Record<string, string> = {
  "1 · Chart type": "1 · Diagrammtyp",
  "2 · Data": "2 · Daten",
  "3 · Decorations": "3 · Dekorationen",
  Preview: "Vorschau",
  Elements: "Elemente",
  "Automation (JSON)": "Automatisierung (JSON)",
  Agenda: "Agenda",
  "Insert into slide": "In Folie einfügen",
  "Insert as new": "Als neu einfügen",
  "Edit selected chart": "Ausgewähltes Diagramm bearbeiten",
  "Same scale (deck)": "Gleiche Skala (Deck)",
  "Same scale (selection)": "Gleiche Skala (Auswahl)",
  "Download SVG": "SVG herunterladen",
  "Update chart": "Diagramm aktualisieren",
  "Insert agenda slides": "Agenda-Folien einfügen",
  "Export current": "Aktuelles exportieren",
  Import: "Importieren",
  "Insert batch": "Stapel einfügen",
  "Export style": "Stil exportieren",
  "Import style": "Stil importieren",
  "Save as template": "Als Vorlage speichern",
  Delete: "Löschen",
  "Chart title": "Diagrammtitel",
  "Segment labels": "Segmentbeschriftungen",
  "Series labels": "Reihenbeschriftungen",
  "Column totals": "Säulensummen",
  "Category labels": "Kategoriebeschriftungen",
  "Value axis": "Werteachse",
  Gridlines: "Gitterlinien",
  "Horizontal (bar)": "Horizontal (Balken)",
  "Auto-update chart": "Diagramm automatisch aktualisieren",
  Insert: "Einfügen",
  "Edit it": "Bearbeiten",
};

const DICTS: Record<string, Record<string, string>> = { de: DE };

/** Translate visible UI text in place; no-op for unsupported languages. */
export function localizePane(language: string | undefined): void {
  const dict = DICTS[(language ?? "").slice(0, 2).toLowerCase()];
  if (!dict) return;
  const selector = "h2, button, label, .banner, option, .tagline, figcaption";
  for (const el of document.querySelectorAll<HTMLElement>(selector)) {
    // Only translate elements whose direct text matches, keeping child inputs.
    for (const child of el.childNodes) {
      if (child.nodeType === Node.TEXT_NODE) {
        const t = child.textContent?.trim();
        if (t && dict[t]) child.textContent = child.textContent!.replace(t, dict[t]);
      }
    }
  }
}
