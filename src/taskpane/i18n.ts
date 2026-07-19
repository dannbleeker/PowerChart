/**
 * Minimal pane localization: translate static UI strings when Office reports
 * a matching display language. Chart output localization (separators) is
 * handled separately via NumberFormat.locale.
 */

const DE: Record<string, string> = {
  "1 · Chart type": "1 · Diagrammtyp",
  "2 · Data": "2 · Daten",
  "3 · Decorations": "3 · Dekorationen",
  "Preview & size": "Vorschau & Größe",
  Elements: "Elemente",
  "Automation (JSON)": "Automatisierung (JSON)",
  Agenda: "Agenda",
  "Insert into slide": "In Folie einfügen",
  "Insert as new": "Als neu einfügen",
  "Edit selected chart": "Ausgewähltes Diagramm bearbeiten",
  "Same scale (deck)": "Gleiche Skala (Deck)",
  "Same scale (selection)": "Gleiche Skala (Auswahl)",
  "Download SVG": "SVG herunterladen",
  "Download PNG": "PNG herunterladen",
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
  "Grand total": "Gesamtsumme",
  "Category labels": "Kategoriebeschriftungen",
  "Value axis": "Werteachse",
  Gridlines: "Gitterlinien",
  "Horizontal (bar)": "Horizontal (Balken)",
  "Connector lines": "Verbindungslinien",
  "100% = note": "100%-Hinweis",
  "Auto-update chart": "Diagramm automatisch aktualisieren",
  Insert: "Einfügen",
  "Edit it": "Bearbeiten",
  "Total row": "Summenzeile",
  "Datamark axis (ticks only)": "Datenmarken-Achse (nur Striche)",
  "Use deck theme": "Design der Präsentation",
  // Chart-type families (grouped picker) + its search.
  "Columns & bars": "Säulen & Balken",
  "Line & area": "Linien & Flächen",
  "Parts of a whole": "Anteile am Ganzen",
  Distribution: "Verteilung",
  Correlation: "Korrelation",
  "Matrix & spatial": "Matrix & Raum",
  "Search chart types…": "Diagrammtypen suchen…",
  "No chart type matches that search.": "Kein Diagrammtyp passt zu dieser Suche.",
  // Format groups (Layout is spelled the same in German).
  Labels: "Beschriftungen",
  "Axes & scale": "Achsen & Skala",
  Analysis: "Analyse",
  "Colours & style": "Farben & Stil",
  // Datasheet help.
  "Paste straight from Excel — special data rows": "Direkt aus Excel einfügen — besondere Datenzeilen",
  // Runtime status messages (announced via the aria-live status strip). Routed
  // through t(); interpolated messages that carry a count or an error detail are
  // not keyed and stay in English until the status catalog grows params.
  "Working…": "Arbeite…",
  "Done.": "Fertig.",
  "Chart loaded — edits will update it in place.": "Diagramm geladen — Änderungen aktualisieren es direkt.",
  "Style exported — share the JSON as your corporate style file.":
    "Stil exportiert — teilen Sie das JSON als Ihre Firmen-Stildatei.",
  "Style imported — applied to every chart from this pane.":
    "Stil importiert — auf jedes Diagramm aus diesem Bereich angewendet.",
  "The selection is not a PowerChart — select an inserted chart group first.":
    "Die Auswahl ist kein PowerChart — wählen Sie zuerst eine eingefügte Diagrammgruppe.",
};

const DICTS: Record<string, Record<string, string>> = { de: DE };

let activeDict: Record<string, string> | undefined;

/* The picker families, Format group names, the datasheet-help summary, and the
   search placeholder are added to the base selector; input placeholders are
   translated too. */
// .acc-title is the accordion step heading. It sits in a <span> inside the
// <summary>, and translateTree only rewrites an element's DIRECT text, so
// matching the <summary> alone never reaches it.
const LOCALIZE_SELECTOR =
  "h2, button, label, .banner, option, .tagline, figcaption, summary, .acc-title, .group-label, .fgroup-name, .no-type-result";

function translateTree(root: ParentNode, dict: Record<string, string>): void {
  for (const el of root.querySelectorAll<HTMLElement>(LOCALIZE_SELECTOR)) {
    // Only translate an element's direct text, so child inputs/spans survive.
    for (const child of el.childNodes) {
      if (child.nodeType === Node.TEXT_NODE) {
        const t = child.textContent?.trim();
        if (t && dict[t]) child.textContent = child.textContent!.replace(t, dict[t]);
      }
    }
  }
  for (const input of root.querySelectorAll<HTMLInputElement>("input[placeholder]")) {
    const p = input.placeholder.trim();
    if (p && dict[p]) input.placeholder = dict[p];
  }
}

/** Translate visible UI text in place; no-op for unsupported languages. */
export function localizePane(language: string | undefined): void {
  activeDict = DICTS[(language ?? "").slice(0, 2).toLowerCase()];
  if (activeDict) translateTree(document, activeDict);
}

/** Re-apply the active translation to a freshly-rendered subtree — the chart
 *  gallery and Format groups are rebuilt in English on every render. */
export function localizeTree(root: ParentNode): void {
  if (activeDict) translateTree(root, activeDict);
}

/**
 * Translate a runtime string (a status message) built in code rather than markup
 * — the DOM-sweep translateTree never sees these. Returns the source string
 * unchanged when there's no active language or no matching entry, so callers can
 * wrap unconditionally.
 */
export function t(s: string): string {
  return activeDict?.[s] ?? s;
}
