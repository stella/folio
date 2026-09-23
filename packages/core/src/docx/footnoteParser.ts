/**
 * Footnote/Endnote Parser - Parse footnotes.xml and endnotes.xml
 *
 * Footnotes and endnotes are stored in separate XML files within the DOCX package:
 * - word/footnotes.xml - Contains all footnote definitions
 * - word/endnotes.xml - Contains all endnote definitions
 *
 * Each note contains:
 * - An ID that matches references in document.xml (w:footnoteReference, w:endnoteReference)
 * - A type (normal, separator, continuationSeparator, continuationNotice)
 * - Content (paragraphs)
 *
 * The references in the document body are parsed by runParser as NoteReferenceContent.
 *
 * OOXML Reference:
 * - Footnote: w:footnote[@w:id][@w:type]
 * - Endnote: w:endnote[@w:id][@w:type]
 * - Content: w:p (paragraphs)
 */

import type {
  BlockSdt,
  Footnote,
  Endnote,
  Paragraph,
  Table,
  Theme,
  RelationshipMap,
  MediaFile,
} from "../types/document";
import { PARSE_WARNING_CODES } from "@stll/docx-core/model";

import type { NumberingMap } from "./numberingParser";
import type { ParseContext } from "./parseContext";
import { type PreviewLedger, standalonePreviewLedger } from "./previewBudget";
import { blockPlainText } from "./blockPlainText";
import { parseParagraph } from "./paragraphParser";
import { captureSdtSiblingMarkers, parseSdtProperties } from "./sdtProperties";
import type { StyleMap } from "./styleParser";
import { parseTable } from "./tableParser";
import {
  findChildren,
  findWordprocessingChild,
  getAttributes,
  getChildElements,
  getLocalName,
  parseXml,
  type XmlElement,
} from "./xmlParser";

// ============================================================================
// FOOTNOTE MAP INTERFACE
// ============================================================================

/**
 * Footnote map returned by parseFootnotes
 */
export type FootnoteMap = {
  /** All footnotes indexed by ID */
  byId: Map<number, Footnote>;

  /** Array of all footnotes in document order */
  footnotes: Footnote[];

  /** Get footnote by ID */
  getFootnote: (id: number) => Footnote | undefined;

  /** Check if footnote exists */
  hasFootnote: (id: number) => boolean;

  /** Get all normal (non-separator) footnotes */
  getNormalFootnotes: () => Footnote[];

  /** Get separator footnote if exists */
  getSeparator: () => Footnote | undefined;

  /** Get continuation separator if exists */
  getContinuationSeparator: () => Footnote | undefined;
};

/**
 * Endnote map returned by parseEndnotes
 */
export type EndnoteMap = {
  /** All endnotes indexed by ID */
  byId: Map<number, Endnote>;

  /** Array of all endnotes in document order */
  endnotes: Endnote[];

  /** Get endnote by ID */
  getEndnote: (id: number) => Endnote | undefined;

  /** Check if endnote exists */
  hasEndnote: (id: number) => boolean;

  /** Get all normal (non-separator) endnotes */
  getNormalEndnotes: () => Endnote[];

  /** Get separator endnote if exists */
  getSeparator: () => Endnote | undefined;

  /** Get continuation separator if exists */
  getContinuationSeparator: () => Endnote | undefined;
};

// ============================================================================
// NOTE TYPE PARSING
// ============================================================================

/**
 * Parse note type attribute
 */
function parseNoteType(
  typeAttr: string | null,
): "normal" | "separator" | "continuationSeparator" | "continuationNotice" {
  switch (typeAttr) {
    case "separator":
      return "separator";
    case "continuationSeparator":
      return "continuationSeparator";
    case "continuationNotice":
      return "continuationNotice";
    default:
      return "normal";
  }
}

function getNoteAttribute(element: XmlElement, localName: "id" | "type"): string | null {
  for (const [name, value] of Object.entries(getAttributes(element))) {
    if (getLocalName(name) === localName) {
      return value;
    }
  }

  return null;
}

function parseNoteId(element: XmlElement): number {
  const id = getNoteAttribute(element, "id");
  if (id === null) {
    return 0;
  }

  const parsed = Number.parseInt(id, 10);
  return Number.isNaN(parsed) ? 0 : parsed;
}

// ============================================================================
// FOOTNOTE PARSING
// ============================================================================

function parseNoteBlockContent(
  element: XmlElement,
  styles: StyleMap | null,
  theme: Theme | null,
  numbering: NumberingMap | null,
  rels: RelationshipMap | null,
  media: Map<string, MediaFile> | null,
  previews: PreviewLedger,
): (Paragraph | Table | BlockSdt)[] {
  const blocks: (Paragraph | Table | BlockSdt)[] = [];

  for (const child of getChildElements(element)) {
    const localName = getLocalName(child.name ?? "");
    if (localName === "p") {
      blocks.push(parseParagraph(child, styles, theme, numbering, rels, media, { previews }));
    } else if (localName === "tbl") {
      const table = parseTable(child, styles, theme, numbering, rels, media, { previews });
      if (table) {
        blocks.push(table);
      }
    } else if (localName === "sdt") {
      // Recurse into sdtContent so SDT children inside notes are
      // recognized; otherwise the note body silently drops citation
      // slots and bound metadata controls.
      const sdtPr = findWordprocessingChild(child, "sdtPr");
      const sdtEndPr = findWordprocessingChild(child, "sdtEndPr");
      const sdtContent = findWordprocessingChild(child, "sdtContent");
      const properties = parseSdtProperties(sdtPr, sdtEndPr);
      const captured = captureSdtSiblingMarkers(child);
      if (captured.before.length > 0) {
        properties.rawSdtChildrenBeforeContent = captured.before;
      }
      if (captured.after.length > 0) {
        properties.rawSdtChildrenAfterContent = captured.after;
      }
      blocks.push({
        type: "blockSdt",
        properties,
        content: sdtContent
          ? parseNoteBlockContent(sdtContent, styles, theme, numbering, rels, media, previews)
          : [],
      });
    }
  }

  return blocks;
}

/**
 * Parse a single footnote element (w:footnote)
 */
function parseFootnote(
  element: XmlElement,
  styles: StyleMap | null,
  theme: Theme | null,
  numbering: NumberingMap | null,
  rels: RelationshipMap | null,
  media: Map<string, MediaFile> | null,
  previews: PreviewLedger,
): Footnote {
  const id = parseNoteId(element);
  const typeAttr = getNoteAttribute(element, "type");
  const noteType = parseNoteType(typeAttr);

  const content = parseNoteBlockContent(element, styles, theme, numbering, rels, media, previews);

  return {
    type: "footnote",
    id,
    noteType,
    content,
  };
}

/**
 * Parse footnotes.xml
 *
 * @param footnotesXml - The raw XML content of word/footnotes.xml
 * @param styles - Parsed style map for applying styles
 * @param theme - Parsed theme for color resolution
 * @param numbering - Parsed numbering definitions for lists
 * @param rels - Relationships for resolving hyperlinks
 * @param media - Media files for images
 * @returns FootnoteMap with all footnotes
 */
export function parseFootnotes(
  footnotesXml: string | null,
  styles: StyleMap | null = null,
  theme: Theme | null = null,
  numbering: NumberingMap | null = null,
  rels: RelationshipMap | null = null,
  media: Map<string, MediaFile> | null = null,
  context?: ParseContext,
  previews: PreviewLedger = standalonePreviewLedger(),
): FootnoteMap {
  const byId = new Map<number, Footnote>();
  const footnotes: Footnote[] = [];

  if (!footnotesXml) {
    return createFootnoteMap(byId, footnotes);
  }

  const doc = parseXml(footnotesXml);

  // Find the root footnotes element
  const rootElement = doc.elements?.find(
    (el: XmlElement) =>
      el.type === "element" && (el.name === "w:footnotes" || el.name?.endsWith(":footnotes")),
  );

  if (!rootElement) {
    return createFootnoteMap(byId, footnotes);
  }

  // Parse all footnote elements
  const footnoteElements = findChildren(rootElement, "w", "footnote");

  for (const fnEl of footnoteElements) {
    const footnote = parseFootnote(fnEl, styles, theme, numbering, rels, media, previews);
    // A `w:footnoteReference` names one `w:id`, so a repeat of an id is a note
    // nothing can reference. Word resolves such a reference to the first
    // definition; keeping the first here is also what `mergeFootnoteMaps` does,
    // and it keeps the array and the id index saying the same thing.
    if (byId.has(footnote.id)) {
      context?.warn({
        code: PARSE_WARNING_CODES.duplicateNoteId,
        element: "w:footnote",
        at: `w:id ${String(footnote.id)}`,
      });
      previews.release(footnote);
      continue;
    }
    byId.set(footnote.id, footnote);
    footnotes.push(footnote);
  }

  return createFootnoteMap(byId, footnotes);
}

/**
 * Create FootnoteMap object with helper methods
 */
function createFootnoteMap(byId: Map<number, Footnote>, footnotes: Footnote[]): FootnoteMap {
  return {
    byId,
    footnotes,

    getFootnote(id: number): Footnote | undefined {
      return byId.get(id);
    },

    hasFootnote(id: number): boolean {
      return byId.has(id);
    },

    getNormalFootnotes(): Footnote[] {
      return footnotes.filter((fn) => fn.noteType === "normal");
    },

    getSeparator(): Footnote | undefined {
      return footnotes.find((fn) => fn.noteType === "separator");
    },

    getContinuationSeparator(): Footnote | undefined {
      return footnotes.find((fn) => fn.noteType === "continuationSeparator");
    },
  };
}

// ============================================================================
// ENDNOTE PARSING
// ============================================================================

/**
 * Parse a single endnote element (w:endnote)
 */
function parseEndnote(
  element: XmlElement,
  styles: StyleMap | null,
  theme: Theme | null,
  numbering: NumberingMap | null,
  rels: RelationshipMap | null,
  media: Map<string, MediaFile> | null,
  previews: PreviewLedger,
): Endnote {
  const id = parseNoteId(element);
  const typeAttr = getNoteAttribute(element, "type");
  const noteType = parseNoteType(typeAttr);

  const content = parseNoteBlockContent(element, styles, theme, numbering, rels, media, previews);

  return {
    type: "endnote",
    id,
    noteType,
    content,
  };
}

/**
 * Parse endnotes.xml
 *
 * @param endnotesXml - The raw XML content of word/endnotes.xml
 * @param styles - Parsed style map for applying styles
 * @param theme - Parsed theme for color resolution
 * @param numbering - Parsed numbering definitions for lists
 * @param rels - Relationships for resolving hyperlinks
 * @param media - Media files for images
 * @returns EndnoteMap with all endnotes
 */
export function parseEndnotes(
  endnotesXml: string | null,
  styles: StyleMap | null = null,
  theme: Theme | null = null,
  numbering: NumberingMap | null = null,
  rels: RelationshipMap | null = null,
  media: Map<string, MediaFile> | null = null,
  context?: ParseContext,
  previews: PreviewLedger = standalonePreviewLedger(),
): EndnoteMap {
  const byId = new Map<number, Endnote>();
  const endnotes: Endnote[] = [];

  if (!endnotesXml) {
    return createEndnoteMap(byId, endnotes);
  }

  const doc = parseXml(endnotesXml);

  // Find the root endnotes element
  const rootElement = doc.elements?.find(
    (el: XmlElement) =>
      el.type === "element" && (el.name === "w:endnotes" || el.name?.endsWith(":endnotes")),
  );

  if (!rootElement) {
    return createEndnoteMap(byId, endnotes);
  }

  // Parse all endnote elements
  const endnoteElements = findChildren(rootElement, "w", "endnote");

  for (const enEl of endnoteElements) {
    const endnote = parseEndnote(enEl, styles, theme, numbering, rels, media, previews);
    // First definition wins, as for footnotes above.
    if (byId.has(endnote.id)) {
      context?.warn({
        code: PARSE_WARNING_CODES.duplicateNoteId,
        element: "w:endnote",
        at: `w:id ${String(endnote.id)}`,
      });
      previews.release(endnote);
      continue;
    }
    byId.set(endnote.id, endnote);
    endnotes.push(endnote);
  }

  return createEndnoteMap(byId, endnotes);
}

/**
 * Create EndnoteMap object with helper methods
 */
function createEndnoteMap(byId: Map<number, Endnote>, endnotes: Endnote[]): EndnoteMap {
  return {
    byId,
    endnotes,

    getEndnote(id: number): Endnote | undefined {
      return byId.get(id);
    },

    hasEndnote(id: number): boolean {
      return byId.has(id);
    },

    getNormalEndnotes(): Endnote[] {
      return endnotes.filter((en) => en.noteType === "normal");
    },

    getSeparator(): Endnote | undefined {
      return endnotes.find((en) => en.noteType === "separator");
    },

    getContinuationSeparator(): Endnote | undefined {
      return endnotes.find((en) => en.noteType === "continuationSeparator");
    },
  };
}

// Re-export note properties parsers for backward compatibility
export { parseFootnoteProperties, parseEndnoteProperties } from "./notePropertiesParser";

// ============================================================================
// UTILITY FUNCTIONS
// ============================================================================

/**
 * Get plain text content of a footnote.
 *
 * Uses the accepted tracked-change view and recurses through every note block.
 */
export function getFootnoteText(footnote: Footnote): string {
  return blockPlainText(footnote.content);
}

/**
 * Get plain text content of an endnote.
 */
export function getEndnoteText(endnote: Endnote): string {
  return blockPlainText(endnote.content);
}

/**
 * Check if a footnote is a separator (not regular content)
 */
export function isSeparatorFootnote(footnote: Footnote): boolean {
  return (
    footnote.noteType === "separator" ||
    footnote.noteType === "continuationSeparator" ||
    footnote.noteType === "continuationNotice"
  );
}

/**
 * Check if an endnote is a separator (not regular content)
 */
export function isSeparatorEndnote(endnote: Endnote): boolean {
  return (
    endnote.noteType === "separator" ||
    endnote.noteType === "continuationSeparator" ||
    endnote.noteType === "continuationNotice"
  );
}

// Note: display numbers are NOT derivable from footnotes.xml file order.
// Word numbers notes by reference order in the document body; see
// `computeNoteDisplayNumbers` in `layout-bridge/convert/footnoteLayout.ts`.

/**
 * Create an empty footnote map
 */
export function createEmptyFootnoteMap(): FootnoteMap {
  return createFootnoteMap(new Map(), []);
}

/**
 * Create an empty endnote map
 */
export function createEmptyEndnoteMap(): EndnoteMap {
  return createEndnoteMap(new Map(), []);
}

/**
 * Merge multiple footnote maps (e.g., from different documents)
 */
export function mergeFootnoteMaps(...maps: FootnoteMap[]): FootnoteMap {
  const byId = new Map<number, Footnote>();
  const footnotes: Footnote[] = [];

  for (const map of maps) {
    for (const fn of map.footnotes) {
      if (!byId.has(fn.id)) {
        byId.set(fn.id, fn);
        footnotes.push(fn);
      }
    }
  }

  return createFootnoteMap(byId, footnotes);
}

/**
 * Merge multiple endnote maps (e.g., from different documents)
 */
export function mergeEndnoteMaps(...maps: EndnoteMap[]): EndnoteMap {
  const byId = new Map<number, Endnote>();
  const endnotes: Endnote[] = [];

  for (const map of maps) {
    for (const en of map.endnotes) {
      if (!byId.has(en.id)) {
        byId.set(en.id, en);
        endnotes.push(en);
      }
    }
  }

  return createEndnoteMap(byId, endnotes);
}
