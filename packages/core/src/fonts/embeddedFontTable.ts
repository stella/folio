/**
 * Parses the embedded-font references out of `word/fontTable.xml`.
 *
 * The shared docx model (`@stll/docx-core`) keeps only the embed relationship
 * ids and drops the `w:fontKey` GUID, which is exactly the value needed to
 * de-obfuscate the `.odttf` binary. This parser therefore reads the font table
 * directly and keeps the full embed reference (rel id, key, subsetted flag).
 *
 * Structure:
 * ```xml
 * <w:fonts>
 *   <w:font w:name="My Brand Sans">
 *     <w:embedRegular r:id="rId1" w:fontKey="{GUID}" w:subsetted="true"/>
 *   </w:font>
 * </w:fonts>
 * ```
 *
 * Ported from eigenpal/docx-editor (see NOTICE.md).
 */

import { findChild, findChildren, getAttribute, parseXmlDocument } from "../docx/xmlParser";
import type { XmlElement } from "../docx/xmlParser";

/** A single embedded face referenced from a `w:embed*` element. */
export type EmbeddedFontFaceRef = {
  /** Relationship id (`r:id`) resolved via `word/_rels/fontTable.xml.rels`. */
  relId: string;
  /** Obfuscation key GUID (`w:fontKey`), e.g. `{XXXXXXXX-...}`. */
  fontKey?: string;
  /** Whether the embedded face is subsetted (`w:subsetted`). */
  subsetted?: boolean;
};

/** One `<w:font>` entry with whichever of the four embed faces it declares. */
export type EmbeddedFontEntry = {
  name: string;
  embedRegular?: EmbeddedFontFaceRef;
  embedBold?: EmbeddedFontFaceRef;
  embedItalic?: EmbeddedFontFaceRef;
  embedBoldItalic?: EmbeddedFontFaceRef;
};

/** The four `w:embed*` local names paired with the entry field they populate. */
const EMBED_ELEMENTS = [
  { element: "embedRegular", field: "embedRegular" },
  { element: "embedBold", field: "embedBold" },
  { element: "embedItalic", field: "embedItalic" },
  { element: "embedBoldItalic", field: "embedBoldItalic" },
] as const;

function parseEmbed(font: XmlElement, localName: string): EmbeddedFontFaceRef | undefined {
  const el = findChild(font, "w", localName);
  if (!el) {
    return undefined;
  }

  const relId = getAttribute(el, "r", "id");
  if (!relId) {
    // An embed with no relationship cannot resolve to a binary; skip it.
    return undefined;
  }

  const ref: EmbeddedFontFaceRef = { relId };

  const fontKey = getAttribute(el, "w", "fontKey");
  if (fontKey) {
    ref.fontKey = fontKey;
  }

  const subsetted = getAttribute(el, "w", "subsetted");
  if (subsetted === "true" || subsetted === "1") {
    ref.subsetted = true;
  }

  return ref;
}

function parseFontEntry(font: XmlElement): EmbeddedFontEntry | null {
  const name = getAttribute(font, "w", "name");
  if (!name) {
    return null;
  }

  const entry: EmbeddedFontEntry = { name };
  for (const { element, field } of EMBED_ELEMENTS) {
    const ref = parseEmbed(font, element);
    if (ref) {
      entry[field] = ref;
    }
  }
  return entry;
}

/**
 * Parse `word/fontTable.xml` into the list of font entries that declare at
 * least one embedded face. Entries with no embeds are still returned (callers
 * filter), and missing/unparseable input yields an empty list.
 */
export function parseEmbeddedFontTable(
  fontTableXml: string | null | undefined,
): EmbeddedFontEntry[] {
  if (!fontTableXml || fontTableXml.trim().length === 0) {
    return [];
  }

  const root = parseXmlDocument(fontTableXml);
  if (!root) {
    return [];
  }

  const entries: EmbeddedFontEntry[] = [];
  for (const font of findChildren(root, "w", "font")) {
    const entry = parseFontEntry(font);
    if (entry) {
      entries.push(entry);
    }
  }
  return entries;
}
