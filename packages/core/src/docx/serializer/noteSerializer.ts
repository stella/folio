/**
 * Footnote/Endnote Serializer - Serialize footnotes/endnotes back to OOXML XML
 *
 * Converts the parsed Footnote[] / Endnote[] model back into valid
 * word/footnotes.xml / word/endnotes.xml. Reuses the same paragraph, table,
 * and block-SDT serializers the document body and header/footer serializers
 * use, so note bodies round-trip the full block model (runs, tracked changes,
 * fields, content controls) instead of a flattened subset.
 *
 * OOXML Reference:
 * - Footnotes root: w:footnotes; each note: w:footnote[@w:id][@w:type]
 * - Endnotes root:  w:endnotes;  each note: w:endnote[@w:id][@w:type]
 *
 * The document model only retains the normal (content) notes: the separator /
 * continuationSeparator notes Word requires are dropped during parsing
 * (`getNormalFootnotes`). Callers therefore must NOT overwrite the whole part
 * with this output — it would lose those separators. The selective save path
 * uses this serializer only to extract the edited note paragraph by `paraId`
 * and splices it into the original part, keeping separators and unedited notes
 * byte-exact.
 */

import type { BlockContent, Endnote, Footnote } from "../../types/document";
import { serializeBlockSdt } from "./blockSdtSerializer";
import { serializePartElement, type OoxmlNamespacePrefix } from "./partNamespaces";
import { serializeParagraph } from "./paragraphSerializer";
import { serializeTable } from "./tableSerializer";

// Prefixes a notes part declares whether or not the bodies use them. Mirrors
// the header/footer baseline so note bodies carrying DrawingML, math, or
// raw-replayed SDT extensions land on a root that declares every prefix they
// might use.
const NOTE_BASELINE_PREFIXES = [
  "wpc",
  "mc",
  "o",
  "r",
  "m",
  "v",
  "a",
  "pic",
  "wp14",
  "wp",
  "w10",
  "w",
  "w14",
  "w15",
  "w16",
  "w16cex",
  "w16cid",
  "w16sdtdh",
  "w16se",
  "wpg",
  "wps",
] as const satisfies readonly OoxmlNamespacePrefix[];

const serializeNotePart = (elementName: "footnote" | "endnote", body: string): string =>
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
  serializePartElement({
    partPath: `word/${elementName}s.xml`,
    rootName: `w:${elementName}s`,
    baselinePrefixes: NOTE_BASELINE_PREFIXES,
    sourceBindings: undefined,
    body,
  });

/**
 * Serialize a block content item (paragraph, table, or block-level SDT) for a
 * note body.
 */
function serializeBlock(block: BlockContent): string {
  if (block.type === "paragraph") {
    return serializeParagraph(block);
  }
  if (block.type === "table") {
    return serializeTable(block, serializeParagraph);
  }
  return serializeBlockSdt(block, serializeBlock);
}

/**
 * Serialize a single note element. Footnotes and endnotes share the same
 * structure, differing only in the wrapper element name.
 */
function serializeNote(elementName: "footnote" | "endnote", note: Footnote | Endnote): string {
  const attrs: string[] = [];
  // Word emits w:type before w:id on typed (separator) notes; mirror that
  // ordering. Normal notes carry no type attribute.
  if (note.noteType && note.noteType !== "normal") {
    attrs.push(`w:type="${note.noteType}"`);
  }
  attrs.push(`w:id="${note.id}"`);

  const body = note.content.map((block) => serializeBlock(block)).join("");
  return `<w:${elementName} ${attrs.join(" ")}>${body}</w:${elementName}>`;
}

const NOTE_SEPARATOR_BODY =
  '<w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>';

function serializeRequiredNoteSeparators(elementName: "footnote" | "endnote"): string {
  return (
    `<w:${elementName} w:type="separator" w:id="-1">${NOTE_SEPARATOR_BODY}` +
    `<w:r><w:separator/></w:r></w:p></w:${elementName}>` +
    `<w:${elementName} w:type="continuationSeparator" w:id="0">${NOTE_SEPARATOR_BODY}` +
    `<w:r><w:continuationSeparator/></w:r></w:p></w:${elementName}>`
  );
}

function insertNoteReferenceMark(xml: string, elementName: "footnote" | "endnote"): string {
  const paragraphOpen = /<w:p(?=[\s>])[^>]*>/u.exec(xml);
  if (!paragraphOpen) {
    const referenceParagraph =
      `<w:p><w:r><w:rPr><w:rStyle w:val="${elementName === "footnote" ? "FootnoteReference" : "EndnoteReference"}"/></w:rPr>` +
      `<w:${elementName}Ref/></w:r></w:p>`;
    return `${referenceParagraph}${xml}`;
  }

  const paragraphStart = paragraphOpen.index + paragraphOpen[0].length;
  const paragraphEnd = xml.indexOf("</w:p>", paragraphStart);
  const propertiesEnd = xml.indexOf("</w:pPr>", paragraphStart);
  const insertAt =
    propertiesEnd !== -1 && propertiesEnd < paragraphEnd
      ? propertiesEnd + "</w:pPr>".length
      : paragraphStart;
  const referenceRun =
    `<w:r><w:rPr><w:rStyle w:val="${elementName === "footnote" ? "FootnoteReference" : "EndnoteReference"}"/></w:rPr>` +
    `<w:${elementName}Ref/></w:r>`;
  return `${xml.slice(0, insertAt)}${referenceRun}${xml.slice(insertAt)}`;
}

function serializeNewNotePart(
  elementName: "footnote" | "endnote",
  notes: readonly (Footnote | Endnote)[],
): string {
  const serializedNotes = notes
    .map((note) => insertNoteReferenceMark(serializeNote(elementName, note), elementName))
    .join("");
  return serializeNotePart(
    elementName,
    serializeRequiredNoteSeparators(elementName) + serializedNotes,
  );
}

/**
 * Serialize footnotes to a complete word/footnotes.xml string.
 *
 * @param footnotes - Notes to serialize (the model's normal footnotes).
 * @returns Complete footnotes.xml string.
 */
export function serializeFootnotes(footnotes: readonly Footnote[]): string {
  return serializeNotePart(
    "footnote",
    footnotes.map((fn) => serializeNote("footnote", fn)).join(""),
  );
}

/**
 * Serialize endnotes to a complete word/endnotes.xml string.
 *
 * @param endnotes - Notes to serialize (the model's normal endnotes).
 * @returns Complete endnotes.xml string.
 */
export function serializeEndnotes(endnotes: readonly Endnote[]): string {
  return serializeNotePart("endnote", endnotes.map((en) => serializeNote("endnote", en)).join(""));
}

/** Serialize a brand-new footnote part, including Word's required separators
 * and the automatic reference mark at the start of every normal note. */
export function serializeNewFootnotesPart(footnotes: readonly Footnote[]): string {
  return serializeNewNotePart("footnote", footnotes);
}

/** Serialize a brand-new endnote part, including Word's required separators
 * and the automatic reference mark at the start of every normal note. */
export function serializeNewEndnotesPart(endnotes: readonly Endnote[]): string {
  return serializeNewNotePart("endnote", endnotes);
}
