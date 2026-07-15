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
import { serializeParagraph } from "./paragraphSerializer";
import { serializeTable } from "./tableSerializer";

// Namespaces declared on the notes root. Mirrors the header/footer serializer's
// declared set so note bodies carrying DrawingML, math, or raw-replayed SDT
// extensions land on a root that declares every prefix they might use.
const NAMESPACES: Record<string, string> = {
  wpc: "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas",
  mc: "http://schemas.openxmlformats.org/markup-compatibility/2006",
  o: "urn:schemas-microsoft-com:office:office",
  r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
  m: "http://schemas.openxmlformats.org/officeDocument/2006/math",
  v: "urn:schemas-microsoft-com:vml",
  a: "http://schemas.openxmlformats.org/drawingml/2006/main",
  pic: "http://schemas.openxmlformats.org/drawingml/2006/picture",
  wp14: "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
  wp: "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
  w10: "urn:schemas-microsoft-com:office:word",
  w: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
  w14: "http://schemas.microsoft.com/office/word/2010/wordml",
  w15: "http://schemas.microsoft.com/office/word/2012/wordml",
  w16: "http://schemas.microsoft.com/office/word/2018/wordml",
  w16cex: "http://schemas.microsoft.com/office/word/2018/wordml/cex",
  w16cid: "http://schemas.microsoft.com/office/word/2016/wordml/cid",
  w16sdtdh: "http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash",
  w16se: "http://schemas.microsoft.com/office/word/2015/wordml/symex",
  wpg: "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
  wps: "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
};

function buildNamespaceDeclarations(): string {
  return Object.entries(NAMESPACES)
    .map(([prefix, uri]) => `xmlns:${prefix}="${uri}"`)
    .join(" ");
}

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

/**
 * Serialize footnotes to a complete word/footnotes.xml string.
 *
 * @param footnotes - Notes to serialize (the model's normal footnotes).
 * @returns Complete footnotes.xml string.
 */
export function serializeFootnotes(footnotes: readonly Footnote[]): string {
  const nsDecl = buildNamespaceDeclarations();
  const notes = footnotes.map((fn) => serializeNote("footnote", fn)).join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:footnotes ${nsDecl} mc:Ignorable="w14 w15 wp14">${notes}</w:footnotes>`;
}

/**
 * Serialize endnotes to a complete word/endnotes.xml string.
 *
 * @param endnotes - Notes to serialize (the model's normal endnotes).
 * @returns Complete endnotes.xml string.
 */
export function serializeEndnotes(endnotes: readonly Endnote[]): string {
  const nsDecl = buildNamespaceDeclarations();
  const notes = endnotes.map((en) => serializeNote("endnote", en)).join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:endnotes ${nsDecl} mc:Ignorable="w14 w15 wp14">${notes}</w:endnotes>`;
}
