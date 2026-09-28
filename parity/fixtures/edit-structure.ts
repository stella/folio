/** Stable structural view of a saved synthetic package. */

import JSZip from "jszip";

import { FolioDocxReviewer } from "../../packages/core/src/ai-edits/headless";
import {
  getAttributeByNamespaceUri,
  getChildElements,
  getLocalName,
  getNamespaceUri,
  parseXmlDocument,
  type XmlElement,
} from "../../packages/core/src/docx/xmlParser";

const WORDPROCESSING_NAMESPACES = new Set([
  "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
  "http://purl.oclc.org/ooxml/wordprocessingml/main",
]);

const countPackageCarriers = async (buffer: ArrayBuffer) => {
  const zip = await JSZip.loadAsync(buffer);
  const xml = await zip.file("word/document.xml")?.async("text");
  const root = xml ? parseXmlDocument(xml) : null;
  if (!root) throw new Error("Saved package has no readable document root");
  const counts = { paragraphs: 0, sections: 0, fields: 0, noteReferences: 0 };
  const visit = (node: XmlElement): void => {
    if (WORDPROCESSING_NAMESPACES.has(getNamespaceUri(node) ?? "")) {
      switch (getLocalName(node.name)) {
        case "p":
          counts.paragraphs++;
          break;
        case "sectPr":
          counts.sections++;
          break;
        case "fldChar":
          if (
            getAttributeByNamespaceUri(node, WORDPROCESSING_NAMESPACES, "fldCharType") === "begin"
          ) {
            counts.fields++;
          }
          break;
        case "fldSimple":
          counts.fields++;
          break;
        case "footnoteReference":
        case "endnoteReference":
          counts.noteReferences++;
          break;
      }
    }
    for (const child of getChildElements(node)) visit(child);
  };
  visit(root);
  return counts;
};

export const readEditStructure = async (buffer: ArrayBuffer) => {
  const reviewer = await FolioDocxReviewer.fromBuffer(buffer);
  const blocks = reviewer.getContent();
  const cells = new Map<
    string,
    {
      table: number;
      row: number;
      column: number;
      rowSpan: number;
      columnSpan: number;
      text: string;
    }
  >();
  for (const block of blocks) {
    if (!block.table) continue;
    const { tableIndex, rowIndex, gridColumnIndex, rowSpan, columnSpan } = block.table;
    const key = `${tableIndex}:${rowIndex}:${gridColumnIndex}`;
    const existing = cells.get(key);
    if (existing) {
      existing.text += `\n${block.text}`;
      continue;
    }
    cells.set(key, {
      table: tableIndex,
      row: rowIndex,
      column: gridColumnIndex,
      rowSpan,
      columnSpan,
      text: block.text,
    });
  }
  return {
    package: await countPackageCarriers(buffer),
    blocks: blocks.map(({ kind, text, displayLabel, headingLevel, table }) =>
      Object.assign(
        { kind, text },
        displayLabel === undefined ? {} : { label: displayLabel },
        headingLevel === undefined ? {} : { headingLevel },
        table === undefined
          ? {}
          : {
              table: {
                index: table.tableIndex,
                row: table.rowIndex,
                column: table.gridColumnIndex,
              },
            },
      ),
    ),
    cells: [...cells.values()],
    revisions: reviewer.getChanges().map(({ type, text }) => ({ type, text })),
    comments: reviewer.getComments().map(({ text, anchoredText, replies }) => ({
      text,
      anchor: anchoredText,
      replies: replies.map((reply) => reply.text),
    })),
    notes: reviewer
      .getNotesAsText()
      .split("\n")
      .filter(Boolean)
      .map((line) => line.replace(/^\[(footnote|endnote) #\d+\]/u, "[$1]")),
  };
};

export type EditStructure = Awaited<ReturnType<typeof readEditStructure>>;
