/**
 * Every direct deletion of a paragraph that applies undoes to the document as
 * it was, whatever stands on either side of it: its inverse is an ordinary
 * insertion, which must be able to place the paragraph again.
 */
import { describe, expect, test } from "bun:test";

import type { BlockContent, Document, Paragraph } from "../../model/document";
import { applyDocumentOp, applyDocumentOps } from "../apply";
import { DOCUMENT_OP_TYPES, OP_STORIES } from "../types";

const paragraph = (paraId: string, text: string): Paragraph => ({
  type: "paragraph",
  paraId,
  content:
    text === "" ? [] : [{ type: "run", content: [{ type: "text", text }] }],
});

const table = (paraId: string): BlockContent => ({
  type: "table",
  rows: [
    {
      type: "tableRow",
      cells: [{ type: "tableCell", content: [paragraph(paraId, "Cell.")] }],
    },
  ],
});

const control = (paraId: string, id: number): BlockContent => ({
  type: "blockSdt",
  properties: { sdtType: "richText", id },
  content: [paragraph(paraId, "Held.")],
});

const bookmark = (id: number): BlockContent => ({
  type: "bookmarkStart",
  id,
  name: `Mark${id}`,
});

const BEFORE = {
  nothing: [],
  "a paragraph": [paragraph("00000011", "Before.")],
  "a paragraph ending a section": [
    {
      ...paragraph("00000011", "Before."),
      sectionProperties: { pageWidth: 11_906 },
    },
  ],
  "a table": [table("00000012")],
  "a content control": [control("00000013", 13)],
  "a bookmark": [paragraph("00000011", "Before."), bookmark(14)],
} as const satisfies Record<string, readonly BlockContent[]>;

const AFTER = {
  nothing: [],
  "a paragraph": [paragraph("00000021", "After.")],
  "a table": [table("00000022"), paragraph("00000021", "")],
  "a content control": [control("00000023", 23), paragraph("00000021", "")],
  "a preserved block": [
    { type: "preservedBlock", xml: "<w:altChunk/>" },
    paragraph("00000021", ""),
  ],
  "a bookmark": [bookmark(24), paragraph("00000021", "After.")],
} as const satisfies Record<string, readonly BlockContent[]>;

const DELETED = {
  empty: paragraph("00000001", ""),
  "with words": paragraph("00000001", "Deleted."),
  "ending a section": {
    ...paragraph("00000001", "Deleted."),
    sectionProperties: { pageWidth: 10_000 },
  },
} as const satisfies Record<string, Paragraph>;

/** The sections a body's breaks divide it into, as a reader of it states them. */
const withSections = (content: BlockContent[]): Document => {
  const sections: { content: BlockContent[]; properties: object }[] = [];
  let open: BlockContent[] = [];
  for (const block of content) {
    open.push(block);
    if (block.type !== "paragraph" || block.sectionProperties === undefined)
      continue;
    sections.push({ content: open, properties: block.sectionProperties });
    open = [];
  }
  const finalSectionProperties = { pageWidth: 12_240 };
  if (open.length > 0)
    sections.push({ content: open, properties: finalSectionProperties });
  return {
    package: { document: { content, sections, finalSectionProperties } },
  };
};

describe("a direct paragraph deletion", () => {
  test.each(
    Object.entries(BEFORE).flatMap(([before, blocksBefore]) =>
      Object.entries(AFTER).flatMap(([after, blocksAfter]) =>
        Object.entries(DELETED).map(([kind, deleted]) => ({
          before,
          after,
          kind,
          content: [...blocksBefore, deleted, ...blocksAfter],
        })),
      ),
    ),
  )(
    "of a paragraph $kind between $before and $after undoes exactly",
    ({ content }) => {
      const document = withSections(structuredClone(content));
      const removed = applyDocumentOp(document, {
        type: DOCUMENT_OP_TYPES.DELETE_BLOCKS,
        story: OP_STORIES.MAIN,
        blockIds: ["00000001"],
      });
      if (removed.isErr()) return;
      const undone = applyDocumentOps(
        removed.value.document,
        removed.value.inverse,
      );
      if (undone.isErr()) throw undone.error;
      expect(undone.value.document).toStrictEqual(document);
      const redone = applyDocumentOps(
        undone.value.document,
        undone.value.inverse,
      );
      if (redone.isErr()) throw redone.error;
      expect(redone.value.document).toStrictEqual(removed.value.document);
    },
  );
});
