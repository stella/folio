import { expect, test } from "bun:test";

import { validateDocxPackage } from "@stll/docx-core";
import {
  applyDocumentOp,
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  OP_STORIES,
  paragraphLogicalText,
  REVISION_DECISIONS,
} from "@stll/docx-core/ops";

import type { BlockContent, Document, Paragraph } from "../types/document";
import { createEmptyDocument } from "../utils/createDocument";
import { parseDocx } from "./parser";
import { createDocx } from "./rezip";

// A tracked deletion of a paragraph that no paragraph follows right away, a
// section-ending one before range boundaries or one before a block that
// ends its container, saves and reopens to the same records: resolving the
// reopened file gives what resolving the open one does.

const paragraph = (paraId: string, text: string): Paragraph => ({
  type: "paragraph",
  paraId,
  content: [{ type: "run", content: [{ type: "text", text }] }],
});

const documentOf = (content: BlockContent[]): Document => {
  const document = createEmptyDocument();
  document.package.document.content = content;
  return document;
};

type Outline =
  | { paraId: string | undefined; text: string; ends: boolean; marked: boolean }
  | { control: Outline[] }
  | { type: string };

/** The blocks, their paragraphs' text and section ends, without package-level detail. */
const outline = (blocks: readonly BlockContent[]): Outline[] =>
  blocks.map((block) => {
    if (block.type === "paragraph")
      return {
        paraId: block.paraId,
        text: paragraphLogicalText(block),
        ends: block.sectionProperties !== undefined,
        marked: block.pPrMark !== undefined,
      };
    if (block.type === "blockSdt") return { control: outline(block.content) };
    return { type: block.type };
  });

const apply = (document: Document, op: DocumentOp) => {
  const applied = applyDocumentOp(document, op);
  if (applied.isErr()) throw applied.error;
  return applied.value;
};

const CASES: { name: string; content: BlockContent[] }[] = [
  {
    name: "a section-ending paragraph before range boundaries",
    content: [
      paragraph("00000001", "First"),
      {
        ...paragraph("00000002", "Ends a section"),
        sectionProperties: { pageWidth: 11_906 },
      },
      { type: "bookmarkStart", id: 3, name: "Clause" },
      paragraph("00000003", "Next"),
      { type: "bookmarkEnd", id: 3 },
      paragraph("00000004", "Final"),
    ],
  },
  {
    name: "a paragraph before a block content control that ends the body",
    content: [
      paragraph("00000001", "First"),
      paragraph("00000002", "Before the control"),
      {
        type: "blockSdt",
        properties: { sdtType: "richText", id: 5 },
        content: [paragraph("00000003", "Inside")],
      },
    ],
  },
];

test.each(CASES)(
  "the tracked deletion of $name reopens and resolves as before",
  async ({ content }) => {
    const original = documentOf(content);
    const op = {
      type: DOCUMENT_OP_TYPES.DELETE_BLOCKS,
      story: OP_STORIES.MAIN,
      blockIds: ["00000002"],
    } satisfies DocumentOp;
    const tracked = apply(original, {
      ...op,
      revision: { id: 100, author: "Reviewer", date: "2026-11-27T00:00:00Z" },
      newIds: { revision: [101, 102, 103, 104] },
    });
    const direct = apply(original, op);
    const saved = await createDocx(tracked.document);
    expect(await validateDocxPackage(new Uint8Array(saved))).toEqual({
      valid: true,
    });
    const reopened = await parseDocx(saved, { preloadFonts: false });
    expect(outline(reopened.package.document.content)).toEqual(
      outline(tracked.document.package.document.content),
    );
    for (const decision of Object.values(REVISION_DECISIONS)) {
      const resolve = (document: Document) =>
        apply(document, {
          type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
          story: OP_STORIES.MAIN,
          revisionIds: tracked.revisions,
          decision,
        }).document.package.document.content;
      const expected =
        decision === REVISION_DECISIONS.ACCEPT ? direct.document : original;
      expect(outline(resolve(reopened))).toEqual(
        outline(expected.package.document.content),
      );
    }
  },
);
