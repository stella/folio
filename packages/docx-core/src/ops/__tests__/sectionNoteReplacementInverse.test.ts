/**
 * Replacing a paragraph that ends a section and holds a note's only reference
 * removes the section break and the note. The inverse restores the note part
 * first, then the blocks with their section, and undoes exactly.
 */

import { expect, test } from "bun:test";

import type { Document, Paragraph } from "../../model/document";
import { applyDocumentOp, applyDocumentOps } from "../apply";
import {
  DOCUMENT_OP_TYPES,
  OP_STORIES,
  SECTION_BOUNDARY_POLICIES,
} from "../types";

const ending: Paragraph = {
  type: "paragraph",
  paraId: "00000001",
  content: [{ type: "run", content: [{ type: "footnoteRef", id: 2 }] }],
  sectionProperties: { pageWidth: 11906 },
};
const last: Paragraph = { type: "paragraph", paraId: "00000002", content: [] };

const document: Document = {
  package: {
    document: {
      content: [ending, last],
      sections: [
        { properties: { pageWidth: 11906 }, content: [ending] },
        {
          properties: { pageWidth: 12240, pageHeight: 15840 },
          content: [last],
        },
      ],
    },
    footnotes: [
      {
        type: "footnote",
        id: 2,
        content: [
          {
            type: "paragraph",
            paraId: "0F0F0F0F",
            content: [{ type: "run", content: [{ type: "text", text: "n" }] }],
          },
        ],
      },
    ],
  },
};

test("a replacement taking a section break and a note's reference undoes exactly", () => {
  const applied = applyDocumentOp(document, {
    type: DOCUMENT_OP_TYPES.REPLACE_BLOCKS,
    story: OP_STORIES.MAIN,
    expected: [ending, last],
    blocks: [last],
    sectionBoundaries: SECTION_BOUNDARY_POLICIES.REPLACE,
  }).unwrap();
  expect(applied.document.package.document.sections).toHaveLength(1);
  expect(applied.document.package.footnotes ?? []).toEqual([]);
  expect(
    applied.inverse.find(
      (inverse) => inverse.type === DOCUMENT_OP_TYPES.REPLACE_BLOCKS,
    ),
  ).toHaveProperty("sectionView");
  expect(
    applyDocumentOps(applied.document, applied.inverse).unwrap().document,
  ).toStrictEqual(document);
});
