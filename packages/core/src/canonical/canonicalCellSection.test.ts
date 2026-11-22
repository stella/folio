/**
 * A document that arrived with a section break inside a table cell: an
 * operation has the same answer on it as made, cloned and reopened, and the
 * editor refuses to undo the edit that removed the break.
 */

import { panic } from "better-result";
import { expect, test } from "bun:test";
import { EditorState } from "prosemirror-state";

import {
  applyDocumentOp,
  DOCUMENT_OP_REFUSAL_REASONS,
  type DocumentOp,
  OP_STORIES,
} from "@stll/docx-core/ops";

import { parseDocx } from "../docx/parser";
import { createDocx } from "../docx/rezip";
import { schema } from "../prosemirror/schema";
import type { Document, Paragraph, Table } from "../types/document";
import { createEmptyDocument } from "../utils/createDocument";
import {
  createCanonicalSession,
  publishCanonicalProjection,
} from "./canonicalSession";

const story = OP_STORIES.MAIN;

const paragraph = (paraId: string, text: string): Paragraph => ({
  type: "paragraph",
  paraId,
  content: [{ type: "run", content: [{ type: "text", text }] }],
});

const arrived = (): Document => {
  const document = createEmptyDocument();
  document.package.document.content = [
    paragraph("74000001", "Before"),
    {
      type: "table",
      columnWidths: [4800],
      rows: [
        {
          type: "tableRow",
          cells: [
            {
              type: "tableCell",
              content: [
                {
                  ...paragraph("74000002", "Inner"),
                  sectionProperties: { pageWidth: 11_906, pageHeight: 16_838 },
                },
                paragraph("74000003", "Beside"),
              ],
            },
          ],
        },
      ],
    },
    paragraph("74000004", "After"),
  ];
  document.package.document.sections = undefined;
  return document;
};

/** The document as a package of its own reads back. */
const reopened = async (document: Document): Promise<Document> => {
  const copy = structuredClone(document);
  delete copy.originalBuffer;
  return parseDocx(await createDocx(copy), {
    preloadFonts: false,
    detectVariables: false,
  });
};

const inCell = (document: Document, at: number): Paragraph => {
  const table = document.package.document.content.find(
    (block): block is Table => block.type === "table",
  );
  const block = table?.rows.at(0)?.cells.at(0)?.content.at(at);
  return block?.type === "paragraph" ? block : panic("The cell's paragraph.");
};

/** A paragraph with a section, and one without. */
const toggled = (document: Document, at: number): DocumentOp => {
  const from = inCell(document, at);
  const { sectionProperties, ...without } = from;
  return {
    type: "replaceBlocks",
    story,
    expected: [from],
    blocks: [
      sectionProperties === undefined
        ? { ...from, sectionProperties: { pageWidth: 11_906 } }
        : without,
    ],
    sectionBoundaries: "replace",
  };
};

const answer = (document: Document, op: DocumentOp): string => {
  const result = applyDocumentOp(document, op);
  return result.isOk() ? "applied" : result.error.reason;
};

test("an operation has the same answer on a document as made, cloned and reopened", async () => {
  const made = arrived();
  for (const document of [made, structuredClone(made), await reopened(made)]) {
    expect(answer(document, toggled(document, 0))).toBe("applied");
    expect(answer(document, toggled(document, 1))).toBe(
      DOCUMENT_OP_REFUSAL_REASONS.SECTION_BOUNDARY,
    );
    const removed = applyDocumentOp(document, toggled(document, 0)).unwrap();
    for (const without of [
      removed.document,
      structuredClone(removed.document),
      await reopened(removed.document),
    ])
      expect(answer(without, toggled(without, 0))).toBe(
        DOCUMENT_OP_REFUSAL_REASONS.SECTION_BOUNDARY,
      );
  }
});

test("the editor refuses to undo the edit that removed it, and stays as it is", async () => {
  const session = createCanonicalSession(await reopened(arrived())).unwrap();
  const state = EditorState.create({ schema, doc: session.projection.doc });
  const joined = publishCanonicalProjection({
    session,
    state,
    commit: session
      .prepareIntent(state, {
        type: "joinParagraphs",
        story,
        blockId: "74000002",
        nextBlockId: "74000003",
      })
      .unwrap(),
  }).unwrap().state;
  const after = session.document;
  expect(inCell(after, 0).sectionProperties).toBeUndefined();
  const undo = session.prepareUndo(joined);
  expect(undo.isErr() ? undo.error.message : "undone").toContain(
    "A section cannot end inside a table cell.",
  );
  expect(session.document).toBe(after);
});
