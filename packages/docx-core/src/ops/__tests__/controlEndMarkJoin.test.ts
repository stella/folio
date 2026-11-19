/**
 * Resolving away the mark of the paragraph that ends a block content control
 * joins it with the paragraph after the control: an empty one has nothing to
 * bring into the control and goes; one with content stays where it is.
 */
import { expect, test } from "bun:test";

import type { BlockContent, Document, Paragraph } from "../../model/document";
import { applyDocumentOp, applyDocumentOps } from "../apply";
import {
  DOCUMENT_OP_TYPES,
  OP_STORIES,
  REVISION_DECISIONS,
  type RevisionDecision,
} from "../types";

const paragraph = (
  paraId: string,
  text: string,
  mark?: "ins" | "del",
): Paragraph => ({
  type: "paragraph",
  paraId,
  content:
    text === "" ? [] : [{ type: "run", content: [{ type: "text", text }] }],
  ...(mark === undefined
    ? {}
    : {
        pPrMark: {
          kind: mark,
          info: { id: 9, author: "A", date: "2026-01-02T03:04:05Z" },
        },
      }),
});

const documentOf = (after: Paragraph[], mark: "ins" | "del"): Document => ({
  package: {
    document: {
      content: [
        paragraph("00000001", "One"),
        {
          type: "blockSdt",
          properties: { sdtType: "group", id: 1 },
          content: [
            paragraph("00000002", "Two"),
            paragraph("00000003", "", mark),
          ],
        },
        ...after,
      ],
    },
  },
});

const resolved = (document: Document, decision: RevisionDecision) => {
  const applied = applyDocumentOp(document, {
    type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
    story: OP_STORIES.MAIN,
    revisionIds: [9],
    decision,
  });
  if (applied.isErr()) throw applied.error;
  const undo = applyDocumentOps(applied.value.document, applied.value.inverse);
  if (undo.isErr()) throw undo.error;
  expect(undo.value.document).toStrictEqual(document);
  return applied.value.document.package.document.content;
};

const outline = (blocks: readonly BlockContent[]): unknown[] =>
  blocks.map((block) => {
    if (block.type === "blockSdt") return outline(block.content);
    return block.type === "paragraph" ? block.paraId : block.type;
  });

for (const [mark, decision] of [
  ["ins", REVISION_DECISIONS.REJECT],
  ["del", REVISION_DECISIONS.ACCEPT],
] as const) {
  test(`${decision}ing the control's last mark takes the empty paragraph after it`, () => {
    expect(
      outline(
        resolved(
          documentOf(
            [paragraph("00000004", ""), paragraph("00000005", "Five")],
            mark,
          ),
          decision,
        ),
      ),
    ).toStrictEqual(["00000001", ["00000002", "00000003"], "00000005"]);
  });

  test(`${decision}ing the control's last mark takes an empty final paragraph`, () => {
    expect(
      outline(
        resolved(documentOf([paragraph("00000004", "")], mark), decision),
      ),
    ).toStrictEqual(["00000001", ["00000002", "00000003"]]);
  });

  test(`${decision}ing the control's last mark keeps a paragraph with content after it`, () => {
    expect(
      outline(
        resolved(documentOf([paragraph("00000004", "Four")], mark), decision),
      ),
    ).toStrictEqual(["00000001", ["00000002", "00000003"], "00000004"]);
  });
}
