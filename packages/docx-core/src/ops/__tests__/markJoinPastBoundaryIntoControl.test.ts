/**
 * Rejecting the added mark of a paragraph whose next block, past range
 * boundaries standing between them, is a block content control runs the
 * paragraph into the control's first paragraph, as it does with no boundary
 * between: the boundaries go into that paragraph where the two meet, so what
 * stood before them stays before. An empty paragraph goes, and the
 * boundaries stay where they stand.
 */
import { describe, expect, test } from "bun:test";

import type { BlockContent, Document, Paragraph } from "../../model/document";
import { applyDocumentOp, applyDocumentOps } from "../apply";
import {
  DOCUMENT_OP_TYPES,
  OP_STORIES,
  REVISION_DECISIONS,
  type RevisionDecision,
} from "../types";

const INFO = { author: "A", date: "2026-01-02T03:04:05Z" } as const;

const paragraph = (
  paraId: string,
  text: string,
  mark?: { kind: "ins" | "del"; id: number },
): Paragraph => ({
  type: "paragraph",
  paraId,
  content:
    text === "" ? [] : [{ type: "run", content: [{ type: "text", text }] }],
  ...(mark === undefined
    ? {}
    : { pPrMark: { kind: mark.kind, info: { id: mark.id, ...INFO } } }),
});

const documentOf = (firstText: string): Document => ({
  package: {
    document: {
      content: [
        paragraph("00000001", firstText, { kind: "ins", id: 9 }),
        { type: "bookmarkStart", id: "5", name: "Clause" },
        {
          type: "blockSdt",
          properties: { sdtType: "group", id: 1 },
          content: [paragraph("00000002", "Two", { kind: "del", id: 8 })],
        },
        { type: "bookmarkEnd", id: "5" },
        paragraph("00000003", "Three"),
      ],
    },
  },
});

const resolved = (document: Document, decision: RevisionDecision) => {
  const applied = applyDocumentOp(document, {
    type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
    story: OP_STORIES.MAIN,
    revisionIds: [8, 9],
    decision,
  });
  if (applied.isErr()) throw applied.error;
  const undo = applyDocumentOps(applied.value.document, applied.value.inverse);
  if (undo.isErr()) throw undo.error;
  expect(undo.value.document).toStrictEqual(document);
  return applied.value.document.package.document.content;
};

/** Each block as its id or type; a paragraph's inline boundaries and text. */
const outline = (blocks: readonly BlockContent[]): unknown[] =>
  blocks.map((block) => {
    if (block.type === "blockSdt") return outline(block.content);
    if (block.type !== "paragraph") return block.type;
    const inline = block.content.map((node) =>
      node.type === "run"
        ? node.content
            .map((leaf) => (leaf.type === "text" ? leaf.text : ""))
            .join("")
        : node.type,
    );
    return [block.paraId, ...inline];
  });

/**
 * The first paragraph's words are pending deletion and its mark is deleted:
 * accepting the mark alone runs the still-pending words into the control.
 */
const withPendingWords = (): Document => ({
  package: {
    document: {
      content: [
        {
          type: "paragraph",
          paraId: "00000001",
          content: [
            {
              type: "deletion",
              info: { id: 7, ...INFO },
              content: [
                { type: "run", content: [{ type: "text", text: "Gone " }] },
              ],
            },
          ],
          pPrMark: { kind: "del", info: { id: 9, ...INFO } },
        },
        { type: "bookmarkStart", id: "5", name: "Clause" },
        {
          type: "blockSdt",
          properties: { sdtType: "group", id: 1 },
          content: [paragraph("00000002", "Two")],
        },
        { type: "bookmarkEnd", id: "5" },
        paragraph("00000003", "Three"),
      ],
    },
  },
});

describe("accepting a deleted mark before boundaries and a block content control", () => {
  test("runs the pending words into the control, the boundary where the two meet", () => {
    const document = withPendingWords();
    const applied = applyDocumentOp(document, {
      type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
      story: OP_STORIES.MAIN,
      revisionIds: [9],
      decision: REVISION_DECISIONS.ACCEPT,
    });
    if (applied.isErr()) throw applied.error;
    const undo = applyDocumentOps(
      applied.value.document,
      applied.value.inverse,
    );
    if (undo.isErr()) throw undo.error;
    expect(undo.value.document).toStrictEqual(document);
    expect(
      outline(applied.value.document.package.document.content),
    ).toStrictEqual([
      [["00000002", "deletion", "bookmarkStart", "Two"]],
      "bookmarkEnd",
      ["00000003", "Three"],
    ]);
  });
});

describe("rejecting an added mark before boundaries and a block content control", () => {
  test("keeps a worded paragraph's words and the boundary before the control", () => {
    expect(
      outline(resolved(documentOf("One "), REVISION_DECISIONS.REJECT)),
    ).toStrictEqual([
      ["00000001", "One "],
      "bookmarkStart",
      [["00000002", "Two"]],
      "bookmarkEnd",
      ["00000003", "Three"],
    ]);
  });

  test("takes an empty paragraph and leaves the boundary standing", () => {
    expect(
      outline(resolved(documentOf(""), REVISION_DECISIONS.REJECT)),
    ).toStrictEqual([
      "bookmarkStart",
      [["00000002", "Two"]],
      "bookmarkEnd",
      ["00000003", "Three"],
    ]);
  });

  test("accepting keeps the paragraph and the boundary before the control", () => {
    const blocks = resolved(documentOf("One "), REVISION_DECISIONS.ACCEPT);
    expect(outline(blocks).slice(0, 2)).toStrictEqual([
      ["00000001", "One "],
      "bookmarkStart",
    ]);
  });
});

// Only a boundary follows the last paragraph: nothing follows it, and
// removing it would leave the boundary ending the body, so it stays.
test("rejecting the added mark of the last paragraph before a trailing boundary keeps the paragraph", () => {
  const document: Document = {
    package: {
      document: {
        content: [
          paragraph("00000001", "One"),
          paragraph("00000002", "", { kind: "ins", id: 9 }),
          { type: "bookmarkStart", id: "5", name: "Clause" },
        ],
      },
    },
  };
  const applied = applyDocumentOp(document, {
    type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
    story: OP_STORIES.MAIN,
    revisionIds: [9],
    decision: REVISION_DECISIONS.REJECT,
  });
  if (applied.isErr()) throw applied.error;
  const undo = applyDocumentOps(applied.value.document, applied.value.inverse);
  if (undo.isErr()) throw undo.error;
  expect(undo.value.document).toStrictEqual(document);
  const blocks = applied.value.document.package.document.content;
  expect(outline(blocks)).toStrictEqual([
    ["00000001", "One"],
    ["00000002"],
    "bookmarkStart",
  ]);
  expect(blocks.at(1)).not.toHaveProperty("pPrMark");
});
