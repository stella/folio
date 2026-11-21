import { expect, spyOn, test } from "bun:test";

import type { Document, Paragraph } from "../../model/document";
import { applyDocumentOp, applyDocumentOps } from "../apply";
import * as blocks from "../blocks";
import { formatParaId } from "../idLease";
import * as primitives from "../paragraphPrimitive";
import { DOCUMENT_OP_REFUSAL_REASONS } from "../refusal";
import {
  DOCUMENT_OP_TYPES,
  OP_STORIES,
  type ReplaceBlocksOp,
  type ReplaceInlineOp,
  SECTION_BOUNDARY_POLICIES,
  type SetParagraphReviewOp,
} from "../types";
import { captureDocumentOp } from "../wire";

const paragraph = (id: number, text: string): Paragraph => ({
  type: "paragraph",
  paraId: formatParaId(id),
  content: [{ type: "run", content: [{ type: "text", text }] }],
});

const documentOf = (content: Paragraph[]): Document => ({
  package: { document: { content } },
});

test("native inline preparation shares validation with apply and leaves materialization to its caller", () => {
  const first = paragraph(1, "before");
  const after = paragraph(1, "after");
  const other = Array.from({ length: 255 }, (_, index) =>
    paragraph(index + 2, "untouched"),
  );
  const document = documentOf([first, ...other]);
  const op: ReplaceInlineOp = {
    type: DOCUMENT_OP_TYPES.REPLACE_INLINE,
    story: OP_STORIES.MAIN,
    blockId: "00000001",
    expected: first.content,
    content: after.content,
  };
  const prepare = spyOn(primitives, "prepareResolutionParagraphOp");
  const materialize = spyOn(blocks, "replaceParagraphs");
  try {
    const at = { paragraph: first, index: 0, list: [] };
    for (let attempt = 0; attempt < 64; attempt += 1)
      expect(
        primitives.prepareResolutionParagraphOp({ document, op, at }).isOk(),
      ).toBe(true);
    expect(materialize).toHaveBeenCalledTimes(0);

    const planned = primitives
      .prepareResolutionParagraphOp({ document, op, at })
      .unwrap();
    if (planned.type !== "replacement")
      throw new Error("The inline content changes.");
    expect(planned.incomingParagraphIds).toEqual([]);
    expect(planned.incomingIdentityKeys).toEqual([]);
    prepare.mockClear();
    const applied = applyDocumentOp(document, op).unwrap();
    expect(prepare).toHaveBeenCalledTimes(1);
    expect(materialize).toHaveBeenCalledTimes(1);
    expect(applied.document.package.document.content.at(0)).toStrictEqual(
      planned.paragraph,
    );
    expect(applied.touched).toStrictEqual(planned.touched);
    expect(applied.inverse).toStrictEqual([captureDocumentOp(planned.inverse)]);
    expect(
      applyDocumentOps(applied.document, applied.inverse).unwrap().document,
    ).toStrictEqual(document);
  } finally {
    materialize.mockRestore();
    prepare.mockRestore();
  }
});

test("clearing an imported terminal review mark retains the native block inverse", () => {
  const mark = { kind: "ins", info: { id: 4, author: "Editor" } } as const;
  const original = { ...paragraph(1, "text"), pPrMark: mark };
  const document = documentOf([original]);
  const op: SetParagraphReviewOp = {
    type: DOCUMENT_OP_TYPES.SET_PARAGRAPH_REVIEW,
    story: OP_STORIES.MAIN,
    blockId: "00000001",
    expected: { pPrMark: mark },
    review: {},
  };
  const planned = primitives
    .prepareResolutionParagraphOp({
      document,
      op,
      at: { paragraph: original, index: 0, list: [] },
    })
    .unwrap();
  if (planned.type !== "replacement")
    throw new Error("The review mark is cleared.");
  expect(planned.inverse.type).toBe(DOCUMENT_OP_TYPES.REPLACE_BLOCKS);
  expect(planned.incomingIdentityKeys).toEqual([]);
  const applied = applyDocumentOp(document, op).unwrap();
  expect(applied.inverse).toStrictEqual([captureDocumentOp(planned.inverse)]);
  expect(
    applyDocumentOps(applied.document, applied.inverse).unwrap().document,
  ).toStrictEqual(document);
});

test("nonempty block replacement shares validation and its section-aware inverse", () => {
  const before = [paragraph(1, "retire"), paragraph(2, "survive")];
  const document = documentOf(before);
  const retained = paragraph(2, "retire survive");
  const op: ReplaceBlocksOp = {
    type: DOCUMENT_OP_TYPES.REPLACE_BLOCKS,
    story: OP_STORIES.MAIN,
    expected: before,
    blocks: [retained],
    sectionBoundaries: SECTION_BOUNDARY_POLICIES.REPLACE,
  };
  const validate = spyOn(primitives, "validateBlockReplacement");
  try {
    const planned = primitives.nonemptyBlockReplacementEffects(op, before);
    const applied = applyDocumentOp(document, op).unwrap();
    expect(validate).toHaveBeenCalledTimes(1);
    expect(applied.inverse).toStrictEqual(
      planned.inverse.map(captureDocumentOp),
    );
    expect(applied.touched).toStrictEqual(planned.touched);
    expect(
      applyDocumentOps(applied.document, applied.inverse).unwrap().document,
    ).toStrictEqual(document);
  } finally {
    validate.mockRestore();
  }
});

test("stale inline content is refused before inspecting its invalid replacement", () => {
  const original = paragraph(1, "text");
  const document = documentOf([original]);
  const invalid = [{ type: "run", content: [] }] as const;
  const op: ReplaceInlineOp = {
    type: DOCUMENT_OP_TYPES.REPLACE_INLINE,
    story: OP_STORIES.MAIN,
    blockId: "00000001",
    expected: [],
    content: [...invalid],
  };
  const planned = primitives.prepareResolutionParagraphOp({
    document,
    op,
    at: { paragraph: original, index: 0, list: [] },
  });
  const applied = applyDocumentOp(document, op);
  if (planned.isOk() || applied.isOk())
    throw new Error("The stale primitive must be refused.");
  expect(planned.error.reason).toBe(DOCUMENT_OP_REFUSAL_REASONS.STALE);
  expect(applied.error).toStrictEqual(planned.error);
});
