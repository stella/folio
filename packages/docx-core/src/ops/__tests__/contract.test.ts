import { describe, expect, test } from "bun:test";

import type { Document, Paragraph, Run } from "../../model/document";
import { applyDocumentOp, applyDocumentOps } from "../apply";
import { normalizeForOps, validateOpsDocument } from "../contract";
import { DOCUMENT_OP_REFUSAL_REASONS } from "../refusal";
import { DOCUMENT_OP_TYPES, type DocumentOp, INHERIT_RUN_PROPS, OP_STORIES } from "../types";

const run = (text: string): Run => ({ type: "run", content: [{ type: "text", text }] });

const paragraph = (paraId: string, text: string): Paragraph => ({
  type: "paragraph",
  paraId,
  content: [run(text)],
});

const at = (blockId: string, offset: number) => ({ story: OP_STORIES.MAIN, blockId, offset });

const typing = (blockId: string): DocumentOp => ({
  type: DOCUMENT_OP_TYPES.INSERT_TEXT,
  at: at(blockId, 0),
  text: "x",
  runProps: INHERIT_RUN_PROPS,
});

const reasonOf = (document: Document, op: DocumentOp): string | undefined => {
  const result = applyDocumentOp(document, op);
  return result.isErr() ? result.error.reason : undefined;
};

/** A body whose section view holds the body's own records, as the parser builds it. */
const parsedShape = (...content: Paragraph[]): Document => ({
  package: {
    document: {
      content,
      sections: [{ properties: { pageWidth: 12240 }, content }],
      finalSectionProperties: { pageWidth: 12240 },
    },
  },
});

describe("the id census", () => {
  test("counts a paragraph the section view also holds once", () => {
    const document = parsedShape(paragraph("00000001", "a"), paragraph("00000002", "b"));
    const [first] = document.package.document.content;
    if (first?.type !== "paragraph") throw new Error("fixture");
    const replaced = applyDocumentOp(document, {
      type: DOCUMENT_OP_TYPES.REPLACE_BLOCKS,
      story: OP_STORIES.MAIN,
      expected: [first],
      blocks: [paragraph("00000001", "kept id")],
    });
    expect(replaced.isOk()).toBe(true);
  });

  test("finds a paragraph by its id in either case", () => {
    const document = parsedShape(paragraph("0000ABCD", "a"));
    expect(reasonOf(document, typing("0000abcd"))).toBeUndefined();
  });
});

describe("the seed contract", () => {
  test("refuses a paragraph id another story also uses, before any operation", () => {
    const document: Document = {
      package: {
        document: { content: [paragraph("00000001", "a"), paragraph("00000002", "b")] },
        footnotes: [{ type: "footnote", id: 2, content: [paragraph("00000002", "note")] }],
      },
    };
    // Joining would retire 00000002, and its inverse could not create it again.
    expect(
      reasonOf(document, {
        type: DOCUMENT_OP_TYPES.JOIN_BLOCKS,
        story: OP_STORIES.MAIN,
        blockId: "00000001",
        nextBlockId: "00000002",
      }),
    ).toBe(DOCUMENT_OP_REFUSAL_REASONS.DUPLICATE_BLOCK_ID);
  });

  test("refuses a main-story paragraph without an id", () => {
    const document = parsedShape({ type: "paragraph", content: [run("a")] });
    expect(validateOpsDocument(document).isErr()).toBe(true);
    expect(reasonOf(document, typing("00000001"))).toBe(
      DOCUMENT_OP_REFUSAL_REASONS.MISSING_BLOCK_ID,
    );
  });

  test("refuses two records carrying one revision id", () => {
    const insertion = (text: string) => ({
      type: "insertion" as const,
      info: { id: 5, author: "A" },
      content: [run(text)],
    });
    const document = parsedShape({
      type: "paragraph",
      paraId: "00000001",
      content: [insertion("a"), insertion("b")],
    });
    expect(reasonOf(document, typing("00000001"))).toBe(
      DOCUMENT_OP_REFUSAL_REASONS.DUPLICATE_RECORD_ID,
    );
  });

  test("refuses a section view out of step with the body", () => {
    const document: Document = {
      package: {
        document: {
          content: [paragraph("00000001", "a"), paragraph("00000002", "b")],
          sections: [{ properties: { pageWidth: 12240 }, content: [paragraph("00000001", "a")] }],
        },
      },
    };
    expect(reasonOf(document, typing("00000001"))).toBe(
      DOCUMENT_OP_REFUSAL_REASONS.SECTIONS_OUT_OF_STEP,
    );
  });

  test("normalizing removes empty runs and empty text nodes, and nothing else", () => {
    const document = parsedShape({
      type: "paragraph",
      paraId: "00000001",
      content: [
        { type: "run", content: [] },
        { type: "run", content: [{ type: "text", text: "" }] },
        { type: "hyperlink", children: [] },
        run("a"),
      ],
    });
    expect(reasonOf(document, typing("00000001"))).toBe(DOCUMENT_OP_REFUSAL_REASONS.EMPTY_RECORD);
    const normalized = normalizeForOps(document);
    expect(normalized.package.document.content).toEqual([
      {
        type: "paragraph",
        paraId: "00000001",
        content: [{ type: "hyperlink", children: [] }, run("a")],
      },
    ]);
    expect(normalized.package.document.sections?.[0]?.content).toEqual(
      normalized.package.document.content,
    );
    expect(validateOpsDocument(normalized).isOk()).toBe(true);
  });
});

describe("the section view", () => {
  test("is derived from the body after an edit, whatever records it held", () => {
    const content = [paragraph("00000001", "a"), paragraph("00000002", "b")];
    const document: Document = {
      package: {
        document: {
          content,
          // A view built independently of the body: equal records, no sharing.
          sections: [{ properties: { pageWidth: 12240 }, content: structuredClone(content) }],
          finalSectionProperties: { pageWidth: 12240 },
        },
      },
    };
    const edited = applyDocumentOp(document, typing("00000002"));
    if (edited.isErr()) throw edited.error;
    const body = edited.value.document.package.document;
    expect(body.sections?.[0]?.content).toEqual(body.content);
    for (const [index, block] of (body.sections?.[0]?.content ?? []).entries()) {
      expect(block).toBe(body.content[index]!);
    }
    const undone = applyDocumentOps(edited.value.document, edited.value.inverse);
    expect(undone.isOk() ? undone.value.document : undefined).toEqual(document);
  });
});
