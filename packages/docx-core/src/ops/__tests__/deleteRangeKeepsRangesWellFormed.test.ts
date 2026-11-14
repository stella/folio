import { describe, expect, test } from "bun:test";

import type {
  Document,
  Paragraph,
  ParagraphContent,
} from "../../model/document";
import { applyDocumentOp, applyDocumentOps } from "../apply";
import { deleteBetween } from "../inline";
import { DOCUMENT_OP_TYPES, OP_STORIES } from "../types";

const run = (text: string): ParagraphContent => ({
  type: "run",
  content: [{ type: "text", text }],
});

const paragraph = (paraId: string, content: ParagraphContent[]): Paragraph => ({
  type: "paragraph",
  paraId,
  content,
});

const documentOf = (...content: Paragraph[]): Document => ({
  package: { document: { content } },
});

const position = (blockId: string, offset: number) => ({
  story: OP_STORIES.MAIN,
  blockId,
  offset,
});

/** A comment or move-to range's two ends, as the range's family spells them. */
const FAMILIES = {
  comment: {
    start: { type: "commentRangeStart", id: 1 },
    end: { type: "commentRangeEnd", id: 1 },
  },
  "move-to": {
    start: { type: "moveToRangeStart", id: 1, name: "_Move1", author: "A" },
    end: { type: "moveToRangeEnd", id: 1 },
  },
} as const satisfies Record<
  string,
  { start: ParagraphContent; end: ParagraphContent }
>;

const deleteIn = (
  document: Document,
  blockId: string,
  from: number,
  to: number,
) => {
  const applied = applyDocumentOp(document, {
    type: DOCUMENT_OP_TYPES.DELETE_RANGE,
    from: position(blockId, from),
    to: position(blockId, to),
  });
  if (applied.isErr()) throw new Error(applied.error.message);
  return applied.value;
};

const contentOf = (document: Document, index: number): ParagraphContent[] => {
  const block = document.package.document.content.at(index);
  if (block?.type !== "paragraph") throw new Error("expected a paragraph");
  return block.content;
};

const undone = (applied: ReturnType<typeof deleteIn>): Document => {
  const restored = applyDocumentOps(
    applied.document,
    JSON.parse(JSON.stringify(applied.inverse)),
  );
  if (restored.isErr()) throw new Error(restored.error.message);
  return restored.value.document;
};

describe("a direct deletion keeps comment and move ranges well formed", () => {
  for (const [family, { start, end }] of Object.entries(FAMILIES)) {
    test(`a ${family} range's end inside the deletion stays at its edge`, () => {
      // "ab" [start] "cd" [end] "ef": delete "d", the end and "e".
      const document = documentOf(
        paragraph("00000001", [run("ab"), start, run("cd"), end, run("ef")]),
      );
      const original = structuredClone(document);

      const applied = deleteIn(document, "00000001", 3, 5);

      expect(contentOf(applied.document, 0)).toEqual([
        run("ab"),
        start,
        run("c"),
        end,
        run("f"),
      ]);
      expect(undone(applied)).toStrictEqual(original);

      // Break proof: removing every leaf between the gaps, as the deletion
      // did before, leaves the range with a start and no end.
      const unkept = deleteBetween(
        contentOf(original, 0),
        { offset: 3, zeroWidthBefore: 0 },
        { offset: 5, zeroWidthBefore: 0 },
      ).content;
      expect(unkept.some(({ type }) => type === end.type)).toBe(false);
      expect(unkept.some(({ type }) => type === start.type)).toBe(true);
    });

    test(`a ${family} range's start inside the deletion stays at its edge`, () => {
      const document = documentOf(
        paragraph("00000001", [run("ab"), start, run("cd"), end]),
      );
      const original = structuredClone(document);

      const applied = deleteIn(document, "00000001", 1, 3);

      expect(contentOf(applied.document, 0)).toEqual([
        run("a"),
        start,
        run("d"),
        end,
      ]);
      expect(undone(applied)).toStrictEqual(original);
    });

    test(`a ${family} range whose other end is in another paragraph keeps its start`, () => {
      const document = documentOf(
        paragraph("00000001", [run("ab"), start, run("cd")]),
        paragraph("00000002", [run("x"), end]),
      );

      const applied = deleteIn(document, "00000001", 1, 3);

      expect(contentOf(applied.document, 0)).toEqual([
        run("a"),
        start,
        run("d"),
      ]);
      expect(contentOf(applied.document, 1)).toEqual([run("x"), end]);
    });

    test(`a ${family} range the deletion covers whole goes whole`, () => {
      const document = documentOf(
        paragraph("00000001", [run("ab"), start, run("cd"), end, run("ef")]),
      );

      const applied = deleteIn(document, "00000001", 1, 5);

      expect(contentOf(applied.document, 0)).toEqual([run("a"), run("f")]);
    });

    test(`a ${family} range already inverted loses the end the deletion covers`, () => {
      // Not this deletion's range to repair: the end before its start goes
      // as the rest of the deleted content does.
      const document = documentOf(
        paragraph("00000001", [run("ab"), end, run("cd"), start]),
      );

      const applied = deleteIn(document, "00000001", 1, 3);

      expect(contentOf(applied.document, 0)).toEqual([
        run("a"),
        run("d"),
        start,
      ]);
    });
  }
});
