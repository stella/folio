/**
 * Regression coverage for sequential note numbering.
 *
 * Word numbers footnotes/endnotes by reference order in the body, not by
 * their `w:id` values. Documents with non-contiguous or out-of-order ids
 * (e.g. ids 5, 2, 9 referenced in that order) must display 1, 2, 3 in BOTH
 * the body marker and the footnote area; separator/continuation notes must
 * not consume numbers even when they carry a positive id.
 */

import { describe, expect, test } from "bun:test";

import type {
  FlowBlock,
  ParagraphBlock,
  TableBlock,
  TextBoxBlock,
} from "../../layout-engine/types";
import type { Footnote } from "../../types/document";
import {
  buildFootnoteContentMap,
  computeNoteDisplayNumbers,
  remapNoteMarkerText,
} from "./footnoteLayout";

const normalFootnote = (id: number, text: string): Footnote => ({
  type: "footnote",
  id,
  noteType: "normal",
  content: [{ type: "paragraph", content: [{ type: "run", content: [{ type: "text", text }] }] }],
});

const continuationNotice = (id: number): Footnote => ({
  type: "footnote",
  id,
  noteType: "continuationNotice",
  content: [{ type: "paragraph", content: [] }],
});

function paragraphWithFootnoteRef(id: string, footnoteId: number, pmStart: number): ParagraphBlock {
  return {
    kind: "paragraph",
    id,
    runs: [
      {
        kind: "text",
        text: footnoteId.toString(),
        footnoteRefId: footnoteId,
        pmStart,
        pmEnd: pmStart + footnoteId.toString().length,
      },
    ],
  };
}

describe("computeNoteDisplayNumbers", () => {
  test("numbers non-contiguous out-of-order ids by reference order", () => {
    const notes = [normalFootnote(2, "two"), normalFootnote(5, "five"), normalFootnote(9, "nine")];

    const numbers = computeNoteDisplayNumbers(notes, [5, 2, 9]);

    expect(numbers.get(5)).toBe(1);
    expect(numbers.get(2)).toBe(2);
    expect(numbers.get(9)).toBe(3);
  });

  test("repeated references keep the first number and consume no new one", () => {
    const notes = [normalFootnote(2, "two"), normalFootnote(5, "five"), normalFootnote(9, "nine")];

    const numbers = computeNoteDisplayNumbers(notes, [5, 2, 5, 9]);

    expect(numbers.get(5)).toBe(1);
    expect(numbers.get(2)).toBe(2);
    expect(numbers.get(9)).toBe(3);
  });

  test("positive-id continuation notices and unknown ids do not shift numbering", () => {
    const notes = [continuationNotice(1), normalFootnote(5, "five"), normalFootnote(2, "two")];

    const numbers = computeNoteDisplayNumbers(notes, [1, 5, 7, 2]);

    expect(numbers.has(1)).toBe(false);
    expect(numbers.has(7)).toBe(false);
    expect(numbers.get(5)).toBe(1);
    expect(numbers.get(2)).toBe(2);
  });
});

describe("remapNoteMarkerText", () => {
  test("rewrites body marker text from raw w:id to the display number", () => {
    const blocks: FlowBlock[] = [
      paragraphWithFootnoteRef("p1", 5, 10),
      paragraphWithFootnoteRef("p2", 2, 20),
      paragraphWithFootnoteRef("p3", 9, 30),
    ];
    const footnoteNumbers = new Map([
      [5, 1],
      [2, 2],
      [9, 3],
    ]);

    const remapped = remapNoteMarkerText(blocks, { footnoteNumbers });

    const texts = remapped.map((block) =>
      block.kind === "paragraph" ? block.runs.at(0) : undefined,
    );
    expect(texts.at(0)).toMatchObject({ text: "1", footnoteRefId: 5, pmStart: 10, pmEnd: 11 });
    expect(texts.at(1)).toMatchObject({ text: "2", footnoteRefId: 2, pmStart: 20, pmEnd: 21 });
    expect(texts.at(2)).toMatchObject({ text: "3", footnoteRefId: 9, pmStart: 30, pmEnd: 31 });
  });

  test("returns untouched blocks and runs by reference", () => {
    const plain: ParagraphBlock = {
      kind: "paragraph",
      id: "plain",
      runs: [{ kind: "text", text: "no refs here" }],
    };
    const marker = paragraphWithFootnoteRef("marker", 5, 10);
    const blocks: FlowBlock[] = [plain, marker];

    const remapped = remapNoteMarkerText(blocks, { footnoteNumbers: new Map([[5, 1]]) });

    expect(remapped.at(0)).toBe(plain);
    expect(remapped.at(1)).not.toBe(marker);
  });

  test("returns the input array when there is nothing to remap", () => {
    const blocks: FlowBlock[] = [paragraphWithFootnoteRef("p1", 5, 10)];

    expect(remapNoteMarkerText(blocks, {})).toBe(blocks);
    expect(remapNoteMarkerText(blocks, { footnoteNumbers: new Map() })).toBe(blocks);
  });

  test("leaves markers without a display number unchanged", () => {
    const blocks: FlowBlock[] = [paragraphWithFootnoteRef("p1", 7, 10)];

    const remapped = remapNoteMarkerText(blocks, { footnoteNumbers: new Map([[5, 1]]) });

    expect(remapped.at(0)).toBe(blocks[0]);
  });

  test("recurses into table cells and text boxes", () => {
    const table: TableBlock = {
      kind: "table",
      id: "t1",
      rows: [
        {
          id: "r1",
          cells: [{ id: "c1", blocks: [paragraphWithFootnoteRef("cell-p", 5, 100)] }],
        },
      ],
    };
    const textBox: TextBoxBlock = {
      kind: "textBox",
      id: "tb1",
      width: 200,
      content: [paragraphWithFootnoteRef("tb-p", 9, 200)],
    };
    const footnoteNumbers = new Map([
      [5, 1],
      [9, 2],
    ]);

    const remapped = remapNoteMarkerText([table, textBox], { footnoteNumbers });

    const remappedTable = remapped.at(0);
    if (remappedTable?.kind !== "table") {
      throw new Error("Expected a table block");
    }
    const cellParagraph = remappedTable.rows.at(0)?.cells.at(0)?.blocks.at(0);
    if (cellParagraph?.kind !== "paragraph") {
      throw new Error("Expected a paragraph in the table cell");
    }
    expect(cellParagraph.runs.at(0)).toMatchObject({ text: "1", footnoteRefId: 5 });

    const remappedBox = remapped.at(1);
    if (remappedBox?.kind !== "textBox") {
      throw new Error("Expected a text box block");
    }
    const boxParagraph = remappedBox.content.at(0);
    if (boxParagraph?.kind !== "paragraph") {
      throw new Error("Expected a paragraph in the text box");
    }
    expect(boxParagraph.runs.at(0)).toMatchObject({ text: "2", footnoteRefId: 9 });
  });

  test("remaps endnote markers independently of footnote markers", () => {
    const blocks: FlowBlock[] = [
      paragraphWithFootnoteRef("fn", 5, 10),
      {
        kind: "paragraph",
        id: "en",
        runs: [{ kind: "text", text: "8", endnoteRefId: 8, pmStart: 20, pmEnd: 21 }],
      },
    ];

    const remapped = remapNoteMarkerText(blocks, {
      footnoteNumbers: new Map([[5, 1]]),
      endnoteNumbers: new Map([[8, 1]]),
    });

    const footnoteParagraph = remapped.at(0);
    const endnoteParagraph = remapped.at(1);
    if (footnoteParagraph?.kind !== "paragraph" || endnoteParagraph?.kind !== "paragraph") {
      throw new Error("Expected paragraph blocks");
    }
    expect(footnoteParagraph.runs.at(0)).toMatchObject({ text: "1", footnoteRefId: 5 });
    expect(endnoteParagraph.runs.at(0)).toMatchObject({ text: "1", endnoteRefId: 8 });
  });
});

describe("buildFootnoteContentMap", () => {
  const measureBlocks = (blocks: FlowBlock[]) =>
    blocks.map(() => ({ kind: "paragraph" as const, lines: [], totalHeight: 12 }));

  test("assigns area display numbers by reference order, not id order", () => {
    const footnotes = [
      continuationNotice(1),
      normalFootnote(2, "two"),
      normalFootnote(5, "five"),
      normalFootnote(9, "nine"),
    ];
    const refs = [{ footnoteId: 5 }, { footnoteId: 2 }, { footnoteId: 9 }];

    const contentMap = buildFootnoteContentMap(footnotes, refs, 400, { measureBlocks });

    expect(contentMap.get(5)?.displayNumber).toBe(1);
    expect(contentMap.get(2)?.displayNumber).toBe(2);
    expect(contentMap.get(9)?.displayNumber).toBe(3);
    expect(contentMap.has(1)).toBe(false);

    const firstBlock = contentMap.get(5)?.blocks.at(0);
    if (firstBlock?.kind !== "paragraph") {
      throw new Error("Expected the footnote content to start with a paragraph");
    }
    expect(firstBlock.runs.at(0)).toMatchObject({ kind: "text", text: "1 " });
  });
});
