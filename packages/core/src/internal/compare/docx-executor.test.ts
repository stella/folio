import { describe, expect, test } from "bun:test";
import { EditorState } from "prosemirror-state";

import { buildCleanBlockText } from "../../ai-edits/clean-text";
import { createFolioAIEditSnapshot } from "../../ai-edits/snapshot";
import { schema } from "../../prosemirror/schema";
import {
  executePreflightedDocxComparison,
  preflightDocxComparisonProgram,
} from "./docx-executor";
import {
  DocxComparisonProgram,
  type DocxComparisonInstructionInput,
} from "./docx-program";

const paragraphProperties = Object.freeze({
  styleId: null,
  listLevel: null,
  alignment: null,
  spacing: null,
});

const stateWithParagraphs = (...texts: readonly string[]): EditorState =>
  EditorState.create({
    doc: schema.node(
      "doc",
      null,
      texts.map((text, index) =>
        schema.node(
          "paragraph",
          { paraId: (0xa100_0000 + index).toString(16).toUpperCase() },
          text.length === 0 ? null : [schema.text(text)],
        ),
      ),
    ),
  });

const replacement = ({
  blockId,
  sourceText,
  targetText,
}: {
  readonly blockId: string;
  readonly sourceText: string;
  readonly targetText: string;
}): Extract<DocxComparisonInstructionInput, { readonly type: "replaceText" }> => ({
  type: "replaceText",
  source: {
    blockId,
    kind: "paragraph",
    text: sourceText,
    paragraphProperties,
    structuralBoundaries: [],
    containerPath: [],
  },
  sourceStartOffset: 0,
  range: {
    sourceText,
    targetText,
    segments: [
      {
        type: "del",
        text: sourceText,
        baseStart: 0,
        baseEnd: sourceText.length,
        revisedStart: 0,
        revisedEnd: 0,
      },
      {
        type: "ins",
        text: targetText,
        baseStart: sourceText.length,
        baseEnd: sourceText.length,
        revisedStart: 0,
        revisedEnd: targetText.length,
      },
    ],
    sourceRuns: [{ startOffset: 0, endOffset: sourceText.length, formatting: {} }],
    targetRuns: [{ startOffset: 0, endOffset: targetText.length, formatting: {} }],
    authoredChanges: [],
  },
});

const visibleTexts = (state: EditorState): string[] => {
  const texts: string[] = [];
  state.doc.descendants((node, position) => {
    if (!node.isTextblock) return true;
    texts.push(buildCleanBlockText(node, position).text);
    return false;
  });
  return texts;
};

describe("the dedicated DOCX comparison executor", () => {
  test("preflights every instruction without mutating an earlier valid source", () => {
    const state = stateWithParagraphs("old");
    const snapshot = createFolioAIEditSnapshot(state.doc);
    const blockId = snapshot.blocks[0]?.id ?? "";
    const program = DocxComparisonProgram.create([
      replacement({ blockId, sourceText: "old", targetText: "new" }),
      replacement({ blockId: "missing", sourceText: "elsewhere", targetText: "changed" }),
    ]);

    const prepared = preflightDocxComparisonProgram({
      state,
      snapshot,
      targetTables: new Map(),
      program,
    });

    expect(prepared.supportedInstructionCount).toBe(1);
    expect(prepared.totalInstructionCount).toBe(2);
    expect(prepared.issues).toEqual([
      {
        instructionIndex: 1,
        instructionType: "replaceText",
        reason: "missing-block",
        blockId: "missing",
      },
    ]);
    expect(visibleTexts(state)).toEqual(["old"]);
  });

  test("executes once and returns receipts in canonical instruction order", () => {
    const state = stateWithParagraphs("left", "right");
    const snapshot = createFolioAIEditSnapshot(state.doc);
    const leftId = snapshot.blocks[0]?.id ?? "";
    const rightId = snapshot.blocks[1]?.id ?? "";
    const program = DocxComparisonProgram.create([
      replacement({ blockId: leftId, sourceText: "left", targetText: "LEFT" }),
      replacement({ blockId: rightId, sourceText: "right", targetText: "RIGHT" }),
    ]);
    const prepared = preflightDocxComparisonProgram({
      state,
      snapshot,
      targetTables: new Map(),
      program,
    });

    const executed = executePreflightedDocxComparison({
      state,
      prepared,
      revisionStamp: { idSeed: 700, date: "2026-09-11T00:00:00.000Z" },
      author: "Comparison",
    });
    expect(executed.status).toBe("executed");
    if (executed.status !== "executed") throw new Error("expected execution");
    expect(executed.receipt.instructions.map(({ instructionIndex }) => instructionIndex)).toEqual([
      0, 1,
    ]);
    expect(executed.receipt.instructions.map(({ instructionType }) => instructionType)).toEqual([
      "replaceText",
      "replaceText",
    ]);
    expect(visibleTexts(state.apply(executed.receipt.transaction))).toEqual(["LEFT", "RIGHT"]);

    expect(() =>
      executePreflightedDocxComparison({
        state,
        prepared,
        revisionStamp: { idSeed: 800, date: "2026-09-11T00:00:00.000Z" },
        author: "Comparison",
      }),
    ).toThrow("consumed more than once");
  });

  test("rejects live authored formatting that differs from the captured source", () => {
    const bold = schema.marks["bold"] ?? (() => {
      throw new Error("schema has no bold mark");
    })();
    const state = EditorState.create({
      doc: schema.node("doc", null, [
        schema.node("paragraph", { paraId: "A1000000" }, [
          schema.text("old", [bold.create()]),
        ]),
      ]),
    });
    const snapshot = createFolioAIEditSnapshot(state.doc);
    const blockId = snapshot.blocks[0]?.id ?? "";
    const instruction = replacement({ blockId, sourceText: "old", targetText: "new" });
    Reflect.set(instruction.range.sourceRuns[0]!.formatting, "bold", false);
    const program = DocxComparisonProgram.create([instruction]);

    const prepared = preflightDocxComparisonProgram({
      state,
      snapshot,
      targetTables: new Map(),
      program,
    });
    expect(prepared.supportedInstructionCount).toBe(0);
    expect(prepared.issues[0]).toMatchObject({ reason: "source-formatting-mismatch" });
    expect(visibleTexts(state)).toEqual(["old"]);
  });

  test("rejects stale structural placement before creating a transaction", () => {
    const state = stateWithParagraphs("old");
    const snapshot = createFolioAIEditSnapshot(state.doc);
    const blockId = snapshot.blocks[0]?.id ?? "";
    const instruction = replacement({ blockId, sourceText: "old", targetText: "new" });
    Reflect.set(instruction.source, "structuralBoundaries", [
      { type: "pageBreak", offset: 0 },
    ]);
    const prepared = preflightDocxComparisonProgram({
      state,
      snapshot,
      targetTables: new Map(),
      program: DocxComparisonProgram.create([instruction]),
    });

    expect(prepared.supportedInstructionCount).toBe(0);
    expect(prepared.issues[0]).toMatchObject({ reason: "source-expectation-mismatch" });
    expect(visibleTexts(state)).toEqual(["old"]);
  });

  test("treats duplicate live paragraph identities as ambiguous", () => {
    const snapshotState = stateWithParagraphs("old");
    const snapshot = createFolioAIEditSnapshot(snapshotState.doc);
    const blockId = snapshot.blocks[0]?.id ?? "";
    const paraId = snapshotState.doc.firstChild?.attrs["paraId"];
    const duplicateState = EditorState.create({
      doc: schema.node("doc", null, [
        schema.node("paragraph", { paraId }, [schema.text("old")]),
        schema.node("paragraph", { paraId }, [schema.text("old")]),
      ]),
    });
    const prepared = preflightDocxComparisonProgram({
      state: duplicateState,
      snapshot,
      targetTables: new Map(),
      program: DocxComparisonProgram.create([
        replacement({ blockId, sourceText: "old", targetText: "new" }),
      ]),
    });

    expect(prepared.supportedInstructionCount).toBe(0);
    expect(prepared.issues[0]).toMatchObject({ reason: "changed-block" });
    expect(visibleTexts(duplicateState)).toEqual(["old", "old"]);
  });
});
