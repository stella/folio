import { describe, expect, test } from "bun:test";
import { EditorState } from "prosemirror-state";
import type { Transaction } from "prosemirror-state";

import { buildCleanBlockText } from "../../ai-edits/clean-text";
import { createFolioAIEditSnapshot } from "../../ai-edits/snapshot";
import { updateDocumentContent } from "../../prosemirror/conversion/fromProseDoc";
import { schema } from "../../prosemirror/schema";
import { createEmptyDocument } from "../../utils/createDocument";
import { acceptAllChanges, rejectAllChanges } from "../../prosemirror/commands/comments";
import { createContentComparisonWorkSession } from "../../compare/content";
import { planStoryCompare } from "../../compare/plan";
import { executePreflightedDocxComparison, preflightDocxComparisonProgram } from "./docx-executor";
import { DocxComparisonProgram, type DocxComparisonInstructionInput } from "./docx-program";
import {
  createResolvedDocxStorySnapshot,
  resolvedDocxContentBlocks,
  resolvedDocxContentSnapshot,
  resolvedDocxSourceOperand,
  resolvedDocxTableNodes,
  type ResolvedDocxStorySnapshot,
} from "./resolved-docx-story-snapshot";

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

const resolvedSnapshotOf = (state: EditorState): ResolvedDocxStorySnapshot => {
  const snapshot = createResolvedDocxStorySnapshot({
    document: updateDocumentContent(createEmptyDocument(), state.doc),
    story: { type: "main" },
    operationSnapshot: createFolioAIEditSnapshot(state.doc),
  });
  if (!snapshot) throw new Error("main story projection missing");
  return snapshot;
};

const sourceBlockOf = (snapshot: ResolvedDocxStorySnapshot, blockIndex = 0) => {
  const block = resolvedDocxContentBlocks(snapshot).at(blockIndex);
  if (!block) throw new Error("fixture source block missing");
  return block;
};

type TableParagraph = { readonly id: string; readonly text: string };

const stateWithTableCells = (...cells: readonly (readonly TableParagraph[])[]): EditorState =>
  EditorState.create({
    doc: schema.node("doc", null, [
      schema.node("table", null, [
        schema.node(
          "tableRow",
          null,
          cells.map((paragraphs) =>
            schema.node(
              "tableCell",
              null,
              paragraphs.map(({ id, text }) =>
                schema.node(
                  "paragraph",
                  { paraId: id },
                  text.length === 0 ? null : [schema.text(text)],
                ),
              ),
            ),
          ),
        ),
      ]),
    ]),
  });

const tableCellTexts = (state: EditorState): string[][] => {
  const table = state.doc.firstChild;
  if (table?.type.name !== "table") throw new Error("Expected one table.");
  const row = table.firstChild;
  if (row?.type.name !== "tableRow") throw new Error("Expected one table row.");
  const cells: string[][] = [];
  row.forEach((cell) => {
    const paragraphs: string[] = [];
    cell.forEach((paragraph) => paragraphs.push(paragraph.textContent));
    cells.push(paragraphs);
  });
  return cells;
};

const resolvedState = (state: EditorState, action: "accept" | "reject"): EditorState => {
  const view = {
    state,
    dispatch(transaction: Transaction) {
      view.state = view.state.apply(transaction);
    },
  };
  const command = action === "accept" ? acceptAllChanges() : rejectAllChanges();
  if (!command(view.state, view.dispatch)) {
    throw new Error(`Expected ${action} to resolve tracked changes.`);
  }
  return view.state;
};

const rangeReplacement = ({
  snapshot,
  blockIndex = 0,
  sourceStartOffset,
  sourceText,
  targetText,
}: {
  readonly snapshot: ResolvedDocxStorySnapshot;
  readonly blockIndex?: number;
  readonly sourceStartOffset: number;
  readonly sourceText: string;
  readonly targetText: string;
}): Extract<DocxComparisonInstructionInput, { readonly type: "replaceText" }> => {
  const block = sourceBlockOf(snapshot, blockIndex);
  if (block.text.slice(sourceStartOffset, sourceStartOffset + sourceText.length) !== sourceText) {
    throw new Error("fixture range does not name its source text");
  }
  return {
    type: "replaceText",
    source: resolvedDocxSourceOperand(snapshot, block),
    sourceStartOffset,
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
  };
};

const replacement = ({
  snapshot,
  blockIndex = 0,
  targetText,
}: {
  readonly snapshot: ResolvedDocxStorySnapshot;
  readonly blockIndex?: number;
  readonly targetText: string;
}): Extract<DocxComparisonInstructionInput, { readonly type: "replaceText" }> => {
  const sourceText = sourceBlockOf(snapshot, blockIndex).text;
  return rangeReplacement({
    snapshot,
    blockIndex,
    sourceStartOffset: 0,
    sourceText,
    targetText,
  });
};

const paragraphTarget = (text: string) => ({
  text,
  runs: text.length === 0 ? [] : [{ startOffset: 0, endOffset: text.length, formatting: {} }],
  properties: paragraphProperties,
});

const insertedParagraph = (
  snapshot: ResolvedDocxStorySnapshot,
  position: "after" | "before",
  text: string,
): Extract<DocxComparisonInstructionInput, { readonly type: "insertParagraph" }> => ({
  type: "insertParagraph",
  boundary: {
    type: position === "after" ? "afterParagraph" : "beforeParagraph",
    paragraph: resolvedDocxSourceOperand(snapshot, sourceBlockOf(snapshot)),
  },
  target: paragraphTarget(text),
});

const unchangedRange = (text: string) => ({
  sourceText: text,
  targetText: text,
  segments: [
    {
      type: "equal" as const,
      text,
      baseStart: 0,
      baseEnd: text.length,
      revisedStart: 0,
      revisedEnd: text.length,
    },
  ],
  sourceRuns: [{ startOffset: 0, endOffset: text.length, formatting: {} }],
  targetRuns: [{ startOffset: 0, endOffset: text.length, formatting: {} }],
  authoredChanges: [],
});

type SameBoundarySourceCase = {
  readonly name: string;
  readonly initialTexts: readonly string[];
  readonly sourceInstruction: (
    snapshot: ResolvedDocxStorySnapshot,
  ) => DocxComparisonInstructionInput;
  readonly expectedSource: readonly string[];
  readonly afterSourceIndex: number;
  readonly expectedRunPropertyChanges?: number;
};

const sameBoundarySourceCases: readonly SameBoundarySourceCase[] = [
  {
    name: "replacement",
    initialTexts: ["old"],
    sourceInstruction: (snapshot) => replacement({ snapshot, targetText: "NEW" }),
    expectedSource: ["NEW"],
    afterSourceIndex: 1,
  },
  {
    name: "deletion",
    initialTexts: ["old", "tail"],
    sourceInstruction: (snapshot) => ({
      type: "deleteParagraph",
      source: resolvedDocxSourceOperand(snapshot, sourceBlockOf(snapshot)),
    }),
    expectedSource: ["", "tail"],
    afterSourceIndex: 1,
  },
  {
    name: "formatting",
    initialTexts: ["old"],
    sourceInstruction: (snapshot) => ({
      type: "formatText",
      source: resolvedDocxSourceOperand(snapshot, sourceBlockOf(snapshot)),
      sourceStartOffset: 0,
      range: {
        ...unchangedRange("old"),
        targetRuns: [{ startOffset: 0, endOffset: 3, formatting: { bold: true } }],
        authoredChanges: [
          {
            baseStart: 0,
            baseEnd: 3,
            revisedStart: 0,
            revisedEnd: 3,
            properties: ["bold"],
          },
        ],
      },
    }),
    expectedSource: ["old"],
    afterSourceIndex: 1,
    expectedRunPropertyChanges: 1,
  },
  {
    name: "split",
    initialTexts: ["left right"],
    sourceInstruction: (snapshot) => ({
      type: "splitParagraph",
      source: resolvedDocxSourceOperand(snapshot, sourceBlockOf(snapshot)),
      offset: 4,
      first: unchangedRange("left"),
      second: unchangedRange("right"),
      separatorText: " ",
      separatorRuns: [{ startOffset: 0, endOffset: 1, formatting: {} }],
      firstTarget: paragraphTarget("left"),
      secondTarget: paragraphTarget("right"),
    }),
    expectedSource: ["left", "right"],
    afterSourceIndex: 2,
  },
];

const visibleTexts = (state: EditorState): string[] => {
  const texts: string[] = [];
  state.doc.descendants((node, position) => {
    if (!node.isTextblock) return true;
    texts.push(buildCleanBlockText(node, position).text);
    return false;
  });
  return texts;
};

const runPropertyChangeCount = (state: EditorState): number => {
  let count = 0;
  state.doc.descendants((node) => {
    count += node.marks.filter(({ type }) => type.name === "runPropertyChange").length;
  });
  return count;
};

describe("the dedicated DOCX comparison executor", () => {
  test("refuses every instruction when the live PM source is not the bound source", () => {
    const state = stateWithParagraphs("old", "elsewhere");
    const snapshot = resolvedSnapshotOf(state);
    const changedState = stateWithParagraphs("old", "changed elsewhere");
    const program = DocxComparisonProgram.create(snapshot, [
      replacement({ snapshot, targetText: "new" }),
      replacement({ snapshot, blockIndex: 1, targetText: "changed" }),
    ]);

    const prepared = preflightDocxComparisonProgram({
      state: changedState,
      snapshot,
      targetTables: new Map(),
      program,
    });

    expect(prepared.supportedInstructionCount).toBe(0);
    expect(prepared.totalInstructionCount).toBe(2);
    expect(prepared.issues.map(({ reason }) => reason)).toEqual([
      "source-expectation-mismatch",
      "source-expectation-mismatch",
    ]);
    expect(visibleTexts(changedState)).toEqual(["old", "changed elsewhere"]);
  });

  test("executes once and returns receipts in canonical instruction order", () => {
    const state = stateWithParagraphs("left", "right");
    const snapshot = resolvedSnapshotOf(state);
    const program = DocxComparisonProgram.create(snapshot, [
      replacement({ snapshot, targetText: "LEFT" }),
      replacement({ snapshot, blockIndex: 1, targetText: "RIGHT" }),
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

  for (const reverse of [false, true] as const) {
    test(`maps disjoint source ranges after a prior replacement when reverse=${String(reverse)}`, () => {
      const state = stateWithParagraphs("abcd");
      const snapshot = resolvedSnapshotOf(state);
      const first = rangeReplacement({
        snapshot,
        sourceStartOffset: 0,
        sourceText: "ab",
        targetText: "X",
      });
      const second = rangeReplacement({
        snapshot,
        sourceStartOffset: 2,
        sourceText: "cd",
        targetText: "LONG",
      });
      const prepared = preflightDocxComparisonProgram({
        state,
        snapshot,
        targetTables: new Map(),
        program: DocxComparisonProgram.create(
          snapshot,
          reverse ? [second, first] : [first, second],
        ),
      });
      const executed = executePreflightedDocxComparison({
        state,
        prepared,
        revisionStamp: { idSeed: 700, date: "2026-09-11T00:00:00.000Z" },
        author: "Comparison",
      });

      expect(executed.status).toBe("executed");
      if (executed.status !== "executed") throw new Error("expected execution");
      expect(visibleTexts(state.apply(executed.receipt.transaction))).toEqual(["XLONG"]);
    });
  }

  for (const position of ["before", "after"] as const) {
    test(`preserves canonical order for peer ${position} insertions`, () => {
      const state = stateWithParagraphs("anchor");
      const snapshot = resolvedSnapshotOf(state);
      const prepared = preflightDocxComparisonProgram({
        state,
        snapshot,
        targetTables: new Map(),
        program: DocxComparisonProgram.create(snapshot, [
          insertedParagraph(snapshot, position, "A"),
          insertedParagraph(snapshot, position, "B"),
          insertedParagraph(snapshot, position, "C"),
        ]),
      });
      const executed = executePreflightedDocxComparison({
        state,
        prepared,
        revisionStamp: { idSeed: 700, date: "2026-09-11T00:00:00.000Z" },
        author: "Comparison",
      });

      expect(executed.status).toBe("executed");
      if (executed.status !== "executed") throw new Error("expected execution");
      expect(executed.receipt.executionTaskCount).toBe(1);
      expect(executed.receipt.insertionRunCount).toBe(1);
      expect(executed.receipt.localPositionMappingSteps).toBe(0);
      expect(visibleTexts(state.apply(executed.receipt.transaction))).toEqual(
        position === "before" ? ["A", "B", "C", "anchor"] : ["anchor", "A", "B", "C"],
      );
    });
  }

  test("coalesces a large same-boundary run without accumulated position mapping", () => {
    const state = stateWithParagraphs("anchor");
    const snapshot = resolvedSnapshotOf(state);
    const instructions = Array.from({ length: 256 }, (_, index) =>
      insertedParagraph(snapshot, "after", `inserted-${String(index)}`),
    );
    const prepared = preflightDocxComparisonProgram({
      state,
      snapshot,
      targetTables: new Map(),
      program: DocxComparisonProgram.create(snapshot, instructions),
    });
    const executed = executePreflightedDocxComparison({
      state,
      prepared,
      revisionStamp: { idSeed: 700, date: "2026-09-11T00:00:00.000Z" },
      author: "Comparison",
    });

    if (executed.status !== "executed") throw new Error("expected execution");
    expect(executed.receipt.instructions).toHaveLength(256);
    expect(executed.receipt.executionTaskCount).toBe(1);
    expect(executed.receipt.insertionRunCount).toBe(1);
    expect(executed.receipt.localPositionMappingSteps).toBe(0);
    expect(visibleTexts(state.apply(executed.receipt.transaction))).toEqual([
      "anchor",
      ...instructions.map(({ target }) => target.text),
    ]);
  });

  for (const sourceCase of sameBoundarySourceCases) {
    for (const position of ["before", "after"] as const) {
      for (const sourceSlot of [0, 1, 2] as const) {
        const caseName = `orders peer ${position} insertions with ${sourceCase.name} in source slot ${String(sourceSlot)}`;
        test(caseName, () => {
          const state = stateWithParagraphs(...sourceCase.initialTexts);
          const snapshot = resolvedSnapshotOf(state);
          const source = sourceCase.sourceInstruction(snapshot);
          const instructions: DocxComparisonInstructionInput[] = [
            insertedParagraph(snapshot, position, "A"),
            insertedParagraph(snapshot, position, "B"),
          ];
          instructions.splice(sourceSlot, 0, source);
          const prepared = preflightDocxComparisonProgram({
            state,
            snapshot,
            targetTables: new Map(),
            program: DocxComparisonProgram.create(snapshot, instructions),
          });
          const executed = executePreflightedDocxComparison({
            state,
            prepared,
            revisionStamp: { idSeed: 700, date: "2026-09-11T00:00:00.000Z" },
            author: "Comparison",
          });

          expect(executed.status).toBe("executed");
          if (executed.status !== "executed") throw new Error("expected execution");
          const revised = state.apply(executed.receipt.transaction);
          const insertionIndex = position === "before" ? 0 : sourceCase.afterSourceIndex;
          const expected = [...sourceCase.expectedSource];
          expected.splice(insertionIndex, 0, "A", "B");
          expect(visibleTexts(revised)).toEqual(expected);
          expect(
            executed.receipt.instructions.map(({ instructionIndex }) => instructionIndex),
          ).toEqual([0, 1, 2]);
          if (sourceCase.expectedRunPropertyChanges !== undefined) {
            expect(runPropertyChangeCount(revised)).toBe(sourceCase.expectedRunPropertyChanges);
          }
        });
      }
    }
  }

  test("rejects live authored formatting that differs from the captured source", () => {
    const bold =
      schema.marks["bold"] ??
      (() => {
        throw new Error("schema has no bold mark");
      })();
    const state = EditorState.create({
      doc: schema.node("doc", null, [
        schema.node("paragraph", { paraId: "A1000000" }, [schema.text("old", [bold.create()])]),
      ]),
    });
    const snapshot = resolvedSnapshotOf(state);
    const instruction = replacement({ snapshot, targetText: "new" });
    Reflect.set(instruction.range.sourceRuns[0]!.formatting, "bold", false);
    const program = DocxComparisonProgram.create(snapshot, [instruction]);

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

  test("rejects a source operand copied outside its issuing capsule", () => {
    const state = stateWithParagraphs("old");
    const snapshot = resolvedSnapshotOf(state);
    const instruction = replacement({ snapshot, targetText: "new" });
    Reflect.set(instruction, "source", Object.freeze({ ...instruction.source }));
    expect(() => DocxComparisonProgram.create(snapshot, [instruction])).toThrow(
      "was not created by Folio",
    );
    expect(visibleTexts(state)).toEqual(["old"]);
  });

  test("treats duplicate live paragraph identities as ambiguous", () => {
    const snapshotState = stateWithParagraphs("old");
    const snapshot = resolvedSnapshotOf(snapshotState);
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
      program: DocxComparisonProgram.create(snapshot, [
        replacement({ snapshot, targetText: "new" }),
      ]),
    });

    expect(prepared.supportedInstructionCount).toBe(0);
    expect(prepared.issues[0]).toMatchObject({ reason: "source-expectation-mismatch" });
    expect(visibleTexts(duplicateState)).toEqual(["old", "old"]);
  });

  test("keeps inserted paragraphs and final-mark repairs inside their own table cells", () => {
    const baseState = stateWithTableCells(
      [{ id: "A1000000", text: "Alpha" }],
      [
        { id: "B1000000", text: "Keep" },
        { id: "B1000001", text: "Delete" },
      ],
    );
    const targetState = stateWithTableCells(
      [
        { id: "A1000000", text: "Alpha" },
        { id: "A1000001", text: "One" },
        { id: "A1000002", text: "Two" },
      ],
      [{ id: "B1000000", text: "Keep" }],
    );
    const baseSnapshot = resolvedSnapshotOf(baseState);
    const targetSnapshot = resolvedSnapshotOf(targetState);
    const captured = createContentComparisonWorkSession().captureComparison({
      base: resolvedDocxContentSnapshot(baseSnapshot),
      revised: resolvedDocxContentSnapshot(targetSnapshot),
    });
    if (captured.isErr()) throw captured.error;
    const comparison = captured.value.compare();
    if (comparison.isErr()) throw comparison.error;
    const planned = planStoryCompare({
      story: { type: "main" },
      baseSnapshot,
      targetSnapshot,
      comparison: comparison.value,
      maxOperations: 100,
    });
    if (planned.isErr()) throw planned.error;
    const prepared = preflightDocxComparisonProgram({
      state: baseState,
      snapshot: baseSnapshot,
      targetTables: resolvedDocxTableNodes(targetSnapshot),
      program: planned.value.program,
    });
    expect(prepared.issues).toEqual([]);
    const executed = executePreflightedDocxComparison({
      state: baseState,
      prepared,
      revisionStamp: { idSeed: 900, date: "2026-09-11T00:00:00.000Z" },
      author: "Comparison",
    });
    if (executed.status !== "executed") throw new Error("Expected execution.");
    const tracked = baseState.apply(executed.receipt.transaction);

    expect(tableCellTexts(resolvedState(tracked, "accept"))).toEqual([
      ["Alpha", "One", "Two"],
      ["Keep"],
    ]);
    expect(tableCellTexts(resolvedState(tracked, "reject"))).toEqual([
      ["Alpha"],
      ["Keep", "Delete"],
    ]);
  });
});
