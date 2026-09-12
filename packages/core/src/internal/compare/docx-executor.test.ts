import { describe, expect, test } from "bun:test";
import type { Node as PMNode } from "prosemirror-model";
import { EditorState } from "prosemirror-state";
import type { Transaction } from "prosemirror-state";

import { buildCleanBlockText } from "../../ai-edits/clean-text";
import type { FolioDocumentStoryHandle } from "../../ai-edits/headless";
import { createFolioAIEditSnapshot } from "../../ai-edits/snapshot";
import { updateDocumentContent } from "../../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../../prosemirror/conversion/toProseDoc";
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

type IdentifiedParagraph = {
  readonly id: string;
  readonly text: string;
  readonly alignment?: "center" | "left" | "right";
};

type TestStoryHandle = Extract<
  FolioDocumentStoryHandle,
  { readonly type: "main" | "header" }
>;

const stateFromCanonicalDocument = (doc: PMNode): EditorState => {
  const document = updateDocumentContent(createEmptyDocument(), doc);
  return EditorState.create({ doc: toProseDoc(document) });
};

const stateWithIdentifiedParagraphs = (
  ...paragraphs: readonly IdentifiedParagraph[]
): EditorState =>
  stateFromCanonicalDocument(
    schema.node(
      "doc",
      null,
      paragraphs.map(({ id, text, alignment }) =>
        schema.node(
          "paragraph",
          { paraId: id, ...(alignment === undefined ? {} : { alignment }) },
          text.length === 0 ? null : [schema.text(text)],
        ),
      ),
    ),
  );

const stateWithParagraphs = (...texts: readonly string[]): EditorState =>
  stateWithIdentifiedParagraphs(
    ...texts.map((text, index) => ({
      id: (0xa100_0000 + index).toString(16).toUpperCase(),
      text,
    })),
  );

const resolvedSnapshotOf = (
  state: EditorState,
  story: TestStoryHandle = { type: "main" },
): ResolvedDocxStorySnapshot => {
  const mainDocument = updateDocumentContent(createEmptyDocument(), state.doc);
  const document =
    story.type === "main"
      ? mainDocument
      : story.type === "header"
        ? {
            ...mainDocument,
            package: {
              ...mainDocument.package,
              headers: new Map([
                [story.relationshipId, { content: mainDocument.package.document.content }],
              ]),
            },
          }
        : mainDocument;
  const snapshot = createResolvedDocxStorySnapshot({
    document,
    story,
    sourceDocument: state.doc,
  });
  if (!snapshot) throw new Error("main story projection missing");
  return snapshot;
};

const sourceBlockOf = (snapshot: ResolvedDocxStorySnapshot, blockIndex = 0) => {
  const block = resolvedDocxContentBlocks(snapshot).at(blockIndex);
  if (!block) throw new Error("fixture source block missing");
  return block;
};

const plannedComparisonOf = ({
  baseState,
  targetState,
  story = { type: "main" },
}: {
  readonly baseState: EditorState;
  readonly targetState: EditorState;
  readonly story?: TestStoryHandle;
}) => {
  const baseSnapshot = resolvedSnapshotOf(baseState, story);
  const targetSnapshot = resolvedSnapshotOf(targetState, story);
  const captured = createContentComparisonWorkSession().captureComparison({
    base: resolvedDocxContentSnapshot(baseSnapshot),
    revised: resolvedDocxContentSnapshot(targetSnapshot),
  });
  if (captured.isErr()) throw captured.error;
  const comparison = captured.value.compare();
  if (comparison.isErr()) throw comparison.error;
  const planned = planStoryCompare({
    story,
    baseSnapshot,
    targetSnapshot,
    comparison: comparison.value,
    maxOperations: 100,
  });
  if (planned.isErr()) throw planned.error;
  return { baseSnapshot, targetSnapshot, program: planned.value.program };
};

type TableParagraph = IdentifiedParagraph;

const stateWithTableCells = (...cells: readonly (readonly TableParagraph[])[]): EditorState =>
  stateFromCanonicalDocument(
    schema.node("doc", null, [
      schema.node("table", null, [
        schema.node(
          "tableRow",
          null,
          cells.map((paragraphs) =>
            schema.node(
              "tableCell",
              null,
              paragraphs.map(({ id, text, alignment }) =>
                schema.node(
                  "paragraph",
                  { paraId: id, ...(alignment === undefined ? {} : { alignment }) },
                  text.length === 0 ? null : [schema.text(text)],
                ),
              ),
            ),
          ),
        ),
      ]),
    ]),
  );

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

const terminalMovedParagraph = ({
  snapshot,
  predecessor,
  source,
  anchor,
}: {
  readonly snapshot: ResolvedDocxStorySnapshot;
  readonly predecessor: number;
  readonly source: number;
  readonly anchor: number;
}): Extract<DocxComparisonInstructionInput, { readonly type: "moveTerminalParagraph" }> => {
  const predecessorBlock = sourceBlockOf(snapshot, predecessor);
  const sourceBlock = sourceBlockOf(snapshot, source);
  const anchorBlock = sourceBlockOf(snapshot, anchor);
  return {
    type: "moveTerminalParagraph",
    predecessor: resolvedDocxSourceOperand(snapshot, predecessorBlock),
    source: resolvedDocxSourceOperand(snapshot, sourceBlock),
    carrierTargetProperties: paragraphProperties,
    boundary: {
      type: "beforeParagraph",
      paragraph: resolvedDocxSourceOperand(snapshot, anchorBlock),
    },
    target: paragraphTarget(sourceBlock.text),
  };
};

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
    const state = stateFromCanonicalDocument(
      schema.node("doc", null, [
        schema.node("paragraph", { paraId: "A1000000" }, [schema.text("old", [bold.create()])]),
      ]),
    );
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

  for (const story of [
    { type: "main" } as const,
    { type: "header", relationshipId: "rId1" } as const,
  ]) {
    test(`moves a terminal ${story.type} paragraph through its predecessor carrier`, () => {
      const alpha = { id: "A1000000", text: "Alpha", alignment: "left" } as const;
      const beta = { id: "B1000000", text: "Beta", alignment: "right" } as const;
      const gamma = { id: "C1000000", text: "Gamma", alignment: "center" } as const;
      const baseState = stateWithIdentifiedParagraphs(alpha, beta, gamma);
      const targetState = stateWithIdentifiedParagraphs(gamma, alpha, beta);
      const { baseSnapshot, targetSnapshot, program } = plannedComparisonOf({
        baseState,
        targetState,
        story,
      });
      const prepared = preflightDocxComparisonProgram({
        state: baseState,
        snapshot: baseSnapshot,
        targetTables: resolvedDocxTableNodes(targetSnapshot),
        program,
      });
      expect(prepared.issues).toEqual([]);
      const executed = executePreflightedDocxComparison({
        state: baseState,
        prepared,
        revisionStamp: { idSeed: 950, date: "2026-09-12T00:00:00.000Z" },
        author: "Comparison",
      });
      if (executed.status !== "executed") throw new Error("Expected execution.");
      expect(executed.receipt.instructions.map(({ instructionType }) => instructionType)).toContain(
        "moveTerminalParagraph",
      );
      const tracked = baseState.apply(executed.receipt.transaction);
      expect(tracked.doc.lastChild?.attrs["pPrMark"]).toBeNull();
      expect(tracked.doc.lastChild?.attrs["_propertyChanges"]).not.toBeNull();
      const accepted = resolvedState(tracked, "accept");
      const rejected = resolvedState(tracked, "reject");

      expect(visibleTexts(accepted)).toEqual(["Gamma", "Alpha", "Beta"]);
      expect(visibleTexts(rejected)).toEqual(["Alpha", "Beta", "Gamma"]);
      expect(createFolioAIEditSnapshot(accepted.doc).blocks.at(-1)?.directAlignment).toBe("right");
      expect(createFolioAIEditSnapshot(rejected.doc).blocks.at(-1)?.directAlignment).toBe("center");
    });
  }

  test("uses the successor-owned move branch when the source is not final", () => {
    const alpha = { id: "A1000000", text: "Alpha" } as const;
    const beta = { id: "B1000000", text: "Beta" } as const;
    const gamma = { id: "C1000000", text: "Gamma" } as const;
    const baseState = stateWithIdentifiedParagraphs(alpha, beta, gamma);
    const targetState = stateWithIdentifiedParagraphs(beta, alpha, gamma);
    const { baseSnapshot, targetSnapshot, program } = plannedComparisonOf({
      baseState,
      targetState,
    });
    const prepared = preflightDocxComparisonProgram({
      state: baseState,
      snapshot: baseSnapshot,
      targetTables: resolvedDocxTableNodes(targetSnapshot),
      program,
    });
    expect(prepared.issues).toEqual([]);
    const executed = executePreflightedDocxComparison({
      state: baseState,
      prepared,
      revisionStamp: { idSeed: 960, date: "2026-09-12T00:00:00.000Z" },
      author: "Comparison",
    });
    if (executed.status !== "executed") throw new Error("Expected execution.");
    expect(executed.receipt.instructions.map(({ instructionType }) => instructionType)).toContain(
      "moveParagraph",
    );
    const tracked = baseState.apply(executed.receipt.transaction);

    expect(visibleTexts(resolvedState(tracked, "accept"))).toEqual(["Beta", "Alpha", "Gamma"]);
    expect(visibleTexts(resolvedState(tracked, "reject"))).toEqual(["Alpha", "Beta", "Gamma"]);
  });

  test("keeps a terminal move and its paragraph-mark carrier inside one table cell", () => {
    const alpha = { id: "A1000000", text: "Alpha", alignment: "left" } as const;
    const beta = { id: "B1000000", text: "Beta", alignment: "right" } as const;
    const gamma = { id: "C1000000", text: "Gamma", alignment: "center" } as const;
    const baseState = stateWithTableCells([alpha, beta, gamma]);
    const targetState = stateWithTableCells([gamma, alpha, beta]);
    const { baseSnapshot, targetSnapshot, program } = plannedComparisonOf({
      baseState,
      targetState,
    });
    const prepared = preflightDocxComparisonProgram({
      state: baseState,
      snapshot: baseSnapshot,
      targetTables: resolvedDocxTableNodes(targetSnapshot),
      program,
    });
    expect(prepared.issues).toEqual([]);
    const executed = executePreflightedDocxComparison({
      state: baseState,
      prepared,
      revisionStamp: { idSeed: 975, date: "2026-09-12T00:00:00.000Z" },
      author: "Comparison",
    });
    if (executed.status !== "executed") throw new Error("Expected execution.");
    const tracked = baseState.apply(executed.receipt.transaction);
    const accepted = resolvedState(tracked, "accept");
    const rejected = resolvedState(tracked, "reject");

    expect(tableCellTexts(accepted)).toEqual([["Gamma", "Alpha", "Beta"]]);
    expect(tableCellTexts(rejected)).toEqual([["Alpha", "Beta", "Gamma"]]);
    expect(createFolioAIEditSnapshot(accepted.doc).blocks.at(-1)?.directAlignment).toBe("right");
    expect(createFolioAIEditSnapshot(rejected.doc).blocks.at(-1)?.directAlignment).toBe("center");
  });

  test("preflight rejects a terminal-move branch whose source is not final", () => {
    const state = stateWithIdentifiedParagraphs(
      { id: "A1000000", text: "Alpha" },
      { id: "B1000000", text: "Beta" },
      { id: "C1000000", text: "Gamma" },
    );
    const snapshot = resolvedSnapshotOf(state);
    const prepared = preflightDocxComparisonProgram({
      state,
      snapshot,
      targetTables: new Map(),
      program: DocxComparisonProgram.create(snapshot, [
        terminalMovedParagraph({ snapshot, predecessor: 0, source: 1, anchor: 0 }),
      ]),
    });

    expect(prepared.supportedInstructionCount).toBe(0);
    expect(prepared.issues).toEqual([
      {
        instructionIndex: 0,
        instructionType: "moveTerminalParagraph",
        reason: "unrepresentable-paragraph-boundary",
        blockId: "B1000000",
      },
    ]);
    expect(visibleTexts(state)).toEqual(["Alpha", "Beta", "Gamma"]);
  });

  test("preflight rejects a terminal-move predecessor from another cell", () => {
    const state = stateWithTableCells(
      [
        { id: "A1000000", text: "Alpha" },
        { id: "B1000000", text: "Beta" },
      ],
      [{ id: "C1000000", text: "Gamma" }],
    );
    const snapshot = resolvedSnapshotOf(state);
    const prepared = preflightDocxComparisonProgram({
      state,
      snapshot,
      targetTables: new Map(),
      program: DocxComparisonProgram.create(snapshot, [
        terminalMovedParagraph({ snapshot, predecessor: 1, source: 2, anchor: 0 }),
      ]),
    });

    expect(prepared.supportedInstructionCount).toBe(0);
    expect(prepared.issues.at(0)).toMatchObject({
      instructionType: "moveTerminalParagraph",
      reason: "unrepresentable-paragraph-boundary",
      blockId: "C1000000",
    });
    expect(tableCellTexts(state)).toEqual([["Alpha", "Beta"], ["Gamma"]]);
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
