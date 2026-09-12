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
import {
  createResolvedDocxStorySnapshot,
  type ResolvedDocxStorySnapshot,
} from "./resolved-docx-story-snapshot";
import {
  compareResolvedDocxStoryPair,
  createResolvedDocxStoryPair,
  type ResolvedDocxStoryComparison,
} from "./resolved-docx-story-comparison";

const boldMark =
  schema.marks["bold"] ??
  (() => {
    throw new Error("schema has no bold mark");
  })();

type IdentifiedParagraph = {
  readonly id: string;
  readonly text: string;
  readonly alignment?: "center" | "left" | "right";
  readonly bold?: boolean;
};

type TestStoryHandle = Extract<FolioDocumentStoryHandle, { readonly type: "main" | "header" }>;

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
      paragraphs.map(({ id, text, alignment, bold }) =>
        schema.node(
          "paragraph",
          { paraId: id, ...(alignment === undefined ? {} : { alignment }) },
          text.length === 0 ? null : [schema.text(text, bold === true ? [boldMark.create()] : [])],
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

const comparisonOf = (
  baseSnapshot: ResolvedDocxStorySnapshot,
  targetSnapshot: ResolvedDocxStorySnapshot = baseSnapshot,
): ResolvedDocxStoryComparison => {
  const compared = compareResolvedDocxStoryPair({
    pair: createResolvedDocxStoryPair({ baseSnapshot, targetSnapshot }),
    workSession: createContentComparisonWorkSession(),
  });
  if (compared.isErr()) throw compared.error;
  return compared.value;
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
  const comparison = comparisonOf(baseSnapshot, targetSnapshot);
  const planned = planStoryCompare({
    comparison,
    maxOperations: 1_000,
  });
  if (planned.isErr()) throw planned.error;
  return { baseSnapshot, comparison, targetSnapshot, program: planned.value.program };
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

type SameBoundaryCase = {
  readonly name: string;
  readonly base: readonly IdentifiedParagraph[];
  readonly revised: readonly IdentifiedParagraph[];
  readonly expectedRunPropertyChanges?: number;
};

const sameBoundaryCases: readonly SameBoundaryCase[] = [
  {
    name: "replacement",
    base: [{ id: "A1000000", text: "old" }],
    revised: [{ id: "A1000000", text: "NEW" }],
  },
  {
    name: "deletion",
    base: [
      { id: "A1000000", text: "old" },
      { id: "A1000001", text: "tail" },
    ],
    revised: [{ id: "A1000001", text: "tail" }],
  },
  {
    name: "formatting",
    base: [{ id: "A1000000", text: "old" }],
    revised: [{ id: "A1000000", text: "old", bold: true }],
    expectedRunPropertyChanges: 1,
  },
  {
    name: "split",
    base: [{ id: "A1000000", text: "left right" }],
    revised: [
      { id: "A1000000", text: "left" },
      { id: "A1000001", text: "right" },
    ],
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
    const targetState = stateWithParagraphs("new", "changed");
    const changedState = stateWithParagraphs("old", "changed elsewhere");
    const { program } = plannedComparisonOf({ baseState: state, targetState });

    const prepared = preflightDocxComparisonProgram({
      state: changedState,
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

  test("refuses a structurally equal PM state that is not the captured source", () => {
    const capturedState = stateWithParagraphs("old");
    const targetState = stateWithParagraphs("new");
    const structurallyEqualState = stateWithParagraphs("old");
    expect(structurallyEqualState.doc.eq(capturedState.doc)).toBe(true);
    expect(structurallyEqualState.doc).not.toBe(capturedState.doc);

    const prepared = preflightDocxComparisonProgram({
      state: structurallyEqualState,
      program: plannedComparisonOf({ baseState: capturedState, targetState }).program,
    });

    expect(prepared.supportedInstructionCount).toBe(0);
    expect(prepared.issues).toEqual([
      {
        instructionIndex: 0,
        instructionType: "replaceText",
        reason: "source-expectation-mismatch",
      },
    ]);
    expect(visibleTexts(structurallyEqualState)).toEqual(["old"]);
  });

  test("executes once and returns receipts in canonical instruction order", () => {
    const state = stateWithParagraphs("left", "right");
    const targetState = stateWithParagraphs("LEFT", "RIGHT");
    const { program } = plannedComparisonOf({ baseState: state, targetState });
    const prepared = preflightDocxComparisonProgram({
      state,
      program,
    });
    expect(prepared.supportedChangeCount).toBe(2);

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
    expect(executed.receipt.changes.map(({ kind }) => kind)).toEqual(["replace", "replace"]);
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

  test("executes disjoint text edits from one canonical pair relation", () => {
    const state = stateWithParagraphs("abcd");
    const targetState = stateWithParagraphs("XcLONG");
    const prepared = preflightDocxComparisonProgram({
      state,
      program: plannedComparisonOf({ baseState: state, targetState }).program,
    });
    expect(prepared.totalInstructionCount).toBe(1);
    const executed = executePreflightedDocxComparison({
      state,
      prepared,
      revisionStamp: { idSeed: 700, date: "2026-09-11T00:00:00.000Z" },
      author: "Comparison",
    });

    expect(executed.status).toBe("executed");
    if (executed.status !== "executed") throw new Error("expected execution");
    expect(visibleTexts(state.apply(executed.receipt.transaction))).toEqual(["XcLONG"]);
  });

  for (const position of ["before", "after"] as const) {
    test(`preserves canonical order for peer ${position} insertions`, () => {
      const anchor = { id: "A1000000", text: "anchor" } as const;
      const inserted = [
        { id: "A1000001", text: "A" },
        { id: "A1000002", text: "B" },
        { id: "A1000003", text: "C" },
      ] as const;
      const state = stateWithIdentifiedParagraphs(anchor);
      const targetState = stateWithIdentifiedParagraphs(
        ...(position === "before" ? [...inserted, anchor] : [anchor, ...inserted]),
      );
      const prepared = preflightDocxComparisonProgram({
        state,
        program: plannedComparisonOf({ baseState: state, targetState }).program,
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
    const anchor = { id: "A1000000", text: "anchor" } as const;
    const inserted = Array.from({ length: 256 }, (_, index) => ({
      id: `inserted-${String(index)}`,
      text: `inserted-${String(index)}`,
    }));
    const state = stateWithIdentifiedParagraphs(anchor);
    const targetState = stateWithIdentifiedParagraphs(anchor, ...inserted);
    const prepared = preflightDocxComparisonProgram({
      state,
      program: plannedComparisonOf({ baseState: state, targetState }).program,
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
      ...inserted.map(({ text }) => text),
    ]);
  });

  for (const sourceCase of sameBoundaryCases) {
    for (const position of ["before", "after"] as const) {
      test(`resolves ${sourceCase.name} with peer ${position} insertions`, () => {
        const inserted = [
          { id: "B1000000", text: "A" },
          { id: "B1000001", text: "B" },
        ] as const;
        const revised =
          position === "before"
            ? [...inserted, ...sourceCase.revised]
            : [...sourceCase.revised, ...inserted];
        const state = stateWithIdentifiedParagraphs(...sourceCase.base);
        const targetState = stateWithIdentifiedParagraphs(...revised);
        const prepared = preflightDocxComparisonProgram({
          state,
          program: plannedComparisonOf({ baseState: state, targetState }).program,
        });
        expect(prepared.issues).toEqual([]);
        const executed = executePreflightedDocxComparison({
          state,
          prepared,
          revisionStamp: { idSeed: 700, date: "2026-09-11T00:00:00.000Z" },
          author: "Comparison",
        });

        expect(executed.status).toBe("executed");
        if (executed.status !== "executed") throw new Error("expected execution");
        const tracked = state.apply(executed.receipt.transaction);
        expect(visibleTexts(resolvedState(tracked, "accept"))).toEqual(
          revised.map(({ text }) => text),
        );
        expect(visibleTexts(resolvedState(tracked, "reject"))).toEqual(
          sourceCase.base.map(({ text }) => text),
        );
        if (sourceCase.expectedRunPropertyChanges !== undefined) {
          expect(runPropertyChangeCount(tracked)).toBe(sourceCase.expectedRunPropertyChanges);
        }
      });
    }
  }

  test("treats duplicate live paragraph identities as ambiguous", () => {
    const snapshotState = stateWithParagraphs("old");
    const targetState = stateWithParagraphs("new");
    const paraId = snapshotState.doc.firstChild?.attrs["paraId"];
    const duplicateState = EditorState.create({
      doc: schema.node("doc", null, [
        schema.node("paragraph", { paraId }, [schema.text("old")]),
        schema.node("paragraph", { paraId }, [schema.text("old")]),
      ]),
    });
    const prepared = preflightDocxComparisonProgram({
      state: duplicateState,
      program: plannedComparisonOf({ baseState: snapshotState, targetState }).program,
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
      const { program } = plannedComparisonOf({
        baseState,
        targetState,
        story,
      });
      const prepared = preflightDocxComparisonProgram({
        state: baseState,
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
    const { program } = plannedComparisonOf({
      baseState,
      targetState,
    });
    const prepared = preflightDocxComparisonProgram({
      state: baseState,
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
    const { program } = plannedComparisonOf({
      baseState,
      targetState,
    });
    const prepared = preflightDocxComparisonProgram({
      state: baseState,
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

  test("preflight rejects every required instruction in an incomplete semantic group", () => {
    const state = stateFromCanonicalDocument(
      schema.node("doc", null, [
        schema.node(
          "paragraph",
          {
            paraId: "A1000000",
            alignment: "left",
            _propertyChanges: [
              {
                type: "paragraphPropertyChange",
                info: {
                  id: 4,
                  author: "Existing",
                  date: "2026-09-01T00:00:00.000Z",
                },
                previousFormatting: { alignment: "right" },
              },
            ],
          },
          [schema.text("Alpha")],
        ),
      ]),
    );
    const targetState = stateWithIdentifiedParagraphs({
      id: "A1000000",
      text: "Revised",
      alignment: "center",
    });
    const { program } = plannedComparisonOf({ baseState: state, targetState });
    const prepared = preflightDocxComparisonProgram({
      state,
      program,
    });

    expect(prepared.supportedInstructionCount).toBe(0);
    expect(prepared.supportedChangeCount).toBe(0);
    expect(prepared.issues.map(({ instructionIndex }) => instructionIndex)).toEqual([0, 1]);
    expect(prepared.issues).toContainEqual({
      instructionIndex: 0,
      instructionType: "replaceText",
      reason: "semantic-group-incomplete",
    });
    expect(prepared.issues).toContainEqual({
      instructionIndex: 1,
      instructionType: "setParagraphProperties",
      reason: "pending-paragraph-change",
      blockId: "A1000000",
    });
    const executed = executePreflightedDocxComparison({
      state,
      prepared,
      revisionStamp: { idSeed: 980, date: "2026-09-12T00:00:00.000Z" },
      author: "Comparison",
    });
    if (executed.status !== "executed") throw new Error("Expected execution.");
    expect(executed.receipt.instructions).toEqual([]);
    expect(executed.receipt.changes).toEqual([]);
    expect(executed.receipt.transaction.steps).toEqual([]);
    expect(visibleTexts(state.apply(executed.receipt.transaction))).toEqual(["Alpha"]);
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
    const planned = planStoryCompare({
      comparison: comparisonOf(baseSnapshot, targetSnapshot),
      maxOperations: 100,
    });
    if (planned.isErr()) throw planned.error;
    const prepared = preflightDocxComparisonProgram({
      state: baseState,
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
