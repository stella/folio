import { describe, expect, test } from "bun:test";
import { Fragment, type Node as PMNode } from "prosemirror-model";
import { EditorState } from "prosemirror-state";
import type { Transaction } from "prosemirror-state";

import { buildCleanBlockText } from "../../ai-edits/clean-text";
import { stripBlockIdentityAttrs } from "../../ai-edits/block-identity";
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
import { DocxComparisonProgram } from "./docx-program";
import {
  createResolvedDocxStorySnapshot,
  resolvedDocxSourceOperand,
  type ResolvedDocxStorySnapshot,
} from "./resolved-docx-story-snapshot";
import {
  compareResolvedDocxStoryPair,
  createResolvedDocxStoryPair,
  resolvedDocxStoryComparisonPayload,
  resolvedDocxTableComponents,
  resolvedDocxTableStructureOperand,
  type ResolvedDocxStoryComparison,
} from "./resolved-docx-story-comparison";
import {
  DEFAULT_TABLE_STRUCTURE_PREFLIGHT_LIMITS,
  preflightTableStructureComponents,
} from "./table-structure-program";

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
  let document = mainDocument;
  if (story.type === "header") {
    document = {
      ...mainDocument,
      package: {
        ...mainDocument.package,
        headers: new Map([
          [story.relationshipId, { content: mainDocument.package.document.content }],
        ]),
      },
    };
  }
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

const tableWithTextGrid = (
  rows: readonly (readonly string[])[],
  attrs?: {
    readonly width?: number;
    readonly widthType?: "dxa";
    readonly justification?: "left" | "center" | "right";
  },
): PMNode =>
  schema.node(
    "table",
    attrs ?? null,
    rows.map((texts, rowIndex) =>
      schema.node(
        "tableRow",
        null,
        texts.map((text, columnIndex) =>
          schema.node("tableCell", null, [
            schema.node(
              "paragraph",
              { paraId: `T${String(rowIndex)}${String(columnIndex)}00000` },
              text.length === 0 ? null : [schema.text(text)],
            ),
          ]),
        ),
      ),
    ),
  );

const tableWithCellGrid = (
  rows: readonly (readonly { readonly id: string; readonly text: string }[])[],
): PMNode =>
  schema.node(
    "table",
    null,
    rows.map((cells) =>
      schema.node(
        "tableRow",
        null,
        cells.map(({ id, text }) => schema.node("tableCell", null, [paragraphNode(id, text)])),
      ),
    ),
  );

const paragraphNode = (id: string, text: string): PMNode =>
  schema.node("paragraph", { paraId: id }, text.length === 0 ? null : [schema.text(text)]);

const stateWithBlocks = (...blocks: readonly PMNode[]): EditorState =>
  stateFromCanonicalDocument(schema.node("doc", null, blocks));

const withoutBlockIdentities = (node: PMNode): PMNode => {
  if (node.isText) return node;
  const children: PMNode[] = [];
  node.forEach((child) => children.push(withoutBlockIdentities(child)));
  return node.type.create(
    node.isTextblock ? stripBlockIdentityAttrs(node.attrs) : node.attrs,
    Fragment.fromArray(children),
    node.marks,
  );
};

const expectResolvedViews = (
  baseState: EditorState,
  targetState: EditorState,
  tracked: EditorState,
): void => {
  expect(
    withoutBlockIdentities(resolvedState(tracked, "accept").doc).eq(
      withoutBlockIdentities(targetState.doc),
    ),
  ).toBe(true);
  expect(
    withoutBlockIdentities(resolvedState(tracked, "reject").doc).eq(
      withoutBlockIdentities(baseState.doc),
    ),
  ).toBe(true);
};

const executePlannedComparison = (baseState: EditorState, targetState: EditorState) => {
  const prepared = preflightDocxComparisonProgram({
    state: baseState,
    program: plannedComparisonOf({ baseState, targetState }).program,
  });
  expect(prepared.issues).toEqual([]);
  const executed = executePreflightedDocxComparison({
    state: baseState,
    prepared,
    revisionStamp: { idSeed: 1_100, date: "2026-09-12T00:00:00.000Z" },
    author: "Comparison",
  });
  if (executed.status !== "executed") throw new Error("Expected table comparison execution.");
  return { executed, tracked: baseState.apply(executed.receipt.transaction) };
};

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

  test("executes a terminal table replacement and its final carrier as one obligation", () => {
    const mergedTarget = schema.node("table", null, [
      schema.node("tableRow", null, [
        schema.node("tableCell", { colspan: 2 }, [paragraphNode("T2000000", "New")]),
      ]),
    ]);
    const baseState = stateWithBlocks(
      tableWithTextGrid([["Old", "Other"]]),
      paragraphNode("C1000000", ""),
    );
    const targetState = stateWithBlocks(mergedTarget);
    const { executed, tracked } = executePlannedComparison(baseState, targetState);

    expect(executed.receipt.instructions.map(({ instructionType }) => instructionType)).toEqual([
      "replaceTable",
    ]);
    expect(executed.receipt.instructions.at(0)?.revisionIds).toHaveLength(3);
    expectResolvedViews(baseState, targetState, tracked);
  });

  test("whole table insertion, deletion, and ordinary replacement reconstruct both views", () => {
    const prefix = paragraphNode("A1000000", "Before");
    const suffix = paragraphNode("B1000000", "After");
    const cases = [
      {
        base: stateWithBlocks(prefix, suffix),
        target: stateWithBlocks(prefix, tableWithTextGrid([["Added"]]), suffix),
      },
      {
        base: stateWithBlocks(prefix, tableWithTextGrid([["Removed"]]), suffix),
        target: stateWithBlocks(prefix, suffix),
      },
      {
        base: stateWithBlocks(prefix, tableWithTextGrid([["Old"]]), suffix),
        target: stateWithBlocks(prefix, tableWithTextGrid([["New"]]), suffix),
      },
    ];

    for (const { base, target } of cases) {
      const { tracked } = executePlannedComparison(base, target);
      expectResolvedViews(base, target, tracked);
    }
  });

  test("table formatting is one reported and reversible semantic operation", () => {
    const baseState = stateWithBlocks(
      tableWithTextGrid([["Same"]], {
        width: 6_000,
        widthType: "dxa",
        justification: "left",
      }),
    );
    const targetState = stateWithBlocks(
      tableWithTextGrid([["Same"]], {
        width: 7_200,
        widthType: "dxa",
        justification: "center",
      }),
    );
    const { executed, tracked } = executePlannedComparison(baseState, targetState);

    expect(executed.receipt.changes).toEqual([
      {
        kind: "table-format",
        location: { story: { type: "main" } },
        scope: "table",
        base: { tableIndex: 0 },
        target: { tableIndex: 0 },
        properties: [
          {
            key: "justification",
            base: { type: "present", value: "left" },
            revised: { type: "present", value: "center" },
          },
          {
            key: "width",
            base: {
              type: "present",
              value: {
                type: "object",
                entries: [
                  { key: "type", value: "dxa" },
                  { key: "value", value: 6_000 },
                ],
              },
            },
            revised: {
              type: "present",
              value: {
                type: "object",
                entries: [
                  { key: "type", value: "dxa" },
                  { key: "value", value: 7_200 },
                ],
              },
            },
          },
        ],
      },
    ]);
    expect(Object.hasOwn(executed.receipt.changes[0] ?? {}, "owner")).toBe(false);
    expect(executed.receipt.instructions.map(({ instructionType }) => instructionType)).toEqual([
      "matchTableFormatting",
    ]);
    expectResolvedViews(baseState, targetState, tracked);
  });

  test("executes independent table-format components in one deterministic schedule", () => {
    const formattedTable = (id: string, text: string, width: number) => {
      const content = tableWithCellGrid([[{ id, text }]]);
      return content.type.create({ ...content.attrs, width, widthType: "dxa" }, content.content);
    };
    const baseState = stateWithBlocks(
      formattedTable("F1000000", "First", 4_000),
      paragraphNode("A1000000", "Between"),
      formattedTable("E1000000", "Second", 5_000),
    );
    const targetState = stateWithBlocks(
      formattedTable("F1000000", "First", 4_400),
      paragraphNode("A1000000", "Between"),
      formattedTable("E1000000", "Second", 5_500),
    );
    const { executed, tracked } = executePlannedComparison(baseState, targetState);

    expect(executed.receipt.instructions.map(({ instructionType }) => instructionType)).toEqual([
      "matchTableFormatting",
      "matchTableFormatting",
    ]);
    expect(executed.receipt.executionTaskCount).toBe(2);
    expectResolvedViews(baseState, targetState, tracked);
    const rerun = executePlannedComparison(baseState, targetState);
    expect(rerun.tracked.doc.eq(tracked.doc)).toBe(true);
    expect(rerun.executed.receipt.instructions).toEqual(executed.receipt.instructions);
  });

  test("enforces one global structure-preflight budget across canonical components", () => {
    const formattedTable = (id: string, text: string, width: number) => {
      const content = tableWithCellGrid([[{ id, text }]]);
      return content.type.create({ ...content.attrs, width, widthType: "dxa" }, content.content);
    };
    const baseState = stateWithBlocks(
      formattedTable("F1000000", "First", 4_000),
      paragraphNode("A1000000", "Between"),
      formattedTable("E1000000", "Second", 5_000),
    );
    const targetState = stateWithBlocks(
      formattedTable("F1000000", "First", 4_400),
      paragraphNode("A1000000", "Between"),
      formattedTable("E1000000", "Second", 5_500),
    );
    const planned = plannedComparisonOf({ baseState, targetState });
    const operands = planned.program
      .consume()
      .instructions.flatMap((instruction) =>
        instruction.type === "tableStructure" || instruction.type === "tableFormat"
          ? [instruction.operation]
          : [],
      );
    const components = resolvedDocxTableComponents({
      comparison: planned.comparison,
      operands,
    });

    expect(components).toHaveLength(2);
    expect(
      preflightTableStructureComponents({
        doc: baseState.doc,
        comparison: planned.comparison,
        components,
        limits: { ...DEFAULT_TABLE_STRUCTURE_PREFLIGHT_LIMITS, maxOperands: 1 },
      }),
    ).toEqual({
      status: "unsupported",
      issue: {
        reason: "limit-exceeded",
        limit: "maxOperands",
        maximum: 1,
        actual: 2,
      },
    });
  });

  test("row and column structure edits reconstruct both views through the comparison pipeline", () => {
    const a = { id: "A1000000", text: "A" } as const;
    const b = { id: "B1000000", text: "B" } as const;
    const c = { id: "C1000000", text: "C" } as const;
    const d = { id: "D1000000", text: "D" } as const;
    const x = { id: "E1000000", text: "X" } as const;
    const y = { id: "F1000000", text: "Y" } as const;
    const rowBase = stateWithBlocks(
      tableWithCellGrid([
        [a, b],
        [c, d],
      ]),
    );
    const columnBase = stateWithBlocks(
      tableWithCellGrid([
        [a, b],
        [c, d],
      ]),
    );
    const cases = [
      {
        base: rowBase,
        target: stateWithBlocks(
          tableWithCellGrid([
            [a, b],
            [x, y],
            [c, d],
          ]),
        ),
      },
      {
        base: stateWithBlocks(
          tableWithCellGrid([
            [a, b],
            [x, y],
            [c, d],
          ]),
        ),
        target: rowBase,
      },
      {
        base: columnBase,
        target: stateWithBlocks(
          tableWithCellGrid([
            [a, x, b],
            [c, y, d],
          ]),
        ),
      },
      {
        base: stateWithBlocks(
          tableWithCellGrid([
            [a, x, b],
            [c, y, d],
          ]),
        ),
        target: columnBase,
      },
    ];

    for (const { base, target } of cases) {
      const { tracked } = executePlannedComparison(base, target);
      expectResolvedViews(base, target, tracked);
    }
  });

  test("rejects copied and cross-comparison table operands before preflight", () => {
    const baseState = stateWithBlocks(
      paragraphNode("A1000000", "Before"),
      paragraphNode("B1000000", "After"),
    );
    const targetState = stateWithBlocks(
      paragraphNode("A1000000", "Before"),
      tableWithTextGrid([["New"]]),
      paragraphNode("B1000000", "After"),
    );
    const first = plannedComparisonOf({ baseState, targetState });
    const consumed = first.program.consume();
    const instruction = consumed.instructions.find(
      (candidate) =>
        candidate.type === "tableStructure" && candidate.operation.type === "insertTable",
    );
    if (!instruction || instruction.type !== "tableStructure") {
      throw new Error("Fixture did not produce a table insertion operand.");
    }
    const operand = instruction.operation;
    const copied = Object.freeze({ ...operand });
    expect(() =>
      Reflect.apply(DocxComparisonProgram.create, DocxComparisonProgram, [
        first.comparison,
        [{ type: "tableStructure", operation: copied }],
      ]),
    ).toThrow("was not created by Folio");

    const secondBase = stateWithBlocks(
      paragraphNode("C1000000", "Before"),
      paragraphNode("D1000000", "After"),
    );
    const secondTarget = stateWithBlocks(
      paragraphNode("C1000000", "Before"),
      tableWithTextGrid([["Changed"]]),
      paragraphNode("D1000000", "After"),
    );
    const second = plannedComparisonOf({ baseState: secondBase, targetState: secondTarget });
    expect(() =>
      DocxComparisonProgram.create(second.comparison, [
        { type: "tableStructure", operation: operand },
      ]),
    ).toThrow("belongs to another story comparison");
  });

  test("refuses hidden-row table structure before mutating the transaction", () => {
    const visible = schema.node("tableRow", null, [
      schema.node("tableCell", null, [paragraphNode("A1000000", "Visible")]),
    ]);
    const hidden = schema.node("tableRow", { hidden: true }, [
      schema.node("tableCell", null, [paragraphNode("H1000000", "Hidden")]),
    ]);
    const baseState = stateWithBlocks(schema.node("table", null, [visible, hidden]));
    const targetState = stateWithBlocks(schema.node("table", null, [visible]));
    const prepared = preflightDocxComparisonProgram({
      state: baseState,
      program: plannedComparisonOf({ baseState, targetState }).program,
    });

    expect(prepared.supportedInstructionCount).toBe(0);
    expect(prepared.issues).toEqual([
      {
        instructionIndex: 0,
        instructionType: "matchTableFormatting",
        reason: "unrepresentable-table-structure",
        tableStructure: { reason: "unprojected-table-structure", side: "source" },
      },
    ]);
    expect(baseState.doc.childCount).toBe(1);
    expect(baseState.doc.firstChild?.childCount).toBe(2);
  });

  test("isolates an unsupported table from an independent table-format program", () => {
    const changedBase = tableWithTextGrid([["Independent"]], {
      width: 6_000,
      widthType: "dxa",
      justification: "left",
    });
    const changedTarget = tableWithTextGrid([["Independent"]], {
      width: 7_200,
      widthType: "dxa",
      justification: "center",
    });
    const visible = schema.node("tableRow", null, [
      schema.node("tableCell", null, [paragraphNode("A1000000", "Visible")]),
    ]);
    const hidden = schema.node("tableRow", { hidden: true }, [
      schema.node("tableCell", null, [paragraphNode("H1000000", "Hidden")]),
    ]);
    const baseState = stateWithBlocks(
      changedBase,
      paragraphNode("B1000000", "Between"),
      schema.node("table", null, [visible, hidden]),
    );
    const targetState = stateWithBlocks(
      changedTarget,
      paragraphNode("B1000000", "Between"),
      schema.node("table", null, [visible]),
    );
    const prepared = preflightDocxComparisonProgram({
      state: baseState,
      program: plannedComparisonOf({ baseState, targetState }).program,
    });

    expect(prepared.supportedInstructionCount).toBe(1);
    expect(prepared.issues).toHaveLength(1);
    expect(prepared.issues[0]).toEqual(
      expect.objectContaining({
        instructionType: "matchTableFormatting",
        reason: "unrepresentable-table-structure",
        tableStructure: { reason: "unprojected-table-structure", side: "source" },
      }),
    );
    const executed = executePreflightedDocxComparison({
      state: baseState,
      prepared,
      revisionStamp: { idSeed: 950, date: "2026-09-12T00:00:00.000Z" },
      author: "Comparison",
    });
    if (executed.status !== "executed") throw new Error("Expected partial table execution.");
    const tracked = baseState.apply(executed.receipt.transaction);
    const accepted = resolvedState(tracked, "accept");
    const rejected = resolvedState(tracked, "reject");
    const acceptedFirstTable = accepted.doc.firstChild;
    const rejectedFirstTable = rejected.doc.firstChild;
    expect(acceptedFirstTable?.attrs["width"]).toBe(7_200);
    expect(acceptedFirstTable?.attrs["justification"]).toBe("center");
    expect(rejectedFirstTable?.attrs["width"]).toBe(6_000);
    expect(rejectedFirstTable?.attrs["justification"]).toBe("left");
    expect(
      withoutBlockIdentities(acceptedFirstTable ?? changedBase).eq(
        withoutBlockIdentities(targetState.doc.child(0)),
      ),
    ).toBe(true);
    expect(
      withoutBlockIdentities(rejectedFirstTable ?? changedTarget).eq(
        withoutBlockIdentities(baseState.doc.child(0)),
      ),
    ).toBe(true);
    expect(accepted.doc.child(2).eq(baseState.doc.child(2))).toBe(true);
    expect(rejected.doc.child(2).eq(baseState.doc.child(2))).toBe(true);
    expect(executed.receipt.instructions.map(({ instructionType }) => instructionType)).toEqual([
      "matchTableFormatting",
    ]);
  });

  test("executes a valid table structure component beside an unsupported one", () => {
    const validBase = tableWithCellGrid([
      [{ id: "A1000000", text: "A" }],
      [{ id: "B1000000", text: "B" }],
    ]);
    const validTarget = tableWithCellGrid([
      [{ id: "A1000000", text: "A" }],
      [{ id: "C1000000", text: "Inserted" }],
      [{ id: "B1000000", text: "B" }],
    ]);
    const spanningRow = (id: string, text: string) =>
      schema.node("tableRow", null, [
        schema.node("tableCell", { colspan: 2 }, [paragraphNode(id, text)]),
      ]);
    const invalidBase = schema.node("table", null, [
      spanningRow("D1000000", "D"),
      spanningRow("E1000000", "E"),
    ]);
    const invalidTarget = schema.node("table", null, [
      spanningRow("D1000000", "D"),
      spanningRow("F1000000", "Unsupported"),
      spanningRow("E1000000", "E"),
    ]);
    const baseState = stateWithBlocks(validBase, paragraphNode("A2000000", "Between"), invalidBase);
    const targetState = stateWithBlocks(
      validTarget,
      paragraphNode("A2000000", "Between"),
      invalidTarget,
    );
    const prepared = preflightDocxComparisonProgram({
      state: baseState,
      program: plannedComparisonOf({ baseState, targetState }).program,
    });

    expect(prepared.supportedInstructionCount).toBe(1);
    expect(prepared.issues).toEqual([
      expect.objectContaining({
        instructionType: "insertTableRow",
        reason: "unrepresentable-table-structure",
        tableStructure: { reason: "unrepresentable-span", side: "source" },
      }),
    ]);
    const executed = executePreflightedDocxComparison({
      state: baseState,
      prepared,
      revisionStamp: { idSeed: 975, date: "2026-09-12T00:00:00.000Z" },
      author: "Comparison",
    });
    if (executed.status !== "executed") throw new Error("Expected partial table execution.");
    const tracked = baseState.apply(executed.receipt.transaction);
    const accepted = resolvedState(tracked, "accept");
    const rejected = resolvedState(tracked, "reject");
    expect(
      withoutBlockIdentities(accepted.doc.child(0)).eq(
        withoutBlockIdentities(targetState.doc.child(0)),
      ),
    ).toBe(true);
    expect(
      withoutBlockIdentities(rejected.doc.child(0)).eq(
        withoutBlockIdentities(baseState.doc.child(0)),
      ),
    ).toBe(true);
    expect(accepted.doc.child(2).eq(baseState.doc.child(2))).toBe(true);
    expect(rejected.doc.child(2).eq(baseState.doc.child(2))).toBe(true);
    expect(executed.receipt.instructions.map(({ instructionType }) => instructionType)).toEqual([
      "insertTableRow",
    ]);
  });

  test("fails every obligation in one shared table component atomically", () => {
    const spanningRow = (id: string, text: string) =>
      schema.node("tableRow", null, [
        schema.node("tableCell", { colspan: 2 }, [paragraphNode(id, text)]),
      ]);
    const baseState = stateWithBlocks(
      schema.node("table", { width: 6_000, widthType: "dxa", justification: "left" }, [
        spanningRow("A1000000", "A"),
        spanningRow("B1000000", "B"),
      ]),
    );
    const targetState = stateWithBlocks(
      schema.node("table", { width: 7_200, widthType: "dxa", justification: "center" }, [
        spanningRow("A1000000", "A"),
        spanningRow("C1000000", "Inserted"),
        spanningRow("B1000000", "B"),
      ]),
    );
    const before = baseState.doc.toJSON();
    const prepared = preflightDocxComparisonProgram({
      state: baseState,
      program: plannedComparisonOf({ baseState, targetState }).program,
    });

    expect(prepared.supportedInstructionCount).toBe(0);
    expect(prepared.issues).toHaveLength(2);
    expect(prepared.issues.map(({ instructionType }) => instructionType).toSorted()).toEqual([
      "insertTableRow",
      "matchTableFormatting",
    ]);
    expect(
      prepared.issues.every(
        ({ tableStructure }) => tableStructure?.reason === "unrepresentable-span",
      ),
    ).toBe(true);
    expect(baseState.doc.toJSON()).toEqual(before);
  });

  test("generated rectangular table edits remain reversible without global mapping", () => {
    for (let height = 2; height <= 4; height++) {
      for (let width = 2; width <= 4; width++) {
        const grid = Array.from({ length: height }, (_, rowIndex) =>
          Array.from({ length: width }, (_unused, columnIndex) => ({
            id: `${String(rowIndex)}${String(columnIndex)}ABCDEF`,
            text: `${String(rowIndex)}:${String(columnIndex)}`,
          })),
        );
        const inserted = Array.from({ length: width }, (_unused, columnIndex) => ({
          id: `I${String(height)}${String(width)}${String(columnIndex)}0000`,
          text: `new:${String(columnIndex)}`,
        }));
        const targetGrid = grid.map((row) => [...row]);
        targetGrid.splice(1, 0, inserted);
        const baseState = stateWithBlocks(tableWithCellGrid(grid));
        const targetState = stateWithBlocks(tableWithCellGrid(targetGrid));
        const { executed, tracked } = executePlannedComparison(baseState, targetState);

        expect(executed.receipt.localPositionMappingSteps).toBe(0);
        expectResolvedViews(baseState, targetState, tracked);
      }
    }
  });

  test("does not compose lookalike table changes across an unrelated stream event", () => {
    const splitTable = (leftId: string, rightId: string, label: string) =>
      schema.node("table", null, [
        schema.node("tableRow", null, [
          schema.node("tableCell", null, [paragraphNode(leftId, `${label} left`)]),
          schema.node("tableCell", null, [paragraphNode(rightId, `${label} right`)]),
        ]),
      ]);
    const mergedTable = (id: string, label: string) =>
      schema.node("table", null, [
        schema.node("tableRow", null, [
          schema.node("tableCell", { colspan: 2 }, [paragraphNode(id, label)]),
        ]),
      ]);
    const prefix = paragraphNode("E0000000", "Before tables");
    const separator = paragraphNode("E1000000", "Between tables");
    const suffix = paragraphNode("F1000000", "After tables");
    const baseState = stateWithBlocks(
      prefix,
      splitTable("A1000000", "A1000001", "First"),
      separator,
      suffix,
    );
    const targetState = stateWithBlocks(
      prefix,
      separator,
      mergedTable("C1000000", "First"),
      suffix,
    );
    const baseSnapshot = resolvedSnapshotOf(baseState);
    const comparison = comparisonOf(baseSnapshot, resolvedSnapshotOf(targetState));
    const structural = resolvedDocxStoryComparisonPayload(comparison).comparison.events.flatMap(
      (event) => (event.type === "structural" && event.memberIndex === 0 ? [event.change] : []),
    );
    const deleted = structural.filter((change) => change.type === "table-delete");
    const inserted = structural.filter((change) => change.type === "table-insert");
    const firstDeleted = deleted.at(0);
    const separatedInserted = inserted.at(0);
    if (firstDeleted?.type !== "table-delete" || separatedInserted?.type !== "table-insert") {
      throw new Error(
        `Fixture did not produce independent table changes: ${JSON.stringify(
          resolvedDocxStoryComparisonPayload(comparison).comparison.events.map((event) =>
            event.type === "structural" ? event.change.type : event.type,
          ),
        )}`,
      );
    }
    const sourceBlock =
      firstDeleted.blocks.at(0) ??
      (() => {
        throw new Error("Fixture table deletion has no source block.");
      })();

    expect(() =>
      resolvedDocxTableStructureOperand(comparison, {
        type: "replaceTable",
        source: resolvedDocxSourceOperand(baseSnapshot, sourceBlock),
        owner: { type: "structural-pair", deleted: firstDeleted, inserted: separatedInserted },
      }),
    ).toThrow(/container ownership|canonical event sequence/u);
  });
});
