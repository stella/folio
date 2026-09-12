import { describe, expect, test } from "bun:test";
import { EditorState } from "prosemirror-state";

import { buildCleanBlockText } from "../../ai-edits/clean-text";
import { createFolioAIEditSnapshot } from "../../ai-edits/snapshot";
import { updateDocumentContent } from "../../prosemirror/conversion/fromProseDoc";
import { schema } from "../../prosemirror/schema";
import { createEmptyDocument } from "../../utils/createDocument";
import {
  executePreflightedDocxComparison,
  preflightDocxComparisonProgram,
} from "./docx-executor";
import {
  DocxComparisonProgram,
  type DocxComparisonInstructionInput,
} from "./docx-program";
import {
  createResolvedDocxStorySnapshot,
  resolvedDocxContentBlocks,
  resolvedDocxSourceOperand,
  type ResolvedDocxStorySnapshot,
} from "./resolved-docx-story-snapshot";

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

const replacement = ({
  snapshot,
  blockIndex = 0,
  targetText,
}: {
  readonly snapshot: ResolvedDocxStorySnapshot;
  readonly blockIndex?: number;
  readonly targetText: string;
}): Extract<DocxComparisonInstructionInput, { readonly type: "replaceText" }> => {
  const block = resolvedDocxContentBlocks(snapshot).at(blockIndex);
  if (!block) throw new Error("fixture source block missing");
  const sourceText = block.text;
  return {
    type: "replaceText",
    source: resolvedDocxSourceOperand(snapshot, block),
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
  };
};

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
});
