import { describe, expect, test } from "bun:test";
import { EditorState } from "prosemirror-state";

import { createFolioAIEditSnapshot } from "../../ai-edits/snapshot";
import { updateDocumentContent } from "../../prosemirror/conversion/fromProseDoc";
import { schema } from "../../prosemirror/schema";
import { createEmptyDocument } from "../../utils/createDocument";

import {
  DocxComparisonProgram,
  type DocxComparisonInstructionInput,
} from "./docx-program";
import {
  createResolvedDocxStorySnapshot,
  resolvedDocxContentBlocks,
  resolvedDocxSourceOperand,
  type ResolvedDocxSourceOperand,
  type ResolvedDocxStorySnapshot,
} from "./resolved-docx-story-snapshot";

const sourceFixture = (): {
  readonly snapshot: ResolvedDocxStorySnapshot;
  readonly source: ResolvedDocxSourceOperand;
} => {
  const bold = schema.marks["bold"];
  if (!bold) throw new Error("schema has no bold mark");
  const state = EditorState.create({
    doc: schema.node("doc", null, [
      schema.node("paragraph", { paraId: "p-1" }, [schema.text("old", [bold.create()])]),
    ]),
  });
  const snapshot = createResolvedDocxStorySnapshot({
    document: updateDocumentContent(createEmptyDocument(), state.doc),
    story: { type: "main" },
    operationSnapshot: createFolioAIEditSnapshot(state.doc),
  });
  if (!snapshot) throw new Error("main story projection missing");
  const block = resolvedDocxContentBlocks(snapshot).at(0);
  if (!block) throw new Error("source block missing");
  return { snapshot, source: resolvedDocxSourceOperand(snapshot, block) };
};

const replacement = (source: ResolvedDocxSourceOperand): Extract<
  DocxComparisonInstructionInput,
  { readonly type: "replaceText" }
> => ({
  type: "replaceText",
  source,
  sourceStartOffset: 0,
  range: {
    sourceText: "old",
    targetText: "new",
    segments: [
      {
        type: "del",
        text: "old",
        baseStart: 0,
        baseEnd: 3,
        revisedStart: 0,
        revisedEnd: 0,
      },
      {
        type: "ins",
        text: "new",
        baseStart: 3,
        baseEnd: 3,
        revisedStart: 0,
        revisedEnd: 3,
      },
    ],
    sourceRuns: [{ startOffset: 0, endOffset: 3, formatting: { bold: true } }],
    targetRuns: [{ startOffset: 0, endOffset: 3, formatting: { italic: true } }],
    authoredChanges: [],
  },
});

describe("DocxComparisonProgram", () => {
  test("captures and deeply freezes the sole instruction payload", () => {
    const { snapshot, source } = sourceFixture();
    const input = replacement(source);
    const program = DocxComparisonProgram.create(snapshot, [input]);

    Reflect.set(input.range.segments[0]!, "text", "corrupted");
    Reflect.set(input.range.sourceRuns[0]!.formatting, "bold", false);

    const [instruction] = program.consume(snapshot);
    expect(instruction?.type).toBe("replaceText");
    if (instruction?.type !== "replaceText") throw new Error("expected replacement");
    expect(instruction.range.sourceText).toBe("old");
    expect(instruction.range.fragments[0]?.text).toBe("old");
    expect(
      instruction.range.fragments.find(({ type }) => type === "del")?.sourceFormatting.bold,
    ).toBe(true);
    expect(Object.isFrozen(instruction)).toBe(true);
    expect(Object.isFrozen(instruction.range.fragments)).toBe(true);
    expect(Object.isFrozen(instruction.range.fragments[0])).toBe(true);
  });

  test("is consumable exactly once", () => {
    const { snapshot, source } = sourceFixture();
    const program = DocxComparisonProgram.create(snapshot, [replacement(source)]);
    expect(program.consume(snapshot)).toHaveLength(1);
    expect(() => program.consume(snapshot)).toThrow("consumed more than once");
  });

  test("rejects a range whose canonical text does not name its source", () => {
    const { snapshot, source } = sourceFixture();
    const input = replacement(source);
    Reflect.set(input.range, "sourceText", "elsewhere");
    expect(() => DocxComparisonProgram.create(snapshot, [input])).toThrow(
      "does not name its exact source and target text",
    );
  });

  test("rejects an authored formatting delta missing from canonical changes", () => {
    const { snapshot, source } = sourceFixture();
    const input = replacement(source);
    Reflect.set(input.range, "segments", [
      {
        type: "equal",
        text: "old",
        baseStart: 0,
        baseEnd: 3,
        revisedStart: 0,
        revisedEnd: 3,
      },
    ]);
    Reflect.set(input.range, "targetText", "old");
    expect(() => DocxComparisonProgram.create(snapshot, [input])).toThrow(
      "Canonical authored-formatting changes do not reconstruct the target run",
    );
  });

  test("rejects oversized instruction and geometry arrays before copying them", () => {
    const { snapshot, source } = sourceFixture();
    expect(() =>
      DocxComparisonProgram.create(
        snapshot,
        Array.from({ length: 10_001 }, () => replacement(source)),
      ),
    ).toThrow("exceeds its instruction limit");

    const pairing = {
      base: { tableIndex: 0, rowIndex: 0, cellIndex: 0 },
      target: { tableIndex: 0, rowIndex: 0, cellIndex: 0 },
    };
    expect(() =>
      DocxComparisonProgram.create(snapshot, [
        {
          type: "matchTableGeometry",
          pairings: Array.from({ length: 10_001 }, () => pairing),
        },
      ]),
    ).toThrow("exceeds its pairing limit");
  });

  test("cannot mix a source operand with another story snapshot", () => {
    const left = sourceFixture();
    const right = sourceFixture();
    expect(() =>
      DocxComparisonProgram.create(right.snapshot, [replacement(left.source)]),
    ).toThrow("cannot mix source story snapshots");
  });
});
