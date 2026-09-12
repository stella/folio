import { describe, expect, test } from "bun:test";
import { EditorState } from "prosemirror-state";

import { createContentComparisonWorkSession } from "../../compare/content";
import { updateDocumentContent } from "../../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../../prosemirror/conversion/toProseDoc";
import { schema } from "../../prosemirror/schema";
import { createEmptyDocument } from "../../utils/createDocument";

import { DocxComparisonProgram, type DocxComparisonInstructionInput } from "./docx-program";
import {
  createResolvedDocxStorySnapshot,
  resolvedDocxContentBlocks,
  resolvedDocxSourceOperand,
  type ResolvedDocxSourceOperand,
  type ResolvedDocxStorySnapshot,
} from "./resolved-docx-story-snapshot";
import {
  compareResolvedDocxStoryPair,
  createResolvedDocxStoryPair,
  resolvedDocxPairRangeOperand,
  resolvedDocxStoryComparisonPayload,
  resolvedDocxTargetBlockOperand,
  type ResolvedDocxStoryComparison,
} from "./resolved-docx-story-comparison";

const sourceFixture = (): {
  readonly snapshot: ResolvedDocxStorySnapshot;
  readonly targetSnapshot: ResolvedDocxStorySnapshot;
  readonly source: ResolvedDocxSourceOperand;
  readonly comparison: ResolvedDocxStoryComparison;
} => {
  const bold = schema.marks["bold"];
  if (!bold) throw new Error("schema has no bold mark");
  const state = EditorState.create({
    doc: schema.node("doc", null, [
      schema.node("paragraph", { paraId: "p-1" }, [schema.text("old", [bold.create()])]),
    ]),
  });
  const document = updateDocumentContent(createEmptyDocument(), state.doc);
  const snapshot = createResolvedDocxStorySnapshot({
    document,
    story: { type: "main" },
    sourceDocument: toProseDoc(document),
  });
  if (!snapshot) throw new Error("main story projection missing");
  const targetState = EditorState.create({
    doc: schema.node("doc", null, [
      schema.node("paragraph", { paraId: "p-1" }, [schema.text("new", [bold.create()])]),
    ]),
  });
  const targetDocument = updateDocumentContent(createEmptyDocument(), targetState.doc);
  const targetSnapshot = createResolvedDocxStorySnapshot({
    document: targetDocument,
    story: { type: "main" },
    sourceDocument: toProseDoc(targetDocument),
  });
  if (!targetSnapshot) throw new Error("target story projection missing");
  const compared = compareResolvedDocxStoryPair({
    pair: createResolvedDocxStoryPair({ baseSnapshot: snapshot, targetSnapshot }),
    workSession: createContentComparisonWorkSession(),
  });
  if (compared.isErr()) throw compared.error;
  const block = resolvedDocxContentBlocks(snapshot).at(0);
  if (!block) throw new Error("source block missing");
  return {
    snapshot,
    targetSnapshot,
    source: resolvedDocxSourceOperand(snapshot, block),
    comparison: compared.value,
  };
};

const replacement = (
  comparison: ResolvedDocxStoryComparison,
): Extract<DocxComparisonInstructionInput, { readonly type: "replaceText" }> => ({
  type: "replaceText",
  range: (() => {
    const event = resolvedDocxStoryComparisonPayload(comparison).comparison.events.at(0);
    if (event?.type !== "modified") throw new Error("replacement fixture was not modified");
    return resolvedDocxPairRangeOperand(comparison, event.relation);
  })(),
});

const grouped = (...instructions: readonly DocxComparisonInstructionInput[]) =>
  instructions.map((instruction) => ({
    group: { reports: [] },
    instruction,
  }));

describe("DocxComparisonProgram", () => {
  test("derives and deeply freezes the sole instruction payload", () => {
    const { comparison, snapshot, targetSnapshot } = sourceFixture();
    const input = replacement(comparison);
    const program = DocxComparisonProgram.create(comparison, grouped(input));

    Reflect.set(input, "range", Object.freeze({ ...input.range }));

    const consumed = program.consume();
    const [instruction] = consumed.instructions;
    expect(consumed.sourceSnapshot).toBe(snapshot);
    expect(consumed.targetSnapshot).toBe(targetSnapshot);
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
    const { comparison } = sourceFixture();
    const program = DocxComparisonProgram.create(comparison, grouped(replacement(comparison)));
    expect(program.consume().instructions).toHaveLength(1);
    expect(() => program.consume()).toThrow("consumed more than once");
  });

  test("rejects duplicate canonical report sequence ownership", () => {
    const { comparison } = sourceFixture();
    const report = {
      sequence: 0,
      change: {
        kind: "replace",
        location: { story: { type: "main" } },
        baseBlockId: "p-1",
        targetBlockId: "p-1",
        before: "old",
        after: "new",
      },
    } as const;

    expect(() =>
      DocxComparisonProgram.create(comparison, [
        { group: { reports: [report] }, instruction: replacement(comparison) },
        { group: { reports: [report] }, instruction: replacement(comparison) },
      ]),
    ).toThrow("invalid canonical sequence");
  });

  test("rejects copied and cross-comparison range operands", () => {
    const left = sourceFixture();
    const right = sourceFixture();
    const copied = Object.freeze({ ...replacement(left.comparison).range });
    expect(() =>
      Reflect.apply(DocxComparisonProgram.create, DocxComparisonProgram, [
        left.comparison,
        grouped({ type: "replaceText", range: copied }),
      ]),
    ).toThrow("was not created by Folio");
    expect(() =>
      DocxComparisonProgram.create(right.comparison, grouped(replacement(left.comparison))),
    ).toThrow("belongs to another story comparison");
  });

  test("rejects an oversized instruction graph before compiling it", () => {
    const { comparison } = sourceFixture();
    expect(() =>
      DocxComparisonProgram.create(
        comparison,
        grouped(...Array.from({ length: 10_001 }, () => replacement(comparison))),
      ),
    ).toThrow("exceeds its instruction limit");
  });

  test("cannot mix a source or target operand with another story comparison", () => {
    const left = sourceFixture();
    const right = sourceFixture();
    expect(() =>
      DocxComparisonProgram.create(
        right.comparison,
        grouped({ type: "deleteParagraph", source: left.source }),
      ),
    ).toThrow("cannot mix source story snapshots");
    const targetBlock = resolvedDocxStoryComparisonPayload(left.comparison)
      .comparison.events.flatMap((event) =>
        event.type === "modified" ? [event.relation.revised.block] : [],
      )
      .at(0);
    if (!targetBlock) throw new Error("target block missing");
    expect(() =>
      DocxComparisonProgram.create(
        right.comparison,
        grouped({
          type: "insertParagraph",
          boundary: { type: "afterParagraph", paragraph: right.source },
          target: resolvedDocxTargetBlockOperand(left.comparison, targetBlock),
        }),
      ),
    ).toThrow("belongs to another story comparison");
  });

  test("compiled fragments reconstruct both canonical text views", () => {
    const { comparison } = sourceFixture();
    const [instruction] = DocxComparisonProgram.create(
      comparison,
      grouped(replacement(comparison)),
    ).consume().instructions;
    if (instruction?.type !== "replaceText") throw new Error("expected replacement");
    expect(
      instruction.range.fragments
        .filter(({ type }) => type !== "ins")
        .map(({ text }) => text)
        .join(""),
    ).toBe("old");
    expect(
      instruction.range.fragments
        .filter(({ type }) => type !== "del")
        .map(({ text }) => text)
        .join(""),
    ).toBe("new");
  });
});
