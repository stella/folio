import { describe, expect, test } from "bun:test";
import { EditorState } from "prosemirror-state";

import { createContentComparisonWorkSession } from "../../compare/content";
import { updateDocumentContent } from "../../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../../prosemirror/conversion/toProseDoc";
import { schema } from "../../prosemirror/schema";
import { createEmptyDocument } from "../../utils/createDocument";

import { DocxComparisonProgram, type DocxComparisonOperationInput } from "./docx-program";
import {
  createResolvedDocxStorySnapshot,
  type ResolvedDocxStorySnapshot,
} from "./resolved-docx-story-snapshot";
import {
  compareResolvedDocxStoryPair,
  createResolvedDocxStoryPair,
  resolvedDocxPairedEventOperand,
  resolvedDocxStoryComparisonPayload,
  type ResolvedDocxStoryComparison,
} from "./resolved-docx-story-comparison";

const sourceFixture = (): {
  readonly snapshot: ResolvedDocxStorySnapshot;
  readonly targetSnapshot: ResolvedDocxStorySnapshot;
  readonly comparison: ResolvedDocxStoryComparison;
} => {
  const bold = schema.marks["bold"];
  if (!bold) throw new Error("schema has no bold mark");
  const state = EditorState.create({
    doc: schema.node("doc", null, [
      schema.node("paragraph", { paraId: "p-1" }, [schema.text("old shared", [bold.create()])]),
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
      schema.node("paragraph", { paraId: "p-1" }, [schema.text("new shared")]),
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
  return {
    snapshot,
    targetSnapshot,
    comparison: compared.value,
  };
};

const replacement = (
  comparison: ResolvedDocxStoryComparison,
): Extract<DocxComparisonOperationInput, { readonly type: "pairedBlock" }> => ({
  type: "pairedBlock",
  event: (() => {
    const event = resolvedDocxStoryComparisonPayload(comparison).comparison.events.at(0);
    if (event?.type !== "modified") throw new Error("replacement fixture was not modified");
    return resolvedDocxPairedEventOperand(comparison, event);
  })(),
});

describe("DocxComparisonProgram", () => {
  test("derives and deeply freezes the sole instruction payload", () => {
    const { comparison, snapshot, targetSnapshot } = sourceFixture();
    const input = replacement(comparison);
    const program = DocxComparisonProgram.create(comparison, [input]);

    Reflect.set(input, "event", Object.freeze({ ...input.event }));

    const consumed = program.consume();
    const [instruction] = consumed.instructions;
    expect(consumed.sourceSnapshot).toBe(snapshot);
    expect(consumed.targetSnapshot).toBe(targetSnapshot);
    expect(
      consumed.semanticGroups.flatMap(({ reports }) => reports.map(({ change }) => change.kind)),
    ).toEqual(["replace", "format"]);
    expect(instruction?.type).toBe("replaceText");
    if (instruction?.type !== "replaceText") throw new Error("expected replacement");
    expect(instruction.range.sourceText).toBe("old shared");
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
    const program = DocxComparisonProgram.create(comparison, [replacement(comparison)]);
    expect(program.consume().instructions).toHaveLength(1);
    expect(() => program.consume()).toThrow("consumed more than once");
  });

  test("rejects duplicate canonical report sequence ownership", () => {
    const { comparison } = sourceFixture();
    expect(() =>
      DocxComparisonProgram.create(comparison, [replacement(comparison), replacement(comparison)]),
    ).toThrow("invalid canonical sequence");
  });

  test("rejects copied and cross-comparison event operands", () => {
    const left = sourceFixture();
    const right = sourceFixture();
    const copied = Object.freeze({ ...replacement(left.comparison).event });
    expect(() =>
      Reflect.apply(DocxComparisonProgram.create, DocxComparisonProgram, [
        left.comparison,
        [{ type: "pairedBlock", event: copied }],
      ]),
    ).toThrow("was not created by Folio");
    expect(() =>
      DocxComparisonProgram.create(right.comparison, [replacement(left.comparison)]),
    ).toThrow("belongs to another story comparison");
  });

  test("rejects an oversized instruction graph before compiling it", () => {
    const { comparison } = sourceFixture();
    expect(() =>
      DocxComparisonProgram.create(
        comparison,
        Array.from({ length: 10_001 }, () => replacement(comparison)),
      ),
    ).toThrow("exceeds its instruction limit");
  });

  test("cannot mix an event operand with another story comparison", () => {
    const left = sourceFixture();
    const right = sourceFixture();
    expect(() =>
      DocxComparisonProgram.create(right.comparison, [replacement(left.comparison)]),
    ).toThrow("belongs to another story comparison");
  });

  test("compiled fragments reconstruct both canonical text views", () => {
    const { comparison } = sourceFixture();
    const [instruction] = DocxComparisonProgram.create(comparison, [
      replacement(comparison),
    ]).consume().instructions;
    if (instruction?.type !== "replaceText") throw new Error("expected replacement");
    expect(
      instruction.range.fragments
        .filter(({ type }) => type !== "ins")
        .map(({ text }) => text)
        .join(""),
    ).toBe("old shared");
    expect(
      instruction.range.fragments
        .filter(({ type }) => type !== "del")
        .map(({ text }) => text)
        .join(""),
    ).toBe("new shared");
  });
});
