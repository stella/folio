import { describe, expect, test } from "bun:test";
import { EditorState } from "prosemirror-state";

import { createContentComparisonWorkSession } from "../../compare/content";
import { planStoryCompare, type PlanStoryCompareOptions } from "../../compare/plan";
import { updateDocumentContent } from "../../prosemirror/conversion/fromProseDoc";
import { headerFooterToProseDoc, toProseDoc } from "../../prosemirror/conversion/toProseDoc";
import { schema } from "../../prosemirror/schema";
import { createEmptyDocument } from "../../utils/createDocument";
import {
  compareResolvedDocxStoryPair,
  createResolvedDocxStoryPair,
  resolvedDocxDeletedEventOperand,
  resolvedDocxParagraphMarkDispositions,
  resolvedDocxPairedEventOperand,
  resolvedDocxStoryComparisonPayload,
  resolvedDocxTerminalTransitionForEvent,
  resolvedDocxTerminalTransitionOperandPayload,
} from "./resolved-docx-story-comparison";
import {
  createResolvedDocxStorySnapshot,
  type ResolvedDocxStorySnapshot,
} from "./resolved-docx-story-snapshot";

const PLAN_OPTION_KEYS = Object.freeze({
  comparison: true,
  maxOperations: true,
} as const satisfies Record<keyof PlanStoryCompareOptions, true>);

const stateOf = (text: string): EditorState =>
  EditorState.create({
    doc: schema.node("doc", null, [
      schema.node("paragraph", { paraId: "A1000000" }, [schema.text(text)]),
    ]),
  });

const stateWithParagraphs = (
  paragraphs: readonly { readonly id: string; readonly text: string }[],
): EditorState =>
  EditorState.create({
    doc: schema.node(
      "doc",
      null,
      paragraphs.map(({ id, text }) =>
        schema.node("paragraph", { paraId: id }, text.length === 0 ? null : [schema.text(text)]),
      ),
    ),
  });

const snapshotOf = (
  text: string,
  story: { readonly type: "main" } | { readonly type: "header"; readonly relationshipId: string },
): ResolvedDocxStorySnapshot => {
  const state = stateOf(text);
  const mainDocument = updateDocumentContent(createEmptyDocument(), state.doc);
  const document =
    story.type === "main"
      ? mainDocument
      : {
          ...mainDocument,
          package: {
            ...mainDocument.package,
            headers: new Map([
              [story.relationshipId, { content: mainDocument.package.document.content }],
            ]),
          },
        };
  const conversionOptions = {
    ...(document.package.styles !== undefined && { styles: document.package.styles }),
    ...(document.package.theme !== undefined && { theme: document.package.theme }),
  };
  const snapshot = createResolvedDocxStorySnapshot({
    document,
    story,
    sourceDocument:
      story.type === "main"
        ? toProseDoc(document, conversionOptions)
        : headerFooterToProseDoc(mainDocument.package.document.content, conversionOptions),
  });
  if (!snapshot) throw new Error("fixture story projection missing");
  return snapshot;
};

const snapshotWithParagraphs = (
  paragraphs: readonly { readonly id: string; readonly text: string }[],
): ResolvedDocxStorySnapshot => {
  const state = stateWithParagraphs(paragraphs);
  const document = updateDocumentContent(createEmptyDocument(), state.doc);
  const snapshot = createResolvedDocxStorySnapshot({
    document,
    story: { type: "main" },
    sourceDocument: toProseDoc(document),
  });
  if (!snapshot) throw new Error("fixture story projection missing");
  return snapshot;
};

const comparisonOf = (
  baseSnapshot: ResolvedDocxStorySnapshot,
  targetSnapshot: ResolvedDocxStorySnapshot,
) => {
  const compared = compareResolvedDocxStoryPair({
    pair: createResolvedDocxStoryPair({ baseSnapshot, targetSnapshot }),
    workSession: createContentComparisonWorkSession(),
  });
  if (compared.isErr()) throw compared.error;
  return compared.value;
};

describe("resolved DOCX story comparison provenance", () => {
  test("pairs compatible story kinds without equating package-local handles", () => {
    const comparison = comparisonOf(
      snapshotOf("base", { type: "header", relationshipId: "rId1" }),
      snapshotOf("target", { type: "header", relationshipId: "rId99" }),
    );

    expect(resolvedDocxStoryComparisonPayload(comparison)).toMatchObject({
      baseStory: { type: "header", relationshipId: "rId1" },
      targetStory: { type: "header", relationshipId: "rId99" },
    });
  });

  test("rejects incompatible story kinds before comparison", () => {
    const baseSnapshot = snapshotOf("base", { type: "main" });
    const targetSnapshot = snapshotOf("target", {
      type: "header",
      relationshipId: "rId1",
    });

    expect(() => createResolvedDocxStoryPair({ baseSnapshot, targetSnapshot })).toThrow(
      "compatible story kinds",
    );
  });

  test("rejects a structurally copied pair through the runtime boundary", () => {
    const pair = createResolvedDocxStoryPair({
      baseSnapshot: snapshotOf("base", { type: "main" }),
      targetSnapshot: snapshotOf("target", { type: "main" }),
    });
    const counterfeit = Object.freeze({ ...pair });

    expect(() =>
      Reflect.apply(compareResolvedDocxStoryPair, undefined, [
        { pair: counterfeit, workSession: createContentComparisonWorkSession() },
      ]),
    ).toThrow("was not created by Folio");
  });

  test("the planner rejects copied comparison tokens and has no snapshot join", () => {
    expect(Object.keys(PLAN_OPTION_KEYS)).toEqual(["comparison", "maxOperations"]);
    const comparison = comparisonOf(
      snapshotOf("base", { type: "main" }),
      snapshotOf("target", { type: "main" }),
    );
    const counterfeit = Object.freeze({ ...comparison });

    expect(() =>
      Reflect.apply(planStoryCompare, undefined, [{ comparison: counterfeit, maxOperations: 100 }]),
    ).toThrow("was not created by Folio");
  });

  test("event operands require the exact canonical event and its matching branch", () => {
    const comparison = comparisonOf(
      snapshotOf("base", { type: "main" }),
      snapshotOf("target", { type: "main" }),
    );
    const event = resolvedDocxStoryComparisonPayload(comparison).comparison.events.at(0);
    if (event?.type !== "modified") throw new Error("fixture comparison was not modified");

    expect(() =>
      Reflect.apply(resolvedDocxPairedEventOperand, undefined, [
        comparison,
        Object.freeze({ ...event }),
      ]),
    ).toThrow("exact canonical event");
    expect(() =>
      Reflect.apply(resolvedDocxDeletedEventOperand, undefined, [comparison, event]),
    ).toThrow("must name a deleted event");
  });

  test("terminal transitions are issued once from canonical base order", () => {
    const comparison = comparisonOf(
      snapshotWithParagraphs([
        { id: "A1000000", text: "Retained opening." },
        { id: "B1000000", text: "First removed clause." },
        { id: "C1000000", text: "Second removed clause." },
      ]),
      snapshotWithParagraphs([{ id: "D1000000", text: "Retained opening." }]),
    );
    const deleted = resolvedDocxStoryComparisonPayload(comparison).comparison.events.filter(
      (event) => event.type === "deleted",
    );
    const firstEvent = deleted.at(0);
    const secondEvent = deleted.at(1);
    if (!firstEvent || !secondEvent || deleted.length !== 2) {
      throw new Error("fixture did not produce two deletions");
    }
    const first = resolvedDocxTerminalTransitionForEvent(comparison, firstEvent);
    const second = resolvedDocxTerminalTransitionForEvent(comparison, secondEvent);
    if (!first || !second) throw new Error("terminal deletions lost their transition");
    expect(first.operation).toBe(second.operation);
    expect(Number(first.isOwner) + Number(second.isOwner)).toBe(1);
    const payload = resolvedDocxTerminalTransitionOperandPayload(first.operation, comparison);
    if (payload.type !== "paragraph") throw new Error("expected a paragraph transition");
    expect(
      payload.sourceMembers.map((member) =>
        member.type === "deleted"
          ? member.occurrence.event.block.identity.id
          : member.occurrence.event.move.relation.base.block.identity.id,
      ),
    ).toEqual(["B1000000", "C1000000"]);
    expect(Object.isFrozen(payload.sourceMembers)).toBe(true);
    const dispositions = resolvedDocxParagraphMarkDispositions(comparison);
    expect(
      dispositions.map((disposition) =>
        disposition.type === "terminal-member"
          ? {
              blockId: disposition.blockId,
              type: disposition.type,
              operation: disposition.operation,
              terminalRole: disposition.terminalRole,
              effect: disposition.effect,
            }
          : disposition,
      ),
    ).toEqual([
      {
        blockId: "A1000000",
        type: "terminal-member",
        operation: first.operation,
        terminalRole: "chain-start",
        effect: { type: "revision", sourceMemberIndex: 0, kind: "del" },
      },
      {
        blockId: "B1000000",
        type: "terminal-member",
        operation: first.operation,
        terminalRole: "source-member",
        effect: { type: "revision", sourceMemberIndex: 1, kind: "del" },
      },
      {
        blockId: "C1000000",
        type: "terminal-member",
        operation: first.operation,
        terminalRole: "terminal-carrier",
        effect: { type: "preserved-final" },
      },
    ]);
    const counterfeit = Object.freeze({ ...first.operation });
    expect(() =>
      Reflect.apply(resolvedDocxTerminalTransitionOperandPayload, undefined, [
        counterfeit,
        comparison,
      ]),
    ).toThrow("was not created by Folio");
  });

  test("nonterminal lookalikes cannot be enrolled into a terminal program", () => {
    const comparison = comparisonOf(
      snapshotWithParagraphs([
        {
          id: "A1000000",
          text: "The Supplier shall deliver the Goods within thirty days of the order.",
        },
        { id: "B1000000", text: "An unrelated closing paragraph." },
      ]),
      snapshotWithParagraphs([
        { id: "C1000000", text: "An unrelated closing paragraph." },
        {
          id: "D1000000",
          text: "The Supplier shall not be liable for any indirect loss however it arises.",
        },
      ]),
    );
    const events = resolvedDocxStoryComparisonPayload(comparison).comparison.events;
    const deleted = events.find((event) => event.type === "deleted");
    const inserted = events.find((event) => event.type === "inserted");
    if (!deleted || !inserted) {
      throw new Error(
        `fixture did not produce the terminal delete/insert events: ${events.map(({ type }) => type).join(",")}`,
      );
    }
    expect(resolvedDocxTerminalTransitionForEvent(comparison, deleted)).toBeNull();
    expect(resolvedDocxTerminalTransitionForEvent(comparison, inserted)).toBeNull();
    const dispositions = resolvedDocxParagraphMarkDispositions(comparison);
    expect(dispositions.map(({ blockId }) => blockId).toSorted()).toEqual(["A1000000", "B1000000"]);
    expect(new Set(dispositions.map(({ blockId }) => blockId)).size).toBe(dispositions.length);
    expect(Object.isFrozen(dispositions)).toBe(true);
    expect(Reflect.set(dispositions, "0", { blockId: "forged", type: "terminal-member" })).toBe(
      false,
    );
  });
});
