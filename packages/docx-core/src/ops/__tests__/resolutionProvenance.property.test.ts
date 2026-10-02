import { expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import { panic } from "better-result";

import { assertExactModel } from "../../../../../test/exactModel";
import { assertProperty, propertyTestTimeout } from "../../../../../test/property-testing";
import type { Document, ParagraphContent, Run } from "../../model/document";
import { applyDocumentOp, applyDocumentOps, type AppliedDocumentOp } from "../apply";
import { normalizeForOps } from "../contract";
import { asParagraphContent, childNodes, type InlineNode, rebuildNode } from "../leaves";
import { allocateEditorIntentIds, compileEditorIntent, type EditorIntent } from "../editorIntent";
import { childrenOf, isInlineContainer, paragraphLogicalText } from "../offsets";
import { DOCUMENT_OP_REFUSAL_REASONS, type DocumentOpRefusalReason } from "../refusal";
import {
  DOCUMENT_OP_TYPES,
  OP_STORIES,
  REVISION_DECISIONS,
  SPLIT_HALVES,
  type DocumentOp,
} from "../types";

setDefaultTimeout(propertyTestTimeout(30_000));

const shapeArbitrary = fc.record({
  token: fc.constantFrom("a", "é", "😀", "tab"),
  count: fc.integer({ min: 2, max: 4 }),
  wrapper: fc.constantFrom("plain", "link", "control", "nested"),
  formatting: fc.constantFrom("absent", "empty", "italic"),
  priorReview: fc.boolean(),
});

type Shape = fc.ArbitraryValue<typeof shapeArbitrary>;

const fixture = ({ token, count, wrapper, formatting, priorReview }: Shape): Document => {
  const runs: Run[] = [];
  for (let index = 0; index < count; index += 1) {
    const run: Run = {
      type: "run",
      content:
        token === "tab"
          ? [{ type: "tab" }, { type: "tab" }]
          : [{ type: "text", text: token.repeat(3) }],
    };
    if (formatting === "empty") run.formatting = {};
    if (formatting === "italic") run.formatting = { italic: true };
    if (priorReview) {
      run.propertyChanges = [
        {
          type: "runPropertyChange",
          info: { id: 10 + index, author: "Earlier" },
          previousFormatting: { italic: false },
          boundaryJoins: [],
        },
      ];
    }
    runs.push(run);
  }
  let content: ParagraphContent[] = runs;
  if (wrapper === "link") content = [{ type: "hyperlink", anchor: "target", children: runs }];
  if (wrapper === "control")
    content = [{ type: "inlineSdt", properties: { id: 100 }, content: runs }];
  if (wrapper === "nested") {
    content = [
      {
        type: "inlineSdt",
        properties: { id: 100 },
        content: [{ type: "hyperlink", anchor: "target", children: runs }],
      },
    ];
  }
  const secondContent = structuredClone(content);
  const remint = (nodes: readonly ParagraphContent[]): void => {
    for (const node of nodes) {
      if (node.type === "run") {
        for (const change of node.propertyChanges ?? []) change.info.id += 20;
      }
      if (node.type === "inlineSdt" && node.properties.id !== undefined) node.properties.id += 100;
      if (isInlineContainer(node)) remint(childrenOf(node));
    }
  };
  remint(secondContent);
  return normalizeForOps({
    package: {
      document: {
        content: [
          { type: "paragraph", paraId: "00000001", content },
          { type: "paragraph", paraId: "00000002", content: secondContent },
        ],
      },
    },
  });
};

const position = ({ offset, blockId = "00000001" }: { offset: number; blockId?: string }) => ({
  story: OP_STORIES.MAIN,
  blockId,
  offset,
});

const lengthOf = (document: Document, blockId = "00000001"): number => {
  const paragraph = document.package.document.content.find(
    (node) => node.type === "paragraph" && node.paraId === blockId,
  );
  if (paragraph?.type !== "paragraph") panic("The fixture must contain its paragraph.");
  return paragraphLogicalText(paragraph).length;
};

const expectInverse = (applied: AppliedDocumentOp, original: Document): void => {
  const undo = applyDocumentOps(applied.document, applied.inverse);
  if (undo.isErr()) throw undo.error;
  assertExactModel(undo.value.document, original);
  const redo = applyDocumentOps(undo.value.document, undo.value.inverse);
  if (redo.isErr()) throw redo.error;
  assertExactModel(redo.value.document, applied.document);
};

type ResolveOptions = { document: Document; revisionIds: number[]; decision: "accept" | "reject" };
const resolve = ({ document, revisionIds, decision }: ResolveOptions): Document => {
  const result = applyDocumentOp(document, {
    type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
    story: OP_STORIES.MAIN,
    revisionIds,
    decision,
  });
  if (result.isErr()) throw result.error;
  expectInverse(result.value, document);
  return result.value.document;
};

type CompileOptions = {
  document: Document;
  intent: EditorIntent;
  mode: "editing" | "suggesting";
  allocation?: ReturnType<typeof allocateEditorIntentIds>;
  refusals: Map<DocumentOpRefusalReason, number>;
};
const compile = ({ document, intent, mode, allocation: supplied, refusals }: CompileOptions) => {
  const allocation = supplied ?? allocateEditorIntentIds(document, intent);
  const result = compileEditorIntent(document, {
    intent,
    mode:
      mode === "suggesting"
        ? {
            type: "suggesting",
            revision: { id: allocation.revisionId, author: "Editor", date: "2026-10-02T00:00:00Z" },
            newIds: allocation.newIds,
          }
        : { type: "editing", newIds: allocation.newIds },
  });
  if (result.isErr()) {
    refusals.set(result.error.reason, (refusals.get(result.error.reason) ?? 0) + 1);
    throw result.error;
  }
  const applied = applyDocumentOps(document, result.value.ops);
  if (applied.isErr()) {
    refusals.set(applied.error.reason, (refusals.get(applied.error.reason) ?? 0) + 1);
    throw applied.error;
  }
  expectInverse(applied.value, document);
  return applied.value;
};

// The broad intent property tested each action in isolation. These sequences
// retain adjacent equal authored records while same-stamp wrappers accumulate.
test("generated same-author insertion bursts preserve exact authored seams", () => {
  const refusals = new Map<DocumentOpRefusalReason, number>();
  assertProperty(
    fc.property(
      shapeArbitrary,
      fc.array(
        fc.record({
          at: fc.nat(30),
          atom: fc.boolean(),
          paragraph: fc.integer({ min: 1, max: 2 }),
        }),
        { minLength: 2, maxLength: 4 },
      ),
      (shape, steps) => {
        const original = fixture(shape);
        let direct = original;
        let tracked = original;
        const revisions: number[] = [];
        const journal: AppliedDocumentOp[] = [];
        for (const [index, step] of steps.entries()) {
          const target = index === 1 ? (steps.at(0) ?? panic("A burst has a first step.")) : step;
          const blockId = target.paragraph === 1 ? "00000001" : "00000002";
          const width = lengthOf(direct, blockId);
          // Include adjacent atoms inside a run, then arbitrary authored boundaries.
          const paragraph = direct.package.document.content.find(
            (node) => node.type === "paragraph" && node.paraId === blockId,
          );
          if (paragraph?.type !== "paragraph") panic("The fixture must retain its paragraph.");
          const gaps = [0];
          for (const character of paragraphLogicalText(paragraph)) {
            gaps.push((gaps.at(-1) ?? 0) + character.length);
          }
          const initialGap = shape.token === "😀" ? 2 : 1;
          const offset =
            index < 2
              ? Math.min(initialGap + index, width)
              : (gaps.at(step.at % gaps.length) ?? panic("A generated gap must exist."));
          const at = position({ offset, blockId });
          const intent: EditorIntent = step.atom
            ? { type: "insertAtom", from: at, to: at, atom: { type: "tab" } }
            : { type: "replaceText", from: at, to: at, text: "a" };
          // Fresh IDs are caller input. Both modes consume one authoritative
          // allocation; their pending revisions occupy different package IDs.
          const allocation = allocateEditorIntentIds(tracked, intent);
          const edited = compile({
            document: direct,
            intent,
            mode: "editing",
            allocation,
            refusals,
          });
          const suggested = compile({
            document: tracked,
            intent,
            mode: "suggesting",
            allocation,
            refusals,
          });
          direct = edited.document;
          tracked = suggested.document;
          revisions.push(...suggested.revisions);
          journal.push(suggested);
          expect(
            resolve({
              document: tracked,
              revisionIds: revisions,
              decision: REVISION_DECISIONS.ACCEPT,
            }),
          ).toStrictEqual(direct);
          expect(
            resolve({
              document: tracked,
              revisionIds: revisions,
              decision: REVISION_DECISIONS.REJECT,
            }),
          ).toStrictEqual(original);
        }
        for (const edit of journal.toReversed()) {
          const undo = applyDocumentOps(tracked, edit.inverse);
          if (undo.isErr()) throw undo.error;
          tracked = undo.value.document;
        }
        expect(tracked).toStrictEqual(original);
      },
    ),
    { numRuns: 50 },
  );
  expect([...refusals]).toStrictEqual([]);
});

test("generated deletion cuts accept and reject at the recorded wrapper depth", () => {
  const refusals = new Map<DocumentOpRefusalReason, number>();
  assertProperty(
    fc.property(shapeArbitrary, fc.nat(50), fc.nat(50), (shape, start, extent) => {
      const original = fixture(shape);
      const width = lengthOf(original);
      const unit = shape.token === "😀" ? 2 : 1;
      const from = (start % (width / unit)) * unit;
      const to = from + (1 + (extent % ((width - from) / unit))) * unit;
      const intent: EditorIntent = {
        type: "replaceText",
        from: position({ offset: from }),
        to: position({ offset: to }),
        text: "",
      };
      const direct = compile({ document: original, intent, mode: "editing", refusals });
      const tracked = compile({ document: original, intent, mode: "suggesting", refusals });
      expect(
        resolve({
          document: tracked.document,
          revisionIds: tracked.revisions,
          decision: REVISION_DECISIONS.ACCEPT,
        }),
      ).toStrictEqual(direct.document);
      expect(
        resolve({
          document: tracked.document,
          revisionIds: tracked.revisions,
          decision: REVISION_DECISIONS.REJECT,
        }),
      ).toStrictEqual(original);
    }),
    { numRuns: 50 },
  );
  expect([...refusals]).toStrictEqual([]);
});

const runsIn = (content: readonly ParagraphContent[]): Run[] => {
  const runs: Run[] = [];
  for (const node of content) {
    if (node.type === "run") runs.push(node);
    else if (isInlineContainer(node)) runs.push(...runsIn(childrenOf(node)));
  }
  return runs;
};

test("generated later property rejection retains earlier cut provenance", () => {
  const refusals = new Map<DocumentOpRefusalReason, number>();
  assertProperty(
    fc.property(shapeArbitrary, fc.nat(10), fc.nat(10), (shape, start, extent) => {
      const original = fixture(shape);
      const width = lengthOf(original);
      const unit = shape.token === "😀" ? 2 : 1;
      const from = (start % (width / unit)) * unit;
      const to = from + (1 + (extent % ((width - from) / unit))) * unit;
      const intent: EditorIntent = {
        type: "formatRun",
        from: position({ offset: from }),
        to: position({ offset: to }),
        patch: { bold: true },
      };
      const direct = compile({ document: original, intent, mode: "editing", refusals });
      const first = compile({ document: original, intent, mode: "suggesting", refusals });
      expect(
        resolve({
          document: first.document,
          revisionIds: first.revisions,
          decision: REVISION_DECISIONS.ACCEPT,
        }),
      ).toStrictEqual(direct.document);
      const second = compile({
        document: first.document,
        intent: {
          type: "formatRun",
          from: position({ offset: from }),
          to: position({ offset: to }),
          patch: { underline: { style: "single" } },
        },
        mode: "suggesting",
        refusals,
      });
      const paragraph = second.document.package.document.content.at(0);
      if (paragraph?.type !== "paragraph") panic("The fixture must retain its paragraph.");
      const newlyRecorded = runsIn(paragraph.content)
        .flatMap((run) => run.propertyChanges ?? [])
        .filter((record) => second.revisions.includes(record.info.id));
      expect(newlyRecorded.length).toBeGreaterThan(0);
      expect(newlyRecorded.every((record) => record.info.author === "Editor")).toBe(true);
      const rejectedFirst = resolve({
        document: second.document,
        revisionIds: first.revisions,
        decision: REVISION_DECISIONS.REJECT,
      });
      expect(
        resolve({
          document: rejectedFirst,
          revisionIds: second.revisions,
          decision: REVISION_DECISIONS.REJECT,
        }),
      ).toStrictEqual(original);
    }),
    { numRuns: 50 },
  );
  expect([...refusals]).toStrictEqual([]);
});

// A later review adds slots to the retained run. Prefix identity transfers
// restore only the recorded slots and preserve the appended review action.
test("generated retained source identities survive a later pending formatting action", () => {
  const refusals = new Map<DocumentOpRefusalReason, number>();
  assertProperty(
    fc.property(shapeArbitrary, (shape) => {
      const original = fixture(Object.assign({}, shape, { priorReview: true }));
      const deletion: EditorIntent = {
        type: "replaceText",
        from: position({ offset: 0 }),
        to: position({ offset: shape.token === "😀" ? 2 : 1 }),
        text: "",
      };
      const allocation = allocateEditorIntentIds(original, deletion);
      const directDeleted = compile({
        document: original,
        intent: deletion,
        mode: "editing",
        allocation,
        refusals,
      });
      const trackedDeleted = compile({
        document: original,
        intent: deletion,
        mode: "suggesting",
        allocation,
        refusals,
      });
      const directFormatting: EditorIntent = {
        type: "formatRun",
        from: position({ offset: 0 }),
        to: position({ offset: lengthOf(directDeleted.document) }),
        patch: { bold: true },
      };
      const trackedFormatting: EditorIntent = {
        type: "formatRun",
        from: position({ offset: shape.token === "😀" ? 2 : 1 }),
        to: position({ offset: lengthOf(trackedDeleted.document) }),
        patch: { bold: true },
      };
      const formattingAllocation = allocateEditorIntentIds(
        trackedDeleted.document,
        trackedFormatting,
      );
      const direct = compile({
        document: directDeleted.document,
        intent: directFormatting,
        mode: "editing",
        allocation: formattingAllocation,
        refusals,
      });
      const tracked = compile({
        document: trackedDeleted.document,
        intent: trackedFormatting,
        mode: "suggesting",
        allocation: formattingAllocation,
        refusals,
      });
      const acceptedDeletion = resolve({
        document: tracked.document,
        revisionIds: trackedDeleted.revisions,
        decision: REVISION_DECISIONS.ACCEPT,
      });
      expect(
        resolve({
          document: acceptedDeletion,
          revisionIds: tracked.revisions,
          decision: REVISION_DECISIONS.ACCEPT,
        }),
      ).toStrictEqual(direct.document);
      const rejectedFormatting = resolve({
        document: tracked.document,
        revisionIds: tracked.revisions,
        decision: REVISION_DECISIONS.REJECT,
      });
      expect(
        resolve({
          document: rejectedFormatting,
          revisionIds: trackedDeleted.revisions,
          decision: REVISION_DECISIONS.REJECT,
        }),
      ).toStrictEqual(original);
    }),
    { numRuns: 25 },
  );
  expect([...refusals]).toStrictEqual([]);
});

test("generated source-review acceptance retires its pending deletion transfer", () => {
  const refusals = new Map<DocumentOpRefusalReason, number>();
  assertProperty(
    fc.property(shapeArbitrary, (shape) => {
      const original = fixture(Object.assign({}, shape, { priorReview: true }));
      const intent: EditorIntent = {
        type: "replaceText",
        from: position({ offset: 0 }),
        to: position({ offset: shape.token === "😀" ? 2 : 1 }),
        text: "",
      };
      const deleted = compile({ document: original, intent, mode: "suggesting", refusals });
      const acceptedSource = resolve({
        document: deleted.document,
        revisionIds: [10],
        decision: REVISION_DECISIONS.ACCEPT,
      });
      const expected = structuredClone(acceptedSource);
      const strip = (nodes: readonly InlineNode[]): InlineNode[] =>
        nodes.flatMap((node) => {
          if (node.type === "deletion" && deleted.revisions.includes(node.info.id)) return [];
          const children = childNodes(node);
          return [children === undefined ? node : rebuildNode(node, strip(children))];
        });
      const paragraph = expected.package.document.content.at(0);
      if (paragraph?.type !== "paragraph") panic("The fixture retains its first paragraph.");
      paragraph.content = asParagraphContent(strip(paragraph.content));
      expect(
        resolve({
          document: acceptedSource,
          revisionIds: deleted.revisions,
          decision: REVISION_DECISIONS.ACCEPT,
        }),
      ).toStrictEqual(expected);
    }),
    { numRuns: 25 },
  );
  expect([...refusals]).toStrictEqual([]);
});

test("malformed journal resolution provenance refuses atomically", () => {
  const original = fixture({
    token: "a",
    count: 2,
    wrapper: "plain",
    formatting: "absent",
    priorReview: true,
  });
  const refusals = new Map<DocumentOpRefusalReason, number>();
  const intent: EditorIntent = {
    type: "replaceText",
    from: position({ offset: 0 }),
    to: position({ offset: 1 }),
    text: "",
  };
  const tracked = compile({ document: original, intent, mode: "suggesting", refusals });
  for (const malformed of [
    null,
    {},
    [null],
    [{ depth: -1, source: [], target: [] }],
    [{ depth: 0, source: null, target: [] }],
    [{ depth: 0, source: [null], target: [null] }],
  ]) {
    const input = structuredClone(tracked.document);
    const paragraph = input.package.document.content.at(0);
    if (paragraph?.type !== "paragraph") panic("The fixture retains its first paragraph.");
    const deletion = paragraph.content.find((node) => node.type === "deletion");
    if (deletion?.type !== "deletion" || deletion.resolutionJoins === undefined)
      panic("The fixture has deletion provenance.");
    // The protocol boundary can receive shapes the TypeScript constructor forbids.
    Reflect.set(deletion.resolutionJoins, "retainedAfter", malformed);
    const before = structuredClone(input);
    const result = applyDocumentOp(input, {
      type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
      story: OP_STORIES.MAIN,
      revisionIds: tracked.revisions,
      decision: REVISION_DECISIONS.ACCEPT,
    });
    expect(result.isErr()).toBe(true);
    if (result.isErr())
      expect(result.error.reason).toBe(DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH);
    expect(input).toStrictEqual(before);
  }
});

// The original range property never split an already pending deletion.
// Rebinding must follow its exact cut pieces across paragraph boundaries.
test("generated pending deletion splits preserve resolution order and exact histories", () => {
  const refusals = new Map<DocumentOpRefusalReason, number>();
  assertProperty(
    fc.property(shapeArbitrary, (shape) => {
      const original = fixture(Object.assign({}, shape, { priorReview: true }));
      const unit = shape.token === "😀" ? 2 : 1;
      const first = compile({
        document: original,
        mode: "suggesting",
        refusals,
        intent: {
          type: "replaceText",
          from: position({ offset: 0 }),
          to: position({ offset: 2 * unit }),
          text: "",
        },
      });
      const second = compile({
        document: first.document,
        mode: "suggesting",
        refusals,
        intent: { type: "splitParagraph", at: position({ offset: unit }), newBlockId: "00000003" },
      });
      const third = compile({
        document: second.document,
        mode: "suggesting",
        refusals,
        intent: {
          type: "replaceText",
          from: position({ offset: 0 }),
          to: position({ offset: 0 }),
          text: "x",
        },
      });
      const ids = [...first.revisions, ...second.revisions, ...third.revisions];
      for (const decision of [REVISION_DECISIONS.ACCEPT, REVISION_DECISIONS.REJECT]) {
        const together = resolve({ document: third.document, revisionIds: ids, decision });
        let separately = third.document;
        for (const revisionIds of [first.revisions, second.revisions, third.revisions])
          separately = resolve({ document: separately, revisionIds, decision });
        assertExactModel(separately, together);
      }
    }),
    { numRuns: 25 },
  );
  expect([...refusals]).toStrictEqual([]);
});

// Comment anchors are deliberately retained by replacement in either mode.
// Atom width remains physical, including references between selected text.
test("generated comment-bearing replacements share the planned anchor policy", () => {
  const refusals = new Map<DocumentOpRefusalReason, number>();
  assertProperty(
    fc.property(fc.integer({ min: 0, max: 3 }), fc.integer({ min: 0, max: 3 }), (before, after) => {
      const content: ParagraphContent[] = [];
      if (before > 0)
        content.push({ type: "run", content: [{ type: "text", text: "a".repeat(before) }] });
      content.push(
        { type: "commentRangeStart", id: 7 },
        { type: "commentReference", id: 7 },
        { type: "commentRangeEnd", id: 7 },
      );
      if (after > 0)
        content.push({ type: "run", content: [{ type: "text", text: "b".repeat(after) }] });
      const original = normalizeForOps({
        package: {
          document: {
            content: [{ type: "paragraph", paraId: "00000001", content }],
            comments: [
              {
                id: 7,
                author: "Earlier",
                content: [{ type: "paragraph", paraId: "7FFFFFF0", content: [] }],
              },
            ],
          },
        },
      });
      const intent: EditorIntent = {
        type: "replaceText",
        from: position({ offset: 0 }),
        to: position({ offset: before + 1 + after }),
        text: "edited",
      };
      const allocation = allocateEditorIntentIds(original, intent);
      const direct = compile({ document: original, intent, allocation, mode: "editing", refusals });
      const tracked = compile({
        document: original,
        intent,
        allocation,
        mode: "suggesting",
        refusals,
      });
      assertExactModel(
        resolve({
          document: tracked.document,
          revisionIds: tracked.revisions,
          decision: REVISION_DECISIONS.ACCEPT,
        }),
        direct.document,
      );
      assertExactModel(
        resolve({
          document: tracked.document,
          revisionIds: tracked.revisions,
          decision: REVISION_DECISIONS.REJECT,
        }),
        original,
      );
      const paragraph = direct.document.package.document.content.at(0);
      if (paragraph?.type !== "paragraph") panic("The replacement must retain its paragraph.");
      expect(paragraph.content.filter((node) => node.type.startsWith("comment"))).toStrictEqual(
        content.filter((node) => node.type.startsWith("comment")),
      );
    }),
    { numRuns: 25 },
  );
  expect([...refusals]).toStrictEqual([]);
});

// Earlier replacement generators selected whole text-only paragraphs. A
// zero-width leaf at the trailing endpoint moves to the leading offset when
// direct deletion runs; insertion must keep the input gap's affinity.
test("generated replacements preserve input affinity around zero-width leaves", () => {
  const refusals = new Map<DocumentOpRefusalReason, number>();
  assertProperty(
    fc.property(
      fc.integer({ min: 2, max: 6 }),
      fc.nat(20),
      fc.nat(20),
      fc.nat(20),
      fc.constantFrom("pageBreak", "bookmark"),
      fc.integer({ min: 1, max: 2 }),
      (width, anchor, head, target, kind, paragraphNumber) => {
        const markerAt = target % (width + 1);
        const content: ParagraphContent[] = [];
        const addText = (text: string): void => {
          if (text !== "") content.push({ type: "run", content: [{ type: "text", text }] });
        };
        addText("a".repeat(markerAt));
        if (kind === "pageBreak")
          content.push({ type: "run", content: [{ type: "renderedPageBreak" }] });
        else
          content.push(
            { type: "bookmarkStart", id: 7, name: "range" },
            { type: "bookmarkEnd", id: 7 },
          );
        addText("a".repeat(width - markerAt));
        const blockId = paragraphNumber === 1 ? "00000001" : "00000002";
        const paragraphs = [
          { type: "paragraph", paraId: "00000001", content: [] },
          { type: "paragraph", paraId: "00000002", content: [] },
        ] satisfies Document["package"]["document"]["content"];
        const chosen = paragraphs.at(paragraphNumber - 1);
        if (chosen === undefined) panic("The generated replacement target must exist.");
        const original = normalizeForOps({
          package: {
            document: {
              content: paragraphs.map((paragraph) =>
                paragraph === chosen ? Object.assign({}, paragraph, { content }) : paragraph,
              ),
            },
          },
        });
        const from = Math.min(anchor % (width + 1), head % (width + 1));
        const to = Math.max(anchor % (width + 1), head % (width + 1));
        const intent: EditorIntent = {
          type: "replaceText",
          from: position({ blockId, offset: from }),
          to: position({ blockId, offset: to }),
          text: "edited",
        };
        const allocation = allocateEditorIntentIds(original, intent);
        const direct = compile({
          document: original,
          intent,
          allocation,
          mode: "editing",
          refusals,
        });
        const tracked = compile({
          document: original,
          intent,
          allocation,
          mode: "suggesting",
          refusals,
        });
        assertExactModel(
          resolve({
            document: tracked.document,
            revisionIds: tracked.revisions,
            decision: REVISION_DECISIONS.ACCEPT,
          }),
          direct.document,
        );
        assertExactModel(
          resolve({
            document: tracked.document,
            revisionIds: tracked.revisions,
            decision: REVISION_DECISIONS.REJECT,
          }),
          original,
        );
        const paragraph = direct.document.package.document.content.find(
          (node) => node.type === "paragraph" && node.paraId === blockId,
        );
        if (paragraph?.type !== "paragraph") panic("Replacement must retain its target.");
        const actual: string[] = [];
        for (const node of paragraph.content) {
          if (node.type === "bookmarkStart") actual.push("marker");
          if (node.type !== "run") continue;
          for (const leaf of node.content) {
            if (leaf.type === "text") actual.push(...leaf.text);
            if (leaf.type === "renderedPageBreak") actual.push("marker");
          }
        }
        const expected = [..."a".repeat(from), ..."edited", ..."a".repeat(width - to)];
        if (markerAt <= from || markerAt >= to) {
          let markerIndex = markerAt;
          if (markerAt > from || (markerAt === from && kind === "bookmark"))
            markerIndex = from + "edited".length + Math.max(0, markerAt - to);
          expected.splice(markerIndex, 0, "marker");
        }
        expect(actual).toStrictEqual(expected);
      },
    ),
    { numRuns: 50 },
  );
  expect([...refusals]).toStrictEqual([]);
});

// Primitive seam inverses cannot reconstruct repartitioned retained-slot
// hints. Exercise every cut producer against an existing pending deletion.
test("generated provenance-bearing cuts capture exact inverse source state", () => {
  const refusals = new Map<DocumentOpRefusalReason, number>();
  assertProperty(
    fc.property(
      shapeArbitrary,
      fc.constantFrom("insert", "inlineSplit", "delete", "format", "blockSplit"),
      (shape, kind) => {
        const original = fixture(Object.assign({}, shape, { priorReview: true }));
        const unit = shape.token === "😀" ? 2 : 1;
        const pending = compile({
          document: original,
          mode: "suggesting",
          refusals,
          intent: {
            type: "replaceText",
            from: position({ offset: 0 }),
            to: position({ offset: 2 * unit }),
            text: "",
          },
        });
        const at = position({ offset: unit });
        const allocation = allocateEditorIntentIds(pending.document, {
          type: "splitParagraph",
          at,
          newBlockId: "00000003",
        });
        const makeOp = () => {
          switch (kind) {
            case "insert":
              return {
                type: DOCUMENT_OP_TYPES.INSERT_CONTENT,
                at,
                slice: {
                  openStart: 0,
                  openEnd: 0,
                  content: [{ type: "run", content: [{ type: "text", text: "x" }] }],
                },
                newIds: allocation.newIds,
              } as const satisfies DocumentOp;
            case "inlineSplit":
              return {
                type: DOCUMENT_OP_TYPES.SPLIT_INLINE,
                at,
                depth: 1,
                newIds: allocation.newIds,
              } as const satisfies DocumentOp;
            case "delete":
              return {
                type: DOCUMENT_OP_TYPES.DELETE_RANGE,
                from: at,
                to: position({ offset: 2 * unit }),
                newIds: allocation.newIds,
              } as const satisfies DocumentOp;
            case "format":
              return {
                type: DOCUMENT_OP_TYPES.SET_RUN_PROPS,
                from: at,
                to: position({ offset: 2 * unit }),
                patch: { bold: true },
                newIds: allocation.newIds,
              } as const satisfies DocumentOp;
            case "blockSplit":
              return {
                type: DOCUMENT_OP_TYPES.SPLIT_BLOCK,
                at,
                newBlockId: "00000003",
                newHalf: SPLIT_HALVES.FIRST,
                newIds: allocation.newIds,
              } as const satisfies DocumentOp;
            default: {
              const unreachable: never = kind;
              return unreachable;
            }
          }
        };
        const changed = applyDocumentOp(pending.document, makeOp());
        if (changed.isErr()) {
          refusals.set(changed.error.reason, (refusals.get(changed.error.reason) ?? 0) + 1);
          throw changed.error;
        }
        expectInverse(changed.value, pending.document);
      },
    ),
    { numRuns: 50 },
  );
  expect([...refusals]).toStrictEqual([]);
});

test("generated replacement retains prior removed revisions and visible inserted ownership", () => {
  const refusals = new Map<DocumentOpRefusalReason, number>();
  assertProperty(
    fc.property(
      fc.integer({ min: 2, max: 6 }),
      fc.nat(20),
      fc.nat(20),
      fc.constantFrom("deletion", "moveFrom"),
      (width, anchor, head, kind) => {
        const original = normalizeForOps({
          package: {
            document: {
              content: [
                {
                  type: "paragraph",
                  paraId: "00000001",
                  content: [
                    {
                      type: kind,
                      info: { id: 10, author: "Earlier" },
                      content: [
                        { type: "run", content: [{ type: "text", text: "a".repeat(width) }] },
                        { type: "bookmarkEnd", id: 7 },
                      ],
                    },
                  ],
                },
              ],
            },
          },
        });
        const from = Math.min(anchor % (width + 1), head % (width + 1));
        const to = Math.max(anchor % (width + 1), head % (width + 1));
        const intent: EditorIntent = {
          type: "replaceText",
          from: position({ offset: from }),
          to: position({ offset: to }),
          text: "edited",
        };
        const allocation = allocateEditorIntentIds(original, intent);
        const direct = compile({
          document: original,
          intent,
          allocation,
          mode: "editing",
          refusals,
        });
        const tracked = compile({
          document: original,
          intent,
          allocation,
          mode: "suggesting",
          refusals,
        });
        assertExactModel(
          resolve({
            document: tracked.document,
            revisionIds: tracked.revisions,
            decision: REVISION_DECISIONS.ACCEPT,
          }),
          direct.document,
        );
        assertExactModel(
          resolve({
            document: tracked.document,
            revisionIds: tracked.revisions,
            decision: REVISION_DECISIONS.REJECT,
          }),
          original,
        );
        const paragraph = direct.document.package.document.content.at(0);
        if (paragraph?.type !== "paragraph") panic("Replacement must retain its target paragraph.");
        const visible = paragraph.content
          .flatMap((node) =>
            node.type === "run"
              ? node.content.flatMap((leaf) => (leaf.type === "text" ? [leaf.text] : []))
              : [],
          )
          .join("");
        expect(visible).toBe("edited");
        expect(paragraphLogicalText(paragraph)).toBe(
          "a".repeat(from) + "edited" + "a".repeat(width - from),
        );
      },
    ),
    { numRuns: 50 },
  );
  expect([...refusals]).toStrictEqual([]);
});

// Joining independent wrappers can produce a pending record whose inverse
// split would manufacture source-slot transfers absent from either source.
test("generated provenance-bearing joins restore independent empty transfer metadata", () => {
  assertProperty(
    fc.property(shapeArbitrary, (shape) => {
      const original = fixture(Object.assign({}, shape, { priorReview: true }));
      for (const [index, paragraph] of original.package.document.content.entries()) {
        if (paragraph.type !== "paragraph") panic("The join fixture must contain two paragraphs.");
        paragraph.content = [
          {
            type: "deletion",
            info: { id: index + 1, author: "Earlier" },
            resolutionJoins: { before: 0, after: 0, remove: 0, retainedAfter: [] },
            content: paragraph.content,
          },
        ];
      }
      let innerDepth = 2;
      if (shape.wrapper === "link" || shape.wrapper === "control") innerDepth += 1;
      if (shape.wrapper === "nested") innerDepth += 2;
      // Tabs are atoms: merge their run, leaving the two tab leaves distinct.
      if (shape.token === "tab") innerDepth -= 1;
      const joined = applyDocumentOp(original, {
        type: DOCUMENT_OP_TYPES.JOIN_BLOCKS,
        story: OP_STORIES.MAIN,
        blockId: "00000001",
        nextBlockId: "00000002",
        depth: innerDepth + 1,
      });
      if (joined.isErr()) throw joined.error;
      expectInverse(joined.value, original);
    }),
    { numRuns: 50 },
  );
});
