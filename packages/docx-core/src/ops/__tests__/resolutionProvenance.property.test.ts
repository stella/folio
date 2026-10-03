import { expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import { assertFreshIdentityEquivalent } from "./freshIdentityOracle";
import { panic } from "better-result";

import { assertExactModel } from "../../../../../test/exactModel";
import { assertProperty, propertyTestTimeout } from "../../../../../test/property-testing";
import type { Document, ParagraphContent, Run } from "../../model/document";
import { applyDocumentOp, applyDocumentOps, type AppliedDocumentOp } from "../apply";
import { storyBody, storyParagraphs } from "../blocks";
import { normalizeForOps } from "../contract";
import { planTrackedReplace } from "../rangeReplacement";
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
    content = [{ type: "inlineSdt", properties: { id: 100, sdtType: "richText" }, content: runs }];
  if (wrapper === "nested") {
    content = [
      {
        type: "inlineSdt",
        properties: { id: 100, sdtType: "richText" },
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

const authoredHyperlinkBookmarkNames = (document: Document): string[][] => {
  const names: string[][] = [];
  const visit = (nodes: readonly ParagraphContent[]): void => {
    for (const node of nodes) {
      if (node.type === "hyperlink") {
        names.push(
          node.children
            .filter((child) => child.type === "bookmarkStart")
            .map((child) => child.name),
        );
      }
      if (isInlineContainer(node)) visit(childrenOf(node));
    }
  };
  for (const { paragraph } of storyParagraphs(storyBody(document, OP_STORIES.MAIN)))
    visit(paragraph.content);
  return names;
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
  author?: string;
};
const compile = ({
  document,
  intent,
  mode,
  allocation: supplied,
  refusals,
  author = "Editor",
}: CompileOptions) => {
  const allocation = supplied ?? allocateEditorIntentIds(document, intent);
  const result = compileEditorIntent(document, {
    intent,
    mode:
      mode === "suggesting"
        ? {
            type: "suggesting",
            revision: { id: allocation.revisionId, author, date: "2026-10-02T00:00:00Z" },
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

test("generated replacement preserves authored hyperlink cuts after emptying an unselected wrapper", () => {
  assertProperty(
    fc.property(fc.constantFrom("edited", "revised"), (replacementText) => {
      for (const wrapper of ["insertion", "moveTo"] as const) {
        for (const location of ["body", "table"] as const) {
          const oldWrapper = {
            type: wrapper,
            info: { id: 1, author: "Earlier" },
            content: [
              {
                type: "mathEquation" as const,
                display: "inline" as const,
                ommlXml: "<m:oMath><m:r><m:t>x</m:t></m:r></m:oMath>",
                plainText: "x",
              },
            ],
          };
          const paragraph = {
            type: "paragraph" as const,
            paraId: "00000001",
            content: [
              {
                type: "hyperlink" as const,
                anchor: "same-target",
                children: [{ type: "bookmarkStart" as const, id: 1, name: "left" }],
              },
              oldWrapper,
              {
                type: "hyperlink" as const,
                anchor: "same-target",
                children: [
                  { type: "bookmarkStart" as const, id: 2, name: "right" },
                  { type: "run" as const, content: [{ type: "text" as const, text: "tail" }] },
                ],
              },
            ],
          };
          const content =
            location === "body"
              ? [paragraph]
              : [
                  {
                    type: "table" as const,
                    rows: [
                      {
                        type: "tableRow" as const,
                        cells: [{ type: "tableCell" as const, content: [paragraph] }],
                      },
                    ],
                  },
                ];
          const original = normalizeForOps({ package: { document: { content } } });
          const intent = {
            type: "replaceText",
            from: position({ offset: 0 }),
            to: position({ offset: 1 }),
            text: replacementText,
          } as const satisfies EditorIntent;
          const allocation = allocateEditorIntentIds(original, intent);
          const refusals = new Map<DocumentOpRefusalReason, number>();
          const direct = compile({
            document: original,
            intent,
            mode: "editing",
            allocation,
            refusals,
          });
          const tracked = compile({
            document: original,
            intent,
            mode: "suggesting",
            allocation,
            refusals,
          });
          const accepted = resolve({
            document: tracked.document,
            revisionIds: tracked.revisions,
            decision: REVISION_DECISIONS.ACCEPT,
          });
          expect(authoredHyperlinkBookmarkNames(accepted)).toStrictEqual([["left"], ["right"]]);
          expect(authoredHyperlinkBookmarkNames(direct.document)).toStrictEqual([
            ["left"],
            ["right"],
          ]);
          assertFreshIdentityEquivalent({
            actual: accepted,
            expected: direct.document,
            original,
            allocated: allocation.newIds,
          });
        }
      }
    }),
    { numRuns: 20 },
  );
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
        if (decision === REVISION_DECISIONS.REJECT) assertExactModel(together, original);
      }
    }),
    { numRuns: 25 },
  );
  expect([...refusals]).toStrictEqual([]);
});

// A replacement can strand a deletion's inner source seam at a control edge.
// Resolving the deletion first must retain that fact for the payload blocker.
test("generated nested replacements preserve reviewed source seams in every resolution order", () => {
  const refusals = new Map<DocumentOpRefusalReason, number>();
  assertProperty(
    fc.property(
      shapeArbitrary,
      fc.nat(30),
      fc.nat(30),
      fc.integer({ min: 1, max: 2 }),
      fc.boolean(),
      (shape, start, end, paragraph, retainedPayloadBoundary) => {
        const original = fixture(
          Object.assign({}, shape, {
            wrapper: shape.wrapper === "nested" ? "nested" : "control",
          }),
        );
        const blockId = paragraph === 1 ? "00000001" : "00000002";
        const source = original.package.document.content.find(
          (node) => node.type === "paragraph" && node.paraId === blockId,
        );
        if (source?.type !== "paragraph") panic("The generated source paragraph exists.");
        const gaps = [0];
        for (const character of paragraphLogicalText(source))
          gaps.push((gaps.at(-1) ?? 0) + character.length);
        const lower = start % (gaps.length - 1);
        const upper = lower + 1 + (end % (gaps.length - lower - 1));
        const from = position({
          blockId,
          offset: gaps.at(lower) ?? panic("The source start is a generated gap."),
        });
        const to = position({
          blockId,
          offset: gaps.at(upper) ?? panic("The source end is a generated gap."),
        });
        const intent = { type: "replaceText", from, to, text: "paste" } as const;
        const allocation = allocateEditorIntentIds(original, intent);
        const direct = compile({
          document: original,
          intent,
          mode: "editing",
          allocation,
          refusals,
        });
        const tracked = compile({
          document: original,
          intent,
          mode: "suggesting",
          allocation,
          refusals,
        });
        let trackedDocument = tracked.document;
        if (retainedPayloadBoundary) {
          // Clipboard planning closes the removed deletion seam when retained
          // payload separates its fragments. Exercise that same primitive
          // snapshot postpass without a second compiler in the parent package.
          const deletionIds = new Set(tracked.revisions);
          const closeRemovedSeam = (node: InlineNode): InlineNode => {
            if (
              node.type === "deletion" &&
              deletionIds.has(node.info.id) &&
              node.resolutionJoins !== undefined
            )
              return Object.assign({}, node, {
                resolutionJoins: Object.assign({}, node.resolutionJoins, { remove: 0 }),
              });
            const children = childNodes(node);
            return children === undefined
              ? node
              : rebuildNode(node, children.map(closeRemovedSeam));
          };
          const selected = trackedDocument.package.document.content.find(
            (node) => node.type === "paragraph" && node.paraId === blockId,
          );
          if (selected?.type !== "paragraph") panic("Tracked replacement retains its paragraph.");
          const captured = applyDocumentOp(trackedDocument, {
            type: DOCUMENT_OP_TYPES.REPLACE_INLINE,
            story: OP_STORIES.MAIN,
            blockId,
            expected: selected.content,
            content: asParagraphContent(selected.content.map(closeRemovedSeam)),
          });
          if (captured.isErr()) throw captured.error;
          expectInverse(captured.value, trackedDocument);
          trackedDocument = captured.value.document;
        }
        for (const decision of [REVISION_DECISIONS.ACCEPT, REVISION_DECISIONS.REJECT]) {
          const together = resolve({
            document: trackedDocument,
            revisionIds: tracked.revisions,
            decision,
          });
          if (decision === REVISION_DECISIONS.ACCEPT) {
            assertFreshIdentityEquivalent({
              actual: together,
              expected: direct.document,
              original,
              allocated: allocation.newIds,
            });
          } else assertExactModel(together, original);
          for (const order of [tracked.revisions, tracked.revisions.toReversed()]) {
            let separately = trackedDocument;
            for (const id of order)
              separately = resolve({ document: separately, revisionIds: [id], decision });
            assertExactModel(separately, together);
          }
        }
        // Mixed decisions also commute: an accepted deletion cannot donate
        // its rejected-source seam to a subsequently rejected payload.
        const firstId = tracked.revisions.at(0);
        const secondId = tracked.revisions.at(1);
        if (firstId === undefined || secondId === undefined)
          panic("A nonempty tracked replacement stamps deletion and insertion.");
        for (const firstDecision of [REVISION_DECISIONS.ACCEPT, REVISION_DECISIONS.REJECT]) {
          const secondDecision =
            firstDecision === REVISION_DECISIONS.ACCEPT
              ? REVISION_DECISIONS.REJECT
              : REVISION_DECISIONS.ACCEPT;
          const firstThenSecond = resolve({
            document: resolve({
              document: trackedDocument,
              revisionIds: [firstId],
              decision: firstDecision,
            }),
            revisionIds: [secondId],
            decision: secondDecision,
          });
          const secondThenFirst = resolve({
            document: resolve({
              document: trackedDocument,
              revisionIds: [secondId],
              decision: secondDecision,
            }),
            revisionIds: [firstId],
            decision: firstDecision,
          });
          assertExactModel(firstThenSecond, secondThenFirst);
        }
      },
    ),
    { numRuns: 100 },
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

// Paragraph rejection previously merged every alike inner authored wrapper,
// and a later pending insertion could erase the source-cut seam altogether.
test("generated paragraph cut depths preserve authored wrappers and deferred split seams", () => {
  const refusals = new Map<DocumentOpRefusalReason, number>();
  assertProperty(
    fc.property(
      fc.boolean(),
      fc.integer({ min: 0, max: 3 }),
      fc.integer({ min: 1, max: 2 }),
      fc.boolean(),
      fc.boolean(),
      (nested, cut, count, foreign, pendingSource) => {
        const siblings: ParagraphContent[] = [
          {
            type: "inlineWrapper",
            kind: "bidi",
            control: "embedding",
            content: [
              { type: "preservedInline", xml: "<w:proofErr w:type='spellStart'/>", text: "" },
            ],
          },
          {
            type: "inlineWrapper",
            kind: "bidi",
            control: "embedding",
            content: [{ type: "run", formatting: {}, content: [{ type: "text", text: "aaa" }] }],
          },
        ];
        const authoredContent: ParagraphContent[] = nested
          ? [{ type: "hyperlink", anchor: "target", children: siblings }]
          : siblings;
        const sourceContent: ParagraphContent[] = pendingSource
          ? [
              {
                type: "insertion",
                info: { id: 17, author: "Earlier" },
                content: authoredContent,
                resolutionJoins: { before: 0, after: 0, remove: 0 },
              },
            ]
          : authoredContent;
        const original = normalizeForOps({
          package: {
            document: {
              content: [
                {
                  type: "paragraph",
                  paraId: "00000001",
                  content: sourceContent,
                },
                { type: "paragraph", paraId: "00000002", content: [] },
              ],
            },
          },
        });
        const intent: EditorIntent = {
          type: "splitParagraph",
          at: position({ offset: cut }),
          newBlockId: "00000003",
        };
        const allocation = allocateEditorIntentIds(original, intent);
        const direct = compile({
          document: original,
          intent,
          allocation,
          mode: "editing",
          refusals,
        });
        const split = compile({
          document: original,
          intent,
          allocation,
          mode: "suggesting",
          refusals,
        });
        assertExactModel(
          resolve({ document: split.document, revisionIds: split.revisions, decision: "accept" }),
          direct.document,
        );
        assertExactModel(
          resolve({ document: split.document, revisionIds: split.revisions, decision: "reject" }),
          original,
        );
        let current = split.document;
        const blockers: number[][] = [];
        for (let index = 0; index < count; index += 1) {
          const inserted = compile({
            document: current,
            intent: {
              type: "replaceText",
              from: position({ offset: 0 }),
              to: position({ offset: 0 }),
              text: "x",
            },
            mode: "suggesting",
            author: foreign ? "Later" : "Editor",
            refusals,
          });
          current = inserted.document;
          blockers.push(inserted.revisions);
        }
        const groups = [split.revisions, ...blockers];
        const together = resolve({
          document: current,
          revisionIds: groups.flat(),
          decision: "reject",
        });
        assertExactModel(together, original);
        for (const order of [groups, groups.toReversed(), [...blockers, split.revisions]]) {
          let separate = current;
          for (const revisionIds of order)
            separate = resolve({ document: separate, revisionIds, decision: "reject" });
          assertExactModel(separate, together);
          if (pendingSource) {
            const sourceIds: number[] = [];
            const collect = (nodes: readonly InlineNode[]): void => {
              for (const node of nodes) {
                if (node.type === "insertion" && node.info.author === "Earlier")
                  sourceIds.push(node.info.id);
                collect(childNodes(node) ?? []);
              }
            };
            for (const block of current.package.document.content) {
              if (block.type === "paragraph") collect(block.content);
            }
            const allTogether = resolve({
              document: current,
              revisionIds: [...groups.flat(), ...sourceIds],
              decision: "reject",
            });
            const allSeparate = resolve({
              document: separate,
              revisionIds: [17],
              decision: "reject",
            });
            assertExactModel(allSeparate, allTogether);
          }
        }
        // Accepting any intervening insertion breaks only its deferred source
        // seam; rejecting the others must preserve that accepted authored payload.
        if (blockers.length === 2) {
          const acceptIds = blockers.at(0) ?? panic("Two blockers contain a first group.");
          const rejectIds = blockers.at(1) ?? panic("Two blockers contain a second group.");
          let left = resolve({
            document: current,
            revisionIds: split.revisions,
            decision: "reject",
          });
          left = resolve({ document: left, revisionIds: acceptIds, decision: "accept" });
          left = resolve({ document: left, revisionIds: rejectIds, decision: "reject" });
          let right = resolve({ document: current, revisionIds: rejectIds, decision: "reject" });
          right = resolve({ document: right, revisionIds: acceptIds, decision: "accept" });
          right = resolve({ document: right, revisionIds: split.revisions, decision: "reject" });
          assertExactModel(left, right);
        }
      },
    ),
    { numRuns: 50 },
  );
  expect([...refusals]).toStrictEqual([]);
});

// Earlier broad fixtures supplied ample IDs and tested only isolated actions.
// Action stamps must never displace the IDs assigned to inherited source cuts.
test("generated compound intents reserve source IDs before additional action stamps", () => {
  const refusals = new Map<DocumentOpRefusalReason, number>();
  assertProperty(
    fc.property(
      fc.integer({ min: 1, max: 30 }),
      fc.integer({ min: 1, max: 3 }),
      fc.boolean(),
      (sourceId, paragraphCount, marker) => {
        const content: ParagraphContent[] = marker
          ? [
              {
                type: "insertion",
                info: { id: sourceId, author: "Earlier" },
                content: [
                  { type: "bookmarkEnd", id: 3 },
                  {
                    type: "hyperlink",
                    anchor: "target",
                    children: [{ type: "bookmarkStart", id: 1, name: "target" }],
                  },
                ],
              },
              { type: "run", content: [{ type: "text", text: "a" }] },
            ]
          : [
              {
                type: "run",
                content: [{ type: "tab" }, { type: "renderedPageBreak" }],
                propertyChanges: [
                  {
                    type: "runPropertyChange",
                    info: { id: sourceId, author: "Earlier" },
                    previousFormatting: { italic: true },
                  },
                ],
              },
            ];
        const blocks = [];
        for (let index = 0; index < paragraphCount; index += 1)
          blocks.push({
            type: "paragraph" as const,
            paraId: (index + 1).toString(16).padStart(8, "0").toUpperCase(),
            content: [],
          });
        const finalId = (paragraphCount + 1).toString(16).padStart(8, "0").toUpperCase();
        const original = normalizeForOps({
          package: {
            document: { content: [...blocks, { type: "paragraph", paraId: finalId, content }] },
          },
        });
        const intent: EditorIntent = marker
          ? {
              type: "replaceText",
              from: position({ offset: 0, blockId: finalId }),
              to: position({ offset: 1, blockId: finalId }),
              text: "x",
            }
          : {
              type: "formatRun",
              from: position({ offset: 0 }),
              to: position({ offset: 1, blockId: finalId }),
              patch: { bold: true },
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
            decision: "accept",
          }),
          direct.document,
        );
        assertExactModel(
          resolve({
            document: tracked.document,
            revisionIds: tracked.revisions,
            decision: "reject",
          }),
          original,
        );
        const before = structuredClone(original);
        const short = compileEditorIntent(original, {
          intent,
          mode: {
            type: "suggesting",
            revision: { id: allocation.revisionId, author: "Editor" },
            newIds: { revision: [], control: allocation.newIds.control },
          },
        });
        expect(short.isErr()).toBe(true);
        if (short.isOk()) panic("This compound intent needs source-cut or additional action IDs.");
        expect(short.error.reason).toBe(DOCUMENT_OP_REFUSAL_REASONS.NEEDS_NEW_IDS);
        assertExactModel(original, before);
      },
    ),
    { numRuns: 50 },
  );
  expect([...refusals]).toStrictEqual([]);
});

// Identity restoration previously updated physical wrappers and retained
// slots while leaving deferred blocker references and selected IDs stale.
test("generated retained wrapper transfers rebase deferred references and selected lineage", () => {
  assertProperty(
    fc.property(fc.integer({ min: 5, max: 30 }), fc.boolean(), (sourceId, ownEmptyFormatting) => {
      const targetId = sourceId + 1;
      const deletionId = sourceId + 2;
      const original = normalizeForOps({
        package: {
          document: {
            content: [
              {
                type: "paragraph",
                paraId: "00000001",
                pPrMark: {
                  kind: "ins",
                  info: { id: sourceId + 3, author: "Split" },
                  resolutionJoin: 2,
                },
                content: [
                  {
                    type: "inlineSdt",
                    properties: { id: targetId },
                    content: [{ type: "run", content: [{ type: "text", text: "kept" }] }],
                  },
                  {
                    type: "deletion",
                    info: { id: deletionId, author: "Earlier" },
                    content: [
                      {
                        type: "insertion",
                        info: { id: sourceId, author: "Author" },
                        content: [{ type: "run", content: [{ type: "text", text: "source" }] }],
                      },
                    ],
                    resolutionJoins: {
                      before: 0,
                      after: 0,
                      remove: 0,
                      retainedAfter: [
                        {
                          depth: 0,
                          source: [{ space: "revision", id: sourceId }],
                          target: [{ space: "revision", id: targetId }],
                        },
                      ],
                    },
                  },
                  {
                    type: "insertion",
                    info: { id: targetId, author: "Author" },
                    content: [
                      {
                        type: "run",
                        ...(ownEmptyFormatting ? { formatting: {} } : {}),
                        content: [{ type: "text", text: "target" }],
                      },
                    ],
                    resolutionJoins: {
                      before: 0,
                      after: 0,
                      remove: 0,
                      deferredRemove: [{ depth: 0, blockers: [targetId] }],
                    },
                  },
                ],
              },
              {
                type: "paragraph",
                paraId: "00000002",
                content: [{ type: "run", content: [{ type: "text", text: "following" }] }],
              },
            ],
          },
        },
      });
      const first = resolve({ document: original, revisionIds: [deletionId], decision: "accept" });
      const paragraph = first.package.document.content.at(0);
      if (paragraph?.type !== "paragraph") panic("The source transfer retains its paragraph.");
      const retained = paragraph.content.at(1);
      if (retained?.type !== "insertion") panic("The source transfer retains its target wrapper.");
      expect(retained.info.id).toBe(sourceId);
      expect(retained.resolutionJoins?.deferredRemove).toStrictEqual([
        { depth: 0, blockers: [sourceId] },
      ]);
      const separate = resolve({ document: first, revisionIds: [sourceId], decision: "accept" });
      const together = resolve({
        document: original,
        revisionIds: [deletionId, targetId],
        decision: "accept",
      });
      assertExactModel(separate, together);
      const selectedSource = resolve({
        document: original,
        revisionIds: [deletionId, sourceId],
        decision: "accept",
      });
      assertExactModel(selectedSource, together);
      const selectedParagraph = selectedSource.package.document.content.at(0);
      if (selectedParagraph?.type !== "paragraph")
        panic("The selected lineage retains its paragraph.");
      expect(selectedParagraph.pPrMark?.resolutionJoin).toBe(0);
      assertExactModel(
        resolve({ document: selectedSource, revisionIds: [sourceId + 3], decision: "reject" }),
        resolve({ document: together, revisionIds: [sourceId + 3], decision: "reject" }),
      );
      resolve({ document: first, revisionIds: [sourceId], decision: "reject" });
    }),
    { numRuns: 25 },
  );
});

// The range law already generates replacement shapes; this sequence isolates
// a deletion cut exposed by rejecting the subsequently inserted paragraph.
test("generated replacement breaks restore deletion cut depths across paragraphs", () => {
  assertProperty(
    fc.property(
      fc.integer({ min: 3, max: 12 }),
      fc.nat(20),
      fc.boolean(),
      fc.boolean(),
      (width, cutPick, linked, explicitFormatting) => {
        const run: Run = {
          type: "run",
          content: [{ type: "text", text: "a".repeat(width) }],
          ...(explicitFormatting ? { formatting: { bold: false, italic: false } } : {}),
        };
        const content: ParagraphContent[] = linked
          ? [{ type: "hyperlink", anchor: "target", children: [run] }]
          : [run];
        const original = normalizeForOps({
          package: {
            document: {
              content: [
                { type: "paragraph", paraId: "00000001", content },
                { type: "paragraph", paraId: "00000002", content: [] },
              ],
            },
          },
        });
        const revisionIds = Array.from({ length: 32 }, (_, index) => 1001 + index);
        const planned = planTrackedReplace(original, {
          from: position({ offset: 1 + (cutPick % (width - 1)) }),
          to: position({ blockId: "00000002", offset: 0 }),
          revision: { id: 1000, author: "Reviewer" },
          newIds: { revision: revisionIds },
          replacement: {
            paragraphs: [{ type: "paragraph", paraId: "00000003", content: [] }],
            tail: { content: [], openStart: 0, openEnd: 0 },
          },
        });
        if (planned.isErr()) throw planned.error;
        const applied = applyDocumentOps(original, planned.value);
        if (applied.isErr()) throw applied.error;
        expectInverse(applied.value, original);
        const rejected = resolve({
          document: applied.value.document,
          revisionIds: [1000, ...revisionIds],
          decision: REVISION_DECISIONS.REJECT,
        });
        assertExactModel(rejected, original);
      },
    ),
    { numRuns: 50 },
  );
});
