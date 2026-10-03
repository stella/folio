import { panic } from "better-result";
import { expect, test } from "bun:test";

import type { BlockContent, Document, Paragraph, ParagraphContent } from "../../model/document";
import { applyDocumentOp, applyDocumentOps, type AppliedDocumentOp } from "../apply";
import { endsItsContainer, storyParagraphs } from "../blocks";
import {
  allocateEditorIntentIds,
  compileEditorIntent,
  editorParagraphGroups,
  physicalPositionAtEditorOffset,
} from "../editorIntent";
import { identityKeysIn } from "../ids";
import { paragraphLength, paragraphLogicalText } from "../offsets";
import { planTrackedDeletion } from "../plan";
import { planTrackedReplace, type PlanTrackedReplaceOptions } from "../rangeReplacement";
import { DOCUMENT_OP_REFUSAL_REASONS } from "../refusal";
import {
  DOCUMENT_OP_TYPES,
  OP_STORIES,
  REVISION_DECISIONS,
  type DocumentOp,
  type RevisionDecision,
} from "../types";

const run = (text: string): ParagraphContent => ({
  type: "run",
  content: [{ type: "text", text }],
});
const paragraph = (paraId: string, text: string): Paragraph => ({
  type: "paragraph",
  paraId,
  content: text === "" ? [] : [run(text)],
});
const at = (blockId: string, offset: number) => ({ story: OP_STORIES.MAIN, blockId, offset });
const revision = { id: 100, author: "Reviewer", date: "2026-02-03T04:05:06Z" };
const newIds = { revision: Array.from({ length: 64 }, (_, index) => 101 + index) };
const documentOf = (content: BlockContent[]): Document => ({ package: { document: { content } } });
const wrap = (content: BlockContent[], container: "body" | "cell"): BlockContent[] =>
  container === "body"
    ? content
    : [{ type: "table", rows: [{ type: "tableRow", cells: [{ type: "tableCell", content }] }] }];
const apply = (document: Document, ops: readonly DocumentOp[]) => {
  const applied = applyDocumentOps(document, ops);
  if (applied.isErr()) throw applied.error;
  return applied.value;
};
const resolve = (applied: AppliedDocumentOp, decision: RevisionDecision) => {
  const result = applyDocumentOp(applied.document, {
    type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
    story: OP_STORIES.MAIN,
    revisionIds: applied.revisions,
    decision,
  });
  if (result.isErr()) throw result.error;
  return result.value;
};
const paragraphs = (document: Document) =>
  storyParagraphs(document.package.document).map(({ paragraph: block }) => block);
const expectUndo = (applied: AppliedDocumentOp, original: Document) =>
  expect(apply(applied.document, applied.inverse).document).toStrictEqual(original);

test.each(["body", "cell"] as const)(
  "cross-paragraph deletion retains endpoint identity and authored formatting in a %s",
  (container) => {
    const first = {
      ...paragraph("00000001", "alpha"),
      formatting: { alignment: "center" },
    } satisfies Paragraph;
    const middle = {
      ...paragraph("00000002", "middle"),
      formatting: { alignment: "start" },
    } satisfies Paragraph;
    const final = {
      ...paragraph("00000003", "omega"),
      formatting: { alignment: "end", runProperties: { bold: true } },
    } satisfies Paragraph;
    const document = documentOf(wrap([first, middle, final], container));
    for (const start of [0, 2]) {
      const planned = planTrackedDeletion(document, {
        from: at("00000001", start),
        to: at("00000003", 3),
        revision,
        newIds,
      });
      if (planned.isErr()) throw planned.error;
      const tracked = apply(document, planned.value);
      const accepted = resolve(tracked, REVISION_DECISIONS.ACCEPT);
      const rejected = resolve(tracked, REVISION_DECISIONS.REJECT);
      expect(paragraphs(accepted.document).map(paragraphLogicalText)).toEqual([
        `${start === 0 ? "" : "al"}ga`,
      ]);
      expect(paragraphs(accepted.document).at(0)?.paraId).toBe(final.paraId);
      expect(paragraphs(accepted.document).at(0)?.formatting).toEqual(
        start === 0 ? final.formatting : { alignment: "center", runProperties: { bold: true } },
      );
      expect(rejected.document).toStrictEqual(document);
      expectUndo(tracked, document);
      expectUndo(accepted, tracked.document);
      expectUndo(rejected, tracked.document);
      expect(new Set(tracked.revisions).size).toBe(tracked.revisions.length);
      expect(
        storyParagraphs(tracked.document.package.document).filter(
          (location) =>
            endsItsContainer(tracked.document.package.document, location) &&
            location.paragraph.pPrMark !== undefined,
        ),
      ).toEqual([]);
    }
  },
);

test.each(["body", "cell"] as const)(
  "whole-range replacement tracks every original unit and restores final properties in a %s",
  (container) => {
    for (const text of ["", "final"]) {
      const first = {
        ...paragraph("00000001", "original"),
        formatting: { alignment: "center" },
      } satisfies Paragraph;
      const final = {
        ...paragraph("00000002", text),
        formatting: { alignment: "end", runProperties: { italic: true } },
      } satisfies Paragraph;
      const document = documentOf(wrap([first, final], container));
      const options = {
        from: at("00000001", 0),
        to: at("00000002", paragraphLength(final)),
        revision,
        newIds,
        replacement: {
          paragraphs: [
            { ...paragraph("00000003", "first new"), formatting: { alignment: "start" } },
          ],
          tail: { content: [run("last new")], openStart: 0, openEnd: 0 },
        },
      } satisfies PlanTrackedReplaceOptions;
      const planned = planTrackedReplace(document, options);
      if (planned.isErr()) throw planned.error;
      const tracked = apply(document, planned.value);
      const accepted = resolve(tracked, REVISION_DECISIONS.ACCEPT);
      const rejected = resolve(tracked, REVISION_DECISIONS.REJECT);
      expect(paragraphs(accepted.document).map(paragraphLogicalText)).toEqual([
        "first new",
        "last new",
      ]);
      expect(paragraphs(accepted.document).map(({ paraId }) => paraId)).toEqual([
        "00000003",
        "00000002",
      ]);
      expect(paragraphs(accepted.document).at(-1)?.formatting).toEqual(final.formatting);
      for (const source of paragraphs(tracked.document).filter(
        ({ paraId }) => paraId !== "00000003",
      )) {
        const selected = source.content.filter((node) => node.type === "deletion");
        expect(
          selected
            .flatMap((node) => (node.type === "deletion" ? node.content : []))
            .map((node) => node.type),
        ).toEqual(source.paraId === "00000001" || text !== "" ? ["run"] : []);
      }
      expect(rejected.document).toStrictEqual(document);
      expectUndo(tracked, document);
      expectUndo(accepted, tracked.document);
      expectUndo(rejected, tracked.document);
    }
  },
);

test("cross-block planning retracts own insertions and preserves comment anchors", () => {
  const document = documentOf([
    {
      type: "paragraph",
      paraId: "00000001",
      content: [
        run("ab"),
        { type: "commentRangeStart", id: 1 },
        {
          type: "insertion",
          info: { id: 9, author: "Reviewer" },
          content: [{ type: "run", content: [{ type: "text", text: "own" }] }],
        },
      ],
    },
    {
      type: "paragraph",
      paraId: "00000002",
      content: [run("end"), { type: "commentRangeEnd", id: 1 }],
    },
  ]);
  const planned = planTrackedDeletion(document, {
    from: at("00000001", 1),
    to: { ...at("00000002", 3), zeroWidthBefore: 1 },
    revision,
    newIds,
  });
  if (planned.isErr()) throw planned.error;
  const tracked = apply(document, planned.value);
  expect(paragraphs(tracked.document).map(paragraphLogicalText)).toEqual(["ab", "end"]);
  expect(identityKeysIn(tracked.document).includes("revision:9")).toBe(false);
  const rejected = resolve(tracked, REVISION_DECISIONS.REJECT);
  expect(paragraphs(rejected.document).map(paragraphLogicalText)).toEqual(["ab", "end"]);
  expect(paragraphs(rejected.document).at(0)?.content.at(-1)?.type).toBe("commentRangeStart");
  expect(paragraphs(rejected.document).at(-1)?.content.at(-1)?.type).toBe("commentRangeEnd");
  expectUndo(tracked, document);
});

test("planner refuses unordered ranges", () => {
  const first = paragraph("00000001", "first");
  const final = paragraph("00000002", "final");
  const cases = [
    {
      document: documentOf([first, final]),
      from: at("00000002", 0),
      to: at("00000001", 1),
      reason: DOCUMENT_OP_REFUSAL_REASONS.NOT_ADJACENT,
    },
  ] satisfies {
    document: Document;
    from: ReturnType<typeof at>;
    to: ReturnType<typeof at>;
    reason: string;
  }[];
  for (const { document, from, to, reason } of cases) {
    const planned = planTrackedDeletion(document, { from, to, revision, newIds });
    expect(planned.isErr() && planned.error.reason).toBe(reason);
  }
});

test.each(["body", "cell"] as const)(
  "new range suggestions fold paragraph reviews to their original baseline in a %s",
  (container) => {
    for (const start of [0, 2]) {
      for (const reviewed of ["leading", "trailing", "both"] as const) {
        const first = {
          ...paragraph("00000001", "alpha"),
          formatting: { alignment: "center" },
          ...(reviewed === "trailing"
            ? {}
            : {
                propertyChanges: [
                  {
                    type: "paragraphPropertyChange",
                    info: { id: 7, author: "Original author" },
                    previousFormatting: { alignment: "start" },
                  },
                ],
              }),
        } satisfies Paragraph;
        const final = {
          ...paragraph("00000002", "omega"),
          formatting: { alignment: "end", runProperties: { bold: true } },
          ...(reviewed === "leading"
            ? {}
            : {
                propertyChanges: [
                  {
                    type: "paragraphPropertyChange",
                    info: { id: 8, author: "Original author" },
                    previousFormatting: { alignment: "start" },
                  },
                ],
              }),
        } satisfies Paragraph;
        const document = documentOf(wrap([first, final], container));
        const direct = apply(document, [
          {
            type: DOCUMENT_OP_TYPES.DELETE_RANGE,
            from: at("00000001", start),
            to: at("00000001", 5),
          },
          { type: DOCUMENT_OP_TYPES.DELETE_RANGE, from: at("00000002", 0), to: at("00000002", 3) },
          {
            type: DOCUMENT_OP_TYPES.JOIN_BLOCKS,
            story: OP_STORIES.MAIN,
            blockId: "00000001",
            nextBlockId: "00000002",
          },
        ]);
        const planned = planTrackedDeletion(document, {
          from: at("00000001", start),
          to: at("00000002", 3),
          revision,
          newIds,
        });
        if (planned.isErr()) throw planned.error;
        const tracked = apply(document, planned.value);
        const foldedIds = paragraphs(tracked.document).flatMap((block) =>
          (block.propertyChanges ?? [])
            .filter(
              ({ info }) =>
                info.id === 8 && info.author === revision.author && info.date === revision.date,
            )
            .map(({ info }) => info.id),
        );
        expect(foldedIds).toEqual(reviewed !== "leading" ? [8] : []);
        const resolveAffected = (
          source: Document,
          resolution: { decision: RevisionDecision; ids: readonly number[] },
        ) =>
          apply(source, [
            {
              type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
              story: OP_STORIES.MAIN,
              revisionIds: resolution.ids,
              decision: resolution.decision,
            },
          ]);
        const accepted = resolveAffected(tracked.document, {
          decision: REVISION_DECISIONS.ACCEPT,
          ids: [...tracked.revisions, ...foldedIds],
        });
        const rejected = resolveAffected(tracked.document, {
          decision: REVISION_DECISIONS.REJECT,
          ids: [...tracked.revisions, ...foldedIds],
        });
        expect(accepted.document).toStrictEqual(
          resolveAffected(direct.document, { decision: REVISION_DECISIONS.ACCEPT, ids: foldedIds })
            .document,
        );
        expect(rejected.document).toStrictEqual(
          resolveAffected(document, { decision: REVISION_DECISIONS.REJECT, ids: foldedIds })
            .document,
        );
        for (const block of paragraphs(tracked.document)) {
          expect(block.propertyChanges?.length ?? 0).toBeLessThanOrEqual(1);
          for (const change of block.propertyChanges ?? []) {
            expect(change.previousFormatting).toEqual({
              alignment: change.info.id === 7 || change.info.id === 8 ? "start" : "end",
            });
          }
        }
        expect(tracked.revisions.every((id) => id !== 7 && id !== 8)).toBe(true);
        expectUndo(tracked, document);
        expectUndo(accepted, tracked.document);
        expectUndo(rejected, tracked.document);
      }
    }
  },
);

test("tracked range deletion across section boundaries preserves section facts and exact inverses", () => {
  for (const start of [0, 2]) {
    const first = {
      ...paragraph("00000001", "alpha"),
      sectionProperties: { pageWidth: 12000, pageHeight: 16000 },
    } satisfies Paragraph;
    const middle = {
      ...paragraph("00000002", "middle"),
      sectionProperties: { pageWidth: 14000, pageHeight: 17000 },
    } satisfies Paragraph;
    const final = paragraph("00000003", "omega");
    const document = documentOf([first, middle, final]);
    const direct = apply(document, [
      { type: DOCUMENT_OP_TYPES.DELETE_RANGE, from: at("00000001", start), to: at("00000001", 5) },
      { type: DOCUMENT_OP_TYPES.DELETE_RANGE, from: at("00000002", 0), to: at("00000002", 6) },
      { type: DOCUMENT_OP_TYPES.DELETE_RANGE, from: at("00000003", 0), to: at("00000003", 3) },
      {
        type: DOCUMENT_OP_TYPES.JOIN_BLOCKS,
        story: OP_STORIES.MAIN,
        blockId: "00000002",
        nextBlockId: "00000003",
        sectionBoundary: "remove",
      },
      {
        type: DOCUMENT_OP_TYPES.JOIN_BLOCKS,
        story: OP_STORIES.MAIN,
        blockId: "00000001",
        nextBlockId: "00000003",
        sectionBoundary: "remove",
      },
    ]);
    const planned = planTrackedDeletion(document, {
      from: at("00000001", start),
      to: at("00000003", 3),
      revision,
      newIds,
    });
    if (planned.isErr()) throw planned.error;
    const tracked = apply(document, planned.value);
    const accepted = resolve(tracked, REVISION_DECISIONS.ACCEPT);
    const rejected = resolve(tracked, REVISION_DECISIONS.REJECT);
    expect(accepted.document).toStrictEqual(direct.document);
    expect(rejected.document).toStrictEqual(document);
    expect(paragraphs(tracked.document).map(({ sectionProperties }) => sectionProperties)).toEqual([
      first.sectionProperties,
      middle.sectionProperties,
      undefined,
    ]);
    expectUndo(direct, document);
    expectUndo(tracked, document);
    expectUndo(accepted, tracked.document);
    expectUndo(rejected, tracked.document);
  }
});

test.each([
  ["00000001", "00000003"],
  ["00000002", "00000004"],
  ["00000003", "00000005"],
  ["00000001", "00000005"],
] as const)("cross-list text range %s to %s preserves container geometry", (fromId, toId) => {
  for (const start of [0, 1]) {
    for (const end of [1, 3]) {
      for (const text of ["", "replacement"]) {
        const document = documentOf([
          paragraph("00000001", "before"),
          {
            type: "table",
            rows: [
              {
                type: "tableRow",
                cells: [
                  {
                    type: "tableCell",
                    content: [paragraph("00000002", "left"), paragraph("00000003", "second")],
                  },
                  { type: "tableCell", content: [paragraph("00000004", "right")] },
                ],
              },
            ],
          },
          paragraph("00000005", "after"),
        ]);
        const intent = {
          type: "replaceText",
          from: at(fromId, start),
          to: at(toId, end),
          text,
        } as const;
        const direct = compileEditorIntent(document, { intent, mode: { type: "editing" } });
        if (direct.isErr()) throw direct.error;
        const ids = allocateEditorIntentIds(document, intent);
        const planned = compileEditorIntent(document, {
          intent,
          mode: {
            type: "suggesting",
            revision: { ...revision, id: ids.revisionId },
            newIds: ids.newIds,
          },
        });
        if (planned.isErr()) throw planned.error;
        const edited = apply(document, direct.value.ops);
        const tracked = apply(document, planned.value.ops);
        const accepted = resolve(tracked, REVISION_DECISIONS.ACCEPT);
        const rejected = resolve(tracked, REVISION_DECISIONS.REJECT);
        expect(accepted.document).toStrictEqual(edited.document);
        expect(rejected.document).toStrictEqual(document);
        const table = accepted.document.package.document.content.find(
          (block) => block.type === "table",
        );
        expect(table?.rows).toHaveLength(1);
        expect(table?.rows.at(0)?.cells).toHaveLength(2);
        expect(editorParagraphGroups(tracked.document, OP_STORIES.MAIN).length).toBe(
          editorParagraphGroups(edited.document, OP_STORIES.MAIN).length,
        );
        expectUndo(edited, document);
        expectUndo(tracked, document);
        expectUndo(accepted, tracked.document);
        expectUndo(rejected, tracked.document);
      }
    }
  }
});

test.each(["splitGrouped", "joinGrouped"] as const)(
  "grouped intent %s preserves survivor formatting and exact rejection",
  (scenario) => {
    const document = documentOf(
      Array.from({ length: 4 }, (_, index) => {
        let alignment: "center" | "start" | "end" = "end";
        if (index === 0) alignment = "center";
        if (index === 2) alignment = "start";
        return {
          ...paragraph((index + 1).toString(16).padStart(8, "0"), "abcd"),
          formatting: {
            alignment,
            runProperties: { italic: index % 2 === 0 },
          },
        } satisfies Paragraph;
      }),
    );
    let direct = document;
    let tracked = document;
    const trackedOps: DocumentOp[] = [];
    const steps = scenario === "joinGrouped" ? [2, 0, 0] : [0];
    for (const groupIndex of steps) {
      const directGroups = editorParagraphGroups(direct, OP_STORIES.MAIN);
      const trackedGroups = editorParagraphGroups(tracked, OP_STORIES.MAIN);
      const firstDirect = directGroups.at(groupIndex);
      const nextDirect = directGroups.at(groupIndex + 1);
      const firstTracked = trackedGroups.at(groupIndex);
      const nextTracked = trackedGroups.at(groupIndex + 1);
      if (!firstDirect || !nextDirect || !firstTracked || !nextTracked)
        panic("Expected adjacent fixture groups");
      const directPlan = compileEditorIntent(direct, {
        intent: {
          type: "joinParagraphs",
          story: OP_STORIES.MAIN,
          blockId: firstDirect.blockId,
          nextBlockId: nextDirect.paragraphs.at(0)?.paraId ?? "",
        },
        mode: { type: "editing" },
      });
      const ids = allocateEditorIntentIds(tracked, {
        type: "joinParagraphs",
        story: OP_STORIES.MAIN,
        blockId: firstTracked.blockId,
        nextBlockId: nextTracked.paragraphs.at(0)?.paraId ?? "",
      });
      const trackedPlan = compileEditorIntent(tracked, {
        intent: {
          type: "joinParagraphs",
          story: OP_STORIES.MAIN,
          blockId: firstTracked.blockId,
          nextBlockId: nextTracked.paragraphs.at(0)?.paraId ?? "",
        },
        mode: {
          type: "suggesting",
          revision: { ...revision, id: ids.revisionId },
          newIds: ids.newIds,
        },
      });
      if (directPlan.isErr()) throw directPlan.error;
      if (trackedPlan.isErr()) throw trackedPlan.error;
      direct = apply(direct, directPlan.value.ops).document;
      tracked = apply(tracked, trackedPlan.value.ops).document;
      trackedOps.push(...trackedPlan.value.ops);
    }
    if (scenario === "splitGrouped") {
      const directGroup = editorParagraphGroups(direct, OP_STORIES.MAIN).at(0);
      const trackedGroup = editorParagraphGroups(tracked, OP_STORIES.MAIN).at(0);
      if (!directGroup || !trackedGroup) panic("Expected split fixture group");
      const ids = allocateEditorIntentIds(tracked, {
        type: "splitParagraph",
        at: physicalPositionAtEditorOffset(tracked, at(trackedGroup.blockId, 1)),
      });
      const directPlan = compileEditorIntent(direct, {
        intent: {
          type: "splitParagraph",
          at: at(directGroup.blockId, 1),
          newBlockId: ids.newBlockId,
        },
        mode: { type: "editing" },
      });
      const trackedPlan = compileEditorIntent(tracked, {
        intent: {
          type: "splitParagraph",
          at: physicalPositionAtEditorOffset(tracked, at(trackedGroup.blockId, 1)),
          newBlockId: ids.newBlockId,
        },
        mode: {
          type: "suggesting",
          revision: { ...revision, id: ids.revisionId },
          newIds: ids.newIds,
        },
      });
      if (directPlan.isErr()) throw directPlan.error;
      if (trackedPlan.isErr()) throw trackedPlan.error;
      direct = apply(direct, directPlan.value.ops).document;
      trackedOps.push(...trackedPlan.value.ops);
    }
    const suggested = apply(document, trackedOps);
    const accepted = resolve(suggested, REVISION_DECISIONS.ACCEPT);
    const rejected = resolve(suggested, REVISION_DECISIONS.REJECT);
    expect(accepted.document).toStrictEqual(direct);
    expect(rejected.document).toStrictEqual(document);
    expectUndo(suggested, document);
    expectUndo(accepted, suggested.document);
    expectUndo(rejected, suggested.document);
  },
);
