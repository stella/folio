import { expect, test } from "bun:test";

import type { BlockContent, Document, Paragraph, ParagraphContent } from "../../model/document";
import { applyDocumentOp, applyDocumentOps, type AppliedDocumentOp } from "../apply";
import { endsItsContainer, storyParagraphs } from "../blocks";
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

test("planner refuses unordered and cross-list ranges and existing property reviews", () => {
  const first = paragraph("00000001", "first");
  const final = paragraph("00000002", "final");
  const cases = [
    {
      document: documentOf([first, final]),
      from: at("00000002", 0),
      to: at("00000001", 1),
      reason: DOCUMENT_OP_REFUSAL_REASONS.NOT_ADJACENT,
    },
    {
      document: documentOf([first, ...wrap([final], "cell")]),
      from: at("00000001", 0),
      to: at("00000002", 1),
      reason: DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
    },
    {
      document: documentOf([
        {
          ...first,
          propertyChanges: [{ type: "paragraphPropertyChange", info: { id: 9, author: "A" } }],
        },
        final,
      ]),
      from: at("00000001", 0),
      to: at("00000002", 1),
      reason: DOCUMENT_OP_REFUSAL_REASONS.REVISION_CONFLICT,
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
