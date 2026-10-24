import { panic } from "better-result";
import { expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../../test/property-testing";
import { projectReview } from "../../../../../test/reviewProjection";
import type { BlockContent, Document, Paragraph } from "../../model/document";
import { applyDocumentOp, applyDocumentOps } from "../apply";
import { storyParagraphs } from "../blocks";
import { normalizeForOps } from "../contract";
import {
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  OP_STORIES,
  REVISION_DECISIONS,
} from "../types";

setDefaultTimeout(propertyTestTimeout(30_000));

const apply = (document: Document, operation: DocumentOp) => {
  const edit = applyDocumentOp(document, operation).unwrap();
  const replayed: DocumentOp[] = JSON.parse(JSON.stringify(edit.inverse));
  expect(
    applyDocumentOps(edit.document, replayed).unwrap().document,
  ).toStrictEqual(document);
  return edit.document;
};

const resolve = (document: Document, revisionIds: number[]) =>
  apply(document, {
    type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
    story: OP_STORIES.MAIN,
    revisionIds,
    decision: REVISION_DECISIONS.ACCEPT,
  });

const fixture = (length: number, container: "body" | "cell") => {
  const paragraphs = Array.from({ length }, (_, index): Paragraph => ({
    type: "paragraph",
    paraId: (index + 1).toString(16).padStart(8, "0"),
    content: index === 0 ? [{ type: "commentRangeStart", id: 7 }] : [],
    ...(index === length - 1
      ? { pPrMark: { kind: "del", info: { id: 1, author: "Earlier" } } }
      : {}),
  }));
  const blocks: BlockContent[] = [
    ...paragraphs,
    {
      type: "table",
      rows: [
        {
          type: "tableRow",
          cells: [
            {
              type: "tableCell",
              content: [{ type: "paragraph", paraId: "00000100", content: [] }],
            },
          ],
        },
      ],
    },
    {
      type: "paragraph",
      paraId: "00000200",
      content: [
        { type: "commentRangeEnd", id: 7 },
        { type: "commentReference", id: 7 },
      ],
    },
  ];
  const source = normalizeForOps({
    package: {
      document: {
        content:
          container === "body"
            ? blocks
            : [
                {
                  type: "table",
                  rows: [
                    {
                      type: "tableRow",
                      cells: [{ type: "tableCell", content: blocks }],
                    },
                  ],
                },
                { type: "paragraph", paraId: "00000300", content: [] },
              ],
        comments: [
          {
            id: 7,
            author: "Reviewer",
            content: [{ type: "paragraph", paraId: "00000400", content: [] }],
          },
        ],
      },
    },
  });
  const operations: DocumentOp[] = [];
  for (let index = 0; index < paragraphs.length - 1; index += 1) {
    operations.push({
      type: DOCUMENT_OP_TYPES.JOIN_BLOCKS,
      story: OP_STORIES.MAIN,
      blockId:
        paragraphs.at(index)?.paraId ?? panic("A chain predecessor exists."),
      nextBlockId:
        paragraphs.at(index + 1)?.paraId ?? panic("A chain successor exists."),
      depth: 0,
      revision: {
        id: index + 2,
        author: "Reviewer",
        date: "2026-03-04T05:00:00Z",
      },
    });
  }
  operations.push({
    type: DOCUMENT_OP_TYPES.INSERT_TEXT,
    at: {
      story: OP_STORIES.MAIN,
      blockId:
        paragraphs.at(-1)?.paraId ?? panic("A chain has a final paragraph."),
      offset: 0,
    },
    text: "x",
    revision: {
      id: length + 1,
      author: "Reviewer",
      date: "2026-03-04T05:01:00Z",
    },
  });
  return { source, operations };
};

test("accepting a paragraph mark chain before a table is independent of resolution batches", () => {
  assertProperty(
    fc.property(
      fc.integer({ min: 2, max: 6 }),
      fc.constantFrom("body", "cell"),
      fc.shuffledSubarray([1, 2, 3, 4, 5, 6, 7], {
        minLength: 7,
        maxLength: 7,
      }),
      (length, container, order) => {
        const { source, operations } = fixture(length, container);
        let tracked = source;
        for (const operation of operations) tracked = apply(tracked, operation);
        const revisions = Array.from(
          { length: length + 1 },
          (_, index) => index + 1,
        );
        const together = resolve(tracked, revisions);
        let incremental = tracked;
        for (const revision of revisions.slice(1)) {
          incremental = resolve(incremental, [revision]);
        }
        incremental = resolve(incremental, [1]);
        expect(
          projectReview({ document: together, projection: "π′" }),
        ).toStrictEqual(
          projectReview({ document: incremental, projection: "π′" }),
        );
        let reordered = tracked;
        for (const revision of order.filter((id) => revisions.includes(id))) {
          reordered = resolve(reordered, [revision]);
        }
        expect(
          projectReview({ document: together, projection: "π′" }),
        ).toStrictEqual(
          projectReview({ document: reordered, projection: "π′" }),
        );
        expect(together.package.document.comments).toStrictEqual(
          source.package.document.comments,
        );
        expect(
          storyParagraphs(together.package.document).find(
            ({ paragraph }) => paragraph.paraId === "00000100",
          )?.paragraph.content,
        ).toStrictEqual([
          { type: "commentRangeStart", id: 7 },
          { type: "run", content: [{ type: "text", text: "x" }] },
        ]);
      },
    ),
    {
      numRuns: 20,
      id: "paragraph mark chains before tables preserve acceptance batch equivalence",
    },
  );
});
