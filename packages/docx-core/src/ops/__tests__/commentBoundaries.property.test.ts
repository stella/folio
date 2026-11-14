/**
 * A comment whose range opens or closes between blocks (between paragraphs,
 * beside a content control, at the start or end of a table cell or of the
 * story, beside another comment's marker) takes every comment operation:
 * removal, a reply, resolution, and the exact inverse of each, which also
 * crosses the wire.
 */
import { expect, test } from "bun:test";
import fc from "fast-check";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../../test/property-testing";
import { setDocumentBodyComments } from "../../model/canonical";
import type {
  BlockContent,
  Comment,
  Document,
  Paragraph,
  ParagraphContent,
} from "../../model/document";
import { applyDocumentOp, applyDocumentOps } from "../apply";
import { commentAnchorsIn, commentDocumentIssue } from "../comments";
import { validateDocumentBatch } from "../sequencing/envelope";
import {
  DOCUMENT_OP_SCHEMA_VERSION,
  DOCUMENT_OP_TYPES,
  type DocumentOp,
} from "../types";

const START_PLACEMENTS = [
  "inline",
  "betweenParagraphs",
  "besideControl",
  "cellStart",
  "storyStart",
] as const;
const END_PLACEMENTS = [
  "inline",
  "betweenParagraphs",
  "cellEnd",
  "storyEnd",
] as const;

type Layout = {
  start: (typeof START_PLACEMENTS)[number];
  end: (typeof END_PLACEMENTS)[number];
  /** Another comment's start between blocks beside this one's. */
  neighbour: boolean;
};

const run = (text: string): ParagraphContent => ({
  type: "run",
  content: [{ type: "text", text }],
});
const paragraph = (paraId: string, content: ParagraphContent[]): Paragraph => ({
  type: "paragraph",
  paraId,
  content,
});
const comment = (id: number, paraId: string): Comment => ({
  id,
  author: "Reviewer",
  date: "2024-01-01T00:00:00Z",
  content: [paragraph(paraId, [run(`Comment ${id}`)])],
});

const documentFor = ({ start, end, neighbour }: Layout): Document => {
  const startBlock: BlockContent[] =
    start === "inline"
      ? []
      : [
          start === "besideControl"
            ? {
                type: "commentRangeStart",
                id: 1,
                displacedByCustomXml: "next",
              }
            : { type: "commentRangeStart", id: 1 },
        ];
  const endBlock: BlockContent[] =
    end === "inline" ? [] : [{ type: "commentRangeEnd", id: 1 }];
  const other: BlockContent[] = neighbour
    ? [{ type: "commentRangeStart", id: 2 }]
    : [];
  const commented = paragraph("00000002", [
    ...(start === "inline"
      ? [{ type: "commentRangeStart", id: 1 } as const]
      : []),
    run("Please remit funds."),
    ...(end === "inline" ? [{ type: "commentRangeEnd", id: 1 } as const] : []),
    { type: "commentReference", id: 1 },
    ...(neighbour
      ? ([
          { type: "commentRangeEnd", id: 2 },
          { type: "commentReference", id: 2 },
        ] as const)
      : []),
  ]);
  const inner: BlockContent =
    start === "besideControl"
      ? { type: "blockSdt", properties: {}, content: [commented] }
      : commented;
  const inCell = start === "cellStart" || end === "cellEnd";
  const group: BlockContent[] = inCell
    ? [
        ...(start === "cellStart" ? [] : [...other, ...startBlock]),
        {
          type: "table",
          columnWidths: [2400, 2400],
          rows: [
            {
              type: "tableRow",
              cells: [
                {
                  type: "tableCell",
                  content: [
                    ...(start === "cellStart" ? [...other, ...startBlock] : []),
                    inner,
                    ...(end === "cellEnd" ? endBlock : []),
                  ],
                },
                {
                  type: "tableCell",
                  content: [paragraph("00000004", [run("Beside")])],
                },
              ],
            },
          ],
        },
        ...(end === "cellEnd" ? [] : endBlock),
      ]
    : [...other, ...startBlock, inner, ...endBlock];
  const document: Document = {
    package: {
      document: {
        content: [
          ...(start === "storyStart"
            ? []
            : [paragraph("00000001", [run("Opening words")])]),
          ...group,
          ...(end === "storyEnd"
            ? []
            : [paragraph("00000003", [run("Closing words")])]),
        ],
      },
    },
  };
  setDocumentBodyComments(document.package.document, [
    comment(1, "00000011"),
    ...(neighbour ? [comment(2, "00000012")] : []),
  ]);
  return document;
};

const layoutArbitrary: fc.Arbitrary<Layout> = fc
  .record({
    start: fc.constantFrom(...START_PLACEMENTS),
    end: fc.constantFrom(...END_PLACEMENTS),
    neighbour: fc.boolean(),
  })
  // The story's own start and a cell's start are the same marker's two places.
  .filter(({ start, end }) => !(start === "storyStart" && end === "cellEnd"));

const anchorsOf = (document: Document, id: number) =>
  commentAnchorsIn(document)
    .filter(({ node }) => "id" in node && node.id === id)
    .map(({ node, blockId, betweenBlocks }) => ({
      type: node.type,
      blockId,
      betweenBlocks: betweenBlocks === true,
    }));

/** Every text of the document's stories, in order. */
const texts = (document: Document): string[] =>
  JSON.stringify(document.package.document.content)
    .match(/"text":"[^"]*"/gu)
    ?.map((text) => text) ?? [];

/** The inverse crosses the wire as a batch of its own. */
const expectSendable = (ops: readonly DocumentOp[]) =>
  expect(
    validateDocumentBatch(
      JSON.parse(
        JSON.stringify({
          schema: DOCUMENT_OP_SCHEMA_VERSION,
          opId: "inverse",
          actor: "peer",
          baseRev: 0,
          ops,
        }),
      ),
    ).isOk(),
  ).toBe(true);

const apply = (document: Document, op: DocumentOp) => {
  const result = applyDocumentOp(document, op);
  if (result.isErr()) throw result.error;
  expect(commentDocumentIssue(result.value.document)).toBeUndefined();
  expectSendable(result.value.inverse);
  const undone = applyDocumentOps(
    result.value.document,
    result.value.inverse,
  ).unwrap();
  expect(undone.document).toStrictEqual(document);
  return result.value;
};

test(
  "a comment opening or closing between blocks is removed, replied to, resolved and restored exactly",
  async () => {
    await assertProperty(
      fc.asyncProperty(layoutArbitrary, async (layout) => {
        const document = documentFor(layout);
        expect(commentDocumentIssue(document)).toBeUndefined();
        const others = anchorsOf(document, 2);

        // Removal takes every marker of the thread, between blocks too.
        const removed = apply(document, {
          type: DOCUMENT_OP_TYPES.DELETE_COMMENT,
          id: 1,
          scope: "thread",
        });
        expect(anchorsOf(removed.document, 1)).toEqual([]);
        expect(anchorsOf(removed.document, 2)).toEqual(others);
        expect(texts(removed.document)).toEqual(texts(document));

        // A reply's markers stand in paragraphs, beside the owner's.
        const replied = apply(document, {
          type: DOCUMENT_OP_TYPES.CREATE_COMMENT,
          comment: comment(3, "00000013"),
          anchor: { kind: "reply", parentId: 1 },
        });
        expect(anchorsOf(replied.document, 3).map(({ type }) => type)).toEqual([
          "commentRangeStart",
          "commentRangeEnd",
          "commentReference",
        ]);
        expect(
          anchorsOf(replied.document, 3).some(
            ({ betweenBlocks }) => betweenBlocks,
          ),
        ).toBe(false);
        expect(anchorsOf(replied.document, 1)).toEqual(anchorsOf(document, 1));

        // The whole thread goes with its root, and comes back with it.
        const thread = apply(replied.document, {
          type: DOCUMENT_OP_TYPES.DELETE_COMMENT,
          id: 1,
          scope: "thread",
        });
        expect(anchorsOf(thread.document, 1)).toEqual([]);
        expect(anchorsOf(thread.document, 3)).toEqual([]);

        // The reply alone goes, the owner's markers stay.
        const reply = apply(replied.document, {
          type: DOCUMENT_OP_TYPES.DELETE_COMMENT,
          id: 3,
          scope: "reply",
        });
        expect(reply.document).toStrictEqual(document);

        apply(document, {
          type: DOCUMENT_OP_TYPES.SET_COMMENT_RESOLUTION,
          id: 1,
          status: "resolved",
        });

        // Redo after undo gives the removal again.
        const undone = applyDocumentOps(
          removed.document,
          removed.inverse,
        ).unwrap();
        expect(
          applyDocumentOps(undone.document, undone.inverse).unwrap().document,
        ).toStrictEqual(removed.document);
      }),
      {
        id: "a comment opening or closing between blocks is removed, replied to, resolved and restored exactly",
        examples: [
          [{ start: "betweenParagraphs", end: "inline", neighbour: false }],
          [{ start: "besideControl", end: "inline", neighbour: false }],
          [{ start: "cellStart", end: "cellEnd", neighbour: true }],
          [{ start: "storyStart", end: "storyEnd", neighbour: true }],
        ],
      },
    );
  },
  propertyTestTimeout(30_000),
);
