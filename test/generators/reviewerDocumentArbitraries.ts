/** Documents a headless reviewer reads, for properties that compare its engines. */
import fc from "fast-check";

import type {
  BlockContent,
  Document,
  Paragraph,
  ParagraphContent,
} from "../../packages/docx-core/src/model/document";
import { reviewDocumentArbitrary } from "../../packages/docx-core/src/ops/__tests__/documentArbitraries";
import { withCoherentSyntheticReviewComments } from "./coherentReviewerComments";
import { packageDocumentArbitrary } from "./packageOperationArbitraries";

const coherentReviewDocumentArbitrary = reviewDocumentArbitrary.map(
  withCoherentSyntheticReviewComments,
);

/**
 * Every story a reviewer lists (headers, footers, notes, a comment thread,
 * content controls and a table), or a body with review marks of every shape
 * the review properties draw.
 */
export const reviewerDocumentArbitrary: fc.Arbitrary<Document> = fc.oneof(
  packageDocumentArbitrary,
  coherentReviewDocumentArbitrary,
);

/**
 * Range markers that a package holds only in pairs, by their pair. A bookmark
 * boundary without a partner is source content a package keeps, so the
 * generator's lone bookmarks stay.
 */
const PAIRED_MARKERS: Readonly<Record<string, string>> = {
  moveToRangeStart: "moveTo",
  moveToRangeEnd: "moveTo",
  moveFromRangeStart: "moveFrom",
  moveFromRangeEnd: "moveFrom",
};

/** A plain object: a Map, a Date or a buffer is kept as it is. */
const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" &&
  value !== null &&
  Object.getPrototypeOf(value) === Object.prototype;

/** The pair a marker belongs to and its id, or nothing for any other node. */
const markerOf = (node: unknown): { key: string; end: boolean } | undefined => {
  if (!isRecord(node) || typeof node["type"] !== "string") return undefined;
  const pair = PAIRED_MARKERS[node["type"]];
  if (pair === undefined) return undefined;
  return {
    key: `${pair}:${String(node["id"])}`,
    end: node["type"].endsWith("End"),
  };
};

/**
 * The document with every move range marker whose partner is missing left
 * out, so that the writer packages it: the review generator draws markers one
 * by one.
 */
const withPairedMarkers = (document: Document): Document => {
  const starts = new Set<string>();
  const ends = new Set<string>();
  const collect = (value: unknown): void => {
    if (Array.isArray(value)) {
      for (const item of value) collect(item);
      return;
    }
    if (!isRecord(value)) return;
    const marker = markerOf(value);
    if (marker) (marker.end ? ends : starts).add(marker.key);
    for (const child of Object.values(value)) collect(child);
  };
  const keep = (value: unknown): unknown => {
    if (Array.isArray(value))
      return value
        .filter((item) => {
          const marker = markerOf(item);
          return (
            marker === undefined ||
            (starts.has(marker.key) && ends.has(marker.key))
          );
        })
        .map(keep);
    if (!isRecord(value)) return value;
    return Object.fromEntries(
      Object.entries(value).map(([key, child]) => [key, keep(child)]),
    );
  };
  collect(document);
  // SAFETY: only whole marker nodes leave their arrays; every record keeps
  // its shape.
  return keep(document) as Document;
};

const PENDING_JOIN_HOLDINGS = [
  "nothing",
  "anchors",
  "bookmark",
  "words",
] as const;

type PendingJoinPiece =
  | {
      type: "paragraph";
      holds: (typeof PENDING_JOIN_HOLDINGS)[number];
      mark: "none" | "ins" | "del";
      changed: boolean;
    }
  | { type: "bookmark" | "altChunk" | "table" };

// Paragraphs without words and marks pending are drawn most often.
const pendingJoinPieceArbitrary: fc.Arbitrary<PendingJoinPiece> = fc.oneof(
  {
    weight: 3,
    arbitrary: fc.record({
      type: fc.constant("paragraph" as const),
      holds: fc.constantFrom(...PENDING_JOIN_HOLDINGS, "nothing", "anchors"),
      mark: fc.constantFrom("none" as const, "ins" as const, "del" as const),
      changed: fc.boolean(),
    }),
  },
  {
    weight: 2,
    arbitrary: fc.record({
      type: fc.constantFrom(
        "bookmark" as const,
        "altChunk" as const,
        "table" as const,
      ),
    }),
  },
);

const pendingJoinContent = (
  holds: (typeof PENDING_JOIN_HOLDINGS)[number],
): ParagraphContent[] => {
  switch (holds) {
    case "nothing":
      return [];
    case "anchors":
      return [
        { type: "commentRangeStart", id: 8 },
        { type: "commentRangeEnd", id: 8 },
      ];
    case "bookmark":
      return [
        { type: "bookmarkStart", id: 30, name: "_Ref30" },
        { type: "bookmarkEnd", id: 30 },
      ];
    case "words":
      return [{ type: "run", content: [{ type: "text", text: "Alpha beta" }] }];
  }
};

/**
 * Short bodies around pending paragraph joins: paragraphs that are empty, hold
 * only range boundaries or hold words, under pending insertions or deletions
 * of their marks and property changes, beside block-level bookmarks,
 * preserved blocks and a one-cell table. As a review leaves them, a paragraph
 * ending the body has no pending mark.
 */
export const pendingJoinDocumentArbitrary: fc.Arbitrary<Document> = fc
  .array(pendingJoinPieceArbitrary, { minLength: 2, maxLength: 5 })
  .map((pieces): Document => {
    let paragraphs = 0;
    let revisions = 0;
    const paraId = () =>
      (++paragraphs).toString(16).toUpperCase().padStart(8, "0");
    const info = () => ({ id: ++revisions, author: "A" });
    const content = pieces.map((piece, index): BlockContent => {
      switch (piece.type) {
        case "bookmark":
          return { type: "bookmarkStart", id: 90, name: "_GoBack" };
        case "altChunk":
          return { type: "preservedBlock", xml: "<w:altChunk/>" };
        case "table":
          return {
            type: "table",
            columnWidths: [2400],
            rows: [
              {
                type: "tableRow",
                cells: [
                  {
                    type: "tableCell",
                    content: [
                      { type: "paragraph", content: [], paraId: paraId() },
                    ],
                  },
                ],
              },
            ],
          };
        case "paragraph":
          break;
      }
      const paragraph: Paragraph = {
        type: "paragraph",
        content: pendingJoinContent(piece.holds),
        paraId: paraId(),
      };
      if (piece.changed)
        paragraph.propertyChanges = [
          {
            type: "paragraphPropertyChange",
            info: info(),
            previousFormatting: { alignment: "end" },
          },
        ];
      if (piece.mark !== "none" && index < pieces.length - 1)
        paragraph.pPrMark = { kind: piece.mark, info: info() };
      return paragraph;
    });
    return withCoherentSyntheticReviewComments({
      package: {
        document: {
          content,
          finalSectionProperties: { pageWidth: 12240, pageHeight: 15840 },
          comments: [
            {
              id: 7,
              author: "A",
              content: [{ type: "paragraph", paraId: "7FFFFFF0", content: [] }],
            },
          ],
        },
      },
    });
  });

/**
 * Bodies with review marks of every shape that the writer packages, for
 * properties about the body's edits and resolutions. Headers, footers, notes
 * and text boxes have their own story properties.
 */
export const bodyReviewerDocumentArbitrary: fc.Arbitrary<Document> =
  coherentReviewDocumentArbitrary.map(withPairedMarkers);
