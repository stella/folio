import { expect, test } from "bun:test";

import {
  type BlockContent,
  type Document,
  type Paragraph,
  type ParagraphMarkChange,
  type ParagraphMarkDeletion,
  type ParagraphPropertyChange,
  REVIEW_CARRIERS,
} from "../../model/document";
import {
  type AppliedDocumentOp,
  applyDocumentOp,
  applyDocumentOps,
} from "../apply";
import {
  DOCUMENT_OP_REFUSAL_REASONS,
  type DocumentOpRefusalReason,
} from "../refusal";
import {
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  type InsertBlocksOp,
  OP_STORIES,
  REVISION_DECISIONS,
  type RevisionDecision,
} from "../types";

// A tracked paragraph put after a paragraph that stands before a block
// content control takes over that paragraph's mark, and the review records on
// the mark go with it, as a tracked split's second half takes them. Before a
// table whose first row is pending it holds its own inserted mark and leaves
// the paragraph before it as it was. A break pending deletion or moving away
// is not taken over.

const OTHER = { author: "Other", date: "2026-01-01T00:00:00Z" };
const stamp = { id: 100, author: "Reviewer", date: "2026-02-03T04:05:06Z" };
const PENDING_CHANGE = 7;
const PENDING_MARK = 8;
const PENDING_MARK_DELETION = 9;
const PENDING_TEXT = 11;
const PENDING_ROW = 20;
const PENDING_CELL = 21;
const TABLE = [PENDING_ROW, PENDING_CELL];

const run = (text: string) => ({
  type: "run" as const,
  content: [{ type: "text" as const, text }],
});

const paragraph = (paraId: string, text: string): Paragraph => ({
  type: "paragraph",
  paraId,
  content: text === "" ? [] : [run(text)],
});

// A pending first row: the anchor's mark cannot run into the table, and the
// new paragraph's own mark goes without crossing it.
const table: BlockContent = {
  type: "table",
  rows: [
    {
      type: "tableRow",
      structuralChange: {
        type: "tableRowInsertion",
        info: { id: PENDING_ROW, ...OTHER },
      },
      cells: [
        {
          type: "tableCell",
          content: [
            {
              ...paragraph("00000010", ""),
              content: [
                {
                  type: "insertion",
                  info: { id: PENDING_CELL, ...OTHER },
                  content: [run("cell")],
                },
              ],
            },
          ],
        },
      ],
    },
  ],
};

const control: BlockContent = {
  type: "blockSdt",
  properties: { sdtType: "richText" },
  content: [paragraph("00000010", "controlled")],
};

/** The block after the anchor, and whose mark a new paragraph there holds. */
const FOLLOWERS = [
  { name: "a table", follower: table, takesOver: false, records: TABLE },
  { name: "a control", follower: control, takesOver: true, records: [] },
] as const;

const propertyChange: ParagraphPropertyChange = {
  type: "paragraphPropertyChange",
  info: { id: PENDING_CHANGE, ...OTHER },
  previousFormatting: { alignment: "start" },
};

const insertedMark: ParagraphMarkChange = {
  kind: "ins",
  info: { id: PENDING_MARK, ...OTHER },
};

const markDeletion: ParagraphMarkDeletion = {
  info: { id: PENDING_MARK_DELETION, ...OTHER },
};

type AnchorCase = {
  name: string;
  anchor: Paragraph;
  /** The anchor's own pending ids, which this insertion does not take. */
  pending: readonly number[];
  /** The mark of the paragraph that takes over the anchor's. */
  taker?: ParagraphMarkChange;
};

const ANCHOR = {
  ...paragraph("00000001", "kept"),
  formatting: { alignment: "center" },
} satisfies Paragraph;

const ANCHORS: readonly AnchorCase[] = [
  {
    name: "a pending property change",
    anchor: { ...ANCHOR, propertyChanges: [propertyChange] },
    pending: [PENDING_CHANGE],
  },
  {
    name: "a pending inserted break",
    anchor: {
      ...ANCHOR,
      content: [
        {
          type: "insertion",
          info: { id: PENDING_TEXT, ...OTHER },
          content: [run("kept")],
        },
      ],
      pPrMark: insertedMark,
    },
    pending: [PENDING_TEXT, PENDING_MARK],
    taker: insertedMark,
  },
  {
    name: "a pending property change and inserted break",
    anchor: {
      ...ANCHOR,
      propertyChanges: [propertyChange],
      pPrMark: insertedMark,
    },
    pending: [PENDING_CHANGE, PENDING_MARK],
    taker: insertedMark,
  },
];

const documentOf = (
  anchor: Paragraph,
  follower: BlockContent = table,
): Document => ({
  package: {
    document: { content: [anchor, follower, paragraph("00000004", "after")] },
  },
});

const insertion = (inserted: Paragraph): InsertBlocksOp => ({
  type: DOCUMENT_OP_TYPES.INSERT_BLOCKS,
  story: OP_STORIES.MAIN,
  at: { type: "after", blockId: "00000001" },
  blocks: [inserted],
  revision: stamp,
  newIds: { revision: Array.from({ length: 32 }, (_, index) => 101 + index) },
});

const apply = (document: Document, op: DocumentOp) => {
  const result = applyDocumentOp(document, op);
  if (result.isErr()) throw result.error;
  return result.value;
};

const resolve = (
  document: Document,
  revisionIds: readonly number[],
  decision: RevisionDecision,
) =>
  apply(document, {
    type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
    story: OP_STORIES.MAIN,
    revisionIds,
    decision,
  }).document;

const undo = (applied: AppliedDocumentOp) => {
  const restored = applyDocumentOps(applied.document, applied.inverse);
  if (restored.isErr()) throw restored.error;
  return restored.value.document;
};

/**
 * The document without its paragraph ids. Rejecting an added break the new
 * paragraph took over together with the anchor's new one leaves the joined
 * paragraph under the anchor's id or the next one's, by which goes first, as
 * after a tracked split at the anchor's end.
 */
const withoutParagraphIds = (document: Document): unknown =>
  JSON.parse(
    JSON.stringify(document, (key, value: unknown) =>
      key === "paraId" ? undefined : value,
    ),
  );

const DECISIONS = Object.values(REVISION_DECISIONS);

const FORMATTINGS = ["the anchor's", "other"] as const;

test.each(
  FOLLOWERS.flatMap((place) =>
    ANCHORS.flatMap((source) =>
      FORMATTINGS.map((formatting) => ({
        name: `before ${place.name}, after one with ${source.name}, with ${formatting} properties,`,
        place,
        source,
        formatting,
      })),
    ),
  ),
)(
  "a tracked paragraph $name reviews on its own",
  ({ place, source, formatting }) => {
    const { follower, takesOver, records } = place;
    const { anchor, pending, taker } = source;
    const document = documentOf(anchor, follower);
    const inserted = {
      ...paragraph("00000002", "new"),
      formatting:
        formatting === "other"
          ? { alignment: "end" as const }
          : ANCHOR.formatting,
    } satisfies Paragraph;
    const op = insertion(inserted);
    const tracked = apply(document, op);
    const [first, second] = tracked.document.package.document.content;
    if (first?.type !== "paragraph" || second?.type !== "paragraph")
      throw new Error("The anchor and the new paragraph stand first.");
    if (takesOver) {
      // The anchor's mark is the new one; the new paragraph has its former
      // mark and the records it held.
      expect(first.pPrMark).toEqual({
        kind: "ins",
        info: { ...stamp, id: first.pPrMark?.info.id ?? -1 },
        survivor: "first",
        ...(taker === undefined ? {} : { takenOverMark: taker.info.id }),
      });
      expect(first.propertyChanges).toBeUndefined();
      expect(second.pPrMark).toEqual(taker);
      const changes = second.propertyChanges;
      if (anchor.propertyChanges === undefined)
        expect(changes === undefined).toBe(formatting !== "other");
      else
        // One change from the anchor's original properties, under its id.
        expect(
          changes?.map(({ info, previousFormatting }) => [
            info.id,
            previousFormatting,
          ]),
        ).toEqual([[PENDING_CHANGE, { alignment: "start" }]]);
    } else {
      expect(first).toEqual(anchor);
      expect(second).toEqual({
        ...inserted,
        content: [
          {
            type: "insertion",
            info: { ...stamp, id: expect.any(Number) },
            content: [run("new")],
          },
        ],
        pPrMark: { kind: "ins", info: { ...stamp, id: expect.any(Number) } },
      });
    }
    expect(undo(tracked)).toEqual(document);
    const own = tracked.revisions.filter((id) => !pending.includes(id));
    // Rejecting the insertion alone gives the anchor back with its records,
    // unless other properties folded into the pending change it took over.
    const folded =
      takesOver &&
      formatting === "other" &&
      anchor.propertyChanges !== undefined;
    if (!folded)
      expect(resolve(tracked.document, own, REVISION_DECISIONS.REJECT)).toEqual(
        document,
      );
    const direct = apply(document, { ...op, revision: undefined }).document;
    const theirs = [...pending, ...records];
    const addedBreak = taker !== undefined;
    for (const decision of DECISIONS) {
      const accepting = decision === REVISION_DECISIONS.ACCEPT;
      const expected = resolve(accepting ? direct : document, theirs, decision);
      const same = (resolved: Document) =>
        accepting || !addedBreak
          ? expect(resolved).toEqual(expected)
          : expect(withoutParagraphIds(resolved)).toEqual(
              withoutParagraphIds(expected),
            );
      // All at once, the anchor's records first, or the insertion first.
      same(resolve(tracked.document, [...own, ...theirs], decision));
      same(resolve(resolve(tracked.document, theirs, decision), own, decision));
      same(resolve(resolve(tracked.document, own, decision), theirs, decision));
    }
  },
);

test("a new final paragraph does not take a pending property change", () => {
  const refused = applyDocumentOp(
    {
      package: {
        document: {
          content: [
            paragraph("00000004", "before"),
            { ...ANCHOR, propertyChanges: [propertyChange] },
          ],
        },
      },
    },
    insertion(paragraph("00000002", "new")),
  );
  expect(refused.isErr() && refused.error.reason).toBe(
    DOCUMENT_OP_REFUSAL_REASONS.PENDING_PARAGRAPH_PROPERTY_CHANGE,
  );
});

/** A story whose last paragraph is `anchor`. */
const endingWith = (anchor: Paragraph): Document => ({
  package: {
    document: { content: [paragraph("00000004", "before"), anchor] },
  },
});

test.each([
  [insertedMark, DOCUMENT_OP_REFUSAL_REASONS.PENDING_PARAGRAPH_MARK_INSERTION],
  [
    { kind: "del", info: insertedMark.info },
    DOCUMENT_OP_REFUSAL_REASONS.PENDING_PARAGRAPH_MARK_DELETION,
  ],
  [
    { kind: "moveFrom", info: insertedMark.info },
    DOCUMENT_OP_REFUSAL_REASONS.PENDING_PARAGRAPH_MARK_DELETION,
  ],
  [
    { ...insertedMark, deletion: markDeletion },
    DOCUMENT_OP_REFUSAL_REASONS.PENDING_PARAGRAPH_MARK_DELETION,
  ],
] satisfies [ParagraphMarkChange, DocumentOpRefusalReason][])(
  "a new final paragraph does not take a mark under review (%o), refused by what it has pending",
  (pPrMark, reason) => {
    const refused = applyDocumentOp(
      endingWith({ ...ANCHOR, pPrMark }),
      insertion(paragraph("00000002", "new")),
    );
    expect(refused.isErr() && refused.error.reason).toBe(reason);
  },
);

test("a new final paragraph does not take a table's review carrier yet", () => {
  const refused = applyDocumentOp(
    endingWith({ ...ANCHOR, reviewCarrier: REVIEW_CARRIERS.TERMINAL_TABLE }),
    insertion(paragraph("00000002", "new")),
  );
  expect(refused.isErr() && refused.error.reason).toBe(
    DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
  );
});

test.each([
  { kind: "del", info: insertedMark.info },
  { kind: "moveFrom", info: insertedMark.info },
  { ...insertedMark, deletion: markDeletion },
] satisfies ParagraphMarkChange[])(
  "a break pending deletion or moving away is not taken over ($kind)",
  (pPrMark) => {
    const refused = applyDocumentOp(
      documentOf({ ...ANCHOR, pPrMark }),
      insertion(paragraph("00000002", "new")),
    );
    expect(refused.isErr() && refused.error.reason).toBe(
      DOCUMENT_OP_REFUSAL_REASONS.PENDING_BREAK_TAKEOVER,
    );
  },
);

test("a suggestion does not fold other properties into another's pending change", () => {
  const refused = applyDocumentOp(
    documentOf({ ...ANCHOR, propertyChanges: [propertyChange] }, control),
    {
      ...insertion({
        ...paragraph("00000002", "new"),
        formatting: { alignment: "end" },
      }),
      revision: { ...stamp, suggestionId: "suggestion" },
    },
  );
  expect(refused.isErr() && refused.error.reason).toBe(
    DOCUMENT_OP_REFUSAL_REASONS.PENDING_PARAGRAPH_PROPERTY_CHANGE,
  );
});
