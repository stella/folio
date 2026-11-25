/**
 * Tracked deletions of a container's last paragraph and of the paragraph
 * before it, whose properties the last one's deletion gives it. Each order,
 * one at a time or batched, accepts to the same deletions made directly in
 * that order, rejects to the document, and undoes exactly. With a paragraph
 * before them, whose break the two retire, every order records the same
 * document: deleting the earlier paragraph second takes back the property
 * change the first deletion recorded. Back to the container's start or a
 * table, the run settles on its first paragraph when the last one goes first,
 * as it does directly.
 */
import { expect, test } from "bun:test";

import type { BlockContent, Document, Paragraph } from "../../model/document";
import {
  type AppliedDocumentOp,
  applyDocumentOp,
  applyDocumentOps,
} from "../apply";
import { IDENTITY_SPACES, packageIdentityKeys } from "../ids";
import { DOCUMENT_OP_REFUSAL_REASONS } from "../refusal";
import {
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  OP_STORIES,
  REVISION_DECISIONS,
  type RevisionDecision,
  type RevisionStamp,
} from "../types";

const KEPT = "00000001";
const ITEM = "00000002";
const CLOSING = "00000003";

const paragraph = (
  paraId: string,
  text: string,
  formatting: Paragraph["formatting"],
): Paragraph => ({
  type: "paragraph",
  paraId,
  content: [{ type: "run", content: [{ type: "text", text }] }],
  ...(formatting === undefined ? {} : { formatting }),
});

const FORMATTINGS = {
  none: undefined,
  centered: { alignment: "center" },
  numbered: { numPr: { kind: "reference", numId: 2, ilvl: 0 } },
  "right-aligned": { alignment: "end" },
} as const satisfies Record<string, Paragraph["formatting"]>;

type FormattingName = keyof typeof FORMATTINGS;

/** What stands before the item, and whether it is a paragraph whose break goes. */
const BEFORE = {
  nothing: { blocks: [], retiresBreak: false },
  "a table": {
    blocks: [
      {
        type: "table",
        rows: [
          {
            type: "tableRow",
            cells: [
              {
                type: "tableCell",
                content: [paragraph("00000009", "Cell.", undefined)],
              },
            ],
          },
        ],
      },
    ],
    retiresBreak: false,
  },
  "a plain paragraph": {
    blocks: [paragraph(KEPT, "Kept.", undefined)],
    retiresBreak: true,
  },
  "a right-aligned paragraph": {
    blocks: [paragraph(KEPT, "Kept.", FORMATTINGS["right-aligned"])],
    retiresBreak: true,
  },
  "a centered paragraph": {
    blocks: [paragraph(KEPT, "Kept.", FORMATTINGS.centered)],
    retiresBreak: true,
  },
} as const satisfies Record<
  string,
  { blocks: readonly BlockContent[]; retiresBreak: boolean }
>;

const ITEMS = ["numbered", "right-aligned"] as const;

const CLOSINGS = ["none", "centered"] as const;

const DATE = "2026-02-03T04:05:06Z";

const MODES = {
  tracked: { author: "Reviewer", date: DATE },
  suggested: { author: "Reviewer", date: DATE, suggestionId: "suggestion" },
} as const satisfies Record<string, Omit<RevisionStamp, "id">>;

type DocumentOptions = {
  before: readonly BlockContent[];
  item: FormattingName;
  closing: FormattingName;
};

const documentOf = ({ before, item, closing }: DocumentOptions): Document => ({
  package: {
    document: {
      content: [
        ...structuredClone(before),
        paragraph(ITEM, "Item.", FORMATTINGS[item]),
        paragraph(CLOSING, "Closing.", FORMATTINGS[closing]),
      ],
    },
  },
});

const apply = (document: Document, op: DocumentOp): AppliedDocumentOp => {
  const result = applyDocumentOp(document, op);
  if (result.isErr()) throw result.error;
  return result.value;
};

const applyAll = (document: Document, ops: readonly DocumentOp[]) => {
  const result = applyDocumentOps(document, ops);
  if (result.isErr()) throw result.error;
  return result.value;
};

/** Each inverse, as a client stores and replays it. */
const undone = (applied: AppliedDocumentOp) =>
  applyAll(applied.document, JSON.parse(JSON.stringify(applied.inverse)))
    .document;

const revisionIdsIn = (document: Document): number[] => {
  const prefix = `${IDENTITY_SPACES.REVISION}:`;
  return packageIdentityKeys(document.package)
    .filter((key) => key.startsWith(prefix))
    .map((key) => Number(key.slice(prefix.length)));
};

const resolve = (document: Document, decision: RevisionDecision) =>
  apply(document, {
    type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
    story: OP_STORIES.MAIN,
    revisionIds: revisionIdsIn(document),
    decision,
  }).document;

/**
 * The document with its revision ids numbered by first appearance: which
 * operation allocated a record's id depends on the order the deletions ran.
 */
const renumbered = (document: Document): unknown => {
  const ids = new Map<number, number>();
  const walk = (value: unknown): unknown => {
    if (Array.isArray(value)) return value.map(walk);
    if (typeof value !== "object" || value === null) return value;
    return Object.fromEntries(
      Object.entries(value).map(([key, field]) => {
        if (
          (key !== "id" && key !== "resolvesWith") ||
          typeof field !== "number"
        )
          return [key, walk(field)];
        if (!ids.has(field)) ids.set(field, ids.size + 1);
        return [key, ids.get(field)];
      }),
    );
  };
  return walk(document);
};

const CASES = Object.entries(BEFORE).flatMap(([before, place]) =>
  ITEMS.flatMap((item) =>
    CLOSINGS.flatMap((closing) =>
      Object.entries(MODES).map(([mode, stamp]) => ({
        before,
        place,
        item,
        closing,
        mode,
        stamp,
      })),
    ),
  ),
);

/** The deletion orders, by the ids each step names once the earlier ran directly. */
const ORDERS = {
  "the last paragraph first": {
    tracked: [CLOSING, ITEM],
    direct: [CLOSING, CLOSING],
  },
  "the item first": { tracked: [ITEM, CLOSING], direct: [ITEM, CLOSING] },
} as const;

test.each(CASES)(
  "$mode: deleting a $item paragraph after $before and its $closing last paragraph resolves as directly, in either order",
  ({ place, item, closing, stamp }) => {
    const document = documentOf({ before: place.blocks, item, closing });
    let next = 100;
    const deletion = (blockIds: readonly string[]): DocumentOp => {
      const id = next;
      next += 10;
      return {
        type: DOCUMENT_OP_TYPES.DELETE_BLOCKS,
        story: OP_STORIES.MAIN,
        blockIds: [...blockIds],
        revision: { ...stamp, id },
        newIds: {
          revision: Array.from({ length: 9 }, (_, index) => id + 1 + index),
        },
      };
    };
    const recorded: unknown[] = [];
    for (const order of Object.values(ORDERS)) {
      // One at a time, each step undoing exactly.
      let tracked = document;
      for (const blockId of order.tracked) {
        const step = apply(tracked, deletion([blockId]));
        expect(undone(step)).toEqual(tracked);
        tracked = step.document;
      }
      const batch = applyAll(
        document,
        order.tracked.map((blockId) => deletion([blockId])),
      );
      expect(undone(batch)).toEqual(document);
      expect(renumbered(batch.document)).toEqual(renumbered(tracked));
      let direct = document;
      for (const blockId of order.direct)
        direct = apply(direct, {
          type: DOCUMENT_OP_TYPES.DELETE_BLOCKS,
          story: OP_STORIES.MAIN,
          blockIds: [blockId],
        }).document;
      expect(resolve(tracked, REVISION_DECISIONS.ACCEPT)).toEqual(direct);
      expect(resolve(tracked, REVISION_DECISIONS.REJECT)).toEqual(document);
      recorded.push(renumbered(tracked));
    }
    // With a break before them to retire, the order does not matter, and
    // both deletions at once record the same.
    if (!place.retiresBreak) return;
    const together = apply(document, deletion([ITEM, CLOSING]));
    expect(undone(together)).toEqual(document);
    for (const result of recorded)
      expect(result).toEqual(renumbered(together.document));
  },
);

test.each([
  { first: "an ordinary tracked deletion", firstStamp: {} },
  { first: "another suggestion", firstStamp: { suggestionId: "other" } },
])(
  "a suggestion takes back no property change $first recorded",
  ({ firstStamp }) => {
    const closingDeleted = apply(
      documentOf({
        before: BEFORE["a plain paragraph"].blocks,
        item: "numbered",
        closing: "centered",
      }),
      {
        type: DOCUMENT_OP_TYPES.DELETE_BLOCKS,
        story: OP_STORIES.MAIN,
        blockIds: [CLOSING],
        revision: { id: 100, author: "Other", date: DATE, ...firstStamp },
        newIds: { revision: [101, 102, 103] },
      },
    ).document;
    const refused = applyDocumentOp(closingDeleted, {
      type: DOCUMENT_OP_TYPES.DELETE_BLOCKS,
      story: OP_STORIES.MAIN,
      blockIds: [ITEM],
      revision: { id: 110, ...MODES.suggested },
      newIds: { revision: [111, 112, 113, 114] },
    });
    expect(refused.isErr() && refused.error.reason).toBe(
      DOCUMENT_OP_REFUSAL_REASONS.PENDING_PARAGRAPH_PROPERTY_CHANGE,
    );
  },
);

test("deleting before the already deleted final paragraph keeps its unrelated property revision", () => {
  const document = documentOf({
    before: BEFORE["a plain paragraph"].blocks,
    item: "right-aligned",
    closing: "right-aligned",
  });
  const closing = document.package.document.content.at(-1);
  if (closing?.type !== "paragraph")
    throw new Error("Expected final paragraph");
  const propertyChange = {
    type: "paragraphPropertyChange",
    info: { id: 50, ...MODES.tracked },
    previousFormatting: FORMATTINGS.centered,
    currentFormatting: FORMATTINGS["right-aligned"],
  } satisfies NonNullable<Paragraph["propertyChanges"]>[number];
  closing.propertyChanges = [propertyChange];
  const closingDeleted = apply(document, {
    type: DOCUMENT_OP_TYPES.DELETE_BLOCKS,
    story: OP_STORIES.MAIN,
    blockIds: [CLOSING],
    revision: { id: 100, ...MODES.tracked },
    newIds: { revision: [101, 102, 103] },
  }).document;

  const deleted = apply(closingDeleted, {
    type: DOCUMENT_OP_TYPES.DELETE_BLOCKS,
    story: OP_STORIES.MAIN,
    blockIds: [ITEM],
    revision: { id: 110, ...MODES.tracked },
    newIds: { revision: [111, 112, 113, 114] },
  });
  expect(deleted.document.package.document.content.at(-1)).toMatchObject({
    paraId: CLOSING,
    formatting: FORMATTINGS["right-aligned"],
    propertyChanges: [propertyChange],
  });
  expect(deleted.document.package.document.content.at(0)).toMatchObject({
    paraId: KEPT,
    pPrMark: { kind: "del", info: { id: 110 } },
  });
  expect(undone(deleted)).toEqual(closingDeleted);
});
