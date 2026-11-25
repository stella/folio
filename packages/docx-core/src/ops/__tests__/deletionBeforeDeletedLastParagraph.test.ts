/**
 * Tracked deletions of a container's last paragraph and of the paragraph
 * before it, whose properties the last one's deletion gives it: one at a
 * time or batched, in either order, they record the same document. Deleting
 * the earlier paragraph second takes back the property change the first
 * deletion recorded. Accepting all equals deleting both directly; rejecting
 * all gives the document back; every inverse is exact.
 */
import { expect, test } from "bun:test";

import type { Document, Paragraph } from "../../model/document";
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

/** The paragraph before the item, if any, and its properties. */
const BEFORE = ["nothing", "none", "right-aligned", "centered"] as const;

const ITEMS = ["numbered", "right-aligned"] as const;

const CLOSINGS = ["none", "centered"] as const;

const DATE = "2026-02-03T04:05:06Z";

const MODES = {
  tracked: { author: "Reviewer", date: DATE },
  suggested: { author: "Reviewer", date: DATE, suggestionId: "suggestion" },
} as const satisfies Record<string, Omit<RevisionStamp, "id">>;

type DocumentOptions = {
  before: (typeof BEFORE)[number];
  item: FormattingName;
  closing: FormattingName;
};

const documentOf = ({ before, item, closing }: DocumentOptions): Document => ({
  package: {
    document: {
      content: [
        ...(before === "nothing"
          ? []
          : [paragraph(KEPT, "Kept.", FORMATTINGS[before])]),
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
const undone = (applied: { document: Document; inverse: DocumentOp[] }) =>
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
        if (key !== "id" || typeof field !== "number")
          return [key, walk(field)];
        if (!ids.has(field)) ids.set(field, ids.size + 1);
        return [key, ids.get(field)];
      }),
    );
  };
  return walk(document);
};

const CASES = BEFORE.flatMap((before) =>
  ITEMS.flatMap((item) =>
    CLOSINGS.flatMap((closing) =>
      Object.entries(MODES).map(([mode, stamp]) => ({
        before,
        item,
        closing,
        mode,
        stamp,
      })),
    ),
  ),
);

test.each(CASES)(
  "$mode: deleting a $item paragraph after $before and its $closing last paragraph reads alike in every order",
  ({ before, item, closing, stamp }) => {
    const document = documentOf({ before, item, closing });
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
    /** One at a time, each step undoing exactly. */
    const oneAtATime = (order: readonly string[]) => {
      let current = document;
      for (const blockId of order) {
        const step = apply(current, deletion([blockId]));
        expect(undone(step)).toEqual(current);
        current = step.document;
      }
      return current;
    };
    const batched = (order: readonly string[]) => {
      const batch = applyAll(
        document,
        order.map((blockId) => deletion([blockId])),
      );
      expect(undone(batch)).toEqual(document);
      return batch.document;
    };
    const together = apply(document, deletion([ITEM, CLOSING]));
    expect(undone(together)).toEqual(document);
    const results = [
      oneAtATime([CLOSING, ITEM]),
      oneAtATime([ITEM, CLOSING]),
      batched([CLOSING, ITEM]),
      batched([ITEM, CLOSING]),
    ];
    for (const result of results)
      expect(renumbered(result)).toEqual(renumbered(together.document));
    const direct = apply(document, {
      type: DOCUMENT_OP_TYPES.DELETE_BLOCKS,
      story: OP_STORIES.MAIN,
      blockIds: [ITEM, CLOSING],
    }).document;
    for (const result of results) {
      expect(resolve(result, REVISION_DECISIONS.ACCEPT)).toEqual(direct);
      expect(resolve(result, REVISION_DECISIONS.REJECT)).toEqual(document);
    }
  },
);

test.each([
  { first: "an ordinary tracked deletion", firstStamp: {} },
  { first: "another suggestion", firstStamp: { suggestionId: "other" } },
])(
  "a suggestion takes back no property change $first recorded",
  ({ firstStamp }) => {
    const closingDeleted = apply(
      documentOf({ before: "none", item: "numbered", closing: "centered" }),
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
