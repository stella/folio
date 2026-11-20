/**
 * Generated laws for paragraph marks holding an insertion and a later
 * deletion of it: resolving their records, at once or one at a time, in any
 * order, gives one document, and every resolution is undone exactly.
 */
import { expect, test } from "bun:test";
import fc from "fast-check";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../../test/property-testing";
import type { Document, Paragraph } from "../../model/document";
import { applyDocumentOp, applyDocumentOps } from "../apply";
import {
  DOCUMENT_OP_TYPES,
  OP_STORIES,
  REVISION_DECISIONS,
  type RevisionDecision,
} from "../types";

type MarkShape = "plain" | "ins" | "insDel" | "moveToDel" | "del";

const documentOf = (marks: readonly MarkShape[]): Document => {
  let id = 0;
  const info = (author: string) => ({
    id: ++id,
    author,
    date: `2026-01-0${(id % 9) + 1}T00:00:00Z`,
  });
  const content: Paragraph[] = marks.map((shape, index) => {
    const paragraph: Paragraph = {
      type: "paragraph",
      paraId: (0x10000001 + index).toString(16).toUpperCase(),
      content: [
        { type: "run", content: [{ type: "text", text: `P${index}` }] },
      ],
    };
    if (shape === "plain") return paragraph;
    if (shape === "del")
      return { ...paragraph, pPrMark: { kind: "del", info: info("B") } };
    const kind = shape === "moveToDel" ? "moveTo" : "ins";
    return {
      ...paragraph,
      pPrMark: {
        kind,
        info: info("A"),
        ...(shape === "ins" ? {} : { deletion: { info: info("B") } }),
      },
    };
  });
  // The story's last paragraph ends the body and holds no mark.
  content.push({
    type: "paragraph",
    paraId: "1FFFFFFF",
    content: [{ type: "run", content: [{ type: "text", text: "End" }] }],
  });
  return { package: { document: { content } } };
};

const markIds = (document: Document): number[] =>
  document.package.document.content.flatMap((block) =>
    block.type === "paragraph" && block.pPrMark !== undefined
      ? [
          block.pPrMark.info.id,
          ...(block.pPrMark.deletion ? [block.pPrMark.deletion.info.id] : []),
        ]
      : [],
  );

const resolve = (
  document: Document,
  revisionIds: readonly number[],
  decision: RevisionDecision,
): Document => {
  const applied = applyDocumentOp(document, {
    type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
    story: OP_STORIES.MAIN,
    revisionIds: [...revisionIds],
    decision,
  });
  if (applied.isErr()) throw applied.error;
  const undo = applyDocumentOps(applied.value.document, applied.value.inverse);
  if (undo.isErr()) throw undo.error;
  expect(undo.value.document).toStrictEqual(document);
  return applied.value.document;
};

const texts = (document: Document): string[] =>
  document.package.document.content.map((block) =>
    block.type === "paragraph"
      ? block.content
          .flatMap((item) => (item.type === "run" ? item.content : []))
          .map((item) => (item.type === "text" ? item.text : ""))
          .join("")
      : block.type,
  );

const marksArbitrary = fc.array(
  fc.constantFrom<MarkShape>("plain", "ins", "insDel", "moveToDel", "del"),
  { minLength: 1, maxLength: 5 },
);

test(
  "resolving a run of marks' records at once equals resolving them one at a time, in any order",
  () => {
    assertProperty(
      fc.property(
        marksArbitrary,
        fc.constantFrom(REVISION_DECISIONS.ACCEPT, REVISION_DECISIONS.REJECT),
        fc.array(fc.nat(), { minLength: 12, maxLength: 12 }),
        (marks, decision, keys) => {
          const document = documentOf(marks);
          const ids = markIds(document);
          const atOnce = resolve(document, ids, decision);
          const order = ids
            .map((id, index) => ({ id, key: keys[index] ?? index }))
            .toSorted((left, right) => left.key - right.key)
            .map(({ id }) => id);
          let oneByOne = document;
          for (const id of order)
            if (markIds(oneByOne).includes(id))
              oneByOne = resolve(oneByOne, [id], decision);
          expect(texts(oneByOne)).toEqual(texts(atOnce));
          expect(markIds(atOnce)).toEqual([]);
        },
      ),
      {
        id: "two-record marks resolve at once as one at a time",
        numRuns: 60,
      },
    );
  },
  propertyTestTimeout(30_000),
);

test(
  "accepting and rejecting a run's records in two batches, either order first, agree",
  () => {
    assertProperty(
      fc.property(
        marksArbitrary,
        fc.array(fc.boolean(), { minLength: 12, maxLength: 12 }),
        (marks, accepts) => {
          const document = documentOf(marks);
          const ids = markIds(document);
          const accepted = ids.filter((_, index) => accepts[index] ?? true);
          const rejected = ids.filter((_, index) => !(accepts[index] ?? true));
          const acceptedFirst = resolve(
            resolve(document, accepted, REVISION_DECISIONS.ACCEPT),
            rejected.filter((id) =>
              markIds(
                resolve(document, accepted, REVISION_DECISIONS.ACCEPT),
              ).includes(id),
            ),
            REVISION_DECISIONS.REJECT,
          );
          const afterRejecting = resolve(
            document,
            rejected,
            REVISION_DECISIONS.REJECT,
          );
          const rejectedFirst = resolve(
            afterRejecting,
            accepted.filter((id) => markIds(afterRejecting).includes(id)),
            REVISION_DECISIONS.ACCEPT,
          );
          expect(texts(rejectedFirst)).toEqual(texts(acceptedFirst));
        },
      ),
      {
        id: "two-record marks resolve as two batches in either order",
        numRuns: 60,
      },
    );
  },
  propertyTestTimeout(30_000),
);
