/**
 * A suggested join records the properties it gives the paragraph it keeps
 * as a property change of its own id that belongs to the join's
 * paragraph-mark deletion: deciding the deletion decides that change too,
 * and the change alone still resolves by its own id.
 */
import { expect, test } from "bun:test";

import type { Document, Paragraph } from "../../model/document";
import { applyDocumentOps } from "../apply";
import { storyParagraphs } from "../blocks";
import { compileEditorIntent } from "../editorIntent";
import { freshenIdentities } from "../identity";
import { DOCUMENT_OP_REFUSAL_REASONS } from "../refusal";
import {
  DOCUMENT_OP_TYPES,
  OP_STORIES,
  REVISION_DECISIONS,
  type RevisionDecision,
} from "../types";

const paragraph = (
  paraId: string,
  text: string,
  alignment: "center" | "right",
): Paragraph => ({
  type: "paragraph",
  paraId,
  formatting: { alignment },
  content: [{ type: "run", content: [{ type: "text", text }] }],
});

const SOURCE: Document = {
  package: {
    document: {
      content: [
        paragraph("00000001", "First.", "center"),
        paragraph("00000002", "Second.", "right"),
      ],
    },
  },
};

const suggestedJoin = (document: Document) => {
  const compiled = compileEditorIntent(document, {
    intent: {
      type: "joinParagraphs",
      story: OP_STORIES.MAIN,
      blockId: "00000001",
      nextBlockId: "00000002",
    },
    mode: {
      type: "suggesting",
      revision: { id: 100, author: "Reviewer", date: "2026-02-03T04:05:06Z" },
      newIds: { revision: [101, 102, 103] },
    },
  });
  if (compiled.isErr()) throw compiled.error;
  const applied = applyDocumentOps(document, compiled.value.ops);
  if (applied.isErr()) throw applied.error;
  return applied.value.document;
};

const resolve = (
  document: Document,
  revisionIds: number[],
  decision: RevisionDecision,
) => {
  const resolved = applyDocumentOps(document, [
    {
      type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
      story: OP_STORIES.MAIN,
      revisionIds,
      decision,
    },
  ]);
  if (resolved.isErr()) throw resolved.error;
  return resolved.value.document;
};

const records = (document: Document) => {
  const [first, second] = storyParagraphs(document.package.document);
  return {
    mark: first?.paragraph.pPrMark,
    change: second?.paragraph.propertyChanges?.at(0),
  };
};

test("a tracked join links a new change when the paragraph has an empty change list", () => {
  const document: Document = {
    package: {
      document: {
        content: [
          paragraph("00000001", "First.", "center"),
          { ...paragraph("00000002", "Second.", "right"), propertyChanges: [] },
        ],
      },
    },
  };
  const applied = applyDocumentOps(document, [
    {
      type: DOCUMENT_OP_TYPES.JOIN_BLOCKS,
      story: OP_STORIES.MAIN,
      blockId: "00000001",
      nextBlockId: "00000002",
      revision: { id: 100, author: "Reviewer", date: "2026-02-03T04:05:06Z" },
      newIds: { revision: [101] },
    },
  ]);
  if (applied.isErr()) throw applied.error;
  const joined = applied.value.document;
  const { mark, change } = records(joined);
  if (mark === undefined || change === undefined)
    throw new Error("The join records a mark deletion and a property change.");
  expect(change.resolvesWith).toBe(mark.info.id);
  expect(resolve(joined, [mark.info.id], REVISION_DECISIONS.REJECT)).toEqual(
    SOURCE,
  );
});

test("freshening a copied join keeps each property change with its own mark", () => {
  const joined = suggestedJoin(SOURCE);
  const [first, second] = storyParagraphs(joined.package.document);
  if (first === undefined || second === undefined)
    throw new Error("The join keeps both paragraphs pending review.");
  const originals = [first.paragraph, second.paragraph];
  const copies = [
    { ...first.paragraph, paraId: "00000003" },
    { ...second.paragraph, paraId: "00000004" },
  ];
  for (const draft of [
    [...originals, ...copies],
    [...copies, ...originals],
  ]) {
    const freshened = freshenIdentities({
      before: originals,
      after: draft,
      newIds: { revision: [200, 201] },
      usedElsewhere: () => new Set(),
    });
    if (freshened.kind !== "fresh")
      throw new Error("Both copies have fresh ids.");
    const paragraphs = new Map(freshened.paragraphs.map((p) => [p.paraId, p]));
    expect(
      paragraphs.get("00000002")?.propertyChanges?.at(0)?.resolvesWith,
    ).toBe(100);
    expect(paragraphs.get("00000003")?.pPrMark?.info.id).toBe(200);
    expect(
      paragraphs.get("00000004")?.propertyChanges?.at(0)?.resolvesWith,
    ).toBe(200);
  }
});

test("a later formatting revision cannot take over a pending join's property change", () => {
  const joined = suggestedJoin(SOURCE);
  const changed = applyDocumentOps(joined, [
    {
      type: DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS,
      story: OP_STORIES.MAIN,
      blockId: "00000002",
      patch: { styleId: "Heading2" },
      revision: { id: 200, author: "Reviewer", date: "2026-02-03T04:05:07Z" },
    },
  ]);
  expect(changed.isErr() && changed.error.reason).toBe(
    DOCUMENT_OP_REFUSAL_REASONS.REVISION_CONFLICT,
  );
});

test("the join's property change belongs to its paragraph-mark deletion", () => {
  const joined = suggestedJoin(SOURCE);
  const { mark, change } = records(joined);
  expect(mark?.kind).toBe("del");
  if (mark === undefined || change === undefined)
    throw new Error("The join records a mark deletion and a property change.");
  expect(change.info.id).not.toBe(mark.info.id);
  expect(change.resolvesWith).toBe(mark.info.id);

  // Rejecting the deletion restores both paragraphs as they were.
  expect(
    resolve(joined, [mark.info.id], REVISION_DECISIONS.REJECT),
  ).toStrictEqual(SOURCE);
  // Accepting it accepts the properties it gave too.
  const accepted = resolve(joined, [mark.info.id], REVISION_DECISIONS.ACCEPT);
  const [kept, ...rest] = storyParagraphs(accepted.package.document);
  expect(rest).toEqual([]);
  expect(kept?.paragraph.formatting).toEqual({ alignment: "center" });
  expect(kept?.paragraph.propertyChanges).toBeUndefined();

  // The change alone resolves by its own id and leaves the join pending.
  const unformatted = resolve(
    joined,
    [change.info.id],
    REVISION_DECISIONS.REJECT,
  );
  expect(records(unformatted)).toEqual({ mark, change: undefined });
});
