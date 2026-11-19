/**
 * A container whose own content a deletion took and whose remaining content
 * is an insertion made beside it goes once both are resolved that way,
 * whichever is resolved first: accepting the deletion last empties it, and
 * so does rejecting the insertion last.
 */
import { expect, test } from "bun:test";

import type { Document, ParagraphContent } from "../../model/document";
import { applyDocumentOp } from "../apply";
import {
  type DocumentOp,
  OP_STORIES,
  REVISION_DECISIONS,
  type RevisionDecision,
} from "../types";

const NEW_IDS = {
  revision: [700, 701, 702, 703, 704, 705, 706, 707],
  control: [800, 801, 802],
};

const apply = (document: Document, op: DocumentOp): Document => {
  const applied = applyDocumentOp(document, op);
  if (applied.isErr()) throw applied.error;
  return applied.value.document;
};

const resolve = (
  document: Document,
  id: number,
  decision: RevisionDecision,
): Document =>
  apply(document, {
    type: "resolveRevision",
    story: OP_STORIES.MAIN,
    revisionIds: [id],
    decision,
  });

const tab: ParagraphContent = { type: "run", content: [{ type: "tab" }] };

const containers: Record<string, ParagraphContent> = {
  inlineWrapper: {
    type: "inlineWrapper",
    kind: "bidi",
    control: "embedding",
    content: [tab],
  },
  inlineSdt: {
    type: "inlineSdt",
    properties: { sdtType: "richText", id: 1, tag: "clause" },
    content: [tab],
  },
};

for (const [name, container] of Object.entries(containers))
  test(`${name} emptied by an accepted deletion and a rejected insertion goes in either order`, () => {
    const source: Document = {
      package: {
        document: {
          content: [
            { type: "paragraph", paraId: "00000001", content: [container] },
          ],
        },
      },
    };
    const inserted = apply(source, {
      type: "insertText",
      at: { story: OP_STORIES.MAIN, blockId: "00000001", offset: 1 },
      text: "xx",
      runProps: {},
      newIds: NEW_IDS,
      revision: { id: 1, author: "C", date: "2026-03-04T05:00:00Z" },
    });
    const tracked = apply(inserted, {
      type: "deleteRange",
      from: { story: OP_STORIES.MAIN, blockId: "00000001", offset: 0 },
      to: { story: OP_STORIES.MAIN, blockId: "00000001", offset: 1 },
      newIds: NEW_IDS,
      revision: { id: 2, author: "C", date: "2026-03-04T05:00:01Z" },
    });
    const acceptedFirst = resolve(
      resolve(tracked, 2, REVISION_DECISIONS.ACCEPT),
      1,
      REVISION_DECISIONS.REJECT,
    );
    const rejectedFirst = resolve(
      resolve(tracked, 1, REVISION_DECISIONS.REJECT),
      2,
      REVISION_DECISIONS.ACCEPT,
    );
    for (const resolved of [acceptedFirst, rejectedFirst])
      expect(resolved.package.document.content).toMatchObject([
        { type: "paragraph", content: [] },
      ]);
  });
