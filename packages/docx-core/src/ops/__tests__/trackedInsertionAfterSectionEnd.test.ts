/**
 * A tracked or suggested paragraph put directly after one that ends a section,
 * before a block that is no paragraph, applies as the direct insertion does:
 * it opens the section after the break, which stays on its paragraph. The new
 * paragraphs hold their own inserted marks, so accepting gives the direct
 * edit, rejecting gives the document back, and the inverse is exact.
 */
import { expect, test } from "bun:test";

import type {
  BlockContent,
  Document,
  Paragraph,
  ParagraphMarkChange,
} from "../../model/document";
import {
  type AppliedDocumentOp,
  applyDocumentOp,
  applyDocumentOps,
} from "../apply";
import {
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  type InsertBlocksOp,
  OP_STORIES,
  REVISION_DECISIONS,
  type RevisionDecision,
  type RevisionStamp,
} from "../types";

const OTHER = { author: "Other", date: "2026-01-01T00:00:00Z" };
const PENDING_CHANGE = 7;
const PENDING_MARK = 8;
const PENDING_TEXT = 9;
const PENDING_ROW = 20;
const PENDING_CELL = 21;

const run = (text: string) => ({
  type: "run" as const,
  content: [{ type: "text" as const, text }],
});

const paragraph = (paraId: string, text: string): Paragraph => ({
  type: "paragraph",
  paraId,
  content: text === "" ? [] : [run(text)],
});

const cell = (content: Paragraph) => ({
  type: "tableCell" as const,
  content: [content],
});

/** The block after the section's last paragraph, with its own pending ids. */
const FOLLOWERS = {
  "a content control": {
    block: {
      type: "blockSdt",
      properties: { sdtType: "richText", id: 4 },
      content: [paragraph("00000030", "Held.")],
    },
    records: [],
  },
  "a custom XML block": {
    block: {
      type: "blockCustomXml",
      openingXml: '<w:customXml w:element="held">',
      closingXml: "</w:customXml>",
      content: [paragraph("00000030", "Held.")],
    },
    records: [],
  },
  "a preserved block": {
    block: { type: "preservedBlock", xml: "<w:altChunk/>" },
    records: [],
  },
  "a table": {
    block: {
      type: "table",
      rows: [
        { type: "tableRow", cells: [cell(paragraph("00000030", "Held."))] },
      ],
    },
    records: [],
  },
  "a table whose first row is pending": {
    block: {
      type: "table",
      rows: [
        {
          type: "tableRow",
          structuralChange: {
            type: "tableRowInsertion",
            info: { id: PENDING_ROW, ...OTHER },
          },
          cells: [
            cell({
              ...paragraph("00000030", ""),
              content: [
                {
                  type: "insertion",
                  info: { id: PENDING_CELL, ...OTHER },
                  content: [run("Held.")],
                },
              ],
            }),
          ],
        },
      ],
    },
    records: [PENDING_ROW, PENDING_CELL],
  },
} as const satisfies Record<
  string,
  { block: BlockContent; records: readonly number[] }
>;

const SECTION = { pageWidth: 11_906 };

const ending = (fields: Omit<Paragraph, "type" | "paraId">): Paragraph => ({
  type: "paragraph",
  paraId: "00000001",
  sectionProperties: SECTION,
  ...fields,
});

const insertedMark: ParagraphMarkChange = {
  kind: "ins",
  info: { id: PENDING_MARK, ...OTHER },
};

/** The section's last paragraph, with its own pending ids. */
const ANCHORS = {
  "with words": { anchor: ending({ content: [run("First.")] }), pending: [] },
  empty: { anchor: ending({ content: [] }), pending: [] },
  "with properties": {
    anchor: ending({
      content: [run("First.")],
      formatting: { alignment: "center" },
    }),
    pending: [],
  },
  "with a pending property change": {
    anchor: ending({
      content: [run("First.")],
      formatting: { alignment: "center" },
      propertyChanges: [
        {
          type: "paragraphPropertyChange",
          info: { id: PENDING_CHANGE, ...OTHER },
          previousFormatting: { alignment: "start" },
        },
      ],
    }),
    pending: [PENDING_CHANGE],
  },
  "with an inserted break": {
    anchor: ending({
      content: [
        {
          type: "insertion",
          info: { id: PENDING_TEXT, ...OTHER },
          content: [run("First.")],
        },
      ],
      pPrMark: insertedMark,
    }),
    pending: [PENDING_TEXT, PENDING_MARK],
  },
  "with a break pending deletion": {
    anchor: ending({
      content: [run("First.")],
      pPrMark: { kind: "del", info: { id: PENDING_MARK, ...OTHER } },
    }),
    pending: [PENDING_MARK],
  },
} as const satisfies Record<
  string,
  { anchor: Paragraph; pending: readonly number[] }
>;

const STAMP = { id: 100, author: "Reviewer", date: "2026-02-03T04:05:06Z" };

const MODES = {
  tracked: STAMP,
  suggested: { ...STAMP, suggestionId: "suggestion" },
} as const satisfies Record<string, RevisionStamp>;

const INSERTED = {
  "one paragraph": [paragraph("00000002", "New.")],
  "two paragraphs": [
    paragraph("00000002", "New."),
    {
      ...paragraph("00000003", "Also new."),
      formatting: { alignment: "end" },
    },
  ],
  "an empty paragraph": [paragraph("00000002", "")],
} as const satisfies Record<string, readonly Paragraph[]>;

/** The sections a body's breaks divide it into, as a reader of it states them. */
const withSections = (content: BlockContent[]): Document => {
  const sections: NonNullable<Document["package"]["document"]["sections"]> = [];
  let open: BlockContent[] = [];
  for (const block of content) {
    open.push(block);
    if (block.type !== "paragraph" || block.sectionProperties === undefined)
      continue;
    sections.push({ content: open, properties: block.sectionProperties });
    open = [];
  }
  const finalSectionProperties = { pageWidth: 12_240 };
  if (open.length > 0)
    sections.push({ content: open, properties: finalSectionProperties });
  return {
    package: { document: { content, sections, finalSectionProperties } },
  };
};

const apply = (document: Document, op: DocumentOp): AppliedDocumentOp => {
  const result = applyDocumentOp(document, op);
  if (result.isErr()) throw result.error;
  return result.value;
};

const resolve = (
  document: Document,
  revisionIds: readonly number[],
  decision: RevisionDecision,
): Document =>
  revisionIds.length === 0
    ? document
    : apply(document, {
        type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
        story: OP_STORIES.MAIN,
        revisionIds,
        decision,
      }).document;

/** Each inverse, as a client stores and replays it. */
const undo = (applied: AppliedDocumentOp): AppliedDocumentOp => {
  const inverse: DocumentOp[] = JSON.parse(JSON.stringify(applied.inverse));
  const restored = applyDocumentOps(applied.document, inverse);
  if (restored.isErr()) throw restored.error;
  return restored.value;
};

const CASES = Object.entries(FOLLOWERS).flatMap(([follower, place]) =>
  Object.entries(ANCHORS).flatMap(([anchor, source]) =>
    Object.entries(MODES).flatMap(([mode, revision]) =>
      Object.entries(INSERTED).map(([inserted, blocks]) => ({
        follower,
        place,
        anchor,
        source,
        mode,
        revision,
        inserted,
        blocks,
      })),
    ),
  ),
);

test.each(CASES)(
  "$mode, $inserted after a section's last paragraph $anchor, before $follower, reviews as the direct insertion",
  ({ place, source, revision, blocks }) => {
    const document = withSections(
      structuredClone([
        source.anchor,
        place.block,
        paragraph("00000040", "Last."),
      ]),
    );
    const op: InsertBlocksOp = {
      type: DOCUMENT_OP_TYPES.INSERT_BLOCKS,
      story: OP_STORIES.MAIN,
      at: { type: "after", blockId: "00000001" },
      blocks: structuredClone(blocks),
    };
    const direct = apply(document, op).document;
    const tracked = apply(document, {
      ...op,
      revision,
      newIds: {
        revision: Array.from({ length: 16 }, (_, index) => 101 + index),
      },
    });
    // The break stays on its paragraph, which keeps its records.
    expect(tracked.document.package.document.content.at(0)).toEqual(
      source.anchor,
    );
    const restored = undo(tracked);
    expect(restored.document).toEqual(document);
    expect(undo(restored).document).toEqual(tracked.document);
    const theirs = [...source.pending, ...place.records];
    const own = tracked.revisions.filter((id) => !theirs.includes(id));
    expect(own.length).toBeGreaterThan(0);
    expect(resolve(tracked.document, own, REVISION_DECISIONS.ACCEPT)).toEqual(
      direct,
    );
    expect(resolve(tracked.document, own, REVISION_DECISIONS.REJECT)).toEqual(
      document,
    );
    for (const decision of Object.values(REVISION_DECISIONS)) {
      const expected = resolve(
        decision === REVISION_DECISIONS.ACCEPT ? direct : document,
        theirs,
        decision,
      );
      // All at once, the others' records first, or the insertion first.
      expect(resolve(tracked.document, [...own, ...theirs], decision)).toEqual(
        expected,
      );
      expect(
        resolve(resolve(tracked.document, theirs, decision), own, decision),
      ).toEqual(expected);
      expect(
        resolve(resolve(tracked.document, own, decision), theirs, decision),
      ).toEqual(expected);
    }
  },
);

test.each(Object.values(MODES))(
  "rejecting an insertion after a section break preserves a trailing bookmark in every order",
  (revision) => {
    const document = withSections([
      ANCHORS.empty.anchor,
      { type: "bookmarkStart", id: 90, name: "_GoBack" },
    ]);
    const tracked = apply(document, {
      type: DOCUMENT_OP_TYPES.INSERT_BLOCKS,
      story: OP_STORIES.MAIN,
      at: { type: "after", blockId: "00000001" },
      blocks: [paragraph("00000002", "New.")],
      revision,
      newIds: { revision: [101] },
    });
    const inserted = tracked.document.package.document.content.at(1);
    if (inserted?.type !== "paragraph" || inserted.pPrMark === undefined)
      throw new Error("Expected an inserted paragraph mark");
    const mark = inserted.pPrMark.info.id;
    const text = tracked.revisions.filter((id) => id !== mark);
    expect(text).toHaveLength(1);
    const reject = (current: Document, ids: readonly number[]) =>
      resolve(current, ids, REVISION_DECISIONS.REJECT);
    expect(reject(tracked.document, [mark, ...text])).toEqual(document);
    expect(reject(reject(tracked.document, [mark]), text)).toEqual(document);
    expect(reject(reject(tracked.document, text), [mark])).toEqual(document);
  },
);
