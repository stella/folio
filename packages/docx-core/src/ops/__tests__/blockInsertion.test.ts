import { expect, test } from "bun:test";

import {
  type BlockContent,
  type Document,
  type Paragraph,
  paragraphNumberingReference,
} from "../../model/document";
import {
  type AppliedDocumentOp,
  applyDocumentOp,
  applyDocumentOps,
} from "../apply";
import { storyParagraphs } from "../blocks";
import { revisionIdDemand } from "../plan";
import { DOCUMENT_OP_REFUSAL_REASONS } from "../refusal";
import {
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  type InsertBlocksOp,
  OP_STORIES,
  REVISION_DECISIONS,
  type RevisionDecision,
} from "../types";

const makeParagraph = (paraId: string, text: string): Paragraph => ({
  type: "paragraph",
  paraId,
  content:
    text === "" ? [] : [{ type: "run", content: [{ type: "text", text }] }],
});

const baseDocument: Document = {
  package: { document: { content: [makeParagraph("00000001", "kept")] } },
};

test("inserting paragraphs before an anchor preserves its identity and content", () => {
  const inserted = makeParagraph("00000002", "new");
  const result = applyDocumentOp(baseDocument, {
    type: DOCUMENT_OP_TYPES.INSERT_BLOCKS,
    story: OP_STORIES.MAIN,
    at: { type: "before", blockId: "00000001" },
    blocks: [inserted],
  });
  if (result.isErr()) throw result.error;
  expect(result.value.document.package.document.content).toEqual([
    inserted,
    ...baseDocument.package.document.content,
  ]);
  expect(result.value.document.package.document.content.at(-1)).toBe(
    baseDocument.package.document.content.at(0),
  );
});

const stamp = { id: 100, author: "Reviewer", date: "2026-02-03T04:05:06Z" };

const apply = (document: Document, op: DocumentOp) => {
  const result = applyDocumentOp(document, op);
  if (result.isErr()) throw result.error;
  return result.value;
};

const resolve = (
  applied: Pick<AppliedDocumentOp, "document" | "revisions">,
  decision: RevisionDecision,
) =>
  apply(applied.document, {
    type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
    story: OP_STORIES.MAIN,
    revisionIds: applied.revisions,
    decision,
  });

const undo = (applied: AppliedDocumentOp) => {
  const restored = applyDocumentOps(applied.document, applied.inverse);
  if (restored.isErr()) throw restored.error;
  return restored.value.document;
};

const documentOf = (content: BlockContent[]): Document => ({
  package: { document: { content } },
});

const inCell = (content: BlockContent[]): BlockContent => ({
  type: "table",
  rows: [{ type: "tableRow", cells: [{ type: "tableCell", content }] }],
});

const insertion = (blocks: Paragraph[]): InsertBlocksOp => ({
  type: DOCUMENT_OP_TYPES.INSERT_BLOCKS,
  story: OP_STORIES.MAIN,
  at: { type: "before", blockId: "00000001" },
  blocks,
  revision: stamp,
  newIds: { revision: Array.from({ length: 32 }, (_, index) => 101 + index) },
});

test.each(["body", "cell"] as const)(
  "tracked list insertion in a %s accepts without a leftover paragraph",
  (container) => {
    const original = makeParagraph("00000001", "kept");
    const content = [original];
    const document = documentOf(
      container === "body" ? content : [inCell(content)],
    );
    const listed = [
      makeParagraph("00000002", "first"),
      makeParagraph("00000003", "second"),
    ].map((paragraph) =>
      Object.assign({}, paragraph, {
        formatting: {
          numPr: paragraphNumberingReference({ numId: 1, ilvl: 0 }),
        },
      }),
    );
    const op = insertion(listed);
    const tracked = apply(document, op);
    const direct = apply(document, { ...op, revision: undefined });
    expect(resolve(tracked, REVISION_DECISIONS.ACCEPT).document).toEqual(
      direct.document,
    );
    expect(resolve(tracked, REVISION_DECISIONS.REJECT).document).toEqual(
      document,
    );
    expect(undo(tracked)).toEqual(document);
    const demand = revisionIdDemand(document, op);
    if (demand.isErr()) throw demand.error;
    expect(demand.value).toBe(3);
  },
);

test.each(["", "kept"])(
  "tracked insertion before a final paragraph preserves its formatting and id (%s)",
  (text) => {
    const final = {
      ...makeParagraph("00000001", text),
      formatting: { alignment: "end" },
    } satisfies Paragraph;
    const document = documentOf([
      inCell([final]),
      makeParagraph("00000004", "outside"),
    ]);
    const tracked = apply(
      document,
      insertion([makeParagraph("00000002", "new")]),
    );
    const finalAfter = storyParagraphs(tracked.document.package.document).find(
      ({ paragraph }) => paragraph.paraId === final.paraId,
    );
    expect(finalAfter?.paragraph).toBe(final);
    expect(finalAfter?.paragraph.pPrMark).toBeUndefined();
    expect(resolve(tracked, REVISION_DECISIONS.REJECT).document).toEqual(
      document,
    );
  },
);

test("direct insertion after a final paragraph is exactly reversible", () => {
  const op = {
    ...insertion([makeParagraph("00000002", "new")]),
    at: { type: "after", blockId: "00000001" },
    revision: undefined,
  } as const;
  const inserted = apply(baseDocument, op);
  expect(inserted.document.package.document.content).toEqual([
    ...baseDocument.package.document.content,
    ...op.blocks,
  ]);
  expect(undo(inserted)).toEqual(baseDocument);
  expect(inserted.touched).toEqual({
    modified: [],
    inserted: ["00000002"],
    removed: [],
  });
});

test.each([
  ["at the end of the list", [] as BlockContent[]],
  [
    "before a table",
    [
      inCell([makeParagraph("00000003", "cell")]),
      makeParagraph("00000004", "after"),
    ],
  ],
] as const)(
  "tracked insertion %s accepts to both paragraphs and rejects to the source exactly",
  (_, after) => {
    const anchor = Object.assign({}, makeParagraph("00000001", "kept"), {
      formatting: { alignment: "center" as const },
    });
    const document = documentOf([anchor, ...after]);
    const inserted = Object.assign({}, makeParagraph("00000002", "new"), {
      formatting: { alignment: "right" as const },
    });
    const applied = apply(document, {
      ...insertion([inserted]),
      at: { type: "after", blockId: "00000001" },
    });
    const [first, second] = applied.document.package.document.content;
    if (after.length === 0) {
      // The anchor's mark is the inserted one; the new paragraph holds the
      // mark the anchor had, with the anchor's properties as its prior ones.
      expect(first?.type === "paragraph" && first.pPrMark?.kind).toBe("ins");
      expect(second?.type === "paragraph" && second.pPrMark).toBeUndefined();
      expect(
        second?.type === "paragraph" &&
          second.propertyChanges?.map(
            ({ previousFormatting }) => previousFormatting,
          ),
      ).toEqual([{ alignment: "center" }]);
    } else {
      expect(first?.type === "paragraph" && first.pPrMark).toBeUndefined();
      expect(second?.type === "paragraph" && second.pPrMark?.kind).toBe("ins");
    }
    expect(resolve(applied, REVISION_DECISIONS.REJECT).document).toEqual(
      document,
    );
    const accepted = resolve(
      applied,
      REVISION_DECISIONS.ACCEPT,
    ).document.package.document.content.slice(0, 2);
    expect(
      accepted.map((block) =>
        block.type === "paragraph"
          ? [
              block.paraId,
              block.formatting,
              block.pPrMark,
              block.propertyChanges,
            ]
          : block.type,
      ),
    ).toEqual([
      ["00000001", { alignment: "center" }, undefined, undefined],
      [expect.any(String), { alignment: "right" }, undefined, undefined],
    ]);
    expect(undo(applied)).toEqual(document);
  },
);

test("insertion validates identities, review conflicts and section boundaries", () => {
  const collision = applyDocumentOp(
    baseDocument,
    insertion([makeParagraph("00000001", "duplicate")]),
  );
  expect(collision.isErr() && collision.error.reason).toBe(
    DOCUMENT_OP_REFUSAL_REASONS.ID_COLLISION,
  );
  const changed = {
    ...makeParagraph("00000002", "changed"),
    pPrMark: { kind: "ins", info: { id: 9, author: "A" } },
  } satisfies Paragraph;
  const conflict = applyDocumentOp(baseDocument, insertion([changed]));
  expect(conflict.isErr() && conflict.error.reason).toBe(
    DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
  );
  const section = {
    ...makeParagraph("00000002", "section"),
    sectionProperties: {},
  };
  const boundary = applyDocumentOp(baseDocument, insertion([section]));
  expect(boundary.isErr() && boundary.error.reason).toBe(
    DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
  );
  const insufficient = applyDocumentOp(baseDocument, {
    ...insertion([makeParagraph("00000002", "new")]),
    newIds: {},
  });
  expect(insufficient.isErr() && insufficient.error.reason).toBe(
    DOCUMENT_OP_REFUSAL_REASONS.NEEDS_NEW_IDS,
  );
});

test.each([
  ["no paragraph follows", [] as BlockContent[]],
  ["a paragraph follows", [makeParagraph("00000004", "after")]],
] as const)(
  "tracked insertion before a range boundary between blocks resolves each record alone (%s)",
  (_, after) => {
    const anchor: Paragraph = {
      ...makeParagraph("00000001", "kept"),
      content: [
        { type: "bookmarkStart", id: 1, name: "_Ref1" },
        { type: "run", content: [{ type: "text", text: "kept" }] },
      ],
    };
    const document = documentOf([
      anchor,
      { type: "bookmarkEnd", id: 1 },
      ...after,
    ]);
    const applied = apply(document, {
      ...insertion([
        makeParagraph("00000002", "first"),
        makeParagraph("00000003", "second"),
      ]),
      at: { type: "after", blockId: "00000001" },
    });
    // The boundary stays after the paragraph that holds the anchor's mark.
    expect(
      applied.document.package.document.content.map((block) =>
        block.type === "paragraph" ? "paragraph" : block.type,
      ),
    ).toEqual([
      "paragraph",
      "paragraph",
      "paragraph",
      "bookmarkEnd",
      ...after.map(() => "paragraph"),
    ]);
    expect(resolve(applied, REVISION_DECISIONS.REJECT).document).toEqual(
      document,
    );
    const accepted = resolve(applied, REVISION_DECISIONS.ACCEPT).document;
    expect(
      storyParagraphs(accepted.package.document).map(({ paragraph }) =>
        paragraph.content
          .flatMap((node) => (node.type === "run" ? node.content : []))
          .map((leaf) => (leaf.type === "text" ? leaf.text : ""))
          .join(""),
      ),
    ).toEqual(["kept", "first", "second", ...after.map(() => "after")]);
    for (const id of applied.revisions)
      for (const decision of Object.values(REVISION_DECISIONS)) {
        const alone = applyDocumentOp(applied.document, {
          type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
          story: OP_STORIES.MAIN,
          revisionIds: [id],
          decision,
        });
        if (alone.isErr()) throw alone.error;
      }
    expect(undo(applied)).toEqual(document);
  },
);

test("tracked insertion after a final paragraph keeps the anchor's own records and ids", () => {
  const plain = makeParagraph("00000001", "kept");
  const anchor: Paragraph = {
    ...plain,
    content: [
      { type: "run", content: [{ type: "text", text: "kept" }] },
      {
        type: "insertion",
        info: { id: 5, author: "Other", date: "2026-01-01T00:00:00Z" },
        content: [{ type: "run", content: [{ type: "text", text: " more" }] }],
      },
      {
        type: "inlineSdt",
        properties: { sdtType: "richText", id: 42, tag: "clause" },
        content: [{ type: "run", content: [{ type: "text", text: " field" }] }],
      },
    ],
  };
  const op: InsertBlocksOp = {
    ...insertion([makeParagraph("00000002", "new")]),
    at: { type: "after", blockId: "00000001" },
  };
  const demand = (document: Document) => {
    const counted = revisionIdDemand(document, op);
    if (counted.isErr()) throw counted.error;
    return counted.value;
  };
  const document = documentOf([anchor]);
  // The anchor's own records take no new ids.
  expect(demand(document)).toBe(demand(documentOf([plain])));
  const applied = apply(document, {
    ...op,
    newIds: {
      revision: Array.from(
        { length: demand(document) },
        (_, index) => 101 + index,
      ),
    },
  });
  const [first] = applied.document.package.document.content;
  expect(first?.type === "paragraph" && first.content).toEqual(anchor.content);
  expect(resolve(applied, REVISION_DECISIONS.REJECT).document).toEqual(
    document,
  );
  expect(undo(applied)).toEqual(document);
});

test.each(["body", "cell"] as const)(
  "rejecting only the anchor's mark after a final paragraph keeps every later paragraph's pending review (%s)",
  (container) => {
    const anchor = Object.assign({}, makeParagraph("00000001", "kept"), {
      formatting: { alignment: "center" as const },
    });
    const document = documentOf(
      container === "body" ? [anchor] : [inCell([anchor])],
    );
    const applied = apply(document, {
      ...insertion([
        Object.assign({}, makeParagraph("00000002", "first"), {
          formatting: { alignment: "right" as const },
        }),
        Object.assign({}, makeParagraph("00000003", "second"), {
          formatting: { alignment: "right" as const },
        }),
      ]),
      at: { type: "after", blockId: "00000001" },
    });
    const paragraphsOf = (edited: Document) =>
      storyParagraphs(edited.package.document).map(
        ({ paragraph }) => paragraph,
      );
    const [first, second, third] = paragraphsOf(applied.document);
    const anchorMark = first?.pPrMark?.info.id;
    const insertedMark = second?.pPrMark?.info.id;
    if (anchorMark === undefined || insertedMark === undefined)
      throw new Error("Both inserted breaks carry a mark.");
    // The last new paragraph holds the anchor's mark, with a pending change
    // of the anchor's properties.
    expect(third?.propertyChanges?.length).toBe(1);
    const rejectedAnchor = apply(applied.document, {
      type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
      story: OP_STORIES.MAIN,
      revisionIds: [anchorMark],
      decision: REVISION_DECISIONS.REJECT,
    });
    const joined = paragraphsOf(rejectedAnchor.document);
    // The anchor keeps its identity and takes the next paragraph's words,
    // its properties and its still-pending inserted mark.
    expect(
      joined.map((paragraph) => [
        paragraph.paraId,
        paragraph.formatting,
        paragraph.pPrMark?.kind,
        paragraph.pPrMark?.info.id,
        paragraph.propertyChanges?.length,
      ]),
    ).toEqual([
      ["00000001", { alignment: "right" }, "ins", insertedMark, undefined],
      [third?.paraId, { alignment: "right" }, undefined, undefined, 1],
    ]);
    // What is left resolves, both ways, to what resolving it all at once does.
    const rest = applied.revisions.filter((id) => id !== anchorMark);
    for (const decision of Object.values(REVISION_DECISIONS)) {
      const resolved = resolve(
        { document: rejectedAnchor.document, revisions: rest },
        decision,
      );
      if (decision === REVISION_DECISIONS.REJECT)
        expect(resolved.document).toEqual(document);
      else
        expect(
          paragraphsOf(resolved.document).map((paragraph) => [
            paragraph.formatting,
            paragraph.pPrMark,
            paragraph.propertyChanges,
          ]),
        ).toEqual([
          [{ alignment: "right" }, undefined, undefined],
          [{ alignment: "right" }, undefined, undefined],
        ]);
    }
    expect(undo(rejectedAnchor)).toEqual(applied.document);
  },
);

test("rejecting a break into an anchor that kept its identity folds the emptied control at the seam", () => {
  const info = (id: number) => ({
    id,
    author: "Reviewer",
    date: "2026-02-03T04:05:06Z",
  });
  const control = (id: number, text: string, revision: number) => ({
    type: "inlineSdt" as const,
    properties: { sdtType: "richText" as const, id, tag: "clause" },
    content: [
      {
        type: "insertion" as const,
        info: info(revision),
        content: [
          { type: "run" as const, content: [{ type: "text" as const, text }] },
        ],
      },
    ],
  });
  const run = (text: string) => ({
    type: "run" as const,
    content: [{ type: "text" as const, text }],
  });
  // An inserted break before an empty anchor whose inserted mark keeps it,
  // with the paragraph that holds the anchor's former mark opening with a
  // control only inserted words fill.
  const document = documentOf([
    {
      type: "paragraph",
      paraId: "00000001",
      pPrMark: { kind: "ins", info: info(101) },
      content: [run("z"), control(42, "x", 102)],
    },
    {
      type: "paragraph",
      paraId: "00000002",
      pPrMark: { kind: "ins", info: info(103), survivor: "first" },
      content: [],
    },
    {
      type: "paragraph",
      paraId: "00000003",
      content: [control(43, "y", 104), run("b")],
    },
  ]);
  const rejected = apply(document, {
    type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
    story: OP_STORIES.MAIN,
    revisionIds: [101, 102, 103, 104],
    decision: REVISION_DECISIONS.REJECT,
  });
  // Both controls were emptied, so the one after the seam folds into the one
  // before it, in the paragraph that kept the anchor's identity.
  expect(rejected.document.package.document.content).toEqual([
    {
      type: "paragraph",
      paraId: "00000002",
      content: [
        run("z"),
        {
          type: "inlineSdt",
          properties: { sdtType: "richText", id: 42, tag: "clause" },
          content: [],
        },
        run("b"),
      ],
    },
  ]);
  expect(undo(rejected)).toEqual(document);
});
