import { expect, spyOn, test } from "bun:test";

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
import * as blockOperations from "../blocks";
import * as identityKeys from "../ids";
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

test.each(["before", "after"] as const)(
  "tracked insertion %s a main paragraph validates without a draft splice",
  (type) => {
    const op = {
      ...insertion([
        makeParagraph("00000003", "first"),
        makeParagraph("00000004", "second"),
      ]),
      at: { type, blockId: "00000001" },
    } satisfies InsertBlocksOp;
    const splice = spyOn(blockOperations, "replaceParagraphs");
    const gate = spyOn(
      blockOperations,
      "canSpliceMainBodyWithoutSectionRepair",
    );
    try {
      const validated = apply(baseDocument, op);
      expect(splice).toHaveBeenCalledTimes(1);
      gate.mockReturnValue(false);
      const materialized = apply(baseDocument, op);
      expect(splice).toHaveBeenCalledTimes(3);
      expect(validated).toStrictEqual(materialized);
      expect(undo(validated)).toStrictEqual(baseDocument);
    } finally {
      gate.mockRestore();
      splice.mockRestore();
    }
  },
);

test.each([
  {
    paragraph: {
      type: "paragraph",
      paraId: "bad",
      content: [{ type: "run", content: [] }],
    },
    reason: DOCUMENT_OP_REFUSAL_REASONS.EMPTY_CONTENT,
  },
  {
    paragraph: makeParagraph("00000001", "duplicate"),
    reason: DOCUMENT_OP_REFUSAL_REASONS.ID_COLLISION,
  },
] satisfies { paragraph: Paragraph; reason: string }[])(
  "local preflight preserves $reason before tracked ID demand",
  ({ paragraph, reason }) => {
    const op = { ...insertion([paragraph]), newIds: { revision: [] } };
    const local = applyDocumentOp(baseDocument, op);
    const gate = spyOn(
      blockOperations,
      "canSpliceMainBodyWithoutSectionRepair",
    ).mockReturnValue(false);
    try {
      const materialized = applyDocumentOp(baseDocument, op);
      expect(local.isErr() && local.error.reason).toBe(reason);
      expect(local.isErr() && local.error.message).toBe(
        materialized.isErr() && materialized.error.message,
      );
      expect(materialized.isErr() && materialized.error.reason).toBe(reason);
    } finally {
      gate.mockRestore();
    }
  },
);

test.each(["after", "afterTable"] as const)(
  "tracked insertions at %s do not repeat package identity censuses",
  (type) => {
    const census = spyOn(identityKeys, "packageIdentityKeys");
    try {
      const carrier = makeParagraph("00000001", "kept");
      const tableAnchor = makeParagraph("00000010", "cell");
      let document = documentOf(
        type === "after" ? [carrier] : [inCell([tableAnchor]), carrier],
      );
      const blockId = type === "after" ? "00000001" : "00000010";
      const count = 8;
      for (let index = 0; index < count; index += 1) {
        const paraId = (0x10000000 + index).toString(16);
        const revisionId = stamp.id + index * 4;
        const result = apply(document, {
          ...insertion([makeParagraph(paraId, "inserted")]),
          at: { type, blockId },
          revision: { ...stamp, id: revisionId },
          newIds: {
            revision: [revisionId + 1, revisionId + 2, revisionId + 3],
          },
        });
        expect(result.revisions).toContain(revisionId);
        document = result.document;
      }
      // Contract validation can census the initial package and each new
      // package produced directly by table-side insertion.
      const validationCensuses = type === "afterTable" ? count + 1 : 1;
      expect(census.mock.calls.length).toBeLessThanOrEqual(validationCensuses);
    } finally {
      census.mockRestore();
    }
  },
);

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
    const finalAfter = blockOperations
      .storyParagraphs(tracked.document.package.document)
      .find(({ paragraph }) => paragraph.paraId === final.paraId);
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
    DOCUMENT_OP_REFUSAL_REASONS.PENDING_PARAGRAPH_MARK_INSERTION,
  );
  const section = {
    ...makeParagraph("00000002", "section"),
    sectionProperties: {},
  };
  // The break goes in with the inserted mark it travels with.
  const sectioned = apply(baseDocument, insertion([section]));
  const [opened] = sectioned.document.package.document.content;
  expect(opened?.type === "paragraph" && opened.pPrMark?.kind).toBe("ins");
  expect(opened?.type === "paragraph" && opened.sectionProperties).toEqual({});
  // After the story's last paragraph the new one keeps the story's mark.
  const boundary = applyDocumentOp(baseDocument, {
    ...insertion([section]),
    at: { type: "after", blockId: "00000001" },
  });
  expect(boundary.isErr() && boundary.error.reason).toBe(
    DOCUMENT_OP_REFUSAL_REASONS.SECTION_BOUNDARY,
  );
  const insufficient = applyDocumentOp(baseDocument, {
    ...insertion([makeParagraph("00000002", "new")]),
    newIds: {},
  });
  expect(insufficient.isErr() && insufficient.error.reason).toBe(
    DOCUMENT_OP_REFUSAL_REASONS.NEEDS_NEW_IDS,
  );
});

test("a direct insertion keeps a supplied mark only before its container end", () => {
  const anchor = makeParagraph("0abcdef1", "anchor");
  const wrapper = {
    type: "blockSdt",
    properties: {},
    content: [anchor],
  } satisfies BlockContent;
  const marked = {
    ...makeParagraph("00000002", "marked"),
    pPrMark: { kind: "ins", info: { id: 9, author: "A" } },
  } satisfies Paragraph;
  const op = {
    ...insertion([marked]),
    at: { type: "after", blockId: "0ABCDEF1" },
    revision: undefined,
  } as const;
  const terminal = applyDocumentOp(documentOf([wrapper]), op);
  expect(terminal.isErr() && terminal.error.reason).toBe(
    DOCUMENT_OP_REFUSAL_REASONS.CONTAINER_FINAL_MARK,
  );
  const tail = makeParagraph("00000003", "tail");
  const nonterminal = apply(documentOf([wrapper, tail]), op);
  expect(nonterminal.document.package.document.content).toEqual([
    { ...wrapper, content: [anchor, marked] },
    tail,
  ]);
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
      blockOperations
        .storyParagraphs(accepted.package.document)
        .map(({ paragraph }) =>
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
      blockOperations
        .storyParagraphs(edited.package.document)
        .map(({ paragraph }) => paragraph);
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

const followingTableCases = [
  {
    name: "terminal table",
    after: [inCell([makeParagraph("00000003", "cell")])],
    markOwner: "anchor",
  },
  {
    name: "table with no run-in receiver",
    after: [
      inCell([
        {
          type: "blockSdt",
          properties: {},
          content: [makeParagraph("00000003", "controlled")],
        },
        makeParagraph("00000004", "cell end"),
      ]),
      makeParagraph("00000005", "story end"),
    ],
    markOwner: "anchor",
  },
  {
    name: "table with a run-in receiver and final paragraph",
    after: [
      inCell([makeParagraph("00000003", "cell")]),
      makeParagraph("00000005", "story end"),
    ],
    markOwner: "inserted",
  },
  {
    name: "following paragraph",
    after: [makeParagraph("00000005", "story end")],
    markOwner: "inserted",
  },
] as const satisfies readonly {
  name: string;
  after: readonly BlockContent[];
  markOwner: "anchor" | "inserted";
}[];

for (const { name, after, markOwner } of followingTableCases)
  for (const suggestion of ["ordinary", "suggested"] as const)
    for (const count of [1, 2])
      test(`${suggestion} insertion before ${name}, ${count} paragraphs, owns a resolvable source break`, () => {
        const anchor = {
          ...makeParagraph("00000001", "kept"),
          formatting: { alignment: "center" },
        } satisfies Paragraph;
        const document = documentOf([anchor, ...structuredClone(after)]);
        const blocks = Array.from(
          { length: count },
          (_, index) =>
            ({
              ...makeParagraph(
                (16 + index).toString(16).padStart(8, "0"),
                `new ${index}`,
              ),
              formatting: { alignment: "right" },
            }) satisfies Paragraph,
        );
        const op = {
          ...insertion(blocks),
          at: { type: "after", blockId: "00000001" },
          revision: {
            ...stamp,
            ...(suggestion === "suggested" ? { suggestionId: "proposal" } : {}),
          },
        } satisfies InsertBlocksOp;
        const applied = apply(document, op);
        const [first, inserted] = applied.document.package.document.content;
        expect(first?.type === "paragraph" && first.pPrMark?.survivor).toBe(
          markOwner === "anchor" ? "first" : undefined,
        );
        expect(first?.type === "paragraph" && first.pPrMark?.kind).toBe(
          markOwner === "anchor" ? "ins" : undefined,
        );
        expect(inserted?.type === "paragraph" && inserted.pPrMark?.kind).toBe(
          markOwner === "anchor" && count === 1 ? undefined : "ins",
        );
        expect(undo(applied)).toStrictEqual(document);
        expect(
          applyDocumentOps(
            applied.document,
            JSON.parse(JSON.stringify(applied.inverse)),
          ).unwrap().document,
        ).toStrictEqual(document);
        const direct = apply(document, { ...op, revision: undefined });
        for (const decision of Object.values(REVISION_DECISIONS)) {
          const expected =
            decision === REVISION_DECISIONS.REJECT ? document : direct.document;
          const resolved = resolve(applied, decision);
          expect(resolved.document).toStrictEqual(expected);
          expect(undo(resolved)).toStrictEqual(applied.document);
          for (const ids of [
            applied.revisions,
            applied.revisions.toReversed(),
          ]) {
            let current = applied.document;
            for (const id of ids)
              current = apply(current, {
                type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
                story: OP_STORIES.MAIN,
                revisionIds: [id],
                decision,
              }).document;
            expect(current).toStrictEqual(expected);
          }
        }
      });

/** Room for resolving every order of up to six revisions one at a time. */
const EVERY_ORDER_TIMEOUT_MS = 30_000;

const orderings = (ids: readonly number[]): number[][] =>
  ids.length <= 1
    ? [[...ids]]
    : ids.flatMap((id, index) =>
        orderings(ids.filter((_, other) => other !== index)).map((rest) =>
          [id].concat(rest),
        ),
      );

test.each([
  { count: 1, tailText: "tail" },
  { count: 2, tailText: "tail" },
  { count: 3, tailText: "tail" },
  // An empty paragraph after the control stays: nothing joins across it.
  { count: 1, tailText: "" },
  { count: 2, tailText: "" },
  { count: 3, tailText: "" },
])(
  "tracked insertion after a table that ends a content control stays in the control ($count paragraphs, tail $tailText)",
  ({ count, tailText }) => {
    const control = (content: BlockContent[]): BlockContent => ({
      type: "blockSdt",
      properties: { tag: "control" },
      content,
    });
    const table = inCell([makeParagraph("00000001", "cell")]);
    const tail = makeParagraph("00000005", tailText);
    const document = documentOf([control([table]), tail]);
    const blocks = Array.from({ length: count }, (_, index) =>
      makeParagraph((16 + index).toString(16).padStart(8, "0"), `new ${index}`),
    );
    const op = {
      ...insertion(blocks),
      at: { type: "afterTable", blockId: "00000001" },
    } satisfies InsertBlocksOp;
    const applied = apply(document, op);
    const direct = apply(document, { ...op, revision: undefined });
    expect(direct.document).toStrictEqual(
      documentOf([control([table, ...blocks]), tail]),
    );
    const [held, after] = applied.document.package.document.content;
    expect(after).toStrictEqual(tail);
    expect(
      held?.type === "blockSdt" &&
        held.content.map((block) =>
          block.type === "paragraph" ? block.pPrMark?.kind : block.type,
        ),
    ).toStrictEqual(["table", ...blocks.map(() => "ins")]);
    expect(undo(applied)).toStrictEqual(document);
    expect(
      applyDocumentOps(
        applied.document,
        JSON.parse(JSON.stringify(applied.inverse)),
      ).unwrap().document,
    ).toStrictEqual(document);
    // Each paragraph's words and break resolve one at a time in any order. A
    // rejected break with words still standing has nothing in the control to
    // join and keeps its place only for them: rejecting them removes it.
    const marks = held?.type === "blockSdt" ? held.content : [];
    const pairs = marks.flatMap((block) =>
      block.type === "paragraph" &&
      block.pPrMark !== undefined &&
      block.content[0]?.type === "insertion"
        ? [{ mark: block.pPrMark.info.id, words: block.content[0].info.id }]
        : [],
    );
    expect(pairs).toHaveLength(count);
    const resolveInOrder = (
      ids: readonly number[],
      decision: RevisionDecision,
    ) => {
      let current = applied.document;
      for (const id of ids) {
        const step = apply(current, {
          type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
          story: OP_STORIES.MAIN,
          revisionIds: [id],
          decision,
        });
        expect(undo(step)).toStrictEqual(current);
        current = step.document;
      }
      return current;
    };
    for (const decision of Object.values(REVISION_DECISIONS)) {
      const expected =
        decision === REVISION_DECISIONS.REJECT ? document : direct.document;
      const resolved = resolve(applied, decision);
      expect(resolved.document).toStrictEqual(expected);
      expect(undo(resolved)).toStrictEqual(applied.document);
      for (const ids of orderings(applied.revisions))
        expect(resolveInOrder(ids, decision)).toStrictEqual(expected);
    }
    // The breaks last, from the final one back: each emptied paragraph goes
    // while another inserted one still stands before it.
    expect(
      resolveInOrder(
        [
          ...pairs.map(({ words }) => words),
          ...pairs.map(({ mark }) => mark).toReversed(),
        ],
        REVISION_DECISIONS.REJECT,
      ),
    ).toStrictEqual(document);
  },
  EVERY_ORDER_TIMEOUT_MS,
);

test.each(
  [1, 2, 3].flatMap((count) =>
    (["rejected", "accepted", "pending"] as const).map((words) => ({
      count,
      words,
    })),
  ),
)(
  "the breaks of $count paragraphs inserted after a table that ends a content control resolve alike one at a time in any order and all at once (words $words)",
  ({ count, words }) => {
    const control: BlockContent = {
      type: "blockSdt",
      properties: { tag: "control" },
      content: [inCell([makeParagraph("00000001", "cell")])],
    };
    const document = documentOf([control, makeParagraph("00000005", "tail")]);
    const applied = apply(document, {
      ...insertion(
        Array.from({ length: count }, (_, index) =>
          makeParagraph(
            (16 + index).toString(16).padStart(8, "0"),
            `new ${index}`,
          ),
        ),
      ),
      at: { type: "afterTable", blockId: "00000001" },
    });
    const [held] = applied.document.package.document.content;
    const marks = (held?.type === "blockSdt" ? held.content : []).flatMap(
      (block) =>
        block.type === "paragraph" && block.pPrMark !== undefined
          ? [block.pPrMark.info.id]
          : [],
    );
    expect(marks).toHaveLength(count);
    const wordIds = applied.revisions.filter((id) => !marks.includes(id));
    const start =
      words === "pending"
        ? applied.document
        : apply(applied.document, {
            type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
            story: OP_STORIES.MAIN,
            revisionIds: wordIds,
            decision:
              words === "accepted"
                ? REVISION_DECISIONS.ACCEPT
                : REVISION_DECISIONS.REJECT,
          }).document;
    const resolveSteps = (
      steps: readonly { ids: readonly number[]; decision: RevisionDecision }[],
    ) => {
      let current = start;
      for (const { ids, decision } of steps)
        if (ids.length > 0)
          current = apply(current, {
            type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
            story: OP_STORIES.MAIN,
            revisionIds: ids,
            decision,
          }).document;
      return current;
    };
    for (let accepting = 0; accepting < 1 << count; accepting += 1) {
      const decide = (id: number) =>
        accepting & (1 << marks.indexOf(id))
          ? REVISION_DECISIONS.ACCEPT
          : REVISION_DECISIONS.REJECT;
      const accepted = {
        ids: marks.filter((id) => decide(id) === REVISION_DECISIONS.ACCEPT),
        decision: REVISION_DECISIONS.ACCEPT,
      };
      const rejected = {
        ids: marks.filter((id) => decide(id) === REVISION_DECISIONS.REJECT),
        decision: REVISION_DECISIONS.REJECT,
      };
      const atOnce = resolveSteps([accepted, rejected]);
      expect(resolveSteps([rejected, accepted])).toStrictEqual(atOnce);
      for (const order of orderings(marks))
        expect(
          resolveSteps(
            order.map((id) => ({ ids: [id], decision: decide(id) })),
          ),
        ).toStrictEqual(atOnce);
    }
  },
);

/** Every revision id a document's records carry. */
const revisionIdsIn = (value: unknown, ids = new Set<number>()) => {
  if (Array.isArray(value)) for (const item of value) revisionIdsIn(item, ids);
  else if (typeof value === "object" && value !== null)
    for (const [key, field] of Object.entries(value))
      if (
        key === "info" &&
        typeof field === "object" &&
        field !== null &&
        "id" in field &&
        typeof field.id === "number"
      )
        ids.add(field.id);
      else revisionIdsIn(field, ids);
  return ids;
};

test.each([1, 2, 3])(
  "every decision on the words and breaks of %i paragraphs inserted after a table that ends a content control resolves alike one at a time in any order and all at once",
  (count) => {
    const control: BlockContent = {
      type: "blockSdt",
      properties: { tag: "control" },
      content: [inCell([makeParagraph("00000001", "cell")])],
    };
    const document = documentOf([control, makeParagraph("00000005", "tail")]);
    const applied = apply(document, {
      ...insertion(
        Array.from({ length: count }, (_, index) =>
          makeParagraph(
            (16 + index).toString(16).padStart(8, "0"),
            `new ${index}`,
          ),
        ),
      ),
      at: { type: "afterTable", blockId: "00000001" },
    });
    const [held] = applied.document.package.document.content;
    const pairs = (held?.type === "blockSdt" ? held.content : []).flatMap(
      (block) =>
        block.type === "paragraph" &&
        block.pPrMark !== undefined &&
        block.content[0]?.type === "insertion"
          ? [{ mark: block.pPrMark.info.id, words: block.content[0].info.id }]
          : [],
    );
    expect(pairs).toHaveLength(count);
    const wordIds = pairs.map((pair) => pair.words);
    const insertedIds = new Set(
      (held?.type === "blockSdt" ? held.content : []).flatMap((block) =>
        block.type === "paragraph" && block.pPrMark !== undefined
          ? [block.paraId]
          : [],
      ),
    );
    const withInsertedIdsMasked = (resolved: Document): unknown =>
      JSON.parse(
        JSON.stringify(resolved, (key, value) =>
          key === "paraId" && insertedIds.has(value) ? "inserted" : value,
        ),
      );
    const revisions = pairs.flatMap(({ mark, words }) => [mark, words]);
    // Three paragraphs take their breaks first to last and last to first,
    // each paragraph's words before or after its own break.
    const orders =
      count < 3
        ? orderings(revisions)
        : [pairs, pairs.toReversed()].flatMap((sequence) =>
            Array.from({ length: 1 << count }, (_, after) =>
              sequence.flatMap(({ mark, words }, index) =>
                after & (1 << index) ? [mark, words] : [words, mark],
              ),
            ),
          );
    const resolveBatch = (
      current: Document,
      ids: readonly number[],
      decision: RevisionDecision,
    ) =>
      ids.length === 0
        ? current
        : apply(current, {
            type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
            story: OP_STORIES.MAIN,
            revisionIds: ids,
            decision,
          }).document;
    let compared = 0;
    for (let assignment = 0; assignment < 1 << revisions.length; assignment++) {
      const decide = (id: number) =>
        (assignment >> revisions.indexOf(id)) & 1
          ? REVISION_DECISIONS.REJECT
          : REVISION_DECISIONS.ACCEPT;
      const accepted = revisions.filter(
        (id) => decide(id) === REVISION_DECISIONS.ACCEPT,
      );
      const rejected = revisions.filter(
        (id) => decide(id) === REVISION_DECISIONS.REJECT,
      );
      const atOnce = resolveBatch(
        resolveBatch(applied.document, accepted, REVISION_DECISIONS.ACCEPT),
        rejected,
        REVISION_DECISIONS.REJECT,
      );
      expect(
        resolveBatch(
          resolveBatch(applied.document, rejected, REVISION_DECISIONS.REJECT),
          accepted,
          REVISION_DECISIONS.ACCEPT,
        ),
      ).toStrictEqual(atOnce);
      orders: for (const order of orders) {
        let current = applied.document;
        for (const [step, id] of order.entries()) {
          const resolved = apply(current, {
            type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
            story: OP_STORIES.MAIN,
            revisionIds: [id],
            decision: decide(id),
          });
          expect(undo(resolved)).toStrictEqual(current);
          current = resolved.document;
          // A rejected break rejoins its paragraph's pending words with the
          // alike pending words it meets into one change, the first one's id
          // standing for both, as a split's pieces rejoin. A later decision
          // on that change decides both, so where the two were to be decided
          // apart, single steps differ from deciding at once by design.
          const present = revisionIdsIn(current.package.document.content);
          const later = new Set(order.slice(step + 1));
          const splitDecisions = wordIds.some((absorbed, index) => {
            if (!later.has(absorbed) || present.has(absorbed)) return false;
            const survivor = wordIds
              .slice(0, index)
              .findLast((other) => present.has(other));
            return (
              survivor !== undefined && decide(survivor) !== decide(absorbed)
            );
          });
          if (splitDecisions) continue orders;
        }
        // Which inserted paragraph's id the control's last paragraph keeps
        // follows the order: rejecting an emptied paragraph's break removes
        // it while the one before it still waits, where at once that one
        // joins into it. Everything else is the same.
        expect(withInsertedIdsMasked(current)).toStrictEqual(
          withInsertedIdsMasked(atOnce),
        );
        compared += 1;
      }
    }
    expect(compared).toBeGreaterThan(0);
  },
  EVERY_ORDER_TIMEOUT_MS,
);

test.each([
  {
    name: "ends the body",
    content: (control: BlockContent) => [control],
  },
  {
    name: "ends a cell",
    content: (control: BlockContent) => [
      inCell([control]),
      makeParagraph("00000005", "tail"),
    ],
  },
])(
  "tracked insertion after a table in a content control that $name keeps its container's final paragraph",
  ({ content }) => {
    const document = documentOf(
      content({
        type: "blockSdt",
        properties: {},
        content: [inCell([makeParagraph("00000001", "cell")])],
      }),
    );
    const result = applyDocumentOp(document, {
      ...insertion([makeParagraph("00000010", "new")]),
      at: { type: "afterTable", blockId: "00000001" },
    });
    expect(result.isErr() && result.error.reason).toBe(
      DOCUMENT_OP_REFUSAL_REASONS.CONTAINER_FINAL_MARK,
    );
  },
);

test.each(["paragraph", "control"] as const)(
  "a tracked paragraph before a table in a body ending with a %s resolves either way",
  (ending) => {
    const final = makeParagraph("00000003", "");
    const original = documentOf([
      inCell([makeParagraph("00000001", "")]),
      makeParagraph("00000002", ""),
      ending === "paragraph"
        ? final
        : {
            type: "blockSdt",
            properties: { sdtType: "group", id: 1 },
            content: [final],
          },
    ]);
    const applied = apply(original, {
      ...insertion([makeParagraph("00000004", "revised")]),
      at: { type: "beforeTable", blockId: "00000001" },
    });
    expect(
      applied.document.package.document.content.map(({ type }) => type),
    ).toStrictEqual([
      "paragraph",
      "table",
      "paragraph",
      ending === "paragraph" ? "paragraph" : "blockSdt",
    ]);
    expect(resolve(applied, REVISION_DECISIONS.REJECT).document).toStrictEqual(
      original,
    );
    expect(
      resolve(applied, REVISION_DECISIONS.ACCEPT).document.package.document
        .content[0],
    ).toMatchObject({ paraId: "00000004", content: [{ type: "run" }] });
    expect(undo(applied)).toStrictEqual(original);
  },
);

test.each(["before", "beforeTable"] as const)(
  "tracked insertion %s preserves pasted controls and separate review identities",
  (type) => {
    const original =
      type === "before"
        ? baseDocument
        : documentOf([
            inCell([makeParagraph("00000001", "")]),
            makeParagraph("00000003", "kept"),
          ]);
    const imported = {
      ...makeParagraph("00000002", "pasted"),
      content: [
        {
          type: "inlineSdt",
          properties: { sdtType: "richText", id: 42 },
          content: [
            {
              type: "insertion",
              info: { ...stamp, id: 91 },
              content: [
                { type: "run", content: [{ type: "text", text: "pasted" }] },
              ],
            },
          ],
        },
      ],
    } satisfies Paragraph;
    const op = {
      ...insertion([imported]),
      at: { type, blockId: "00000001" },
    } satisfies InsertBlocksOp;
    const direct = apply(original, {
      type: op.type,
      story: op.story,
      at: op.at,
      blocks: op.blocks,
    });
    const pending = apply(original, op);
    expect(pending.revisions).not.toContain(91);
    expect(resolve(pending, REVISION_DECISIONS.ACCEPT).document).toStrictEqual(
      direct.document,
    );
    expect(resolve(pending, REVISION_DECISIONS.REJECT).document).toStrictEqual(
      original,
    );
    expect(undo(pending)).toStrictEqual(original);
  },
);

test.each([
  { location: "beside paragraph" },
  { location: "inside control" },
] as const)(
  "an empty paragraph replacement $location restores the exact authored snapshot",
  ({ location }) => {
    // An own `formatting` field, even one holding nothing.
    const removed: Paragraph = {
      ...makeParagraph("00000004", ""),
      formatting: {},
    };
    const control = {
      type: "blockSdt",
      properties: { sdtType: "group", id: 4 },
      content: [removed, makeParagraph("00000006", "inside")],
    } satisfies BlockContent;
    const original = documentOf([
      location === "inside control" ? control : removed,
      makeParagraph("00000005", "survivor"),
    ]);
    const result = apply(original, {
      type: DOCUMENT_OP_TYPES.REPLACE_BLOCKS,
      story: OP_STORIES.MAIN,
      expected: [removed],
      blocks: [],
    });
    expect(result.inverse.at(0)?.type).toBe(
      DOCUMENT_OP_TYPES.SET_CONTAINER_BLOCKS,
    );
    expect(undo(result)).toStrictEqual(original);
    const replay = applyDocumentOps(
      result.document,
      JSON.parse(JSON.stringify(result.inverse)),
    );
    if (replay.isErr()) throw replay.error;
    expect(replay.value.document).toStrictEqual(original);
  },
);

test("an empty replacement of a control's only paragraph is refused before it changes anything", () => {
  const removed = makeParagraph("00000004", "");
  const original = documentOf([
    {
      type: "blockSdt",
      properties: { sdtType: "group", id: 4 },
      content: [removed],
    },
    makeParagraph("00000005", "survivor"),
  ]);
  const result = applyDocumentOp(original, {
    type: DOCUMENT_OP_TYPES.REPLACE_BLOCKS,
    story: OP_STORIES.MAIN,
    expected: [removed],
    blocks: [],
  });
  expect(result.isErr() && result.error.reason).toBe(
    DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
  );
});
