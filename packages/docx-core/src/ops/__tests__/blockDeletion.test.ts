import { describe, expect, test } from "bun:test";

import type { BlockContent, Document, Paragraph } from "../../model/document";
import { applyDocumentOp, applyDocumentOps } from "../apply";
import { storyParagraphs } from "../blocks";
import { DOCUMENT_OP_REFUSAL_REASONS } from "../refusal";
import {
  type DeleteBlocksOp,
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  OP_STORIES,
  REVISION_DECISIONS,
  type RevisionDecision,
} from "../types";

const paragraph = (paraId: string, text: string): Paragraph => ({
  type: "paragraph",
  paraId,
  content:
    text === "" ? [] : [{ type: "run", content: [{ type: "text", text }] }],
});
const documentOf = (content: BlockContent[]): Document => ({
  package: { document: { content } },
});
const cell = (content: BlockContent[]): BlockContent => ({
  type: "table",
  rows: [{ type: "tableRow", cells: [{ type: "tableCell", content }] }],
});
const apply = (document: Document, op: DocumentOp) => {
  const result = applyDocumentOp(document, op);
  if (result.isErr()) throw result.error;
  return result.value;
};
const deletion = (blockIds: string[]): DeleteBlocksOp => ({
  type: DOCUMENT_OP_TYPES.DELETE_BLOCKS,
  story: OP_STORIES.MAIN,
  blockIds,
  revision: { id: 100, author: "Reviewer", date: "2026-02-03T04:05:06Z" },
  newIds: { revision: Array.from({ length: 32 }, (_, index) => 101 + index) },
});
const assertLaws = (document: Document, op: DeleteBlocksOp) => {
  const tracked = apply(document, op);
  const direct = apply(document, { ...op, revision: undefined });
  for (const decision of Object.values(REVISION_DECISIONS)) {
    const resolved = apply(tracked.document, {
      type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
      story: OP_STORIES.MAIN,
      revisionIds: tracked.revisions,
      decision,
    });
    expect(resolved.document).toStrictEqual(
      decision === REVISION_DECISIONS.ACCEPT ? direct.document : document,
    );
    const undo = applyDocumentOps(resolved.document, resolved.inverse);
    if (undo.isErr()) throw undo.error;
    expect(undo.value.document).toStrictEqual(tracked.document);
  }
  for (const edit of [tracked, direct]) {
    const undo = applyDocumentOps(edit.document, edit.inverse);
    if (undo.isErr()) throw undo.error;
    expect(undo.value.document).toStrictEqual(document);
  }
  return { tracked, direct };
};

test.each(["body", "cell"] as const)(
  "whole-container deletion retains the last paragraph in a %s",
  (container) => {
    const first = {
      ...paragraph("00000001", "first"),
      formatting: { alignment: "start" },
    } satisfies Paragraph;
    const final = {
      ...paragraph("00000002", "last"),
      formatting: { alignment: "end", runProperties: { bold: true } },
    } satisfies Paragraph;
    const content = [first, final];
    const document = documentOf(
      container === "body" ? content : [cell(content)],
    );
    const { tracked, direct } = assertLaws(
      document,
      deletion(["00000001", "00000002"]),
    );
    expect(
      storyParagraphs(direct.document.package.document).map(
        ({ paragraph: p }) => p,
      ),
    ).toEqual([{ ...final, content: [] }]);
    expect(
      storyParagraphs(tracked.document.package.document).at(-1)?.paragraph
        .pPrMark,
    ).toBeUndefined();
  },
);

test.each(["body", "cell"] as const)(
  "terminal deletion compensates properties and retains the final identity in a %s",
  (container) => {
    for (const text of ["", "final"]) {
      const previous = {
        ...paragraph("00000001", "kept"),
        formatting: { alignment: "center", runProperties: { italic: true } },
      } satisfies Paragraph;
      const final = {
        ...paragraph("00000002", text),
        formatting: { alignment: "end", runProperties: { bold: true } },
      } satisfies Paragraph;
      const content = [previous, final];
      const document = documentOf(
        container === "body" ? content : [cell(content)],
      );
      const { direct } = assertLaws(document, deletion(["00000002"]));
      expect(
        storyParagraphs(direct.document.package.document).at(0)?.paragraph,
      ).toEqual({
        ...final,
        content: previous.content,
        formatting: { alignment: "center", runProperties: { bold: true } },
      });
    }
  },
);

test("interior deletion preserves the following object and reports physical changes", () => {
  const first = paragraph("00000001", "remove");
  const following = paragraph("00000002", "kept");
  const { direct, tracked } = assertLaws(
    documentOf([first, following]),
    deletion(["00000001"]),
  );
  expect(direct.document.package.document.content.at(0)).toBe(following);
  expect(tracked.document.package.document.content.at(-1)).toBe(following);
  expect(direct.touched).toEqual({
    modified: [],
    inserted: [],
    removed: [first.paraId],
  });
  expect(tracked.touched).toEqual({
    modified: [first.paraId],
    inserted: [],
    removed: [],
  });
});

test.each([
  [
    "a bookmark the deleted paragraph holds",
    [
      { type: "bookmarkStart", id: 0, name: "_GoBack" },
      { type: "bookmarkEnd", id: 0 },
    ],
    [],
  ],
  [
    "the start of a bookmark ending in the paragraph after it",
    [{ type: "bookmarkStart", id: 3, name: "_Ref3" }],
    [{ type: "bookmarkEnd", id: 3 }],
  ],
] as const)(
  "direct deletion of a paragraph holding %s declares the paragraph that takes it",
  (_shape, inDeleted, inFollowing) => {
    const before = paragraph("00000001", "before");
    const removed: Paragraph = {
      type: "paragraph",
      paraId: "00000002",
      content: [
        ...inDeleted,
        { type: "run", content: [{ type: "text", text: "remove" }] },
      ],
    };
    const following: Paragraph = {
      type: "paragraph",
      paraId: "00000003",
      content: [
        ...inFollowing,
        { type: "run", content: [{ type: "text", text: "keep" }] },
      ],
    };
    const after = paragraph("00000004", "after");
    const { direct } = assertLaws(
      documentOf([before, removed, following, after]),
      deletion([removed.paraId ?? ""]),
    );
    const content = direct.document.package.document.content;
    expect(content.at(0)).toBe(before);
    expect(content.at(-1)).toBe(after);
    expect(content.at(1)).toStrictEqual({
      ...following,
      content: [...inDeleted, ...following.content],
    });
    expect(direct.touched).toEqual({
      modified: [following.paraId],
      inserted: [],
      removed: [removed.paraId],
    });
  },
);

test("inverse restores a retired comment whose surviving anchors followed the deleted block", () => {
  const removed: Paragraph = {
    ...paragraph("00000001", "remove"),
    content: [
      { type: "commentRangeStart", id: 7 },
      { type: "run", content: [{ type: "text", text: "remove" }] },
    ],
  };
  const following: Paragraph = {
    ...paragraph("00000002", "keep"),
    content: [
      { type: "commentRangeEnd", id: 7 },
      { type: "commentReference", id: 7 },
      { type: "run", content: [{ type: "text", text: "keep" }] },
    ],
  };
  const document: Document = {
    package: {
      document: {
        comments: [
          { id: 7, author: "Reviewer", anchorKind: "range", content: [] },
        ],
        content: [removed, following],
      },
    },
  };

  const edit = apply(document, {
    ...deletion([removed.paraId ?? ""]),
    revision: undefined,
  });
  expect(edit.document.package.document.comments).toBeUndefined();
  expect(edit.touched.modified).toEqual([following.paraId]);

  const restored = applyDocumentOps(edit.document, edit.inverse).unwrap()
    .document;
  expect(restored).toStrictEqual(document);
});

test.each([
  ["inside the paragraph", false],
  ["from the paragraph before", true],
] as const)(
  "a tracked deletion of a paragraph holding a comment's reference, its range %s, leaves the reference standing and resolves both ways",
  (_shape, startsBefore) => {
    const before: Paragraph = {
      ...paragraph("00000001", "kept"),
      content: [
        ...(startsBefore
          ? [{ type: "commentRangeStart" as const, id: 7 }]
          : []),
        { type: "run", content: [{ type: "text", text: "kept" }] },
      ],
    };
    const removed: Paragraph = {
      ...paragraph("00000002", "removed"),
      content: [
        ...(startsBefore
          ? []
          : [{ type: "commentRangeStart" as const, id: 7 }]),
        { type: "run", content: [{ type: "text", text: "removed" }] },
        { type: "commentRangeEnd", id: 7 },
        { type: "commentReference", id: 7 },
      ],
    };
    const document: Document = {
      package: {
        document: {
          comments: [
            { id: 7, author: "Reviewer", anchorKind: "range", content: [] },
          ],
          content: [before, removed, paragraph("00000003", "after")],
        },
      },
    };
    const { tracked } = assertLaws(document, deletion([removed.paraId ?? ""]));
    const marked = storyParagraphs(tracked.document.package.document).find(
      ({ paragraph: { paraId } }) => paraId === removed.paraId,
    );
    expect(marked?.paragraph.content).toContainEqual({
      type: "commentReference",
      id: 7,
    });
  },
);

test("a deletion keeps a comment's reference whose range ends in an earlier paragraph, tracked as directly", () => {
  const before: Paragraph = {
    ...paragraph("00000001", "kept"),
    content: [
      { type: "commentRangeStart", id: 7 },
      { type: "run", content: [{ type: "text", text: "kept" }] },
      { type: "commentRangeEnd", id: 7 },
    ],
  };
  const removed: Paragraph = {
    ...paragraph("00000002", "removed"),
    content: [
      { type: "run", content: [{ type: "text", text: "removed" }] },
      { type: "commentReference", id: 7 },
    ],
  };
  const document: Document = {
    package: {
      document: {
        comments: [
          { id: 7, author: "Reviewer", anchorKind: "range", content: [] },
        ],
        content: [before, removed, paragraph("00000003", "after")],
      },
    },
  };
  const { direct } = assertLaws(document, deletion([removed.paraId ?? ""]));
  expect(
    direct.document.package.document.comments?.map(({ id }) => id),
  ).toEqual([7]);
  expect(JSON.stringify(direct.document.package.document.content)).toContain(
    '"type":"commentReference"',
  );
});

test("a tracked deletion passes over a point comment's reference an earlier deletion already holds", () => {
  const removed: Paragraph = {
    ...paragraph("00000002", "live"),
    content: [
      {
        type: "deletion",
        info: { id: 50, author: "Other", date: "2026-01-01T00:00:00Z" },
        content: [
          { type: "run", content: [{ type: "text", text: "old" }] },
          { type: "commentReference", id: 9 },
        ],
      },
      { type: "run", content: [{ type: "text", text: "live" }] },
    ],
  };
  const document: Document = {
    package: {
      document: {
        comments: [
          { id: 9, author: "Reviewer", anchorKind: "point", content: [] },
        ],
        content: [
          paragraph("00000001", "kept"),
          removed,
          paragraph("00000003", "after"),
        ],
      },
    },
  };
  const result = applyDocumentOp(document, deletion([removed.paraId ?? ""]));
  expect(result.isOk()).toBe(true);
});

test("a sole paragraph after a table is cleared without marking its final break", () => {
  const table = cell([paragraph("00000003", "cell")]);
  const final = paragraph("00000001", "remove");
  const { direct } = assertLaws(
    documentOf([table, final]),
    deletion(["00000001"]),
  );
  expect(direct.document.package.document.content).toEqual([
    table,
    { ...final, content: [] },
  ]);
  expect(direct.document.package.document.content.at(0)).toBe(table);
});

test("deletion refuses invalid selections and boundaries", () => {
  const first = paragraph("00000001", "first");
  const final = paragraph("00000002", "final");
  const cases = [
    {
      document: documentOf([first, final]),
      ids: [],
      reason: DOCUMENT_OP_REFUSAL_REASONS.EMPTY_BLOCK_LIST,
    },
    {
      document: documentOf([first, final]),
      ids: ["00000002", "00000001"],
      reason: DOCUMENT_OP_REFUSAL_REASONS.NOT_ADJACENT,
    },
    {
      document: documentOf([first, cell([final])]),
      ids: ["00000001", "00000002"],
      reason: DOCUMENT_OP_REFUSAL_REASONS.NOT_ADJACENT,
    },
    {
      document: documentOf([first, cell([final])]),
      ids: ["00000001"],
      reason: DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
    },
  ];
  for (const { document, ids, reason } of cases) {
    const result = applyDocumentOp(document, deletion(ids));
    expect(result.isErr() && result.error.reason).toBe(reason);
  }
});

test("a tracked deletion removes a mark its own author's pending insertion added, joining the paragraph after it", () => {
  // Deleting over one's own insertion removes it; nothing records it deleted.
  const ownMark = (block: Paragraph): Paragraph => ({
    ...block,
    pPrMark: { kind: "ins", info: { id: 9, author: "Reviewer" } },
  });
  const texts = (document: Document) =>
    document.package.document.content.map((block) =>
      block.type === "paragraph"
        ? {
            id: block.paraId,
            mark: block.pPrMark?.kind,
            content: block.content.map((node) => node.type),
          }
        : block.type,
    );
  const cases = [
    {
      // The marked paragraph itself, with one after it.
      blocks: [
        ownMark(paragraph("00000001", "gone")),
        paragraph("00000002", "kept"),
      ],
      ids: ["00000001"],
      expected: [
        { id: "00000002", mark: undefined, content: ["deletion", "run"] },
      ],
      removed: ["00000001"],
    },
    {
      // The paragraph ending the body, after the marked one.
      blocks: [
        ownMark(paragraph("00000001", "kept")),
        paragraph("00000002", "gone"),
      ],
      ids: ["00000002"],
      expected: [
        { id: "00000002", mark: undefined, content: ["run", "deletion"] },
      ],
      removed: ["00000001"],
    },
  ];
  for (const { blocks, ids, expected, removed } of cases) {
    const document = documentOf(blocks);
    const applied = applyDocumentOp(document, deletion(ids));
    if (applied.isErr()) throw applied.error;
    expect(texts(applied.value.document)).toEqual(expected);
    expect(applied.value.touched.removed).toEqual(removed);
    const accepted = apply(applied.value.document, {
      type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
      story: OP_STORIES.MAIN,
      revisionIds: applied.value.revisions,
      decision: REVISION_DECISIONS.ACCEPT,
    });
    expect(
      storyParagraphs(accepted.document.package.document).map(
        ({ paragraph: block }) => block.content,
      ),
    ).toEqual(
      storyParagraphs(
        apply(document, { ...deletion(ids), revision: undefined }).document
          .package.document,
      ).map(({ paragraph: block }) => block.content),
    );
    const undo = applyDocumentOps(
      applied.value.document,
      applied.value.inverse,
    );
    if (undo.isErr()) throw undo.error;
    expect(undo.value.document).toStrictEqual(document);
  }
  // One with no paragraph after it stays refused.
  for (const blocks of [
    [
      ownMark(paragraph("00000001", "gone")),
      cell([paragraph("00000003", "cell")]),
      paragraph("00000002", "kept"),
    ],
  ]) {
    const refused = applyDocumentOp(documentOf(blocks), deletion(["00000001"]));
    expect(refused.isErr() && refused.error.reason).toBe(
      DOCUMENT_OP_REFUSAL_REASONS.REVISION_CONFLICT,
    );
  }
});

test("tracked deletion of a paragraph a table follows marks its own mark, and resolves both ways", () => {
  const first = paragraph("00000001", "first");
  const target = paragraph("00000002", "target");
  const after = paragraph("00000004", "after");
  const document = documentOf([
    first,
    target,
    cell([paragraph("00000003", "cell")]),
    after,
  ]);
  const applied = applyDocumentOp(document, deletion(["00000002"]));
  if (applied.isErr()) throw applied.error;
  const [, deleted] = applied.value.document.package.document.content;
  expect(deleted?.type === "paragraph" && deleted.pPrMark?.kind).toBe("del");
  const resolve = (decision: RevisionDecision) => {
    const resolved = applyDocumentOp(applied.value.document, {
      type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
      story: OP_STORIES.MAIN,
      revisionIds: applied.value.revisions,
      decision,
    });
    if (resolved.isErr()) throw resolved.error;
    return resolved.value.document.package.document.content.map((block) =>
      block.type === "paragraph" ? block.paraId : block.type,
    );
  };
  expect(resolve(REVISION_DECISIONS.REJECT)).toEqual([
    "00000001",
    "00000002",
    "table",
    "00000004",
  ]);
  expect(resolve(REVISION_DECISIONS.ACCEPT)).toEqual([
    "00000001",
    "table",
    "00000004",
  ]);
});

test.each([
  [
    "a block content control",
    {
      type: "blockSdt",
      properties: { sdtType: "richText", id: 5 },
      content: [paragraph("00000003", "inside")],
    },
    "blockSdt",
  ],
  [
    "a preserved block",
    { type: "preservedBlock", xml: "<w:altChunk/>" },
    "preservedBlock",
  ],
] as const satisfies readonly (readonly [string, BlockContent, string])[])(
  "tracked deletion of a paragraph %s follows marks its own mark, and resolves both ways",
  (_name, block, kind) => {
    const document = documentOf([
      paragraph("00000001", "first"),
      paragraph("00000002", "target"),
      block,
      paragraph("00000004", "after"),
    ]);
    const tracked = apply(document, deletion(["00000002"]));
    const [, deleted] = tracked.document.package.document.content;
    expect(deleted?.type === "paragraph" && deleted.pPrMark?.kind).toBe("del");
    const resolve = (decision: RevisionDecision) =>
      apply(tracked.document, {
        type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
        story: OP_STORIES.MAIN,
        revisionIds: tracked.revisions,
        decision,
      });
    const shape = (edit: { document: Document }) =>
      edit.document.package.document.content.map((item) =>
        item.type === "paragraph" ? item.paraId : item.type,
      );
    const rejected = resolve(REVISION_DECISIONS.REJECT);
    expect(rejected.document).toStrictEqual(document);
    const accepted = resolve(REVISION_DECISIONS.ACCEPT);
    expect(shape(accepted)).toEqual(["00000001", kind, "00000004"]);
    for (const resolved of [rejected, accepted])
      expect(
        applyDocumentOps(resolved.document, resolved.inverse).unwrap().document,
      ).toStrictEqual(tracked.document);
    expect(
      applyDocumentOps(tracked.document, tracked.inverse).unwrap().document,
    ).toStrictEqual(document);
    // Accepting its mark and its words one at a time, in either order, ends
    // the same way where the block can take the words in the meantime. A
    // preserved block takes nothing in: there the mark accepted first keeps
    // the paragraph, emptied once its words go.
    expect(tracked.revisions).toHaveLength(2);
    for (const order of [tracked.revisions, tracked.revisions.toReversed()]) {
      let inTurn = tracked.document;
      for (const id of order)
        inTurn = apply(inTurn, {
          type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
          story: OP_STORIES.MAIN,
          revisionIds: [id],
          decision: REVISION_DECISIONS.ACCEPT,
        }).document;
      if (block.type === "blockSdt" || order[0] !== tracked.revisions[0])
        expect(inTurn).toStrictEqual(accepted.document);
      else
        expect(
          inTurn.package.document.content.map((item) =>
            item.type === "paragraph" ? item.paraId : item.type,
          ),
        ).toEqual(["00000001", "00000002", kind, "00000004"]);
    }
    // A deleted mark whose paragraph still holds words keeps the paragraph.
    const kept = documentOf([
      paragraph("00000001", "first"),
      {
        ...paragraph("00000002", "target"),
        pPrMark: { kind: "del", info: { id: 9, author: "Other" } },
      },
      block,
      paragraph("00000004", "after"),
    ]);
    const keptAccepted = apply(kept, {
      type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
      story: OP_STORIES.MAIN,
      revisionIds: [9],
      decision: REVISION_DECISIONS.ACCEPT,
    });
    expect(keptAccepted.document.package.document.content).toStrictEqual([
      paragraph("00000001", "first"),
      paragraph("00000002", "target"),
      block,
      paragraph("00000004", "after"),
    ]);
  },
);

test("accepting the marks of paragraphs before a block content control at once equals accepting them in turn", () => {
  const sdt: BlockContent = {
    type: "blockSdt",
    properties: { sdtType: "group", id: 1 },
    content: [paragraph("00000003", "")],
  };
  for (const words of ["", "x"]) {
    const document = documentOf([
      {
        ...paragraph("00000001", ""),
        content:
          words === ""
            ? []
            : [
                {
                  type: "insertion",
                  info: { id: 4, author: "C" },
                  content: [
                    { type: "run", content: [{ type: "text", text: words }] },
                  ],
                },
              ],
        formatting: { styleId: "Heading2" },
        pPrMark: { kind: "del", info: { id: 2, author: "C" } },
      },
      {
        ...paragraph("00000002", ""),
        pPrMark: { kind: "del", info: { id: 1, author: "A" } },
      },
      sdt,
      paragraph("00000004", "after"),
    ]);
    const accept = (current: Document, revisionIds: number[]) =>
      apply(current, {
        type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
        story: OP_STORIES.MAIN,
        revisionIds,
        decision: REVISION_DECISIONS.ACCEPT,
      }).document;
    const atOnce = accept(document, [1, 2, 4]);
    expect(accept(accept(accept(document, [2]), [1]), [4])).toStrictEqual(
      atOnce,
    );
    // The first paragraph joins the second, which keeps what stays, or goes
    // with it when nothing does.
    expect(atOnce.package.document.content).toStrictEqual(
      words === ""
        ? [sdt, paragraph("00000004", "after")]
        : [paragraph("00000002", words), sdt, paragraph("00000004", "after")],
    );
  }
});

test("tracked deletion keeps existing inline review and refuses missing physical ids", () => {
  const first = paragraph("00000001", "first");
  const final = paragraph("00000002", "final");
  const reviewed: Paragraph = {
    ...first,
    content: [
      {
        type: "deletion",
        info: { id: 9, author: "A" },
        content: [{ type: "run", content: [{ type: "text", text: "first" }] }],
      },
    ],
  };
  // Content an earlier deletion already took stays under it; only the
  // paragraph's mark is recorded deleted.
  const source = documentOf([reviewed, final]);
  const tracked = apply(source, deletion(["00000001"]));
  const [kept] = storyParagraphs(tracked.document.package.document);
  expect(kept?.paragraph.content).toStrictEqual(reviewed.content);
  expect(kept?.paragraph.pPrMark?.kind).toBe("del");
  expect(
    apply(tracked.document, {
      type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
      story: OP_STORIES.MAIN,
      revisionIds: tracked.revisions,
      decision: REVISION_DECISIONS.REJECT,
    }).document,
  ).toStrictEqual(source);
  const missing = applyDocumentOp(documentOf([first, final]), {
    ...deletion(["00000001"]),
    newIds: {},
  });
  expect(missing.isErr() && missing.error.reason).toBe(
    DOCUMENT_OP_REFUSAL_REASONS.NEEDS_NEW_IDS,
  );
});

test.each(["body", "cell"] as const)(
  "terminal deletion folds a pending paragraph review without replacing its original baseline in a %s",
  (container) => {
    const pending = {
      type: "paragraphPropertyChange" as const,
      info: { id: 7, author: "Original author" },
      previousFormatting: { alignment: "start" as const },
    };
    const content = [
      { ...paragraph("00000001", "kept"), formatting: { alignment: "center" } },
      {
        ...paragraph("00000002", "deleted"),
        formatting: { alignment: "end" },
        propertyChanges: [pending],
      },
    ] satisfies Paragraph[];
    const original = documentOf(
      container === "body" ? content : [cell(content)],
    );
    const tracked = apply(original, deletion(["00000002"]));
    const survivor = storyParagraphs(tracked.document.package.document).at(
      -1,
    )?.paragraph;
    expect(survivor?.formatting).toEqual({ alignment: "center" });
    expect(survivor?.propertyChanges).toEqual([
      {
        ...pending,
        info: { id: 7, author: "Reviewer", date: "2026-02-03T04:05:06Z" },
        currentFormatting: { alignment: "center" },
      },
    ]);
    const rejected = apply(tracked.document, {
      type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
      story: OP_STORIES.MAIN,
      revisionIds: [7],
      decision: REVISION_DECISIONS.REJECT,
    });
    expect(
      storyParagraphs(rejected.document.package.document).at(-1)?.paragraph
        .formatting,
    ).toEqual({ alignment: "start" });
    expect(
      applyDocumentOps(tracked.document, tracked.inverse).unwrap().document,
    ).toStrictEqual(original);
    expect(
      applyDocumentOps(rejected.document, rejected.inverse).unwrap().document,
    ).toStrictEqual(tracked.document);
  },
);

test("tracked deletion before a table under review empties the paragraph, one before no receiving cell is refused, and each resolves alone", () => {
  const foreign = { id: 9, author: "Other", date: "2026-01-01T00:00:00Z" };
  const inner = cell([paragraph("00000003", "cell")]);
  if (inner.type !== "table") throw new Error("A cell fixture is a table.");
  const row = inner.rows[0];
  if (row === undefined) throw new Error("A cell fixture has a row.");
  const [firstCell] = row.cells;
  if (firstCell === undefined) throw new Error("A cell fixture has a cell.");
  const underReview: BlockContent[] = [
    // The first row is a pending insertion or deletion.
    {
      ...inner,
      rows: [
        {
          ...row,
          structuralChange: { type: "tableRowInsertion", info: foreign },
        },
      ],
    },
    {
      ...inner,
      rows: [
        {
          ...row,
          structuralChange: { type: "tableRowDeletion", info: foreign },
        },
      ],
    },
    // The first cell is a pending structural change.
    {
      ...inner,
      rows: [
        {
          ...row,
          cells: [
            {
              ...firstCell,
              structuralChange: { type: "tableCellInsertion", info: foreign },
            },
          ],
        },
      ],
    },
  ];
  // The words have nowhere to run in, but the deletion takes them all: each
  // decision on it alone resolves, accepting without crossing the table.
  for (const table of underReview) {
    const document = documentOf([
      paragraph("00000001", "first"),
      paragraph("00000002", "target"),
      table,
      paragraph("00000004", "after"),
    ]);
    const tracked = apply(document, deletion(["00000002"]));
    for (const decision of Object.values(REVISION_DECISIONS)) {
      const resolved = applyDocumentOp(tracked.document, {
        type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
        story: OP_STORIES.MAIN,
        revisionIds: tracked.revisions,
        decision,
      });
      if (resolved.isErr()) throw resolved.error;
    }
  }
  const blocked: BlockContent[] = [
    // The first cell opens with no paragraph.
    {
      ...inner,
      rows: [{ ...row, cells: [{ ...firstCell, content: [] }] }],
    },
  ];
  for (const table of blocked) {
    const document = documentOf([
      paragraph("00000001", "first"),
      paragraph("00000002", "target"),
      table,
      paragraph("00000004", "after"),
    ]);
    const refused = applyDocumentOp(document, deletion(["00000002"]));
    expect(refused.isErr() && refused.error.reason).toBe(
      DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
    );
  }
  // A table whose first cell opens with a paragraph receives the mark, and
  // each decision on the deletion alone resolves.
  const document = documentOf([
    paragraph("00000001", "first"),
    paragraph("00000002", "target"),
    inner,
    paragraph("00000004", "after"),
  ]);
  const tracked = apply(document, deletion(["00000002"]));
  for (const decision of Object.values(REVISION_DECISIONS)) {
    const resolved = applyDocumentOp(tracked.document, {
      type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
      story: OP_STORIES.MAIN,
      revisionIds: tracked.revisions,
      decision,
    });
    if (resolved.isErr()) throw resolved.error;
  }
});

const resolveTracked = (
  document: Document,
  revisionIds: number[],
  decision: RevisionDecision,
) =>
  apply(document, {
    type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
    story: OP_STORIES.MAIN,
    revisionIds,
    decision,
  }).document;

test("a tracked deletion of a paragraph whose mark a pending deletion already took keeps that record", () => {
  const joined: Paragraph = {
    ...paragraph("00000001", "Kept item."),
    pPrMark: {
      kind: "del",
      info: { id: 5, author: "Other", date: "2026-01-02T03:04:05Z" },
    },
  };
  const document = documentOf([
    joined,
    paragraph("00000002", "Joined."),
    paragraph("00000003", "Closing."),
  ]);
  const tracked = apply(document, deletion(["00000001"]));
  const [first] = storyParagraphs(tracked.document.package.document);
  expect(first?.paragraph.pPrMark).toStrictEqual(joined.pPrMark);
  expect(first?.paragraph.content).toStrictEqual([
    {
      type: "deletion",
      info: { id: 100, author: "Reviewer", date: "2026-02-03T04:05:06Z" },
      content: [
        { type: "run", content: [{ type: "text", text: "Kept item." }] },
      ],
    },
  ]);
  expect(
    resolveTracked(tracked.document, [100], REVISION_DECISIONS.REJECT),
  ).toStrictEqual(document);
  const undo = applyDocumentOps(tracked.document, tracked.inverse);
  if (undo.isErr()) throw undo.error;
  expect(undo.value.document).toStrictEqual(document);
});

test("a tracked deletion of a paragraph holding another author's insertion deletes it in place", () => {
  const document = documentOf([
    {
      ...paragraph("00000001", "Kept"),
      content: [
        { type: "run", content: [{ type: "text", text: "Kept" }] },
        {
          type: "insertion",
          info: { id: 7, author: "Other", date: "2026-01-02T03:04:05Z" },
          content: [
            { type: "run", content: [{ type: "text", text: " item." }] },
          ],
        },
      ],
    },
    paragraph("00000002", "Closing."),
  ]);
  const tracked = apply(document, deletion(["00000001"]));
  expect(
    resolveTracked(
      tracked.document,
      tracked.revisions,
      REVISION_DECISIONS.REJECT,
    ),
  ).toStrictEqual(document);
  const accepted = resolveTracked(
    tracked.document,
    [...tracked.revisions, 7],
    REVISION_DECISIONS.ACCEPT,
  );
  expect(
    storyParagraphs(accepted.package.document).map(
      ({ paragraph: located }) => located.paraId,
    ),
  ).toStrictEqual(["00000002"]);
  const undo = applyDocumentOps(tracked.document, tracked.inverse);
  if (undo.isErr()) throw undo.error;
  expect(undo.value.document).toStrictEqual(document);
});

test("a tracked deletion of the paragraph whose mark the last paragraph's deletion took for its properties is refused", () => {
  const info = { id: 5, author: "Other", date: "2026-01-02T03:04:05Z" };
  const numbered = {
    numPr: { kind: "reference", numId: 2, ilvl: 0 },
  } as const;
  // Deleting "Closing." deleted the item's mark and gave the closing
  // paragraph the item's properties; deleting the item too would leave them
  // on a paragraph nothing of the item reaches.
  const item: Paragraph = {
    ...paragraph("00000002", "Item."),
    formatting: numbered,
    pPrMark: { kind: "del", info },
  };
  const closing = (rotated: boolean): Paragraph => ({
    type: "paragraph",
    paraId: "00000003",
    ...(rotated
      ? {
          formatting: numbered,
          propertyChanges: [
            {
              type: "paragraphPropertyChange",
              info: { ...info, id: 6 },
              currentFormatting: numbered,
            },
          ],
        }
      : {}),
    content: [
      {
        type: "deletion",
        info: { ...info, id: 7 },
        content: [
          { type: "run", content: [{ type: "text", text: "Closing." }] },
        ],
      },
    ],
  });
  const refused = applyDocumentOp(
    documentOf([paragraph("00000001", "Kept."), item, closing(true)]),
    deletion(["00000002"]),
  );
  expect(refused.isErr() && refused.error.reason).toBe(
    DOCUMENT_OP_REFUSAL_REASONS.REVISION_CONFLICT,
  );
  // A closing paragraph that kept its own properties leaves the earlier
  // record standing.
  const tracked = apply(
    documentOf([paragraph("00000001", "Kept."), item, closing(false)]),
    deletion(["00000002"]),
  );
  expect(
    storyParagraphs(tracked.document.package.document).at(1)?.paragraph.pPrMark,
  ).toStrictEqual(item.pPrMark);
});

test("a tracked deletion of an empty paragraph whose mark a pending deletion already took records nothing and is refused", () => {
  const deleted: Paragraph = {
    ...paragraph("00000001", ""),
    pPrMark: {
      kind: "del",
      info: { id: 5, author: "Other", date: "2026-01-02T03:04:05Z" },
    },
  };
  const refused = applyDocumentOp(
    documentOf([deleted, paragraph("00000002", "After.")]),
    deletion(["00000001"]),
  );
  expect(refused.isErr() && refused.error.reason).toBe(
    DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
  );
});

test.each(["body", "cell"] as const)(
  "a tracked deletion of the empty paragraph alone in a %s applies as its direct deletion does, recording nothing",
  (container) => {
    const lone: Paragraph = {
      ...paragraph("00000001", ""),
      formatting: { runProperties: { italic: false } },
    };
    const document = documentOf(
      container === "body" ? [lone] : [cell([lone]), paragraph("00000002", "")],
    );
    const { tracked, direct } = assertLaws(document, deletion(["00000001"]));
    expect(tracked.revisions).toEqual([]);
    expect(tracked.document).toStrictEqual(document);
    expect(direct.document).toStrictEqual(document);
  },
);

test("a tracked deletion of another author's inserted paragraph records its deletion beside the insertion", () => {
  const inserted: Paragraph = {
    ...paragraph("00000001", "Added."),
    pPrMark: {
      kind: "ins",
      info: { id: 5, author: "Other", date: "2026-01-02T03:04:05Z" },
    },
  };
  const document = documentOf([inserted, paragraph("00000002", "Closing.")]);
  const tracked = apply(document, deletion(["00000001"]));
  const [first] = storyParagraphs(tracked.document.package.document);
  expect(first?.paragraph.pPrMark).toStrictEqual({
    ...inserted.pPrMark,
    deletion: {
      info: { id: 100, author: "Reviewer", date: "2026-02-03T04:05:06Z" },
    },
  });
  expect(
    resolveTracked(tracked.document, [100, 101], REVISION_DECISIONS.REJECT),
  ).toStrictEqual(document);
  const accepted = resolveTracked(
    tracked.document,
    [100, 101, 5],
    REVISION_DECISIONS.ACCEPT,
  );
  expect(
    storyParagraphs(accepted.package.document).map(
      ({ paragraph: located }) => located.paraId,
    ),
  ).toStrictEqual(["00000002"]);
  const undo = applyDocumentOps(tracked.document, tracked.inverse);
  if (undo.isErr()) throw undo.error;
  expect(undo.value.document).toStrictEqual(document);
});

// The pending record taking the paragraph whole, and whether the final
// deletion is a move's source.
test.each(
  (["del", "moveFrom"] as const).flatMap((earlier) =>
    [false, true].map((move) => ({ earlier, move })),
  ),
)(
  "a tracked deletion of the last paragraph after one a pending deletion takes whole retires the break before that one ($earlier before, move $move)",
  ({ earlier, move }) => {
    const wholly: Paragraph = {
      ...paragraph("00000002", ""),
      content: [
        {
          type: earlier === "del" ? "deletion" : "moveFrom",
          info: { id: 5, author: "Other", date: "2026-01-02T03:04:05Z" },
          content: [{ type: "run", content: [{ type: "text", text: "Gone" }] }],
        },
      ],
      pPrMark: {
        kind: earlier,
        info: { id: 4, author: "Other", date: "2026-01-02T03:04:05Z" },
      },
    };
    const heading: Paragraph = {
      ...paragraph("00000001", "Kept"),
      formatting: { alignment: "center" },
    };
    const document = documentOf([
      heading,
      wholly,
      paragraph("00000003", "Last"),
    ]);
    const tracked = apply(document, {
      ...deletion(["00000003"]),
      ...(move ? { move: true as const } : {}),
    });
    const marks = storyParagraphs(tracked.document.package.document).map(
      ({ paragraph: located }) => located.pPrMark?.kind,
    );
    expect(marks).toEqual([move ? "moveFrom" : "del", earlier, undefined]);
    const accepted = resolveTracked(
      tracked.document,
      [4, 5, ...tracked.revisions],
      REVISION_DECISIONS.ACCEPT,
    );
    const [only, ...rest] = storyParagraphs(accepted.package.document);
    expect(rest).toEqual([]);
    expect(only?.paragraph.paraId).toBe("00000003");
    expect(only?.paragraph.formatting?.alignment).toBe("center");
    expect(JSON.stringify(only?.paragraph.content)).toContain("Kept");
    const undo = applyDocumentOps(tracked.document, tracked.inverse);
    if (undo.isErr()) throw undo.error;
    expect(undo.value.document).toStrictEqual(document);
  },
);

describe("deleting a paragraph that ends a section", () => {
  const sectioned = (paraId: string, text: string, pageWidth: number) =>
    ({
      ...paragraph(paraId, text),
      sectionProperties: { pageWidth },
    }) satisfies Paragraph;
  const document = documentOf([
    sectioned("00000001", "first section", 10_000),
    sectioned("00000002", "second section", 12_000),
    paragraph("00000003", "final section"),
  ]);
  const pageWidths = ({ package: { document: body } }: Document) =>
    body.content.flatMap((block) =>
      block.type === "paragraph" && block.sectionProperties !== undefined
        ? [block.sectionProperties.pageWidth]
        : [],
    );

  // The break is the deleted mark's: taking it joins the sections on either
  // side under the properties of the one after it; rejecting keeps both.
  test.each([
    { ids: ["00000001"], accepted: [12_000] },
    { ids: ["00000002"], accepted: [10_000] },
    { ids: ["00000001", "00000002"], accepted: [] },
    { ids: ["00000003"], accepted: [10_000] },
  ])("deleting $ids takes the break its mark carries", ({ ids, accepted }) => {
    const { tracked, direct } = assertLaws(document, deletion(ids));
    expect(pageWidths(direct.document)).toEqual(accepted);
    expect(pageWidths(tracked.document)).toEqual([10_000, 12_000]);
  });

  // Taking back one's own added break would have to decide where its section
  // goes; the deletion recorded beside the insertion leaves that to the
  // insertion's own resolution.
  test("one's own added mark that ends a section records the deletion beside it", () => {
    const added = {
      ...sectioned("00000001", "added", 10_000),
      pPrMark: { kind: "ins", info: { id: 7, author: "Reviewer" } },
    } satisfies Paragraph;
    const own = documentOf([
      paragraph("00000004", "opening"),
      added,
      paragraph("00000003", "final section"),
    ]);
    const tracked = apply(own, deletion(["00000001"]));
    const marked = tracked.document.package.document.content.at(1);
    expect(marked?.type === "paragraph" && marked.pPrMark).toEqual({
      ...added.pPrMark,
      deletion: { info: expect.objectContaining({ author: "Reviewer" }) },
    });
    expect(pageWidths(tracked.document)).toEqual([10_000]);
    const rejected = apply(tracked.document, {
      type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
      story: OP_STORIES.MAIN,
      revisionIds: tracked.revisions,
      decision: REVISION_DECISIONS.REJECT,
    });
    expect(rejected.document).toStrictEqual(own);
  });
});

// A content control in a deleted paragraph is deleted whole with it: tracked,
// the deletion holds it, accepting removes it as the direct deletion does, and
// rejecting gives it back.
test("a tracked deletion of a paragraph holding a content control meets the laws", () => {
  const holder = {
    type: "paragraph",
    paraId: "00000002",
    content: [
      { type: "run", content: [{ type: "text", text: "Pay " }] },
      {
        type: "inlineSdt",
        properties: { sdtType: "richText", id: 7, tag: "amount" },
        content: [{ type: "run", content: [{ type: "text", text: "100" }] }],
      },
      { type: "run", content: [{ type: "text", text: " euros." }] },
    ],
  } satisfies Paragraph;
  const document = documentOf([
    paragraph("00000001", "Before."),
    holder,
    paragraph("00000003", "After."),
  ]);
  const { tracked } = assertLaws(document, deletion(["00000002"]));
  expect(JSON.stringify(tracked.document)).toContain('"tag":"amount"');
});
