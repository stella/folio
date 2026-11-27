/**
 * Tracked changes at the edge of what their records hold: a section-ending
 * mark resolved across range boundaries standing before the next paragraph
 * and copied range boundaries nested deeper than any fixed limit record,
 * resolve and undo. A tracked deletion of a transparent wrapper holding a
 * comment anchor meets the anchor as a bare one, and a pasted paragraph
 * whose copied deletion of its break holds a cut of its own is refused as a
 * suggestion while the direct paste applies.
 */
import { describe, expect, test } from "bun:test";

import type {
  BlockContent,
  Document,
  Paragraph,
  ParagraphContent,
  ParagraphMarkChange,
  Run,
  TrackedRunContent,
} from "../../model/document";
import { applyDocumentOp, applyDocumentOps } from "../apply";
import {
  type CopyOriginAncestorList,
  copyOriginAncestorsOf,
} from "../copyOriginAncestors";
import { compileEditorIntent, type EditorIntent } from "../editorIntent";
import { childNodes, type InlineNode } from "../leaves";
import { DOCUMENT_OP_REFUSAL_REASONS } from "../refusal";
import {
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  OP_STORIES,
  REVISION_DECISIONS,
  type RevisionDecision,
} from "../types";

const run = (text: string): Run => ({
  type: "run",
  content: [{ type: "text", text }],
});

const paragraph = (paraId: string, text: string): Paragraph => ({
  type: "paragraph",
  paraId,
  content: [run(text)],
});

const documentOf = (content: BlockContent[]): Document => ({
  package: { document: { content } },
});

const apply = (document: Document, ops: readonly DocumentOp[]) => {
  const result = applyDocumentOps(document, ops);
  if (result.isErr()) throw result.error;
  return result.value;
};

const resolve = (
  document: Document,
  revisionIds: readonly number[],
  decision: RevisionDecision,
) =>
  apply(document, [
    {
      type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
      story: OP_STORIES.MAIN,
      revisionIds,
      decision,
    },
  ]);

const EARLIER = { author: "Earlier", date: "2026-01-01T00:00:00Z" };

describe("a section-ending mark resolved across a standing range boundary", () => {
  const SECTION = { pageWidth: 11_906 };
  const pending = (): Document =>
    documentOf([
      {
        ...paragraph("00000001", "Alpha"),
        sectionProperties: SECTION,
        pPrMark: { kind: "del", info: { id: 7, ...EARLIER } },
      },
      { type: "bookmarkStart", id: 2, name: "Later" },
      paragraph("00000002", "Beta"),
      { type: "bookmarkEnd", id: 2 },
    ]);

  test("accepting joins the sections and puts the boundary where the resolved section ended", () => {
    const document = pending();
    const accepted = resolve(document, [7], REVISION_DECISIONS.ACCEPT);
    expect(accepted.document).toStrictEqual(
      documentOf([
        {
          type: "paragraph",
          paraId: "00000002",
          content: [
            run("Alpha"),
            { type: "bookmarkStart", id: 2, name: "Later" },
            run("Beta"),
          ],
        },
        { type: "bookmarkEnd", id: 2 },
      ]),
    );
    expect(apply(accepted.document, accepted.inverse).document).toStrictEqual(
      document,
    );
  });

  test("rejecting keeps the break, its section and the boundary between the blocks", () => {
    const document = pending();
    const rejected = resolve(document, [7], REVISION_DECISIONS.REJECT);
    const [first, ...rest] = document.package.document.content;
    if (first?.type !== "paragraph")
      throw new TypeError("The first block is a paragraph.");
    const { pPrMark: _mark, ...kept } = first;
    expect(rejected.document).toStrictEqual(documentOf([kept, ...rest]));
    expect(apply(rejected.document, rejected.inverse).document).toStrictEqual(
      document,
    );
  });

  test("a later section's break stays with the paragraph that ends it", () => {
    const document = documentOf([
      {
        ...paragraph("00000001", "Alpha"),
        sectionProperties: SECTION,
        pPrMark: { kind: "del", info: { id: 7, ...EARLIER } },
      },
      { type: "commentRangeStart", id: 4 },
      { ...paragraph("00000002", "Beta"), sectionProperties: { pageWidth: 1 } },
      { type: "commentRangeEnd", id: 4 },
      paragraph("00000003", "Tail"),
    ]);
    const accepted = resolve(document, [7], REVISION_DECISIONS.ACCEPT);
    const [joined] = accepted.document.package.document.content;
    expect(joined).toStrictEqual({
      type: "paragraph",
      paraId: "00000002",
      sectionProperties: { pageWidth: 1 },
      content: [
        run("Alpha"),
        { type: "commentRangeStart", id: 4 },
        run("Beta"),
      ],
    });
    expect(apply(accepted.document, accepted.inverse).document).toStrictEqual(
      document,
    );
  });
});

/** `depth` insertions nested one in another around `inner`. */
const nested = (
  depth: number,
  inner: TrackedRunContent[],
): TrackedRunContent[] => {
  let content = inner;
  for (let level = depth; level > 0; level -= 1)
    content = [
      { type: "insertion", info: { id: 1_000 + level, ...EARLIER }, content },
    ];
  return content;
};

const bookmarkEnds = (
  nodes: readonly InlineNode[],
): Extract<InlineNode, { type: "bookmarkEnd" }>[] =>
  nodes.flatMap((node) =>
    node.type === "bookmarkEnd" ? [node] : bookmarkEnds(childNodes(node) ?? []),
  );

const contentOf = (document: Document): ParagraphContent[] => {
  const [first] = document.package.document.content;
  if (first?.type !== "paragraph")
    throw new TypeError("The first block is a paragraph.");
  return first.content;
};

const copyOrigins = (document: Document): CopyOriginAncestorList[] =>
  bookmarkEnds(contentOf(document)).flatMap((boundary) => {
    const origins = copyOriginAncestorsOf(boundary);
    return origins === undefined ? [] : [origins];
  });

describe("a copied range boundary nested deeper than any fixed limit", () => {
  // A copy every one of 256 nested insertions owns; copying it once more
  // records a 257th owner.
  const DEPTH = 256;
  const ORIGINS: CopyOriginAncestorList = [
    0,
    ...Array.from({ length: DEPTH - 1 }, (_, index) => index + 1),
  ];
  const source = (): Document =>
    documentOf([
      {
        type: "paragraph",
        paraId: "00000001",
        content: [
          { type: "bookmarkStart", id: 3, name: "Deep" },
          { type: "bookmarkEnd", id: 3 },
          ...nested(DEPTH, [
            run("a"),
            { type: "bookmarkEnd", id: 3, copyOriginAncestors: ORIGINS },
          ]),
          run("z"),
        ],
      },
    ]);
  const paste = (document: Document): DocumentOp => ({
    type: DOCUMENT_OP_TYPES.INSERT_CONTENT,
    at: { story: OP_STORIES.MAIN, blockId: "00000001", offset: 2 },
    slice: {
      content: contentOf(document).slice(2, 3),
      openStart: 0,
      openEnd: 0,
    },
    revision: { id: 50, author: "Reviewer", date: "2026-02-03T04:05:06Z" },
    newIds: {
      revision: Array.from(
        { length: 2 * DEPTH + 8 },
        (_, index) => 2_000 + index,
      ),
    },
  });

  test("a suggested copy records every owner, resolves and undoes", () => {
    const document = source();
    const suggested = apply(document, [paste(document)]);
    expect(copyOrigins(suggested.document)).toEqual([
      ORIGINS,
      [...ORIGINS, DEPTH],
    ]);
    expect(apply(suggested.document, suggested.inverse).document).toStrictEqual(
      document,
    );
    const rejected = resolve(
      suggested.document,
      suggested.revisions,
      REVISION_DECISIONS.REJECT,
    );
    expect(rejected.document).toStrictEqual(document);
    const accepted = resolve(
      suggested.document,
      suggested.revisions,
      REVISION_DECISIONS.ACCEPT,
    );
    expect(copyOrigins(accepted.document)).toEqual([ORIGINS, ORIGINS]);
    expect(apply(accepted.document, accepted.inverse).document).toStrictEqual(
      suggested.document,
    );
  });
});

describe("a tracked deletion of a transparent wrapper holding a comment anchor", () => {
  const document = (): Document =>
    documentOf([
      {
        type: "paragraph",
        paraId: "00000001",
        content: [
          run("x"),
          {
            type: "inlineWrapper",
            kind: "bidi",
            control: "embedding",
            content: [run("a"), { type: "commentReference", id: 9 }, run("b")],
          },
          run("y"),
        ],
      },
    ]);
  const deletion = (revision: boolean): DocumentOp => ({
    type: DOCUMENT_OP_TYPES.DELETE_RANGE,
    from: { story: OP_STORIES.MAIN, blockId: "00000001", offset: 1 },
    to: { story: OP_STORIES.MAIN, blockId: "00000001", offset: 3 },
    ...(revision
      ? {
          revision: {
            id: 50,
            author: "Reviewer",
            date: "2026-02-03T04:05:06Z",
          },
          newIds: { revision: [60, 61, 62, 63] },
        }
      : {}),
  });

  test("goes inside the wrapper and meets the anchor as a bare one", () => {
    const refused = applyDocumentOp(document(), deletion(true));
    expect(refused.isErr() && refused.error.reason).toBe(
      DOCUMENT_OP_REFUSAL_REASONS.DEPENDENT_RECORDS,
    );
    const bare = documentOf([
      {
        type: "paragraph",
        paraId: "00000001",
        content: [run("xa"), { type: "commentReference", id: 9 }, run("by")],
      },
    ]);
    const bareRefused = applyDocumentOp(bare, deletion(true));
    expect(bareRefused.isErr() && bareRefused.error.reason).toBe(
      DOCUMENT_OP_REFUSAL_REASONS.DEPENDENT_RECORDS,
    );
    expect(applyDocumentOp(document(), deletion(false)).isOk()).toBe(true);
  });

  test("an anchor the deletion owns with its range goes in with the wrapper's content", () => {
    const owned = documentOf([
      {
        type: "paragraph",
        paraId: "00000001",
        content: [
          run("x"),
          {
            type: "inlineWrapper",
            kind: "bidi",
            control: "embedding",
            content: [
              run("a"),
              { type: "commentRangeStart", id: 9 },
              { type: "commentRangeEnd", id: 9 },
              { type: "commentReference", id: 9 },
              run("b"),
            ],
          },
          run("y"),
        ],
      },
    ]);
    const tracked = apply(owned, [deletion(true)]);
    expect(apply(tracked.document, tracked.inverse).document).toStrictEqual(
      owned,
    );
    const direct = apply(owned, [deletion(false)]);
    expect(
      resolve(tracked.document, tracked.revisions, REVISION_DECISIONS.ACCEPT)
        .document.package.document.content,
    ).toStrictEqual(direct.document.package.document.content);
    expect(
      resolve(tracked.document, tracked.revisions, REVISION_DECISIONS.REJECT)
        .document,
    ).toStrictEqual(owned);
  });
});

describe("a pasted paragraph whose copied deletion of its break holds a cut", () => {
  const STAMP = { id: 100, author: "Reviewer", date: "2026-02-03T04:05:06Z" };
  const NEW_IDS = {
    revision: Array.from({ length: 16 }, (_, index) => 101 + index),
  };
  const link = (text: string): ParagraphContent => ({
    type: "hyperlink",
    rId: "rId9",
    href: "https://example.org",
    children: [run(text)],
  });
  const document = documentOf([paragraph("00000001", "Here.")]);
  const paste = (pPrMark: ParagraphMarkChange): EditorIntent => ({
    type: "replaceFragment",
    from: { story: OP_STORIES.MAIN, blockId: "00000001", offset: 5 },
    to: { story: OP_STORIES.MAIN, blockId: "00000001", offset: 5 },
    openStart: 0,
    openEnd: 0,
    paragraphs: [
      {
        type: "paragraph",
        paraId: "0000000D",
        pPrMark,
        content: [link("Cop")],
      },
      { type: "paragraph", paraId: "0000000E", content: [link("ied")] },
      { type: "paragraph", content: [run("tail")] },
    ],
  });

  test.each<ParagraphMarkChange>([
    { kind: "del", info: { id: 93, ...EARLIER }, resolutionJoin: 1 },
    {
      kind: "del",
      info: { id: 93, ...EARLIER },
      resolutionJoin: 1,
      splitDate: "2025-12-01T00:00:00Z",
    },
  ])(
    "is refused as a suggestion, its cut being read below the paste's own insertions (%o)",
    (pPrMark) => {
      const intent = paste(pPrMark);
      const suggested = compileEditorIntent(document, {
        intent,
        mode: { type: "suggesting", revision: STAMP, newIds: NEW_IDS },
      });
      expect(suggested.isErr() && suggested.error.reason).toBe(
        DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
      );
      const direct = compileEditorIntent(document, {
        intent,
        mode: { type: "editing", newIds: NEW_IDS },
      });
      if (direct.isErr()) throw direct.error;
      const pasted = apply(document, direct.value.ops);
      expect(apply(pasted.document, pasted.inverse).document).toStrictEqual(
        document,
      );
    },
  );
});
