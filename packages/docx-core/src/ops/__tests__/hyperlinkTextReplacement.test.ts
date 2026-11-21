/**
 * Text replacing words of a hyperlink, or typed between two of its leaves,
 * keeps the link, edited directly or tracked. A range reaching past the
 * link, or text typed at its outer edge, does not extend it. A tracked
 * record cannot stand inside a link, so the tracked text stands in a copy of
 * the link beside the struck one; accepting it gives the direct edit, one
 * link, whichever order its revisions are resolved in, and rejecting it gives
 * the source back. Bookmark and comment markers at the link's edges do not
 * change that.
 */

import { describe, expect, test } from "bun:test";

import type { Document, ParagraphContent } from "../../model/document";
import { applyDocumentOps } from "../apply";
import { normalizeForOps } from "../contract";
import { compileEditorIntent, type EditorIntentMode } from "../editorIntent";
import {
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  OP_STORIES,
  REVISION_DECISIONS,
  type RevisionDecision,
} from "../types";

const LINK = {
  type: "hyperlink",
  rId: "rId9",
  href: "https://example.invalid/terms",
  tooltip: "Terms",
} as const;

const run = (text: string): ParagraphContent => ({
  type: "run",
  content: [{ type: "text", text }],
});

const sourceOf = (
  content: ParagraphContent[],
  anchorKind?: "point" | "range",
) =>
  normalizeForOps({
    package: {
      document: {
        content: [{ type: "paragraph", paraId: "00000002", content }],
        comments: [
          {
            id: 7,
            author: "Owner",
            ...(anchorKind === undefined ? {} : { anchorKind }),
            content: [],
          },
        ],
      },
    },
  } satisfies Document);

const TERMS = "the standard terms";

const source = sourceOf([
  run("See "),
  { ...LINK, children: [run(TERMS)] },
  run(" now."),
]);

const at = (offset: number) => ({
  story: OP_STORIES.MAIN,
  blockId: "00000002",
  offset,
});

const EDITING: EditorIntentMode = { type: "editing" };
const SUGGESTING: EditorIntentMode = {
  type: "suggesting",
  revision: { id: 2, author: "Reviewer", date: "2026-01-01T00:00:00Z" },
  newIds: { revision: [3, 4, 5, 6], control: [] },
};

/** Apply a plan, checking that its inverse restores the document exactly. */
const applied = (document: Document, ops: readonly DocumentOp[]) => {
  const result = applyDocumentOps(document, ops).unwrap();
  expect(
    applyDocumentOps(result.document, result.inverse).unwrap().document,
  ).toStrictEqual(document);
  return result.document;
};

type Edit = { from: number; to: number; text: string };

const replaced = (
  { from, to, text }: Edit,
  mode: EditorIntentMode,
  document = source,
) =>
  applied(
    document,
    compileEditorIntent(document, {
      intent: { type: "replaceText", from: at(from), to: at(to), text },
      mode,
    }).unwrap().ops,
  );

const paragraphContent = (document: Document) => {
  const [paragraph] = document.package.document.content;
  if (paragraph?.type !== "paragraph") throw new Error("Expected a paragraph");
  return paragraph.content;
};

/** The paragraph as text, each link's text in brackets. */
const reading = (document: Document): string => {
  const text = (node: unknown): string =>
    (JSON.stringify(node).match(/"text":"[^"]*"/gu) ?? [])
      .map((match) => match.slice(8, -1))
      .join("");
  const read = (node: ParagraphContent): string => {
    switch (node.type) {
      case "hyperlink":
        return `[${text(node)}]`;
      case "insertion":
      case "deletion":
        return node.content.map(read).join("");
      default:
        return text(node);
    }
  };
  return paragraphContent(document).map(read).join("");
};

const revisionIds = (document: Document): number[] =>
  paragraphContent(document).flatMap((node) =>
    node.type === "insertion" || node.type === "deletion" ? [node.info.id] : [],
  );

const resolved = (
  document: Document,
  decision: RevisionDecision,
  ids = revisionIds(document),
) =>
  applied(document, [
    {
      type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
      story: OP_STORIES.MAIN,
      revisionIds: ids,
      decision,
    },
  ]);

const DECISIONS = [REVISION_DECISIONS.ACCEPT, REVISION_DECISIONS.REJECT];

/**
 * Resolving the revisions one at a time gives the same document in either
 * order, and what resolving them together gives when the decisions agree.
 */
const expectOrderFree = (tracked: Document) => {
  const [first, second, ...rest] = revisionIds(tracked);
  if (first === undefined || second === undefined) return;
  expect(rest).toEqual([]);
  for (const firstDecision of DECISIONS)
    for (const secondDecision of DECISIONS) {
      const forward = resolved(
        resolved(tracked, firstDecision, [first]),
        secondDecision,
        [second],
      );
      const backward = resolved(
        resolved(tracked, secondDecision, [second]),
        firstDecision,
        [first],
      );
      expect(paragraphContent(backward)).toStrictEqual(
        paragraphContent(forward),
      );
      if (firstDecision === secondDecision)
        expect(paragraphContent(forward)).toStrictEqual(
          paragraphContent(resolved(tracked, firstDecision)),
        );
    }
};

describe.each([
  {
    name: "exactly the link's text",
    edit: { from: 4, to: 22, text: "the revised terms" },
    direct: "See [the revised terms] now.",
  },
  {
    name: "words strictly inside the link",
    edit: { from: 8, to: 16, text: "revised" },
    direct: "See [the revised terms] now.",
  },
  {
    name: "the start of the link's text",
    edit: { from: 4, to: 8, text: "our " },
    direct: "See [our standard terms] now.",
  },
  {
    name: "the end of the link's text",
    edit: { from: 16, to: 22, text: " rules" },
    direct: "See [the standard rules] now.",
  },
  {
    name: "a deletion strictly inside the link",
    edit: { from: 8, to: 17, text: "" },
    direct: "See [the terms] now.",
  },
  {
    name: "a caret inside the link",
    edit: { from: 8, to: 8, text: "new " },
    direct: "See [the new standard terms] now.",
  },
  {
    name: "a range reaching past the link",
    edit: { from: 3, to: 23, text: "X" },
    direct: "SeeXnow.",
  },
  {
    name: "a range reaching past the link's start",
    edit: { from: 3, to: 22, text: "X" },
    direct: "SeeX now.",
  },
  {
    name: "a caret at the link's end",
    edit: { from: 22, to: 22, text: "X" },
    direct: "See [the standard terms]X now.",
  },
  {
    name: "a caret at the link's start",
    edit: { from: 4, to: 4, text: "X" },
    direct: "See X[the standard terms] now.",
  },
])("replacing $name", ({ edit, direct }) => {
  test("directly gives the expected link", () => {
    expect(reading(replaced(edit, EDITING))).toBe(direct);
  });

  test("tracked, accepting gives the direct edit and rejecting the source", () => {
    const tracked = replaced(edit, SUGGESTING);
    expect(
      paragraphContent(resolved(tracked, REVISION_DECISIONS.ACCEPT)),
    ).toStrictEqual(paragraphContent(replaced(edit, EDITING)));
    expect(resolved(tracked, REVISION_DECISIONS.REJECT)).toStrictEqual(source);
  });

  test("tracked, its revisions resolve alike in either order", () => {
    expectOrderFree(replaced(edit, SUGGESTING));
  });
});

describe.each([
  { name: "exactly the link's text", from: 4, to: 22 },
  { name: "words strictly inside the link", from: 8, to: 16 },
  { name: "the end of the link's text", from: 16, to: 22 },
  { name: "a caret inside the link", from: 8, to: 8 },
])("the link kept replacing $name", ({ from, to }) => {
  test("is a copy of the source link in both modes", () => {
    for (const mode of [EDITING, SUGGESTING]) {
      const links = paragraphContent(
        replaced({ from, to, text: "new" }, mode),
      ).flatMap((node) =>
        node.type === "insertion" || node.type === "deletion"
          ? node.content
          : [node],
      );
      for (const link of links)
        if (link.type === "hyperlink")
          expect(link).toMatchObject({ ...LINK, type: "hyperlink" });
    }
  });
});

const BOOKMARK_START = { type: "bookmarkStart", id: 1, name: "_Ref1" } as const;
const BOOKMARK_END = { type: "bookmarkEnd", id: 1 } as const;
const COMMENT_START = { type: "commentRangeStart", id: 7 } as const;
const COMMENT_END = { type: "commentRangeEnd", id: 7 } as const;
const COMMENT_REFERENCE = { type: "commentReference", id: 7 } as const;

describe.each([
  {
    name: "a bookmark around the link's text inside it",
    content: [
      run("See "),
      { ...LINK, children: [BOOKMARK_START, run(TERMS), BOOKMARK_END] },
      run(" now."),
    ],
    exact: "See [the revised terms] now.",
  },
  {
    name: "a bookmark starting inside the link",
    content: [
      run("See "),
      { ...LINK, children: [BOOKMARK_START, run(TERMS)] },
      BOOKMARK_END,
      run(" now."),
    ],
    exact: "See [the revised terms] now.",
  },
  {
    name: "a bookmark ending inside the link",
    content: [
      run("See "),
      BOOKMARK_START,
      { ...LINK, children: [run(TERMS), BOOKMARK_END] },
      run(" now."),
    ],
    exact: "See [the revised terms] now.",
  },
  {
    name: "a comment around the link's text inside it",
    content: [
      run("See "),
      { ...LINK, children: [COMMENT_START, run(TERMS), COMMENT_END] },
      COMMENT_REFERENCE,
      run(" now."),
    ],
    exact: "See [the revised terms] now.",
  },
  {
    // The new text goes past the end of the comment whose text it replaces,
    // which is past the link's end: it takes a copy of the link there.
    name: "a comment starting inside the link",
    content: [
      run("See "),
      { ...LINK, children: [COMMENT_START, run(TERMS)] },
      COMMENT_END,
      COMMENT_REFERENCE,
      run(" now."),
    ],
    exact: "See [][the revised terms] now.",
  },
  {
    name: "a comment ending inside the link",
    content: [
      run("See "),
      COMMENT_START,
      { ...LINK, children: [run(TERMS), COMMENT_END] },
      COMMENT_REFERENCE,
      run(" now."),
    ],
    exact: "See [the revised terms] now.",
  },
] satisfies { name: string; content: ParagraphContent[]; exact: string }[])(
  "with $name",
  ({ content, exact }) => {
    const marked = sourceOf(content);
    describe.each([
      {
        name: "exactly the link's text",
        edit: { from: 4, to: 22, text: "the revised terms" },
        direct: exact,
      },
      {
        name: "words strictly inside the link",
        edit: { from: 8, to: 16, text: "revised" },
        direct: "See [the revised terms] now.",
      },
    ])("replacing $name", ({ edit, direct }) => {
      test("directly keeps the text in the link", () => {
        const edited = replaced(edit, EDITING, marked);
        expect(reading(edited)).toBe(direct);
        expect(
          JSON.stringify(paragraphContent(edited))
            .match(/"type":"[a-zA-Z]+"/gu)
            ?.filter((type) => /Start|End|Reference/u.test(type)),
        ).toEqual(
          JSON.stringify(content)
            .match(/"type":"[a-zA-Z]+"/gu)
            ?.filter((type) => /Start|End|Reference/u.test(type)),
        );
      });

      test("tracked, accepting gives the direct edit and rejecting the source", () => {
        const tracked = replaced(edit, SUGGESTING, marked);
        expect(
          paragraphContent(resolved(tracked, REVISION_DECISIONS.ACCEPT)),
        ).toStrictEqual(paragraphContent(replaced(edit, EDITING, marked)));
        expect(resolved(tracked, REVISION_DECISIONS.REJECT)).toStrictEqual(
          marked,
        );
        expectOrderFree(tracked);
      });
    });
  },
);

test("a range taking a bookmark start just before the link keeps the text in it", () => {
  const marked = sourceOf([
    run("See "),
    BOOKMARK_START,
    { ...LINK, children: [run(TERMS)] },
    BOOKMARK_END,
    run(" now."),
  ]);
  const intent = {
    type: "replaceText",
    from: { ...at(4), zeroWidthBefore: 0 },
    to: at(16),
    text: "the revised",
  } as const;
  for (const mode of [EDITING, SUGGESTING]) {
    const edited = applied(
      marked,
      compileEditorIntent(marked, { intent, mode }).unwrap().ops,
    );
    const accepted =
      mode.type === "editing"
        ? edited
        : resolved(edited, REVISION_DECISIONS.ACCEPT);
    expect(reading(accepted)).toBe("See [the revised terms] now.");
  }
});

describe("replacing a paragraph whose words are exactly a link's, after a comment start", () => {
  const whole = normalizeForOps({
    package: {
      document: {
        content: [
          {
            type: "paragraph",
            paraId: "00000002",
            content: [COMMENT_START, { ...LINK, children: [run(TERMS)] }],
          },
          {
            type: "paragraph",
            paraId: "00000003",
            content: [run("Next."), COMMENT_END, COMMENT_REFERENCE],
          },
        ],
        comments: [
          { id: 7, author: "Owner", anchorKind: "range", content: [] },
        ],
      },
    },
  } satisfies Document);
  const intent = {
    type: "replaceText",
    from: { ...at(0), zeroWidthBefore: 1 },
    to: at(18),
    text: "the revised terms",
  } as const;
  const edited = (mode: EditorIntentMode) =>
    applied(whole, compileEditorIntent(whole, { intent, mode }).unwrap().ops);

  test("directly keeps the link on the new text", () => {
    expect(reading(edited(EDITING))).toBe("[the revised terms]");
  });

  test("tracked, accepting gives the direct edit and rejecting the source", () => {
    const tracked = edited(SUGGESTING);
    expect(
      paragraphContent(resolved(tracked, REVISION_DECISIONS.ACCEPT)),
    ).toStrictEqual(paragraphContent(edited(EDITING)));
    expect(resolved(tracked, REVISION_DECISIONS.REJECT)).toStrictEqual(whole);
  });
});
