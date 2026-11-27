/**
 * Suggested changes over records a tracked change has to hold: a point
 * comment's reference in a deleted paragraph, a comment anchor a hyperlink or
 * content control holds there, a tracked change inside a simple field's
 * result, a deleted or inserted paragraph mark that ends a section, and a
 * contents list over paragraphs under review.
 * Each is recorded, accepted and rejected as its direct edit and its source
 * state are, and every step undoes.
 */
import { describe, expect, test } from "bun:test";

import type { Document, Paragraph } from "../../model/document";
import { applyDocumentOp, applyDocumentOps } from "../apply";
import { compileEditorIntent, type EditorIntent } from "../editorIntent";
import { planTrackedReplace } from "../rangeReplacement";
import { DOCUMENT_OP_REFUSAL_REASONS } from "../refusal";
import {
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  OP_STORIES,
  REVISION_DECISIONS,
  type RevisionDecision,
} from "../types";

const STAMP = { id: 100, author: "Reviewer", date: "2026-02-03T04:05:06Z" };
const NEW_IDS = {
  revision: Array.from({ length: 16 }, (_, index) => 101 + index),
};

const run = (text: string) => ({
  type: "run" as const,
  content: [{ type: "text" as const, text }],
});

const apply = (document: Document, ops: readonly DocumentOp[]) => {
  const result = applyDocumentOps(document, ops);
  if (result.isErr()) throw result.error;
  return result.value;
};

const compiled = (
  document: Document,
  intent: EditorIntent,
  suggesting: boolean,
): readonly DocumentOp[] => {
  const result = compileEditorIntent(document, {
    intent,
    mode: suggesting
      ? { type: "suggesting", revision: STAMP, newIds: NEW_IDS }
      : { type: "editing", newIds: NEW_IDS },
  });
  if (result.isErr()) throw result.error;
  return result.value.ops;
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

/**
 * The suggestion against its direct edit: accepting gives the direct result,
 * rejecting gives the source back, and the suggestion and each resolution
 * undo.
 */
const assertSuggestionLaws = (document: Document, intent: EditorIntent) => {
  const direct = apply(document, compiled(document, intent, false));
  const suggested = apply(document, compiled(document, intent, true));
  expect(suggested.revisions.length).toBeGreaterThan(0);
  expect(apply(suggested.document, suggested.inverse).document).toStrictEqual(
    document,
  );
  for (const decision of Object.values(REVISION_DECISIONS)) {
    const resolved = resolve(suggested.document, suggested.revisions, decision);
    expect(resolved.document).toStrictEqual(
      decision === REVISION_DECISIONS.ACCEPT ? direct.document : document,
    );
    expect(apply(resolved.document, resolved.inverse).document).toStrictEqual(
      suggested.document,
    );
  }
  return { direct, suggested };
};

describe("a point comment in a deleted paragraph", () => {
  const document: Document = {
    package: {
      document: {
        comments: [
          { id: 7, author: "Commenter", anchorKind: "point", content: [] },
        ],
        content: [
          {
            type: "paragraph",
            paraId: "00000001",
            content: [
              run("Noted "),
              { type: "commentReference", id: 7 },
              run("here."),
            ],
          },
          { type: "paragraph", paraId: "00000002", content: [run("Kept.")] },
        ],
      },
    },
  };
  const intent: EditorIntent = {
    type: "deleteParagraphs",
    story: OP_STORIES.MAIN,
    blockIds: ["00000001"],
  };

  test("the deletion takes the reference in; accepting retires the comment as the direct deletion does", () => {
    const { direct, suggested } = assertSuggestionLaws(document, intent);
    expect(direct.document.package.document.comments).toBeUndefined();
    const [deleted] = suggested.document.package.document.content;
    expect(deleted?.type === "paragraph" && deleted.content).toEqual([
      {
        type: "deletion",
        info: expect.objectContaining({ author: STAMP.author }),
        content: [
          run("Noted "),
          { type: "commentReference", id: 7 },
          run("here."),
        ],
      },
    ]);
    expect(suggested.document.package.document.comments).toEqual(
      document.package.document.comments,
    );
  });

  test("a raw tracked block deletion records it too", () => {
    const result = applyDocumentOp(document, {
      type: DOCUMENT_OP_TYPES.DELETE_BLOCKS,
      story: OP_STORIES.MAIN,
      blockIds: ["00000001"],
      revision: STAMP,
      newIds: NEW_IDS,
    });
    expect(result.isOk()).toBe(true);
  });
});

describe("a comment anchor a hyperlink or content control holds in a deleted paragraph", () => {
  type Content = Paragraph["content"];
  const start = { type: "commentRangeStart" as const, id: 7 };
  const end = { type: "commentRangeEnd" as const, id: 7 };
  const reference = { type: "commentReference" as const, id: 7 };
  const link = (
    children: Extract<Content[number], { type: "hyperlink" }>["children"],
  ) => ({
    type: "hyperlink" as const,
    href: "https://example.com",
    children,
  });
  const smartTag = (
    content: Extract<Content[number], { type: "inlineWrapper" }>["content"],
  ) => ({
    type: "inlineWrapper" as const,
    kind: "smartTag" as const,
    element: "place",
    content,
  });
  const withDeleted = (
    deleted: Content,
    kept: Content,
    anchorKind?: "point",
  ): Document => ({
    package: {
      document: {
        comments: [
          {
            id: 7,
            author: "Commenter",
            content: [],
            ...(anchorKind === undefined ? {} : { anchorKind }),
          },
        ],
        content: [
          {
            type: "paragraph",
            paraId: "00000001",
            content: [run("a"), ...deleted, run("d")],
          },
          {
            type: "paragraph",
            paraId: "00000002",
            content: [run("e"), ...kept],
          },
        ],
      },
    },
  });
  const intent: EditorIntent = {
    type: "deleteParagraphs",
    story: OP_STORIES.MAIN,
    blockIds: ["00000001"],
  };
  const firstBlock = (document: Document) =>
    document.package.document.content.at(0);

  test.each<[string, Content, Content, Content]>([
    [
      "a range start inside a link",
      [link([run("b"), start, run("c")])],
      [end, reference],
      [link([start]), run("e"), end, reference],
    ],
    [
      "a range start opening a link",
      [link([start, run("b"), run("c")])],
      [end, reference],
      [link([start]), run("e"), end, reference],
    ],
    [
      "a range start closing a link",
      [link([run("b"), start])],
      [end, reference],
      [link([start]), run("e"), end, reference],
    ],
    [
      "a whole range inside a link",
      [link([run("b"), start, run("c"), end, reference])],
      [],
      [link([start]), link([end, reference]), run("e")],
    ],
    [
      "a range start in a smart tag inside a link",
      [link([run("b"), smartTag([run("x"), start, run("y")]), run("c")])],
      [end, reference],
      [link([smartTag([start])]), run("e"), end, reference],
    ],
    [
      "a range start in a smart tag",
      [smartTag([run("b"), start, run("c")])],
      [end, reference],
      [smartTag([start]), run("e"), end, reference],
    ],
    [
      "a range start in a content control",
      [
        {
          type: "inlineSdt",
          properties: { sdtType: "richText", id: 55 },
          content: [run("b"), start, run("c")],
        },
      ],
      [end, reference],
      [
        {
          type: "inlineSdt",
          properties: { sdtType: "richText", id: 55 },
          content: [start],
        },
        run("e"),
        end,
        reference,
      ],
    ],
  ])(
    "%s stands where the content was, as the direct deletion leaves it",
    (_shape, deleted, kept, joined) => {
      const document = withDeleted(deleted, kept);
      const { direct, suggested } = assertSuggestionLaws(document, intent);
      expect(direct.document.package.document.content).toStrictEqual([
        { type: "paragraph", paraId: "00000002", content: joined },
      ]);
      expect(direct.document.package.document.comments).toStrictEqual(
        document.package.document.comments,
      );
      const deletedBlock = firstBlock(suggested.document);
      expect(
        deletedBlock?.type === "paragraph" && deletedBlock.pPrMark?.kind,
      ).toBe("del");
    },
  );

  test("a point comment's reference inside a link goes with the link", () => {
    const document = withDeleted([link([run("b"), reference])], [], "point");
    const { direct, suggested } = assertSuggestionLaws(document, intent);
    expect(direct.document.package.document.comments).toBeUndefined();
    const deletedBlock = firstBlock(suggested.document);
    expect(deletedBlock?.type === "paragraph" && deletedBlock.content).toEqual([
      {
        type: "deletion",
        info: expect.objectContaining({ author: STAMP.author }),
        content: [run("a"), link([run("b"), reference]), run("d")],
      },
    ]);
  });

  test("a raw tracked block deletion records a link's standing anchor too", () => {
    const document = withDeleted(
      [link([run("b"), start, run("c")])],
      [end, reference],
    );
    const result = applyDocumentOp(document, {
      type: DOCUMENT_OP_TYPES.DELETE_BLOCKS,
      story: OP_STORIES.MAIN,
      blockIds: ["00000001"],
      revision: STAMP,
      newIds: NEW_IDS,
    });
    expect(result.isOk()).toBe(true);
  });
});

describe("a tracked change inside a simple field's result", () => {
  const info = (id: number) => ({
    id,
    author: "Earlier",
    date: "2026-01-01T00:00:00Z",
  });
  type FieldContent = Extract<
    Paragraph["content"][number],
    { type: "simpleField" }
  >["content"];
  const withField = (content: FieldContent): Document => ({
    package: {
      document: {
        content: [
          {
            type: "paragraph",
            paraId: "00000001",
            content: [
              run("Page "),
              {
                type: "simpleField",
                instruction: "PAGE",
                fieldType: "PAGE",
                content,
              },
              run(" of the report."),
            ],
          },
        ],
      },
    },
  });
  const inserted = {
    type: "insertion" as const,
    info: info(5),
    content: [run("2")],
  };
  const deleted = {
    type: "deletion" as const,
    info: info(6),
    content: [run("3")],
  };
  const source = withField([run("1"), inserted, deleted]);

  test.each([
    [REVISION_DECISIONS.ACCEPT, [5], [run("1"), run("2"), deleted]],
    [REVISION_DECISIONS.REJECT, [5], [run("1"), deleted]],
    [REVISION_DECISIONS.ACCEPT, [6], [run("1"), inserted]],
    [REVISION_DECISIONS.REJECT, [6], [run("1"), inserted, run("3")]],
    [REVISION_DECISIONS.ACCEPT, [5, 6], [run("1"), run("2")]],
    [REVISION_DECISIONS.REJECT, [5, 6], [run("1"), run("3")]],
  ] as const)(
    "%s of %j resolves it inside the field",
    (decision, ids, shown) => {
      const resolved = resolve(source, ids, decision);
      expect(resolved.document).toStrictEqual(withField([...shown]));
      expect(apply(resolved.document, resolved.inverse).document).toStrictEqual(
        source,
      );
    },
  );

  test("a container the resolution empties goes; the field stays, showing nothing", () => {
    const linked = withField([
      {
        type: "inlineWrapper",
        kind: "smartTag",
        element: "place",
        content: [inserted],
      },
    ]);
    const rejected = resolve(linked, [5], REVISION_DECISIONS.REJECT);
    expect(rejected.document).toStrictEqual(withField([]));
    expect(apply(rejected.document, rejected.inverse).document).toStrictEqual(
      linked,
    );
  });
});

describe("a suggested join at a section break", () => {
  const document: Document = {
    package: {
      document: {
        content: [
          {
            type: "paragraph",
            paraId: "00000001",
            sectionProperties: { pageWidth: 11_906 },
            content: [run("Ends a section.")],
          },
          {
            type: "paragraph",
            paraId: "00000002",
            content: [run("Opens the next.")],
          },
        ],
      },
    },
  };

  // Deleting a section break merges its section into the next one: the
  // joined paragraph keeps the later paragraph's mark and section, as the
  // direct join does.
  test.each<[string, EditorIntent]>([
    [
      "a join",
      {
        type: "joinParagraphs",
        story: OP_STORIES.MAIN,
        blockId: "00000001",
        nextBlockId: "00000002",
      },
    ],
    [
      "a deletion across the break",
      {
        type: "replaceText",
        from: { story: OP_STORIES.MAIN, blockId: "00000001", offset: 4 },
        to: { story: OP_STORIES.MAIN, blockId: "00000002", offset: 5 },
        text: "",
      },
    ],
  ])(
    "%s records the deleted mark and resolves as the direct edit",
    (_shape, intent) => {
      const { direct, suggested } = assertSuggestionLaws(document, intent);
      const [first] = suggested.document.package.document.content;
      expect(first?.type === "paragraph" && first.pPrMark?.kind).toBe("del");
      expect(
        first?.type === "paragraph" && first.sectionProperties,
      ).toStrictEqual({ pageWidth: 11_906 });
      const [joined, ...rest] = direct.document.package.document.content;
      expect(rest).toEqual([]);
      expect(
        joined?.type === "paragraph" ? joined.sectionProperties : "none",
      ).toBeUndefined();
    },
  );
});

describe("a contents list over paragraphs under review", () => {
  const earlier = (id: number) => ({
    id,
    author: "Earlier",
    date: "2026-01-01T00:00:00Z",
  });
  const headingContent: Paragraph["content"] = [
    run("Head"),
    { type: "insertion", info: earlier(7), content: [run("ing")] },
  ];
  const document: Document = {
    package: {
      document: {
        content: [
          {
            type: "paragraph",
            paraId: "00000001",
            content: [
              run("Intro "),
              { type: "insertion", info: earlier(5), content: [run("added")] },
              run(" text."),
            ],
          },
          {
            type: "paragraph",
            paraId: "00000002",
            formatting: { styleId: "Heading1" },
            pPrMark: { kind: "ins", info: earlier(6) },
            content: headingContent,
          },
          { type: "paragraph", paraId: "00000003", content: [run("Body.")] },
        ],
      },
    },
  };
  const intent = {
    type: "generateTOC",
    // Inside the earlier author's insertion, which the contents list cuts.
    at: { story: OP_STORIES.MAIN, blockId: "00000001", offset: 8 },
    title: "Contents",
    headings: [{ blockId: "00000002", text: "Heading", level: 0 }],
    tabPosition: 9360,
  } as const satisfies EditorIntent;

  test("the heading keeps its pending changes; accepting gives the direct list and rejecting leaves only the bookmark", () => {
    const direct = apply(document, compiled(document, intent, false));
    const suggested = apply(document, compiled(document, intent, true));
    expect(apply(direct.document, direct.inverse).document).toStrictEqual(
      document,
    );
    expect(apply(suggested.document, suggested.inverse).document).toStrictEqual(
      document,
    );
    const heading = (resolved: Document) =>
      resolved.package.document.content.find(
        (block) => block.type === "paragraph" && block.paraId === "00000002",
      );
    const [intro, headingBlock, body] = document.package.document.content;
    if (
      intro === undefined ||
      body === undefined ||
      headingBlock?.type !== "paragraph"
    )
      throw new TypeError("The fixture is intro, heading and body");
    const bookmarked: Paragraph = {
      ...headingBlock,
      content: [
        { type: "bookmarkStart", id: 1, name: "_Toc1" },
        ...headingContent,
        { type: "bookmarkEnd", id: 1 },
      ],
    };
    expect(heading(direct.document)).toStrictEqual(bookmarked);
    expect(heading(suggested.document)).toStrictEqual(bookmarked);
    const accepted = resolve(
      suggested.document,
      suggested.revisions,
      REVISION_DECISIONS.ACCEPT,
    );
    expect(accepted.document).toStrictEqual(direct.document);
    const rejected = resolve(
      suggested.document,
      suggested.revisions,
      REVISION_DECISIONS.REJECT,
    );
    expect(rejected.document.package.document.content).toStrictEqual([
      intro,
      bookmarked,
      body,
    ]);
    for (const resolved of [accepted, rejected])
      expect(apply(resolved.document, resolved.inverse).document).toStrictEqual(
        suggested.document,
      );
  });

  test("a contents list outside the body is an invalid request", () => {
    const result = compileEditorIntent(document, {
      intent: {
        ...intent,
        at: { ...intent.at, story: { kind: "footnote", id: 1 } },
      },
      mode: { type: "editing", newIds: NEW_IDS },
    });
    expect(result.isErr() && result.error.reason).toBe(
      DOCUMENT_OP_REFUSAL_REASONS.INVALID_OPERATION,
    );
  });
});

describe("an inserted paragraph that ends a section", () => {
  const document: Document = {
    package: {
      document: {
        content: [
          { type: "paragraph", paraId: "00000001", content: [run("One.")] },
          { type: "paragraph", paraId: "00000002", content: [run("Two.")] },
        ],
        finalSectionProperties: { pageWidth: 12_240 },
      },
    },
  };
  const ending: Paragraph = {
    type: "paragraph",
    paraId: "0000000A",
    sectionProperties: { pageWidth: 11_906 },
    content: [run("Ends.")],
  };
  const insert = (
    at: { type: "before" | "after"; blockId: string },
    suggesting: boolean,
  ): DocumentOp => ({
    type: DOCUMENT_OP_TYPES.INSERT_BLOCKS,
    story: OP_STORIES.MAIN,
    at,
    blocks: [ending],
    ...(suggesting ? { revision: STAMP, newIds: NEW_IDS } : {}),
  });

  test.each([
    { type: "before" as const, blockId: "00000001" },
    { type: "after" as const, blockId: "00000001" },
  ])("inserted %j, the break travels with the inserted mark", (at) => {
    const direct = apply(document, [insert(at, false)]);
    const suggested = apply(document, [insert(at, true)]);
    for (const edit of [direct, suggested])
      expect(apply(edit.document, edit.inverse).document).toStrictEqual(
        document,
      );
    const inserted = suggested.document.package.document.content.find(
      (block) => block.type === "paragraph" && block.paraId === "0000000A",
    );
    expect(inserted?.type === "paragraph" && inserted.pPrMark?.kind).toBe(
      "ins",
    );
    expect(
      inserted?.type === "paragraph" && inserted.sectionProperties,
    ).toStrictEqual({ pageWidth: 11_906 });
    for (const decision of Object.values(REVISION_DECISIONS)) {
      const resolved = resolve(
        suggested.document,
        suggested.revisions,
        decision,
      );
      expect(resolved.document).toStrictEqual(
        decision === REVISION_DECISIONS.ACCEPT ? direct.document : document,
      );
      expect(apply(resolved.document, resolved.inverse).document).toStrictEqual(
        suggested.document,
      );
    }
  });

  test("suggested after the story's last paragraph, the break has no inserted mark to travel with", () => {
    const at = { type: "after" as const, blockId: "00000002" };
    expect(applyDocumentOps(document, [insert(at, false)]).isOk()).toBe(true);
    const result = applyDocumentOps(document, [insert(at, true)]);
    expect(result.isErr() && result.error.reason).toBe(
      DOCUMENT_OP_REFUSAL_REASONS.SECTION_BOUNDARY,
    );
  });

  test("a tracked replacement brings its section break in with its inserted mark", () => {
    const planned = planTrackedReplace(document, {
      sourceContainerPolicy: "separate",
      from: { story: OP_STORIES.MAIN, blockId: "00000001", offset: 4 },
      to: { story: OP_STORIES.MAIN, blockId: "00000001", offset: 4 },
      revision: STAMP,
      newIds: NEW_IDS,
      replacement: {
        paragraphs: [ending],
        tail: { openStart: 0, openEnd: 0, content: [run(" Opens.")] },
      },
    });
    if (planned.isErr()) throw planned.error;
    const suggested = apply(document, planned.value);
    expect(apply(suggested.document, suggested.inverse).document).toStrictEqual(
      document,
    );
    const [first] = suggested.document.package.document.content;
    expect(first?.type === "paragraph" && first.paraId).toBe("0000000A");
    expect(first?.type === "paragraph" && first.pPrMark?.kind).toBe("ins");
    expect(
      first?.type === "paragraph" && first.sectionProperties,
    ).toStrictEqual({ pageWidth: 11_906 });
    const accepted = resolve(
      suggested.document,
      suggested.revisions,
      REVISION_DECISIONS.ACCEPT,
    );
    const [ends] = accepted.document.package.document.content;
    expect(ends?.type === "paragraph" && ends.sectionProperties).toStrictEqual({
      pageWidth: 11_906,
    });
    const rejected = resolve(
      suggested.document,
      suggested.revisions,
      REVISION_DECISIONS.REJECT,
    );
    expect(rejected.document).toStrictEqual(document);
    for (const resolved of [accepted, rejected])
      expect(apply(resolved.document, resolved.inverse).document).toStrictEqual(
        suggested.document,
      );
  });
});
