/**
 * Suggested changes over records a tracked change has to hold: a point
 * comment's reference in a deleted paragraph, a comment anchor a hyperlink or
 * content control holds there, a tracked change inside a simple field's
 * result, a deleted or inserted paragraph mark that ends a section, a
 * contents list over paragraphs under review, a paragraph mark's own run
 * properties, a paragraph put in after a pending inserted mark, and a pasted
 * paragraph whose break is pending deletion or that carries a review carrier.
 * Each is recorded, accepted and rejected as its direct edit and its source
 * state are, and every step undoes.
 */
import { describe, expect, test } from "bun:test";

import {
  type Document,
  type Paragraph,
  type ParagraphFormatting,
  type ParagraphMarkChange,
  REVIEW_CARRIERS,
  type TextFormatting,
} from "../../model/document";
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
  type RevisionStamp,
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

  test.each([
    [REVISION_DECISIONS.ACCEPT, "keeps"],
    [REVISION_DECISIONS.REJECT, "drops"],
  ] as const)(
    "%s of a change holding boundaries, a field and a control %s its content inside the field",
    (decision) => {
      const nested: FieldContent[number] = {
        type: "simpleField",
        instruction: "NUMPAGES",
        fieldType: "NUMPAGES",
        content: [run("9")],
      };
      const control: FieldContent[number] = {
        type: "inlineSdt",
        properties: { sdtType: "richText" },
        content: [run("x")],
      };
      const start: FieldContent[number] = {
        type: "bookmarkStart",
        id: 3,
        name: "_Ref3",
      };
      const end: FieldContent[number] = { type: "bookmarkEnd", id: 3 };
      const holding: FieldContent[number] = {
        type: "insertion",
        info: info(7),
        content: [start, run("4"), nested, control, end],
      };
      const document = withField([run("1"), holding]);
      const resolved = resolve(document, [7], decision);
      // Rejecting the insertion takes its content and leaves the bookmark's
      // boundaries where they stood, as a paragraph's resolution does.
      expect(resolved.document).toStrictEqual(
        withField(
          decision === REVISION_DECISIONS.ACCEPT
            ? [run("1"), start, run("4"), nested, control, end]
            : [run("1"), start, end],
        ),
      );
      expect(apply(resolved.document, resolved.inverse).document).toStrictEqual(
        document,
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

/**
 * A tracked operation against its direct form: accepting gives the direct
 * result, rejecting gives the source back, and every step undoes.
 */
const assertOpLaws = (
  document: Document,
  directOp: DocumentOp,
  trackedOp: DocumentOp,
) => {
  const direct = apply(document, [directOp]);
  const suggested = apply(document, [trackedOp]);
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

const at = (offset: number, zeroWidthBefore?: number) => ({
  story: OP_STORIES.MAIN,
  blockId: "00000001",
  offset,
  ...(zeroWidthBefore === undefined ? {} : { zeroWidthBefore }),
});

const commented = (content: Paragraph["content"]): Document => ({
  package: {
    document: {
      comments: [
        { id: 40, author: "Commenter", anchorKind: "range", content: [] },
      ],
      content: [
        { type: "paragraph", paraId: "00000001", content },
        { type: "paragraph", paraId: "00000002", content: [run("Kept.")] },
      ],
    },
  },
});

describe("comment ranges a deletion of text meets", () => {
  test.each(["opens", "closes"] as const)(
    "a range that %s inside the cut keeps its standing endpoint outside the w:del",
    (boundary) => {
      const document = commented(
        boundary === "opens"
          ? [
              run("ab"),
              { type: "commentRangeStart", id: 40 },
              run("cd"),
              { type: "commentRangeEnd", id: 40 },
              { type: "commentReference", id: 40 },
            ]
          : [
              { type: "commentRangeStart", id: 40 },
              run("ab"),
              { type: "commentRangeEnd", id: 40 },
              run("cd"),
              { type: "commentReference", id: 40 },
            ],
      );
      const { suggested } = assertSuggestionLaws(document, {
        type: "replaceText",
        from: at(1),
        to: at(3),
        text: "",
      });
      const [first] = suggested.document.package.document.content;
      if (first?.type !== "paragraph")
        throw new TypeError("Expected a paragraph.");
      // Both endpoints stay bare in the paragraph, outside every w:del.
      const anchors = first.content.filter(
        (node) =>
          node.type === "commentRangeStart" || node.type === "commentRangeEnd",
      );
      expect(anchors).toHaveLength(2);
    },
  );

  test("a range crossing out of a hyperlink keeps its start inside the link", () => {
    const document = commented([
      run("a"),
      {
        type: "hyperlink",
        anchor: "Clause",
        children: [run("b"), { type: "commentRangeStart", id: 40 }, run("c")],
      },
      run("d"),
      { type: "commentRangeEnd", id: 40 },
      { type: "commentReference", id: 40 },
    ]);
    assertSuggestionLaws(document, {
      type: "replaceText",
      from: at(1),
      to: at(3),
      text: "",
    });
  });

  test("a deletion holding a whole range and its reference takes the comment with it", () => {
    const document = commented([
      run("a"),
      { type: "commentRangeStart", id: 40 },
      run("b"),
      { type: "commentRangeEnd", id: 40 },
      { type: "commentReference", id: 40 },
      run("c"),
      run("d"),
    ]);
    const deleteRange = {
      type: DOCUMENT_OP_TYPES.DELETE_RANGE,
      from: at(0),
      to: at(3),
    } satisfies DocumentOp;
    const { direct, suggested } = assertOpLaws(document, deleteRange, {
      ...deleteRange,
      revision: STAMP,
      newIds: NEW_IDS,
    });
    expect(direct.document.package.document.comments ?? []).toEqual([]);
    const [first] = suggested.document.package.document.content;
    if (first?.type !== "paragraph")
      throw new TypeError("Expected a paragraph.");
    const deletion = first.content.find((node) => node.type === "deletion");
    if (deletion?.type !== "deletion")
      throw new TypeError("Expected a deletion.");
    expect(deletion.content).toContainEqual({
      type: "commentRangeStart",
      id: 40,
    });
  });
});

describe("an empty inline container a deletion selects alone", () => {
  const document: Document = {
    package: {
      document: {
        content: [
          {
            type: "paragraph",
            paraId: "00000001",
            content: [
              { type: "hyperlink", anchor: "Target", children: [] },
              run("text"),
            ],
          },
        ],
      },
    },
  };

  test("the deletion records it as the direct one removes it", () => {
    const { direct } = assertSuggestionLaws(document, {
      type: "replaceText",
      from: at(0, 0),
      to: at(0, 1),
      text: "",
    });
    const [only] = direct.document.package.document.content;
    expect(only?.type === "paragraph" && only.content).toEqual([run("text")]);
  });
});

describe("a tracked join that merges records at the seam", () => {
  test.each([1, 2])("the deleted mark records a join at depth %d", (depth) => {
    const content = (text: string): Paragraph["content"] =>
      depth === 1
        ? [run(text)]
        : [{ type: "hyperlink", anchor: "Clause", children: [run(text)] }];
    const document: Document = {
      package: {
        document: {
          content: [
            { type: "paragraph", paraId: "00000001", content: content("a") },
            { type: "paragraph", paraId: "00000002", content: content("b") },
          ],
        },
      },
    };
    const join = {
      type: DOCUMENT_OP_TYPES.JOIN_BLOCKS,
      story: OP_STORIES.MAIN,
      blockId: "00000001",
      nextBlockId: "00000002",
      depth,
    } satisfies DocumentOp;
    const { suggested } = assertOpLaws(document, join, {
      ...join,
      revision: STAMP,
    });
    const [first] = suggested.document.package.document.content;
    expect(first?.type === "paragraph" && first.pPrMark).toEqual(
      expect.objectContaining({ kind: "del", resolutionJoin: depth }),
    );
  });
});

describe("a pasted paragraph holding review and control records of its own", () => {
  const document: Document = {
    package: {
      document: {
        content: [
          { type: "paragraph", paraId: "00000001", content: [run("Here.")] },
        ],
      },
    },
  };
  const pasted: Paragraph = {
    type: "paragraph",
    paraId: "0000000B",
    content: [
      {
        type: "inlineSdt",
        properties: { sdtType: "richText", id: 42 },
        content: [
          {
            type: "insertion",
            info: { id: 91, author: "Earlier", date: "2026-01-01T00:00:00Z" },
            content: [run("pasted")],
          },
        ],
      },
    ],
  };

  test("the suggestion keeps the pasted insertion its own and resolves as the direct paste", () => {
    const { suggested } = assertSuggestionLaws(document, {
      type: "insertParagraphs",
      story: OP_STORIES.MAIN,
      at: { type: "before", blockId: "00000001" },
      paragraphs: [pasted],
    });
    expect(suggested.revisions).not.toContain(91);
  });
});

describe("a paragraph mark's run properties", () => {
  const mark: Paragraph = {
    type: "paragraph",
    paraId: "00000001",
    formatting: { styleId: "Heading1", runProperties: { bold: true } },
    content: [run("Title")],
  };
  const body: Paragraph = {
    type: "paragraph",
    paraId: "00000002",
    content: [run("Body.")],
  };
  const document: Document = {
    package: { document: { content: [mark, body] } },
  };
  const blockAt = (from: Document, index: number): Paragraph => {
    const block = from.package.document.content.at(index);
    if (block?.type !== "paragraph")
      throw new TypeError("Expected a paragraph.");
    return block;
  };
  const formatMark = (patch: ParagraphFormatting): EditorIntent => ({
    type: "formatParagraph",
    at: at(0),
    patch,
  });
  const setMark = (
    runProperties: TextFormatting,
    revision: RevisionStamp,
  ): DocumentOp => ({
    type: DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS,
    story: OP_STORIES.MAIN,
    blockId: "00000001",
    patch: { runProperties },
    revision,
  });

  test("a suggested change records the mark's previous run properties and resolves as the direct edit", () => {
    const { direct, suggested } = assertSuggestionLaws(
      document,
      formatMark({ runProperties: { italic: true }, runInWithNext: true }),
    );
    const changed = blockAt(suggested.document, 0);
    expect(changed.propertyChanges).toBeUndefined();
    expect(changed.markPropertyChange).toEqual({
      type: "paragraphMarkPropertyChange",
      info: expect.objectContaining({ author: STAMP.author }),
      previousFormatting: { runProperties: { bold: true } },
    });
    expect(blockAt(direct.document, 0).formatting).toStrictEqual({
      styleId: "Heading1",
      runProperties: { italic: true },
      runInWithNext: true,
    });
  });

  test("a mixed change records paragraph and mark properties apart, each rejected on its own", () => {
    const { suggested } = assertSuggestionLaws(
      document,
      formatMark({ alignment: "center", runProperties: { italic: true } }),
    );
    const changed = blockAt(suggested.document, 0);
    const markId = changed.markPropertyChange?.info.id;
    const paragraphId = changed.propertyChanges?.at(0)?.info.id;
    if (markId === undefined || paragraphId === undefined)
      throw new TypeError("Expected both property records.");
    expect(markId).not.toBe(paragraphId);
    expect([...suggested.revisions].toSorted((a, b) => a - b)).toEqual(
      [markId, paragraphId].toSorted((a, b) => a - b),
    );
    const markRejected = resolve(
      suggested.document,
      [markId],
      REVISION_DECISIONS.REJECT,
    );
    expect(blockAt(markRejected.document, 0).formatting).toStrictEqual({
      styleId: "Heading1",
      alignment: "center",
      runProperties: { bold: true },
    });
    const paragraphRejected = resolve(
      suggested.document,
      [paragraphId],
      REVISION_DECISIONS.REJECT,
    );
    expect(blockAt(paragraphRejected.document, 0).formatting).toStrictEqual({
      styleId: "Heading1",
      runProperties: { italic: true },
    });
  });

  test("a later suggestion folds into the pending change, and returning to the previous properties cancels it", () => {
    const first = apply(document, [setMark({ italic: true }, STAMP)]);
    const later = { ...STAMP, id: 200, date: "2026-02-04T00:00:00Z" };
    const folded = apply(first.document, [
      setMark({ bold: true, italic: true }, later),
    ]);
    expect(blockAt(folded.document, 0).markPropertyChange).toStrictEqual({
      type: "paragraphMarkPropertyChange",
      info: { ...later, id: STAMP.id },
      previousFormatting: { runProperties: { bold: true } },
    });
    expect(apply(folded.document, folded.inverse).document).toStrictEqual(
      first.document,
    );
    const cancelled = apply(folded.document, [setMark({ bold: true }, later)]);
    expect(cancelled.document).toStrictEqual(document);
    const other = applyDocumentOp(
      first.document,
      setMark({ bold: true, italic: true }, { ...later, suggestionId: "o" }),
    );
    expect(other.isErr() && other.error.reason).toBe(
      DOCUMENT_OP_REFUSAL_REASONS.PENDING_RUN_PROPERTY_CHANGE,
    );
  });

  test("a tracked paragraph that takes over the mark records its other run properties", () => {
    const beforeControl: Document = {
      package: {
        document: {
          content: [
            mark,
            {
              type: "blockSdt",
              properties: { sdtType: "richText" },
              content: [body],
            },
          ],
        },
      },
    };
    const inserted: Paragraph = {
      type: "paragraph",
      paraId: "00000003",
      formatting: { styleId: "Heading1", runProperties: { italic: true } },
      content: [run("New")],
    };
    const insertion = {
      type: DOCUMENT_OP_TYPES.INSERT_BLOCKS,
      story: OP_STORIES.MAIN,
      at: { type: "after", blockId: "00000001" },
      blocks: [inserted],
    } satisfies DocumentOp;
    const { suggested } = assertOpLaws(beforeControl, insertion, {
      ...insertion,
      revision: STAMP,
      newIds: NEW_IDS,
    });
    const taker = blockAt(suggested.document, 1);
    expect(taker.propertyChanges).toBeUndefined();
    expect(taker.markPropertyChange).toEqual({
      type: "paragraphMarkPropertyChange",
      info: expect.objectContaining({ author: STAMP.author }),
      previousFormatting: { runProperties: { bold: true } },
    });
  });

  test("a pasted paragraph keeps the mark change it was copied with", () => {
    const copied: Paragraph = {
      type: "paragraph",
      paraId: "0000000C",
      formatting: { runProperties: { italic: true } },
      markPropertyChange: {
        type: "paragraphMarkPropertyChange",
        info: { id: 92, author: "Earlier", date: "2026-01-01T00:00:00Z" },
        previousFormatting: { runProperties: { bold: true } },
      },
      content: [run("Copied")],
    };
    const tail: Paragraph = { type: "paragraph", content: [run("tail")] };
    const { direct, suggested } = assertSuggestionLaws(document, {
      type: "replaceFragment",
      from: at(5),
      to: at(5),
      openStart: 0,
      openEnd: 0,
      paragraphs: [copied, tail],
    });
    for (const pasted of [direct, suggested])
      expect(
        pasted.document.package.document.content.some(
          (block) =>
            block.type === "paragraph" &&
            block.markPropertyChange?.previousFormatting?.runProperties
              ?.bold === true,
        ),
      ).toBe(true);
  });
});

describe("a direct paragraph after a pending inserted mark", () => {
  const EARLIER: RevisionStamp = {
    id: 11,
    author: "Earlier",
    date: "2026-01-01T00:00:00Z",
  };
  const paragraph = (paraId: string, text: string): Paragraph => ({
    type: "paragraph",
    paraId,
    content: [run(text)],
  });
  /**
   * A tracked insertion after the story's last paragraph: the anchor's mark
   * is the inserted one, and the last new paragraph holds the mark it had.
   */
  const pending = (count: number) =>
    apply(
      {
        package: {
          document: {
            content: [
              paragraph("00000001", "Before."),
              paragraph("00000002", "Anchor."),
            ],
          },
        },
      },
      [
        {
          type: DOCUMENT_OP_TYPES.INSERT_BLOCKS,
          story: OP_STORIES.MAIN,
          at: { type: "after", blockId: "00000002" },
          blocks: Array.from({ length: count }, (_, index) =>
            paragraph(`0000001${index}`, "Tracked."),
          ),
          revision: EARLIER,
          newIds: { revision: [12, 13, 14, 15] },
        },
      ],
    );
  const insertion = (blocks: readonly Paragraph[]): DocumentOp => ({
    type: DOCUMENT_OP_TYPES.INSERT_BLOCKS,
    story: OP_STORIES.MAIN,
    at: { type: "after", blockId: "00000002" },
    blocks,
  });
  const added = [
    paragraph("00000020", "Direct."),
    paragraph("00000021", "More."),
  ];
  const blockWith = (document: Document, paraId: string): Paragraph => {
    const block = document.package.document.content.find(
      (candidate) =>
        candidate.type === "paragraph" && candidate.paraId === paraId,
    );
    if (block?.type !== "paragraph")
      throw new TypeError(`Expected paragraph ${paraId}.`);
    return block;
  };

  test.each([1, 2])(
    "takes the break to its last paragraph and resolves as the insertion into the resolved document (%i pending paragraphs)",
    (count) => {
      const marked = pending(count);
      const inserted = apply(marked.document, [insertion(added)]);
      expect(blockWith(inserted.document, "00000002").pPrMark).toBeUndefined();
      expect(blockWith(inserted.document, "00000021").pPrMark).toEqual(
        blockWith(marked.document, "00000002").pPrMark,
      );
      expect(apply(inserted.document, inserted.inverse).document).toStrictEqual(
        marked.document,
      );
      for (const decision of Object.values(REVISION_DECISIONS)) {
        const resolved = resolve(inserted.document, marked.revisions, decision);
        expect(resolved.document).toStrictEqual(
          apply(resolve(marked.document, marked.revisions, decision).document, [
            insertion(added),
          ]).document,
        );
        expect(
          apply(resolved.document, resolved.inverse).document,
        ).toStrictEqual(inserted.document);
      }
    },
  );

  test("from the editor, directly or suggested, the insertion resolves with the pending one", () => {
    const marked = pending(1);
    const intent: EditorIntent = {
      type: "insertParagraphs",
      story: OP_STORIES.MAIN,
      at: { type: "after", blockId: "00000002" },
      paragraphs: [paragraph("00000020", "Direct.")],
    };
    const suggested = apply(
      marked.document,
      compiled(marked.document, intent, true),
    );
    expect(suggested.revisions.length).toBeGreaterThan(0);
    expect(apply(suggested.document, suggested.inverse).document).toStrictEqual(
      marked.document,
    );
    const direct = apply(
      marked.document,
      compiled(marked.document, intent, false),
    );
    expect(apply(direct.document, direct.inverse).document).toStrictEqual(
      marked.document,
    );
    expect(
      resolve(direct.document, marked.revisions, REVISION_DECISIONS.REJECT)
        .document.package.document.content.length,
    ).toBe(3);
    const all = [...marked.revisions, ...suggested.revisions];
    expect(
      resolve(suggested.document, all, REVISION_DECISIONS.ACCEPT).document,
    ).toStrictEqual(
      resolve(direct.document, marked.revisions, REVISION_DECISIONS.ACCEPT)
        .document,
    );
    expect(
      resolve(suggested.document, all, REVISION_DECISIONS.REJECT).document,
    ).toStrictEqual(
      resolve(marked.document, marked.revisions, REVISION_DECISIONS.REJECT)
        .document,
    );
  });

  test("a paragraph the rejection would give other properties is refused", () => {
    const marked = pending(1);
    const refused = applyDocumentOp(
      marked.document,
      insertion([
        { ...paragraph("00000020", "Direct."), formatting: { keepNext: true } },
      ]),
    );
    expect(refused.isErr() && refused.error.reason).toBe(
      DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
    );
  });
});

describe("a pasted paragraph whose break is pending deletion", () => {
  const document: Document = {
    package: {
      document: {
        content: [
          { type: "paragraph", paraId: "00000001", content: [run("Here.")] },
        ],
      },
    },
  };
  const earlier = { author: "Earlier", date: "2026-01-01T00:00:00Z" };
  const paste = (pPrMark: ParagraphMarkChange): EditorIntent => ({
    type: "replaceFragment",
    from: at(5),
    to: at(5),
    openStart: 0,
    openEnd: 0,
    paragraphs: [
      {
        type: "paragraph",
        paraId: "0000000D",
        pPrMark,
        content: [run("Copied")],
      },
      { type: "paragraph", content: [run("tail")] },
    ],
  });

  test.each<ParagraphMarkChange>([
    { kind: "del", info: { id: 93, ...earlier } },
    { kind: "moveFrom", info: { id: 93, ...earlier } },
  ])(
    "the paste inserts the break and keeps the copied removal of it beside (%o)",
    (copied) => {
      const removal =
        copied.kind === "moveFrom"
          ? { info: expect.objectContaining(earlier), kind: "moveFrom" }
          : { info: expect.objectContaining(earlier) };
      const { suggested } = assertSuggestionLaws(document, paste(copied));
      const marks = suggested.document.package.document.content.flatMap(
        (block) =>
          block.type === "paragraph" && block.pPrMark?.deletion !== undefined
            ? [block.pPrMark]
            : [],
      );
      expect(marks).toEqual([
        {
          kind: "ins",
          info: expect.objectContaining({ author: STAMP.author }),
          deletion: removal,
        },
      ]);
      const [mark] = marks;
      if (mark?.deletion === undefined)
        throw new TypeError("The pasted break holds the copied removal.");
      expect(suggested.revisions).not.toContain(mark.deletion.info.id);
      // Accepting the paste alone leaves the copied record as the direct
      // paste copies it; resolving that record then removes the break.
      const accepted = resolve(
        suggested.document,
        suggested.revisions,
        REVISION_DECISIONS.ACCEPT,
      );
      const kept = accepted.document.package.document.content.flatMap(
        (block) =>
          block.type === "paragraph" && block.pPrMark !== undefined
            ? [block.pPrMark]
            : [],
      );
      expect(kept).toEqual([
        { kind: copied.kind, info: expect.objectContaining(earlier) },
      ]);
    },
  );

  test.each<ParagraphMarkChange>([
    { kind: "ins", info: { id: 93, ...earlier } },
    { kind: "moveTo", info: { id: 93, ...earlier } },
    {
      kind: "ins",
      info: { id: 93, ...earlier },
      deletion: { info: { id: 94, ...earlier } },
    },
    { kind: "moveFrom", info: { id: 93, ...earlier }, resolutionJoin: 1 },
  ])(
    "a copied mark that is no plain removal beside the paste's insertion is refused, suggested only (%o)",
    (pPrMark) => {
      const intent = paste(pPrMark);
      const refused = compileEditorIntent(document, {
        intent,
        mode: { type: "suggesting", revision: STAMP, newIds: NEW_IDS },
      });
      // An added mark is a format limit: `w:pPr/w:rPr` holds one `w:ins`,
      // the paste's. A removal holding a cut of its own is one the record
      // beside an added mark does not carry.
      expect(refused.isErr() && refused.error.reason).toBe(
        pPrMark.kind === "moveFrom"
          ? DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE
          : DOCUMENT_OP_REFUSAL_REASONS.NO_REVISION_RECORD,
      );
      expect(compiled(document, intent, false).length).toBeGreaterThan(0);
    },
  );

  test.each<
    [
      string,
      Pick<
        Paragraph,
        "formatting" | "pPrMark" | "propertyChanges" | "sectionProperties"
      >,
    ]
  >([
    [
      "a deleted mark",
      { pPrMark: { kind: "del", info: { id: 93, ...earlier } } },
    ],
    [
      "a property change",
      {
        formatting: { alignment: "center" },
        propertyChanges: [
          {
            type: "paragraphPropertyChange",
            info: { id: 93, ...earlier },
            previousFormatting: { alignment: "end" },
          },
        ],
      },
    ],
    ["a section break", { sectionProperties: { pageWidth: 12240 } }],
  ])(
    "copied review on the mark the paste ends in stays refused, suggested only (%s)",
    (_shape, fields) => {
      const followed: Document = {
        package: {
          document: {
            content: [
              ...document.package.document.content,
              {
                type: "paragraph",
                paraId: "00000002",
                content: [run("Next.")],
              },
            ],
          },
        },
      };
      const intent: EditorIntent = {
        type: "replaceFragment",
        from: at(5),
        to: at(5),
        openStart: 0,
        openEnd: 0,
        paragraphs: [
          { type: "paragraph", paraId: "0000000D", content: [run("Copied")] },
          { type: "paragraph", ...fields, content: [run("tail")] },
        ],
      };
      const refused = compileEditorIntent(followed, {
        intent,
        mode: { type: "suggesting", revision: STAMP, newIds: NEW_IDS },
      });
      // The mark a paste ends in is kept, not inserted: no record of the
      // paste holds copied mark review or a section break added to it.
      expect(refused.isErr() && refused.error.reason).toBe(
        DOCUMENT_OP_REFUSAL_REASONS.NO_REVISION_RECORD,
      );
      expect(
        apply(followed, compiled(followed, intent, false)).document,
      ).not.toStrictEqual(followed);
    },
  );
});

describe("a pasted paragraph carrying a table's review carrier", () => {
  const document: Document = {
    package: {
      document: {
        content: [
          { type: "paragraph", paraId: "00000001", content: [run("Here.")] },
        ],
      },
    },
  };

  test("pastes as the plain paragraph it shows, directly and suggested", () => {
    const { direct, suggested } = assertSuggestionLaws(document, {
      type: "replaceFragment",
      from: at(5),
      to: at(5),
      openStart: 0,
      openEnd: 0,
      paragraphs: [
        {
          type: "paragraph",
          paraId: "0000000E",
          reviewCarrier: REVIEW_CARRIERS.TERMINAL_TABLE,
          content: [run("Carried")],
        },
        { type: "paragraph", content: [run("tail")] },
      ],
    });
    for (const pasted of [direct, suggested])
      expect(
        pasted.document.package.document.content.some(
          (block) =>
            block.type === "paragraph" && block.reviewCarrier !== undefined,
        ),
      ).toBe(false);
  });
});
