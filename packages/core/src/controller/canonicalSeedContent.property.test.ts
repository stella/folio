/**
 * The seed check judges each block once and reuses the judgment while the
 * block stays the same object, as every block an operation does not name
 * does. Along a chain of edits that share blocks, its verdict equals one walk
 * of every story. The check of an edit's result judges only the blocks the
 * edit replaced, and the comments only when it changed their anchors or
 * definitions; after a version the session could hold, its verdict equals
 * one walk too, also with range markers inside content controls, table
 * cells and text boxes.
 */
import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import type {
  BlockContent,
  Paragraph,
  RunContent,
} from "@stll/docx-core/model";
import {
  commentAnchorSequence,
  commentDocumentIssue,
  documentStories,
  findStoryBody,
  OP_STORIES,
} from "@stll/docx-core/ops";

import {
  assertProperty,
  propertyConfig,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import { table } from "../docx/server/build";
import type { Document } from "../types/document";
import {
  changedWindow,
  unsupportedEditReason,
  unsupportedSeedReason,
  unsupportedSeedReasonWholeDocument,
} from "./canonicalSeedContent";

// Synchronous generated edit chains use a 15-second base budget.
setDefaultTimeout(propertyTestTimeout(15_000));

// A tab in a text node is one of the inline contents the session refuses.
const texts = fc.constantFrom("a", "b", "c\td");

const commentIds = fc.constantFrom(1, 2, 3);

/** A comment anchor in a paragraph: a range boundary or a reference. */
const inlineAnchorOf = fc.record({
  type: fc.constantFrom(
    "commentRangeStart" as const,
    "commentRangeEnd" as const,
    "commentReference" as const,
  ),
  id: commentIds,
});

const paragraphOf = fc
  .array(fc.oneof({ weight: 3, arbitrary: texts }, inlineAnchorOf), {
    maxLength: 2,
  })
  .map((parts): Paragraph => ({
    type: "paragraph",
    content: parts.map((part) =>
      typeof part === "string"
        ? { type: "run", content: [{ type: "text", text: part }] }
        : part,
    ),
  }));

const blockOf: fc.Arbitrary<BlockContent> = fc.oneof(
  { weight: 4, arbitrary: paragraphOf },
  {
    weight: 1,
    arbitrary: fc
      .array(paragraphOf, { minLength: 1, maxLength: 2 })
      .map((content): BlockContent => {
        const cells = table({ rows: [["x"]] });
        const cell = cells.rows.at(0)?.cells.at(0);
        if (cell !== undefined) cell.content = content;
        return cells;
      }),
  },
  {
    weight: 1,
    arbitrary: fc
      .array(paragraphOf, { minLength: 1, maxLength: 2 })
      .map((content): BlockContent => ({
        type: "blockSdt",
        properties: { sdtType: "richText" },
        content,
      })),
  },
  // A comment range boundary between blocks.
  {
    weight: 1,
    arbitrary: fc.record({
      type: fc.constantFrom(
        "commentRangeStart" as const,
        "commentRangeEnd" as const,
      ),
      id: commentIds,
    }),
  },
  // A block the session cannot hold.
  {
    weight: 1,
    arbitrary: fc
      .array(paragraphOf, { maxLength: 1 })
      .map((content): BlockContent => ({
        type: "blockCustomXml",
        openingXml: "<w:customXml>",
        closingXml: "</w:customXml>",
        content,
      })),
  },
);

type Edit =
  | { type: "replace"; at: number; block: BlockContent }
  | { type: "insert"; at: number; block: BlockContent }
  | { type: "delete"; at: number }
  | { type: "comments"; ids: number[] };

const editOver = (blocks: fc.Arbitrary<BlockContent>): fc.Arbitrary<Edit> =>
  fc.oneof(
    fc.record({
      type: fc.constant("replace" as const),
      at: fc.nat(),
      block: blocks,
    }),
    fc.record({
      type: fc.constant("insert" as const),
      at: fc.nat(),
      block: blocks,
    }),
    fc.record({ type: fc.constant("delete" as const), at: fc.nat() }),
    fc.record({
      type: fc.constant("comments" as const),
      ids: fc.uniqueArray(commentIds, { maxLength: 3 }),
    }),
  );

const editOf = editOver(blockOf);

/** A comment range boundary between blocks. */
const rangeMarkerOf: fc.Arbitrary<BlockContent> = fc.record({
  type: fc.constantFrom(
    "commentRangeStart" as const,
    "commentRangeEnd" as const,
  ),
  id: commentIds,
});

/** A text box holding these blocks, in a run of its own. */
const textBoxRun = (content: BlockContent[]): RunContent => ({
  type: "run",
  content: [
    {
      type: "shape",
      shape: {
        type: "shape",
        shapeType: "textBox",
        size: { width: 1_000_000, height: 500_000 },
        textBody: { content },
      },
    },
  ],
});

/** A paragraph that may hold a text box of paragraphs and range markers. */
const boxedParagraphOf: fc.Arbitrary<Paragraph> = fc
  .tuple(
    paragraphOf,
    fc.option(
      fc.array(fc.oneof(paragraphOf, rangeMarkerOf), {
        minLength: 1,
        maxLength: 3,
      }),
    ),
    fc.nat(),
  )
  .map(([paragraph, box, at]): Paragraph => {
    if (box === null) return paragraph;
    const content = paragraph.content.slice();
    content.splice(at % (content.length + 1), 0, textBoxRun(box));
    // A shrunk case maps the same generated paragraph again: never change it.
    return { type: "paragraph", content };
  });

/** A container's blocks: paragraphs and range markers between them. */
const containedOf = fc.array(fc.oneof(boxedParagraphOf, rangeMarkerOf), {
  minLength: 1,
  maxLength: 3,
});

/**
 * Blocks whose comment anchors may sit inside a content control, a table
 * cell or a text box, range markers between blocks among them.
 */
const nestedBlockOf: fc.Arbitrary<BlockContent> = fc.oneof(
  { weight: 3, arbitrary: boxedParagraphOf },
  {
    weight: 1,
    arbitrary: containedOf.map((content): BlockContent => {
      const cells = table({ rows: [["x"]] });
      const cell = cells.rows.at(0)?.cells.at(0);
      if (cell !== undefined) cell.content = content;
      return cells;
    }),
  },
  {
    weight: 1,
    arbitrary: containedOf.map((content): BlockContent => ({
      type: "blockSdt",
      properties: { sdtType: "richText" },
      content,
    })),
  },
  { weight: 1, arbitrary: rangeMarkerOf },
);

const commentsOf = (ids: readonly number[]) =>
  ids.map((id) => ({ id, author: "Reviewer", content: [] }));

const documentOf = (
  body: BlockContent[],
  note: BlockContent[],
  comments?: Document["package"]["document"]["comments"],
): Document => ({
  package: {
    document: {
      content: body,
      ...(comments === undefined ? {} : { comments }),
    },
    footnotes: [{ type: "footnote", id: 1, content: note }],
  },
});

/** A new document sharing every block the edit does not touch. */
const edited = (document: Document, edit: Edit, inNote: boolean): Document => {
  const note = document.package.footnotes?.at(0)?.content ?? [];
  const comments = document.package.document.comments;
  if (edit.type === "comments")
    return documentOf(
      document.package.document.content,
      note,
      commentsOf(edit.ids),
    );
  const blocks = [...(inNote ? note : document.package.document.content)];
  const at = blocks.length === 0 ? 0 : edit.at % blocks.length;
  if (edit.type === "replace" && blocks.length > 0)
    blocks.splice(at, 1, edit.block);
  else if (edit.type === "insert") blocks.splice(at, 0, edit.block);
  else if (edit.type === "delete") blocks.splice(at, 1);
  return inNote
    ? documentOf(document.package.document.content, blocks, comments)
    : documentOf(blocks, note, comments);
};

/** The stories whose blocks differ between two versions. */
const changedStoriesOf = (previous: Document, document: Document) => [
  ...(previous.package.document.content === document.package.document.content
    ? []
    : [OP_STORIES.MAIN]),
  ...(previous.package.footnotes?.at(0)?.content ===
  document.package.footnotes?.at(0)?.content
    ? []
    : [{ kind: "footnote" as const, id: 1 }]),
];

type EditStep = {
  previous: Document;
  document: Document;
  stories: ReturnType<typeof changedStoriesOf>;
};

/**
 * Each edit of a chain from a version the session could hold, with the
 * version before it and the stories it changed. A refused edit leaves the
 * chain on the version before it, as the session does.
 */
const eachEdit = (
  [body, note, ids, edits]: [
    BlockContent[],
    BlockContent[],
    number[],
    [Edit, boolean][],
  ],
  judge: (step: EditStep) => void,
): void => {
  const plain = (): Paragraph => ({ type: "paragraph", content: [] });
  let document = documentOf(body, note, commentsOf(ids));
  // The session holds only documents the seed check accepts.
  if (unsupportedSeedReasonWholeDocument(document) !== null)
    document = documentOf([plain()], [plain()], commentsOf(ids));
  expect(unsupportedSeedReasonWholeDocument(document)).toBeNull();
  for (const [edit, inNote] of edits) {
    const previous = document;
    const next = edited(document, edit, inNote);
    judge({
      previous,
      document: next,
      stories: changedStoriesOf(previous, next),
    });
    if (unsupportedSeedReasonWholeDocument(next) === null) document = next;
  }
};

const editChainOf = (blocks: fc.Arbitrary<BlockContent>) =>
  fc.tuple(
    fc.array(blocks, { maxLength: 5 }),
    fc.array(blocks, { maxLength: 2 }),
    fc.uniqueArray(commentIds, { maxLength: 3 }),
    fc.array(fc.tuple(editOver(blocks), fc.boolean()), {
      minLength: 1,
      maxLength: 10,
    }),
  );

describe("canonical seed check", () => {
  test("judges shared blocks as one walk of every story does", () => {
    fc.assert(
      fc.property(
        fc.array(blockOf, { maxLength: 5 }),
        fc.array(blockOf, { maxLength: 2 }),
        fc.array(fc.tuple(editOf, fc.boolean()), {
          minLength: 1,
          maxLength: 8,
        }),
        (body, note, edits) => {
          let document = documentOf(body, note);
          expect(unsupportedSeedReason(document)).toBe(
            unsupportedSeedReasonWholeDocument(document),
          );
          for (const [edit, inNote] of edits) {
            document = edited(document, edit, inNote);
            expect(unsupportedSeedReason(document)).toBe(
              unsupportedSeedReasonWholeDocument(document),
            );
          }
        },
      ),
      propertyConfig({ numRuns: 300 }),
    );
  });

  const judgedAsOneWalk = ({ previous, document, stories }: EditStep) =>
    expect(unsupportedEditReason({ previous, document, stories })).toBe(
      unsupportedSeedReasonWholeDocument(document, stories),
    );

  test("judges an edit's result as one walk of every story does", () => {
    let judged = 0;
    assertProperty(
      fc.property(
        fc.array(blockOf, { maxLength: 5 }),
        fc.array(blockOf, { maxLength: 2 }),
        fc.uniqueArray(commentIds, { maxLength: 3 }),
        fc.array(fc.tuple(editOf, fc.boolean()), {
          minLength: 1,
          maxLength: 10,
        }),
        (body, note, ids, edits) =>
          eachEdit([body, note, ids, edits], (step) => {
            judgedAsOneWalk(step);
            judged += 1;
          }),
      ),
      {
        id: "judges an edit's result as one walk of every story does",
        numRuns: 600,
      },
    );
    expect(judged).toBeGreaterThan(300);
  });

  test("judges an edit's result as one walk does with anchors in content controls, cells and text boxes", () => {
    let judged = 0;
    let commentIssues = 0;
    assertProperty(
      fc.property(editChainOf(nestedBlockOf), (chain) =>
        eachEdit(chain, (step) => {
          judgedAsOneWalk(step);
          judged += 1;
          const verdict = unsupportedSeedReasonWholeDocument(
            step.document,
            step.stories,
          );
          if (verdict !== null && /comment/iu.test(verdict)) commentIssues += 1;
        }),
      ),
      {
        id: "judges an edit's result as one walk with nested anchors",
        numRuns: 400,
      },
    );
    expect(judged).toBeGreaterThan(300);
    // The chains reach comment verdicts, not only accepted edits.
    expect(commentIssues).toBeGreaterThan(20);
  });

  describe("judges again an edit that changes the paragraph a marker between blocks attaches to", () => {
    const text = (value: string): Paragraph => ({
      type: "paragraph",
      content: [{ type: "run", content: [{ type: "text", text: value }] }],
    });
    const start: BlockContent = { type: "commentRangeStart", id: 1 };
    const end: BlockContent = { type: "commentRangeEnd", id: 1 };
    const expectJudgedAsOneWalk = (
      previous: Document,
      document: Document,
    ): void => {
      expect(unsupportedSeedReasonWholeDocument(previous)).toBeNull();
      const stories = changedStoriesOf(previous, document);
      const verdict = unsupportedSeedReasonWholeDocument(document, stories);
      expect(verdict).not.toBeNull();
      expect(unsupportedEditReason({ previous, document, stories })).toBe(
        verdict,
      );
    };

    test("deleting the empty paragraph a range opens and closes on", () => {
      // The range starts and ends on the empty paragraph; without it, its
      // end closes on the paragraph before while its start waits for one.
      const first = text("a");
      const empty: Paragraph = { type: "paragraph", content: [] };
      const previous = documentOf(
        [first, start, empty, end],
        [{ type: "paragraph", content: [] }],
        commentsOf([1]),
      );
      expectJudgedAsOneWalk(
        previous,
        edited(previous, { type: "delete", at: 2 }, false),
      );
    });

    test("moving a range start past the empty paragraph it opened on", () => {
      // Both runs hold the same anchors, but the moved start is held past
      // the paragraph, after the end that closes on it.
      const first = text("a");
      const empty: Paragraph = { type: "paragraph", content: [] };
      const note: BlockContent[] = [{ type: "paragraph", content: [] }];
      const previous = documentOf(
        [first, start, empty, end],
        note,
        commentsOf([1]),
      );
      // The same comments: only the anchors decide that the check runs.
      expectJudgedAsOneWalk(
        previous,
        documentOf(
          [first, empty, start, end],
          note,
          previous.package.document.comments,
        ),
      );
    });

    test("moving a range start inside a content control past the paragraph it opened on", () => {
      // The start waits inside the control for its paragraph; moved past it,
      // it waits for a paragraph after the end, which closes on that one.
      const first = text("a");
      const empty: Paragraph = { type: "paragraph", content: [] };
      const control = (content: BlockContent[]): BlockContent => ({
        type: "blockSdt",
        properties: { sdtType: "richText" },
        content,
      });
      const note: BlockContent[] = [{ type: "paragraph", content: [] }];
      const previous = documentOf(
        [first, control([start, empty]), end],
        note,
        commentsOf([1]),
      );
      expectJudgedAsOneWalk(
        previous,
        documentOf(
          [first, control([empty, start]), end],
          note,
          previous.package.document.comments,
        ),
      );
    });

    test("moving a range start from a paragraph into its text box", () => {
      // A story's text boxes follow its body, so the start now comes after
      // the end in the next paragraph.
      const note: BlockContent[] = [{ type: "paragraph", content: [] }];
      const closing: Paragraph = {
        type: "paragraph",
        content: [
          { type: "run", content: [{ type: "text", text: "b" }] },
          { type: "commentRangeEnd", id: 1 },
        ],
      };
      const previous = documentOf(
        [
          {
            type: "paragraph",
            content: [
              { type: "commentRangeStart", id: 1 },
              { type: "run", content: [{ type: "text", text: "a" }] },
              textBoxRun([{ type: "paragraph", content: [] }]),
            ],
          },
          closing,
        ],
        note,
        commentsOf([1]),
      );
      expectJudgedAsOneWalk(
        previous,
        documentOf(
          [
            {
              type: "paragraph",
              content: [
                { type: "run", content: [{ type: "text", text: "a" }] },
                textBoxRun([
                  {
                    type: "paragraph",
                    content: [{ type: "commentRangeStart", id: 1 }],
                  },
                ]),
              ],
            },
            closing,
          ],
          note,
          previous.package.document.comments,
        ),
      );
    });
  });

  test("judges again an edit that turns an anchor into another kind with the same id", () => {
    const anchor = (
      type: "commentRangeStart" | "commentRangeEnd" | "commentReference",
    ): Paragraph => ({ type: "paragraph", content: [{ type, id: 1 }] });
    const previous = documentOf(
      [anchor("commentRangeStart"), anchor("commentRangeEnd")],
      [{ type: "paragraph", content: [] }],
      commentsOf([1]),
    );
    expect(unsupportedSeedReasonWholeDocument(previous)).toBeNull();
    // The range loses its end and gains a reference: unbalanced.
    const document = edited(
      previous,
      { type: "replace", at: 1, block: anchor("commentReference") },
      false,
    );
    const stories = changedStoriesOf(previous, document);
    const verdict = unsupportedSeedReasonWholeDocument(document, stories);
    expect(verdict).not.toBeNull();
    expect(unsupportedEditReason({ previous, document, stories })).toBe(
      verdict,
    );
  });
});

// Deliberate breaks of the decision to skip the comment check, each run
// against the pinned case and both properties. A break is caught when it
// skips the check on a result the comment check refuses.
const ATTACHMENT_MARKERS = new Set([
  "noParagraph",
  "firstParagraph",
  "heldPastLastParagraph",
]);
type Sequence = (blocks: readonly BlockContent[]) => string[];
const keptBy =
  (
    sequence: Sequence,
    { definitions = true, storyCount = true } = {},
  ): ((step: EditStep) => boolean) =>
  ({ previous, document, stories }) =>
    (!definitions ||
      document.package.document.comments ===
        previous.package.document.comments) &&
    (!storyCount ||
      documentStories(document).length === documentStories(previous).length) &&
    stories.every((story) => {
      const { removed, added } = changedWindow(
        findStoryBody(previous, story)?.content ?? [],
        findStoryBody(document, story)?.content ?? [],
      );
      const left = sequence(removed);
      const right = sequence(added);
      return (
        left.length === right.length &&
        left.every((entry, index) => entry === right[index])
      );
    });
const flatten = (blocks: readonly BlockContent[]): BlockContent[] =>
  blocks.map((block) =>
    block.type === "table" || block.type === "blockSdt"
      ? { type: "paragraph", content: [] }
      : block,
  );
const breaks: [string, (step: EditStep) => boolean][] = [
  ["unchanged decision", keptBy(commentAnchorSequence)],
  [
    "no paragraph attachment markers",
    keptBy((blocks) =>
      commentAnchorSequence(blocks).filter(
        (entry) => !ATTACHMENT_MARKERS.has(entry),
      ),
    ),
  ],
  [
    "text boxes ignored",
    keptBy((blocks) => {
      const sequence = commentAnchorSequence(blocks);
      return sequence.slice(0, sequence.indexOf("textBoxes"));
    }),
  ],
  [
    "content controls and cells ignored",
    keptBy((blocks) => commentAnchorSequence(flatten(blocks))),
  ],
  [
    "anchors compared as a set",
    keptBy((blocks) => commentAnchorSequence(blocks).toSorted()),
  ],
  [
    "definitions ignored",
    keptBy(commentAnchorSequence, { definitions: false }),
  ],
  ["story count ignored", keptBy(commentAnchorSequence, { storyCount: false })],
];
test("deliberate breaks of the comment skip", () => {
  const rows: string[] = [];
  for (const [name, kept] of breaks) {
    const judge = (step: EditStep) => {
      if (
        kept(step) &&
        commentDocumentIssue(step.document, "all") !== undefined
      )
        throw new Error("skipped a refused comment state");
    };
    const original = fc.property(
      fc.array(blockOf, { maxLength: 5 }),
      fc.array(blockOf, { maxLength: 2 }),
      fc.uniqueArray(commentIds, { maxLength: 3 }),
      fc.array(fc.tuple(editOf, fc.boolean()), {
        minLength: 1,
        maxLength: 10,
      }),
      (body, note, ids, edits) => eachEdit([body, note, ids, edits], judge),
    );
    const pinned = fc.check(original, {
      seed: -71963594,
      path: "428:2:4:5:6:5:6:6:6:6:6",
    }).failed;
    const generated = fc.check(original, { numRuns: 600, seed: 1 }).failed;
    const nested = fc.check(
      fc.property(editChainOf(nestedBlockOf), (chain) =>
        eachEdit(chain, judge),
      ),
      { numRuns: 400, seed: 1 },
    ).failed;
    rows.push(
      `BREAK ${name}: pinned=${String(pinned)} property=${String(generated)} nested=${String(nested)}`,
    );
  }
  console.log(rows.join("\n"));
  expect(rows.length).toBe(breaks.length);
});
