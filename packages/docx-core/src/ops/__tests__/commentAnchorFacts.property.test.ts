/**
 * The comment anchors of a document read from facts kept per paragraph equal
 * the anchors read from every leaf.
 *
 * Each paragraph's anchors, text edges and text boxes are read once per
 * paragraph object, so a document an edit made is read by walking only the
 * paragraphs the edit replaced. Every version of a random run of edits is
 * read both ways: generated documents through document operations, and
 * documents with text boxes, tables and range markers between blocks through
 * edits that replace, insert and remove blocks while sharing the rest.
 */
import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../../test/property-testing";
import type {
  BlockContent,
  Document,
  Paragraph,
  ParagraphContent,
} from "../../model/document";
import { applyDocumentOp } from "../apply";
import { storyParagraphs } from "../blocks";
import {
  commentAnchorsIn,
  freshParagraphAnchorFacts,
  type StoryAnchorSpan,
} from "../comments";
import { type Gap, isCommentAnchor, leafSpans } from "../leaves";
import { documentStories, storyBody } from "../stories";
import { OP_STORIES, type OpStory } from "../types";
import {
  documentArbitrary,
  opForStory,
  opSeedArbitrary,
  reviewDocumentArbitrary,
  trackedOpFor,
} from "./documentArbitraries";

setDefaultTimeout(propertyTestTimeout(120_000));

type Body = { content: readonly BlockContent[] };
const ORIGIN: Gap = { offset: 0, zeroWidthBefore: 0 };

/** A story's body, then every text box body in it, listed off every paragraph. */
const bodiesOf = (body: Body): Body[] => [
  body,
  ...storyParagraphs({ content: [...body.content] }).flatMap(({ paragraph }) =>
    leafSpans(paragraph.content).flatMap(({ node }) =>
      node.type === "shape" && node.shape.textBody !== undefined
        ? bodiesOf({ content: node.shape.textBody.content })
        : [],
    ),
  ),
];

/**
 * A story's anchors read off every leaf: a marker between blocks takes the
 * start of the next paragraph's text, or the end of the one before.
 */
const anchorsOf = (story: OpStory, body: Body): StoryAnchorSpan[] => {
  const out: StoryAnchorSpan[] = [];
  const marker = (
    node: Extract<
      BlockContent,
      { type: "commentRangeStart" } | { type: "commentRangeEnd" }
    >,
    blockId: string,
    gap: Gap,
  ): StoryAnchorSpan => ({
    node,
    ancestors: [],
    before: gap,
    after: gap,
    story,
    blockId,
    betweenBlocks: true,
  });
  let waiting: Extract<
    BlockContent,
    { type: "commentRangeStart" } | { type: "commentRangeEnd" }
  >[] = [];
  let previous: { blockId: string; end: Gap } | undefined;
  const visit = (blocks: readonly BlockContent[]): void => {
    for (const block of blocks) {
      if (block.type === "paragraph") {
        const blockId = block.paraId ?? "";
        const spans = leafSpans(block.content);
        for (const node of waiting)
          out.push(marker(node, blockId, spans.at(0)?.before ?? ORIGIN));
        waiting = [];
        for (const span of spans)
          if (isCommentAnchor(span.node)) out.push({ story, blockId, ...span });
        previous = { blockId, end: spans.at(-1)?.after ?? ORIGIN };
      } else if (block.type === "table") {
        for (const row of block.rows)
          for (const cell of row.cells) visit(cell.content);
      } else if (block.type === "blockSdt" || block.type === "blockCustomXml") {
        visit(block.content);
      } else if (block.type === "commentRangeStart") {
        waiting.push(block);
      } else if (block.type === "commentRangeEnd") {
        if (previous === undefined) waiting.push(block);
        else out.push(marker(block, previous.blockId, previous.end));
      }
    }
  };
  visit(body.content);
  for (const node of waiting)
    out.push(marker(node, previous?.blockId ?? "", previous?.end ?? ORIGIN));
  return out;
};

/** Every anchor of a document read off every leaf, story by story. */
const everyLeafAnchors = (document: Document): StoryAnchorSpan[] =>
  documentStories(document).flatMap((story) =>
    bodiesOf(storyBody(document, story)).flatMap((body) =>
      anchorsOf(story, body),
    ),
  );

/** The anchors read from kept facts, from fresh facts and off every leaf agree. */
const expectSameAnchors = (document: Document): void => {
  const kept = commentAnchorsIn(document);
  expect(kept).toEqual(commentAnchorsIn(document, freshParagraphAnchorFacts));
  expect(kept).toEqual(everyLeafAnchors(document));
};

const anchorArbitrary: fc.Arbitrary<ParagraphContent> = fc.oneof(
  fc
    .integer({ min: 0, max: 3 })
    .map((id): ParagraphContent => ({ type: "commentRangeStart", id })),
  fc
    .integer({ min: 0, max: 3 })
    .map((id): ParagraphContent => ({ type: "commentRangeEnd", id })),
  fc.integer({ min: 0, max: 3 }).map((id): ParagraphContent => ({
    type: "run",
    content: [{ type: "commentReference", id }],
  })),
);

const textArbitrary: fc.Arbitrary<ParagraphContent> = fc
  .string({ minLength: 1, maxLength: 4 })
  .map((text) => ({ type: "run", content: [{ type: "text", text }] }));

/** A paragraph's content: text and anchors, and maybe a text box of paragraphs. */
const paragraphContentArbitrary = (
  depth: number,
): fc.Arbitrary<ParagraphContent[]> =>
  fc
    .tuple(
      fc.array(fc.oneof(textArbitrary, anchorArbitrary), { maxLength: 4 }),
      depth > 0
        ? fc.option(
            fc.array(paragraphContentArbitrary(depth - 1), {
              minLength: 1,
              maxLength: 2,
            }),
          )
        : fc.constant(null),
      fc.nat(),
    )
    .map(([content, box, at]) => {
      if (box === null) return content;
      const shape = {
        type: "shape",
        shape: {
          textBody: {
            content: box.map((inner, index): Paragraph => ({
              type: "paragraph",
              paraId: `5${String(depth)}${String(index).padStart(6, "0")}`,
              content: inner,
            })),
          },
        },
      };
      const out = content.slice();
      // SAFETY: a shape holding only a text body is all the anchor walk reads.
      out.splice(at % (out.length + 1), 0, {
        type: "run",
        content: [shape],
      } as unknown as ParagraphContent);
      return out;
    });

let nextParagraph = 0x10000000;
const paragraphOf = (content: ParagraphContent[]): Paragraph => {
  nextParagraph += 1;
  return {
    type: "paragraph",
    paraId: nextParagraph.toString(16).toUpperCase(),
    content,
  };
};

const blockArbitrary: fc.Arbitrary<BlockContent> = fc.oneof(
  { weight: 4, arbitrary: paragraphContentArbitrary(2).map(paragraphOf) },
  {
    weight: 1,
    arbitrary: fc
      .array(paragraphContentArbitrary(1), { minLength: 1, maxLength: 3 })
      .map((cells): BlockContent => ({
        type: "table",
        columnWidths: cells.map(() => 1000),
        rows: [
          {
            type: "tableRow",
            cells: cells.map((content) => ({
              type: "tableCell",
              content: [paragraphOf(content)],
            })),
          },
        ],
      })),
  },
  {
    weight: 1,
    arbitrary: fc
      .integer({ min: 0, max: 3 })
      .chain((id) =>
        fc.constantFrom<BlockContent>(
          { type: "commentRangeStart", id },
          { type: "commentRangeEnd", id },
        ),
      ),
  },
);

const editArbitrary = fc.record({
  kind: fc.constantFrom("replace", "insert", "remove", "retext"),
  at: fc.nat(),
  block: blockArbitrary,
});
type Edit = fc.RecordValue<typeof editArbitrary>;

/** The blocks after one edit, sharing every block it left alone. */
const edited = (
  blocks: readonly BlockContent[],
  edit: Edit,
): BlockContent[] => {
  const out = [...blocks];
  const at = edit.at % (blocks.length + 1);
  switch (edit.kind) {
    case "insert":
      out.splice(at, 0, edit.block);
      return out;
    case "remove":
      out.splice(at, 1);
      return out;
    case "replace":
      out.splice(at, 1, edit.block);
      return out;
    case "retext": {
      // A new paragraph object under the same id, as typing makes one,
      // with one more anchor and one more letter.
      const old = out[at];
      if (old?.type !== "paragraph") return out;
      out[at] = {
        ...old,
        content: [
          { type: "run", content: [{ type: "commentReference", id: 0 }] },
          ...old.content,
          { type: "run", content: [{ type: "text", text: "x" }] },
        ],
      };
      return out;
    }
    default: {
      const unhandled: never = edit.kind;
      return unhandled;
    }
  }
};

const documentWith = (content: BlockContent[]): Document =>
  // SAFETY: the anchor walk reads only the stories' block content.
  ({ package: { document: { content, comments: [] } } }) as unknown as Document;

describe("comment anchors read from kept paragraph facts", () => {
  test("equal anchors read from every leaf over a run of operations", async () => {
    await assertProperty(
      fc.property(
        fc.oneof(documentArbitrary, reviewDocumentArbitrary),
        fc.array(
          fc.record({
            seed: opSeedArbitrary,
            tracked: fc.boolean(),
            story: fc.nat(),
          }),
          { minLength: 1, maxLength: 12 },
        ),
        (generated, edits) => {
          let document = generated;
          expectSameAnchors(document);
          edits.forEach(({ seed, tracked, story: storyIndex }, index) => {
            const stories = documentStories(document);
            const story =
              stories[storyIndex % stories.length] ?? OP_STORIES.MAIN;
            const op =
              tracked && story === OP_STORIES.MAIN
                ? trackedOpFor(document, seed, index)
                : opForStory({ document, seed, story });
            const applied = applyDocumentOp(document, op);
            if (applied.isErr()) return;
            document = applied.value.document;
            expectSameAnchors(document);
          });
        },
      ),
      { id: "comment anchor facts follow operations", numRuns: 150 },
    );
  });

  test("equal anchors read from every leaf through text boxes, tables and markers between blocks", async () => {
    await assertProperty(
      fc.property(
        fc.array(blockArbitrary, { minLength: 1, maxLength: 8 }),
        fc.array(editArbitrary, { minLength: 1, maxLength: 16 }),
        (start, edits) => {
          let blocks: BlockContent[] = start;
          expectSameAnchors(documentWith(blocks));
          for (const edit of edits) {
            blocks = edited(blocks, edit);
            expectSameAnchors(documentWith(blocks));
          }
        },
      ),
      { id: "comment anchor facts follow block edits", numRuns: 300 },
    );
  });
});
