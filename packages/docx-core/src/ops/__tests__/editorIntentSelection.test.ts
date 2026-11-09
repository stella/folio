import { expect, test } from "bun:test";

import type {
  BlockContent,
  Document,
  Paragraph,
  ParagraphContent,
} from "../../model/document";
import { applyDocumentOps } from "../apply";
import { storyParagraphs } from "../blocks";
import { compileEditorIntent, type EditorIntent } from "../editorIntent";
import { paragraphLength } from "../offsets";
import { OP_STORIES, type TextPosition } from "../types";

const run = (text: string): ParagraphContent => ({
  type: "run",
  content: [{ type: "text", text }],
});
const paragraph = (paraId: string, text: string): Paragraph => ({
  type: "paragraph",
  paraId,
  content: text === "" ? [] : [run(text)],
});
const documentOf = (content: BlockContent[]): Document => ({
  package: { document: { content } },
});
const wrap = (
  content: BlockContent[],
  container: "body" | "cell",
): BlockContent[] =>
  container === "body"
    ? content
    : [
        {
          type: "table",
          rows: [{ type: "tableRow", cells: [{ type: "tableCell", content }] }],
        },
      ];

/** Compiles `intent` directly and returns the edited document and its caret. */
const compileDirect = (document: Document, intent: EditorIntent) => {
  const compiled = compileEditorIntent(document, {
    intent,
    mode: { type: "editing" },
  });
  if (compiled.isErr()) throw compiled.error;
  const applied = applyDocumentOps(document, compiled.value.ops);
  if (applied.isErr()) throw applied.error;
  return {
    document: applied.value.document,
    selection: compiled.value.selection,
  };
};

/** The paragraph the caret names, which must exist after the edit. */
const caretParagraph = (document: Document, selection: TextPosition) => {
  const found = storyParagraphs(document.package.document).find(
    ({ paragraph: block }) => block.paraId === selection.blockId,
  )?.paragraph;
  expect(found).toBeDefined();
  if (found === undefined) throw new Error("The caret names no paragraph.");
  expect(selection.offset).toBeLessThanOrEqual(paragraphLength(found));
  return found;
};

test.each(["body", "cell"] as const)(
  "a direct join keeping the first paragraph leaves the caret in it at the seam (%s)",
  (container) => {
    const document = documentOf(
      wrap(
        [paragraph("0000000A", "Alpha"), paragraph("0000000B", "Beta")],
        container,
      ),
    );
    const { document: joined, selection } = compileDirect(document, {
      type: "joinParagraphs",
      story: OP_STORIES.MAIN,
      blockId: "0000000A",
      nextBlockId: "0000000B",
      survivor: "first",
    });
    expect(selection).toStrictEqual({
      story: OP_STORIES.MAIN,
      blockId: "0000000A",
      offset: "Alpha".length,
    });
    caretParagraph(joined, selection);
  },
);

test.each(["body", "cell"] as const)(
  "a direct join keeping the second paragraph leaves the caret in it at the seam (%s)",
  (container) => {
    const document = documentOf(
      wrap(
        [paragraph("0000000A", "Alpha"), paragraph("0000000B", "Beta")],
        container,
      ),
    );
    const { document: joined, selection } = compileDirect(document, {
      type: "joinParagraphs",
      story: OP_STORIES.MAIN,
      blockId: "0000000A",
      nextBlockId: "0000000B",
    });
    expect(selection).toStrictEqual({
      story: OP_STORIES.MAIN,
      blockId: "0000000B",
      offset: "Alpha".length,
    });
    caretParagraph(joined, selection);
  },
);

test.each(["body", "cell"] as const)(
  "directly deleting paragraphs leaves the caret in a paragraph that stays (%s)",
  (container) => {
    const blocks = [
      paragraph("0000000F", "Before"),
      paragraph("0000000A", "Alpha"),
      paragraph("0000000B", "Beta"),
      paragraph("0000000C", "Gamma"),
    ];
    const cases: {
      blockIds: string[];
      expected: { blockId: string; offset: number; text: string };
    }[] = [
      // A following paragraph takes the deleted run's place.
      {
        blockIds: ["0000000A", "0000000B"],
        expected: { blockId: "0000000C", offset: 0, text: "Gamma" },
      },
      {
        blockIds: ["0000000F"],
        expected: { blockId: "0000000A", offset: 0, text: "Alpha" },
      },
      // The container's final paragraph stays and takes its predecessor's
      // content, so the caret ends where that content ends.
      {
        blockIds: ["0000000B", "0000000C"],
        expected: {
          blockId: "0000000C",
          offset: "Alpha".length,
          text: "Alpha",
        },
      },
      {
        blockIds: ["0000000C"],
        expected: { blockId: "0000000C", offset: "Beta".length, text: "Beta" },
      },
      {
        blockIds: ["0000000F", "0000000A", "0000000B", "0000000C"],
        expected: { blockId: "0000000C", offset: 0, text: "" },
      },
    ];
    for (const { blockIds, expected } of cases) {
      const document = documentOf(wrap(blocks, container));
      const { document: deleted, selection } = compileDirect(document, {
        type: "deleteParagraphs",
        story: OP_STORIES.MAIN,
        blockIds,
      });
      expect(selection).toStrictEqual({
        story: OP_STORIES.MAIN,
        blockId: expected.blockId,
        offset: expected.offset,
      });
      const kept = caretParagraph(deleted, selection);
      expect(
        kept.content
          .flatMap((node) => (node.type === "run" ? node.content : []))
          .map((leaf) => (leaf.type === "text" ? leaf.text : ""))
          .join(""),
      ).toBe(expected.text);
    }
  },
);
