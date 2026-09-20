/**
 * A save never leaves a note part with half a comment range.
 *
 * A comment can be anchored on a footnote's or endnote's own text, so its
 * range lives in `word/footnotes.xml` / `word/endnotes.xml` and spans that
 * note's paragraphs. Both saves patch a note part by splicing the paragraphs
 * an edit touched and keeping the rest byte-for-byte, so an edit that moves a
 * range's only half out of a spliced paragraph writes the other half alone:
 * invalid OOXML, and a comment anchored to nothing.
 *
 * Which arrangement produces one is the variable, so the property generates
 * the span (any sub-range of the note's paragraphs), where each half sits in
 * its paragraph, and the edit (each operation kind, each mode, any paragraph)
 * rather than pinning one case: the defect needed a range opening at the END
 * of the paragraph it was authored in, so the only commented text is in the
 * next one and the model writes the start there instead.
 *
 * The save must succeed — declining the splice for a wider rewrite, never
 * throwing — and every story in the saved package must be balanced.
 */

import { describe, expect, test } from "bun:test";
import fc from "fast-check";
import JSZip from "jszip";

import { propertyConfig } from "../../../../test/property-testing";

import { FolioDocxReviewer } from "../ai-edits/headless";
import type { FolioAIEditOperation } from "../ai-edits/types";
import type { Comment, Document, Paragraph, ParagraphContent } from "../types/document";
import { unbalancedCommentRangeIds } from "./commentRangeIntegrity";
import { createDocx } from "./rezip";

const NOTE_PARAGRAPH_COUNT = 4;
const NOTE_ID = 2;
const COMMENT_ID = 0;
const NOTE_KINDS = ["footnote", "endnote"] as const;
const EDIT_MODES = ["direct", "tracked-changes"] as const;
const OPERATION_KINDS = ["replaceInBlock", "insertAfterBlock", "deleteBlock"] as const;
/** Whether a range half sits before the paragraph's text or after it. */
const MARKER_PLACEMENTS = ["leading", "trailing"] as const;

type GeneratedCase = {
  noteKind: (typeof NOTE_KINDS)[number];
  span: { first: number; last: number };
  startPlacement: (typeof MARKER_PLACEMENTS)[number];
  endPlacement: (typeof MARKER_PLACEMENTS)[number];
  edit: {
    kind: (typeof OPERATION_KINDS)[number];
    mode: (typeof EDIT_MODES)[number];
    blockIndex: number;
  };
};

const noteParagraphIndex = fc.nat({ max: NOTE_PARAGRAPH_COUNT - 1 });

const generatedCase = fc
  .record({
    noteKind: fc.constantFrom(...NOTE_KINDS),
    bounds: fc.tuple(noteParagraphIndex, noteParagraphIndex),
    startPlacement: fc.constantFrom(...MARKER_PLACEMENTS),
    endPlacement: fc.constantFrom(...MARKER_PLACEMENTS),
    edit: fc.record({
      kind: fc.constantFrom(...OPERATION_KINDS),
      mode: fc.constantFrom(...EDIT_MODES),
      blockIndex: noteParagraphIndex,
    }),
  })
  .map(({ noteKind, bounds: [one, other], startPlacement, endPlacement, edit }): GeneratedCase => {
    const span = { first: Math.min(one, other), last: Math.max(one, other) };
    // A range confined to one paragraph has to open before it closes, so the
    // placements are only free when the span crosses a paragraph boundary.
    const confined = span.first === span.last;
    return {
      noteKind,
      span,
      startPlacement: confined ? "leading" : startPlacement,
      endPlacement: confined ? "trailing" : endPlacement,
      edit,
    };
  });

const paraId = (index: number): string => `5100000${index}`;
const noteParagraphText = (index: number): string => `Note paragraph ${index}.`;

const run = (text: string): ParagraphContent => ({
  type: "run",
  content: [{ type: "text", text }],
});

const comment: Comment = {
  id: COMMENT_ID,
  author: "Dana Lindqvist",
  initials: "DL",
  date: "2024-01-01T00:00:00Z",
  content: [{ type: "paragraph", content: [run("Check this note.")] }],
};

const buildDocument = ({
  noteKind,
  span,
  startPlacement,
  endPlacement,
}: GeneratedCase): Document => {
  const noteContent: Paragraph[] = [];
  for (let index = 0; index < NOTE_PARAGRAPH_COUNT; index++) {
    const beforeText: ParagraphContent[] = [];
    const afterText: ParagraphContent[] = [];
    if (span.first === index) {
      (startPlacement === "leading" ? beforeText : afterText).push({
        type: "commentRangeStart",
        id: COMMENT_ID,
      });
    }
    if (span.last === index) {
      (endPlacement === "leading" ? beforeText : afterText).push(
        { type: "commentRangeEnd", id: COMMENT_ID },
        { type: "commentReference", id: COMMENT_ID },
      );
    }
    noteContent.push({
      type: "paragraph",
      paraId: paraId(index),
      content: [...beforeText, run(noteParagraphText(index)), ...afterText],
    });
  }

  const note = { type: noteKind, id: NOTE_ID, content: noteContent } as const;
  const body: Paragraph[] = [
    {
      type: "paragraph",
      paraId: "52000000",
      content: [
        run("The clause the note hangs off."),
        {
          type: "run",
          content: [
            noteKind === "footnote"
              ? { type: "footnoteRef", id: NOTE_ID }
              : { type: "endnoteRef", id: NOTE_ID },
          ],
        },
      ],
    },
  ];

  return {
    package: {
      document: { comments: [comment], content: body },
      ...(noteKind === "footnote" ? { footnotes: [note] } : { endnotes: [note] }),
    },
  };
};

const operationFor = ({ edit }: GeneratedCase, blockId: string): FolioAIEditOperation => {
  switch (edit.kind) {
    case "replaceInBlock":
      return {
        id: "edit",
        type: "replaceInBlock",
        blockId,
        find: noteParagraphText(edit.blockIndex),
        replace: "Superseded note wording.",
      };
    case "insertAfterBlock":
      return { id: "edit", type: "insertAfterBlock", blockId, text: "An inserted note line." };
    case "deleteBlock":
      return { id: "edit", type: "deleteBlock", blockId };
    default: {
      const unreachable: never = edit.kind;
      throw new Error(`Unhandled operation kind: ${String(unreachable)}`);
    }
  }
};

/** Every story a comment range can live in, so none passes by not being read. */
const STORY_PARTS = ["word/document.xml", "word/footnotes.xml", "word/endnotes.xml"] as const;

describe("a save leaves every note part's comment ranges balanced", () => {
  test("over generated spans, marker placements and edits", async () => {
    await fc.assert(
      fc.asyncProperty(generatedCase, async (generated) => {
        const source = await createDocx(buildDocument(generated));
        const reviewer = await FolioDocxReviewer.fromBuffer(source, { author: "Editor" });
        const story = { type: generated.noteKind, noteId: NOTE_ID } as const;
        const target = reviewer.snapshotStory(story)?.blocks[generated.edit.blockIndex];
        expect(target).toBeDefined();
        if (!target) {
          return;
        }

        reviewer.applyDocumentOperationsToStory({
          story,
          batch: {
            version: 1,
            operations: [operationFor(generated, target.id)],
            mode: generated.edit.mode,
          },
        });

        const saved = await reviewer.toBuffer();
        const zip = await JSZip.loadAsync(saved);
        const parts = await Promise.all(
          STORY_PARTS.map(async (part) => ({ part, xml: await zip.file(part)?.async("text") })),
        );
        for (const { part, xml } of parts) {
          if (xml === undefined) {
            continue;
          }
          expect({ part, unbalanced: [...unbalancedCommentRangeIds(xml)] }).toEqual({
            part,
            unbalanced: [],
          });
        }
      }),
      propertyConfig({ numRuns: 60 }),
    );
  }, 180_000);
});
