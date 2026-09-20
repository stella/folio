/**
 * No edit may leave a range with one half.
 *
 * A `commentRangeStart` without its end, or an end without its start, is
 * invalid OOXML and a comment that anchors nothing; a bookmark's two
 * boundaries owe the same. Which edit produces one is the variable, so the
 * property generates the document (1-5 comments, nested and overlapping,
 * spanning 1-5 paragraphs, plus bookmarks) and the edit (each operation kind,
 * each mode, any block) rather than pinning one arrangement in a fixture: the
 * defect this covers needed a range whose only start sat in the edited block,
 * which no example set had.
 *
 * A surviving comment must also keep what it is — id, author, text — and a
 * range still spanning every paragraph of its original span the edit left
 * standing. It may disappear only once all of its commented text is gone.
 */

import { describe, expect, test } from "bun:test";
import fc from "fast-check";
import JSZip from "jszip";

import { propertyConfig } from "../../../../test/property-testing";

import { parseDocx } from "../docx/parser";
import { createDocx } from "../docx/rezip";
import { findParagraphOffsets } from "../docx/selectiveXmlPatch";
import type { Comment, Document, Paragraph, ParagraphContent } from "../types/document";
import { FolioDocxReviewer } from "./headless";
import type { FolioAIEditOperation } from "./types";

const PARAGRAPH_COUNT = 5;
const AUTHORS = ["Dana Lindqvist", "Ravi Mehrotra", "Sofia Achebe"] as const;
const EDIT_MODES = ["direct", "tracked-changes"] as const;
const OPERATION_KINDS = ["replaceInBlock", "insertAfterBlock", "deleteBlock"] as const;

const FOOTNOTE_ID = 2;

type GeneratedSpan = { first: number; last: number };
type GeneratedComment = GeneratedSpan & { author: string; text: string };
type GeneratedEdit = {
  kind: (typeof OPERATION_KINDS)[number];
  mode: (typeof EDIT_MODES)[number];
  blockIndex: number;
};
type GeneratedCase = {
  comments: readonly GeneratedComment[];
  bookmarks: readonly GeneratedSpan[];
  /** Paragraphs whose text is a hyperlink, and the one holding a note ref. */
  linkedParagraphs: readonly boolean[];
  noteRefParagraph: number | null;
  edit: GeneratedEdit;
};

const paragraphIndex = fc.nat({ max: PARAGRAPH_COUNT - 1 });

const spanArbitrary = fc
  .tuple(paragraphIndex, paragraphIndex)
  .map(([one, other]) => ({ first: Math.min(one, other), last: Math.max(one, other) }));

const generatedCase = fc.record({
  comments: fc.array(
    fc.record({
      span: spanArbitrary,
      author: fc.constantFrom(...AUTHORS),
      text: fc.stringMatching(/^[A-Za-z]{1,12}$/u),
    }),
    { minLength: 1, maxLength: 5 },
  ),
  bookmarks: fc.array(spanArbitrary, { maxLength: 2 }),
  linkedParagraphs: fc.array(fc.boolean(), {
    minLength: PARAGRAPH_COUNT,
    maxLength: PARAGRAPH_COUNT,
  }),
  noteRefParagraph: fc.option(paragraphIndex, { nil: null }),
  edit: fc.record({
    kind: fc.constantFrom(...OPERATION_KINDS),
    mode: fc.constantFrom(...EDIT_MODES),
    blockIndex: paragraphIndex,
  }),
});

const toCase = ({
  comments,
  ...rest
}: typeof generatedCase extends fc.Arbitrary<infer T> ? T : never): GeneratedCase => ({
  ...rest,
  comments: comments.map(({ span, author, text }) => ({ ...span, author, text })),
});

const PARA_ID_PREFIX = "2000000";
const paraId = (index: number): string => `${PARA_ID_PREFIX}${index}`;
const paragraphText = (index: number): string => `Paragraph ${index} of the agreement.`;
const linkTarget = (index: number): string => `https://example.invalid/clause-${index}`;

const run = (text: string): ParagraphContent => ({
  type: "run",
  content: [{ type: "text", text }],
});

const buildDocument = ({
  comments,
  bookmarks,
  linkedParagraphs,
  noteRefParagraph,
}: GeneratedCase): Document => {
  const content: Paragraph[] = [];
  for (let index = 0; index < PARAGRAPH_COUNT; index++) {
    const items: ParagraphContent[] = [];
    for (const [id, { first }] of comments.entries()) {
      if (first === index) {
        items.push({ type: "commentRangeStart", id });
      }
    }
    for (const [id, { first }] of bookmarks.entries()) {
      if (first === index) {
        items.push({ type: "bookmarkStart", id, name: `bookmark${id}` });
      }
    }
    items.push(
      linkedParagraphs[index]
        ? {
            type: "hyperlink",
            href: linkTarget(index),
            children: [{ type: "run", content: [{ type: "text", text: paragraphText(index) }] }],
          }
        : run(paragraphText(index)),
    );
    if (noteRefParagraph === index) {
      items.push({ type: "run", content: [{ type: "footnoteRef", id: FOOTNOTE_ID }] });
    }
    for (const [id, { last }] of bookmarks.entries()) {
      if (last === index) {
        items.push({ type: "bookmarkEnd", id });
      }
    }
    for (const [id, { last }] of comments.entries()) {
      if (last === index) {
        items.push({ type: "commentRangeEnd", id }, { type: "commentReference", id });
      }
    }
    content.push({ type: "paragraph", paraId: paraId(index), content: items });
  }

  const authored: Comment[] = comments.map(({ author, text }, id) => ({
    id,
    author,
    initials: author.slice(0, 2).toUpperCase(),
    date: "2024-01-01T00:00:00Z",
    content: [{ type: "paragraph", content: [run(text)] }],
  }));

  return {
    package: {
      document: { comments: authored, content },
      ...(noteRefParagraph === null
        ? {}
        : {
            footnotes: [
              {
                type: "footnote" as const,
                id: FOOTNOTE_ID,
                content: [{ type: "paragraph" as const, content: [run("A note.")] }],
              },
            ],
          }),
    },
  };
};

const operationFor = (
  { kind, blockIndex }: GeneratedEdit,
  blockId: string,
): FolioAIEditOperation => {
  switch (kind) {
    case "replaceInBlock":
      return {
        id: "edit",
        type: "replaceInBlock",
        blockId,
        find: paragraphText(blockIndex),
        replace: "Superseded wording throughout.",
      };
    case "insertAfterBlock":
      return { id: "edit", type: "insertAfterBlock", blockId, text: "An inserted paragraph." };
    case "deleteBlock":
      return { id: "edit", type: "deleteBlock", blockId };
    default: {
      const unreachable: never = kind;
      throw new Error(`Unhandled operation kind: ${String(unreachable)}`);
    }
  }
};

type MarkerOccurrence = { id: number; index: number };

const markers = (xml: string, marker: string): MarkerOccurrence[] =>
  [...xml.matchAll(new RegExp(`<w:${marker}\\s[^>]*w:id="(\\d+)"`, "gu"))].map((match) => ({
    id: Number(match[1]),
    index: match.index,
  }));

/**
 * Every id present as one half of a range is present as the other, start
 * first. Returns each balanced id's outermost span in the XML.
 */
const assertBalancedRanges = (xml: string, prefix: string): Map<number, GeneratedSpan> => {
  const starts = markers(xml, `${prefix}Start`);
  const ends = markers(xml, `${prefix}End`);
  const spans = new Map<number, GeneratedSpan>();
  for (const id of new Set([...starts, ...ends].map((marker) => marker.id))) {
    const idStarts = starts.filter((marker) => marker.id === id).map(({ index }) => index);
    const idEnds = ends.filter((marker) => marker.id === id).map(({ index }) => index);
    expect({
      marker: prefix,
      id,
      opens: idStarts.length > 0,
      closes: idEnds.length > 0,
      opensFirst: (idStarts.at(0) ?? Number.POSITIVE_INFINITY) < (idEnds.at(-1) ?? -1),
    }).toEqual({ marker: prefix, id, opens: true, closes: true, opensFirst: true });
    // SAFETY: the assertion above fails the test when either half is missing.
    spans.set(id, { first: idStarts[0]!, last: idEnds.at(-1)! });
  }
  return spans;
};

const commentPlainText = (comment: Comment | undefined): string =>
  (comment?.content ?? [])
    .flatMap((block) => (block.type === "paragraph" ? block.content : []))
    .flatMap((item) => (item.type === "run" ? item.content : []))
    .map((item) => (item.type === "text" ? item.text : ""))
    .join("");

describe("an edit leaves every comment and bookmark range balanced and anchored", () => {
  test("over generated documents and operations", async () => {
    await fc.assert(
      fc.asyncProperty(generatedCase, async (raw) => {
        const generated = toCase(raw);
        const source = await createDocx(buildDocument(generated));
        const reviewer = await FolioDocxReviewer.fromBuffer(source, { author: "Editor" });
        const target = reviewer.snapshot().blocks[generated.edit.blockIndex];
        expect(target).toBeDefined();
        if (!target) {
          return;
        }
        reviewer.applyOperations([operationFor(generated.edit, target.id)], {
          mode: generated.edit.mode,
        });

        const saved = await reviewer.toBuffer();
        const zip = await JSZip.loadAsync(saved);
        const xml = (await zip.file("word/document.xml")?.async("text")) ?? "";
        const rels = (await zip.file("word/_rels/document.xml.rels")?.async("text")) ?? "";

        const commentSpans = assertBalancedRanges(xml, "commentRange");
        assertBalancedRanges(xml, "bookmark");

        const survivingParagraphs = new Map(
          Array.from({ length: PARAGRAPH_COUNT }, (_, index) => index)
            .map((index) => [index, findParagraphOffsets(xml, paraId(index))] as const)
            .filter(([, offsets]) => offsets !== null),
        );

        // A link survives a replacement of the text it wrapped, and still
        // wraps it: the mark is non-inclusive like `comment`, and was lost the
        // same way.
        for (const [index, linked] of generated.linkedParagraphs.entries()) {
          const offsets = survivingParagraphs.get(index);
          if (!linked || !offsets) {
            continue;
          }
          const paragraph = xml.slice(offsets.start, offsets.end);
          const visible = [...paragraph.matchAll(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/gu)]
            .map(([, part]) => part ?? "")
            .join("");
          const inside = [...paragraph.matchAll(/<w:hyperlink\b[^>]*>([\s\S]*?)<\/w:hyperlink>/gu)]
            .map(([, part]) => part ?? "")
            .join("");
          const insideVisible = [...inside.matchAll(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/gu)]
            .map(([, part]) => part ?? "")
            .join("");
          expect({
            index,
            target: rels.includes(linkTarget(index)),
            wrapsItsText: visible.length === 0 || insideVisible === visible,
          }).toEqual({ index, target: true, wrapsItsText: true });
        }

        // A note reference marks its own number, not prose, so it follows the
        // text it sat in rather than being carried onto a replacement — which
        // would serialize the replacement as a bare reference and lose it.
        if (generated.noteRefParagraph !== null) {
          const offsets = survivingParagraphs.get(generated.noteRefParagraph);
          const replacedIt =
            generated.edit.kind === "replaceInBlock" &&
            generated.edit.blockIndex === generated.noteRefParagraph;
          if (offsets && replacedIt) {
            expect(xml.slice(offsets.start, offsets.end)).toContain(
              "Superseded wording throughout.",
            );
          }
        }

        const reparsed = await parseDocx(saved, { preloadFonts: false });
        const authored = new Map(
          (reparsed.package.document.comments ?? []).map((entry) => [entry.id, entry]),
        );

        for (const [id, { author, text, first, last }] of generated.comments.entries()) {
          const covered = Array.from({ length: last - first + 1 }, (_, offset) => first + offset)
            .map((index) => survivingParagraphs.get(index))
            .filter((offsets) => offsets !== undefined && offsets !== null);
          if (covered.length === 0) {
            // Every character the comment covered is gone; losing it is correct.
            continue;
          }

          expect({ id, kept: authored.has(id), anchored: commentSpans.has(id) }).toEqual({
            id,
            kept: true,
            anchored: true,
          });
          expect(authored.get(id)?.author).toBe(author);
          expect(commentPlainText(authored.get(id))).toBe(text);

          // The range still reaches the surviving ends of its original span.
          const span = commentSpans.get(id);
          expect({
            id,
            opensInTime: (span?.first ?? Number.POSITIVE_INFINITY) < (covered.at(0)?.end ?? -1),
            closesInTime: (span?.last ?? -1) > (covered.at(-1)?.start ?? Number.POSITIVE_INFINITY),
          }).toEqual({ id, opensInTime: true, closesInTime: true });
        }
      }),
      propertyConfig({ numRuns: 40 }),
    );
  }, 180_000);
});
