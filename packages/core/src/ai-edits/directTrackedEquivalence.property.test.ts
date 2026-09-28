/**
 * An edit applied directly and the same edit applied as tracked changes and
 * then accepted leave the same document: the same text, and every character
 * with the same formatting, link and comments.
 *
 * The two modes plan a replacement differently — the direct edit writes the
 * smallest change, the redline cuts it into readable words — so the rule that
 * allocates formatting to new text has to be one rule, stated once and used by
 * both. A replacement across a formatting boundary once left the last letter
 * of a new word outside the link or bold it took directly, and inside it when
 * accepted.
 *
 * The regressions go through the public reviewer, saved and reopened; the
 * property runs every text-changing operation over paragraphs of bold,
 * italic, linked, commented and field text in memory.
 */

import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import { propertyConfig, propertyTestTimeout } from "../../../../test/property-testing";

import {
  OperationSession,
  openReviewer,
  paragraphsDocx,
  reopened,
  reopenedAccepted,
  reviewComment,
  reviewerPresentation,
  textRun,
} from "../__tests__/operationBatchDocuments";
import { createDocx } from "../docx/rezip";
import { docxToMarkdown } from "../docx/server/docxToMarkdown";
import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperation,
} from "../document-operations";
import { fromMarkdown } from "../markdown";
import type { Comment, ParagraphContent } from "../types/document";
import { changesFromSegments, planTextChanges, type TextChange } from "./minimal-replacement";
import { createFolioAITextRangeHandle } from "./snapshot";
import { diffWordSegments } from "./word-diff";

setDefaultTimeout(propertyTestTimeout(60_000));

type Mode = "direct" | "tracked-changes";

const markdownAfter = async (source: string, mode: Mode, find: string, replace: string) => {
  const reviewer = await openReviewer(await createDocx(fromMarkdown(source)));
  const blockId = reviewer.snapshot().blocks[0]?.id ?? "";
  const result = reviewer.applyDocumentOperations({
    version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
    mode,
    operations: [{ id: "replace", type: "replaceInBlock", blockId, find, replace }],
  });
  const final = mode === "direct" ? await reopened(reviewer) : await reopenedAccepted(reviewer);
  const markdown = await docxToMarkdown(await final.toBuffer(), {
    annotations: "strip",
    trackedChanges: "clean",
    comments: "strip",
    footnotes: "keep",
  });
  return {
    result,
    markdown: (typeof markdown === "string" ? markdown : markdown.markdown).trim(),
  };
};

describe("a replacement across a formatting boundary", () => {
  test.each([
    [
      "a link",
      "The [Supplier](https://example.com/party) agrees.",
      "The [Provider performs](https://example.com/party).",
    ],
    ["bold", "The **Supplier** agrees.", "The **Provider performs**."],
  ])(
    "formats every new word whole, alike in both modes (%s, saved and reopened)",
    async (_, source, expected) => {
      for (const mode of ["direct", "tracked-changes"] as const) {
        const { result, markdown } = await markdownAfter(
          source,
          mode,
          "Supplier agrees",
          "Provider performs",
        );
        expect({ mode, markdown }).toEqual({ mode, markdown: expected });
        expect(result.skipped).toEqual([]);
        // The new text took one formatting over a stretch that had two: the
        // receipt says so, in either mode.
        expect(result.normalizations).toEqual([
          { id: "replace", code: "uniformReplacementFormatting" },
        ]);
      }
    },
  );

  test("keeps the formatting a word it only lengthens already had, alike in both modes", async () => {
    // `Supp` bold, `lier` plain: appending a letter keeps both runs, and the
    // accepted redline, which rewrites the word, allocates it the same way.
    const buffer = await paragraphsDocx([
      [textRun("The "), textRun("Supp", { bold: true }), textRun("lier agrees.")],
    ]);
    const presentations = [];
    for (const mode of ["direct", "tracked-changes"] as const) {
      const reviewer = await openReviewer(buffer.slice(0));
      const blockId = reviewer.snapshot().blocks[0]?.id ?? "";
      reviewer.applyDocumentOperations({
        version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
        mode,
        operations: [
          { id: "r", type: "replaceInBlock", blockId, find: "Supplier", replace: "Suppliers" },
        ],
      });
      const final = mode === "direct" ? await reopened(reviewer) : await reopenedAccepted(reviewer);
      presentations.push(await markdownOf(final));
    }
    expect(presentations[0]).toBe("The **Supp**liers agrees.");
    expect(presentations[1]).toBe(presentations[0]);
  });
});

const markdownOf = async (reviewer: Awaited<ReturnType<typeof openReviewer>>) => {
  const markdown = await docxToMarkdown(await reviewer.toBuffer(), {
    annotations: "strip",
    trackedChanges: "clean",
    comments: "strip",
    footnotes: "keep",
  });
  return (typeof markdown === "string" ? markdown : markdown.markdown).trim();
};

// ---------------------------------------------------------------------------
// The class: every text-changing operation, over every kind of mark boundary.
// ---------------------------------------------------------------------------

const WORDS = ["Seller", "shall", "deliver", "the", "goods", "on", "time", "Buyer", "pays"];

const words = fc
  .array(fc.constantFrom(...WORDS), { minLength: 1, maxLength: 3 })
  .map((picked) => `${picked.join(" ")} `);

type Item =
  | { kind: "run"; text: string; bold: boolean; italic: boolean }
  | { kind: "link"; text: string }
  | { kind: "comment"; text: string }
  | { kind: "field"; result: string };

const itemArbitrary: fc.Arbitrary<Item> = fc.oneof(
  {
    weight: 4,
    arbitrary: fc.record({
      kind: fc.constant("run" as const),
      text: words,
      bold: fc.boolean(),
      italic: fc.boolean(),
    }),
  },
  { weight: 1, arbitrary: fc.record({ kind: fc.constant("link" as const), text: words }) },
  { weight: 1, arbitrary: fc.record({ kind: fc.constant("comment" as const), text: words }) },
  {
    weight: 1,
    arbitrary: fc.record({
      kind: fc.constant("field" as const),
      result: fc.constantFrom("12", "Clause 4"),
    }),
  },
);

const paragraphOf = (
  items: readonly Item[],
): { content: ParagraphContent[]; comments: Comment[] } => {
  const content: ParagraphContent[] = [];
  const comments: Comment[] = [];
  for (const [index, item] of items.entries()) {
    switch (item.kind) {
      case "run":
        content.push(
          textRun(item.text, {
            ...(item.bold && { bold: true }),
            ...(item.italic && { italic: true }),
          }),
        );
        break;
      case "link":
        content.push({
          type: "hyperlink",
          href: `https://example.com/${String(index)}`,
          children: [textRun(item.text)],
        });
        break;
      case "comment":
        comments.push(reviewComment(index, `Note ${String(index)}`));
        content.push(
          { type: "commentRangeStart", id: index },
          textRun(item.text),
          { type: "commentRangeEnd", id: index },
          { type: "commentReference", id: index },
        );
        break;
      case "field":
        content.push({
          type: "simpleField",
          fieldType: "REF",
          instruction: ` REF _Ref${String(index)} \\h `,
          content: [textRun(item.result)],
        });
        break;
    }
  }
  return { content, comments };
};

const editArbitrary = fc.record({
  kind: fc.constantFrom(
    "replaceInBlock" as const,
    "replaceRange" as const,
    "replaceBlock" as const,
    "splitBlock" as const,
    "mergeBlockWithNext" as const,
    "deleteBlock" as const,
    "insertAfterBlock" as const,
  ),
  sliceStart: fc.double({ min: 0, max: 1, noNaN: true }),
  sliceLength: fc.double({ min: 0, max: 1, noNaN: true }),
  mutations: fc.array(
    fc.record({
      at: fc.double({ min: 0, max: 1, noNaN: true }),
      remove: fc.nat({ max: 6 }),
      insert: fc.constantFrom("", "s", "new ", " and", "X", "shall not "),
    }),
    { minLength: 1, maxLength: 3 },
  ),
});

type Edit = typeof editArbitrary extends fc.Arbitrary<infer Value> ? Value : never;

const mutate = (text: string, mutations: Edit["mutations"]): string => {
  let result = text;
  for (const { at, remove, insert } of mutations) {
    const position = Math.floor(at * result.length);
    result = result.slice(0, position) + insert + result.slice(position + remove);
  }
  return result;
};

/** The operation `edit` describes on block `blockId` reading `text`, if it describes one. */
const operationFor = (edit: Edit, blockId: string, text: string): FolioDocumentOperation | null => {
  const start = Math.floor(edit.sliceStart * text.length);
  const end = Math.min(
    text.length,
    start + 1 + Math.floor(edit.sliceLength * (text.length - start)),
  );
  const find = text.slice(start, end);
  switch (edit.kind) {
    case "replaceInBlock": {
      const replace = mutate(find, edit.mutations);
      return find.length === 0 || replace === find || text.indexOf(find) !== text.lastIndexOf(find)
        ? null
        : { id: "edit", type: "replaceInBlock", blockId, find, replace };
    }
    case "replaceRange": {
      const replace = mutate(find, edit.mutations);
      const range = createFolioAITextRangeHandle({
        blockId,
        text,
        startOffset: start,
        endOffset: end,
      });
      return range === null || replace === find
        ? null
        : { id: "edit", type: "replaceRange", range, replace };
    }
    case "replaceBlock": {
      const replaced = mutate(text, edit.mutations);
      return replaced === text
        ? null
        : { id: "edit", type: "replaceBlock", blockId, text: replaced };
    }
    case "splitBlock": {
      const space = text.indexOf(" ", start);
      return space <= 0 || space + 1 >= text.length
        ? null
        : { id: "edit", type: "splitBlock", blockId, offset: space, separator: " " };
    }
    case "mergeBlockWithNext":
      return { id: "edit", type: "mergeBlockWithNext", blockId, separator: " " };
    case "deleteBlock":
      return { id: "edit", type: "deleteBlock", blockId };
    case "insertAfterBlock":
      return {
        id: "edit",
        type: "insertAfterBlock",
        blockId,
        text: mutate("New clause.", edit.mutations),
      };
  }
};

/** Where each character `changes` keeps from `source` lands in the result. */
const keptCharacters = (source: string, changes: readonly TextChange[]): Map<number, number> => {
  const kept = new Map<number, number>();
  let shift = 0;
  let cursor = 0;
  for (const change of changes) {
    for (; cursor < change.start; cursor++) {
      kept.set(cursor, cursor + shift);
    }
    cursor = change.end;
    shift += change.text.length - (change.end - change.start);
  }
  for (; cursor < source.length; cursor++) {
    kept.set(cursor, cursor + shift);
  }
  return kept;
};

/**
 * Whether the redline keeps only characters the direct edit keeps, each where
 * the direct edit puts it. Both then write the same characters, and the one
 * allocation rule gives each the same formatting. Where the redline's
 * readable alignment keeps a character the shortest edit rewrites (a
 * repeated word matched at a different occurrence), the two keep different
 * characters of equal text, and only the text can be compared.
 */
const redlineKeepsWhatDirectKeeps = (source: string, replacement: string): boolean => {
  const direct = keptCharacters(source, planTextChanges(source, replacement));
  const tracked = keptCharacters(
    source,
    changesFromSegments(diffWordSegments(source, replacement)),
  );
  for (const [index, lands] of tracked) {
    if (direct.get(index) !== lands) {
      return false;
    }
  }
  return true;
};

const replacedSpan = (
  operation: FolioDocumentOperation,
  text: string,
): { source: string; replacement: string } | null => {
  switch (operation.type) {
    case "replaceInBlock":
      return { source: operation.find, replacement: operation.replace };
    case "replaceRange":
      return {
        source: text.slice(operation.range.startOffset, operation.range.endOffset),
        replacement: operation.replace,
      };
    case "replaceBlock":
      return { source: text, replacement: operation.text };
    default:
      return null;
  }
};

describe("an edit applied directly and the same edit accepted", () => {
  test("carries disjoint comments across a saved tracked replacement", async () => {
    const { content, comments } = paragraphOf([
      { kind: "comment", text: "Seller " },
      { kind: "field", result: "12" },
      { kind: "comment", text: "Seller " },
      { kind: "run", text: "Seller ", bold: false, italic: false },
    ]);
    const buffer = await paragraphsDocx([content], comments);
    const direct = await openReviewer(buffer.slice(0));
    const tracked = await openReviewer(buffer.slice(0));
    const operation = {
      id: "edit",
      type: "replaceBlock",
      blockId: direct.snapshot().blocks[0]?.id ?? "",
      text: "s12Seller Seller ",
    } as const;
    for (const [mode, reviewer] of [
      ["direct", direct],
      ["tracked-changes", tracked],
    ] as const) {
      expect(
        reviewer.applyDocumentOperations({
          version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
          mode,
          operations: [operation],
        }).skipped,
      ).toEqual([]);
    }
    const pending = await reopened(tracked);
    expect(pending.getComments()).toHaveLength(2);
    for (const comment of pending.getComments()) {
      expect(comment.anchoredText).toContain("s12Seller");
    }
    expect(reviewerPresentation(await reopenedAccepted(tracked))).toEqual(
      reviewerPresentation(await reopened(direct)),
    );
  });

  test("leave the same text, formatting, links and comments", async () => {
    const equivalenceProperty = fc.asyncProperty(
      fc.array(itemArbitrary, { minLength: 1, maxLength: 7 }),
      editArbitrary,
      fc.constantFrom("first", "last"),
      async (items, edit, placement) => {
        const { content, comments } = paragraphOf(items);
        const untouched = [textRun("Untouched paragraph.")];
        const paragraphs = placement === "first" ? [content, untouched] : [untouched, content];
        const base = await openReviewer(await paragraphsDocx(paragraphs, comments));
        const block = base.snapshot().blocks[placement === "first" ? 0 : 1];
        const operation = block && operationFor(edit, block.id, block.text);
        if (!block || !operation) {
          return;
        }
        const direct = new OperationSession(base.state);
        const tracked = new OperationSession(base.state);
        const directResult = direct.apply("direct", [operation]);
        const trackedResult = tracked.apply("tracked-changes", [operation]);
        // Both modes take the edit, or both refuse it for the same reason.
        expect(trackedResult.skipped).toEqual(directResult.skipped);
        expect(trackedResult.normalizations ?? []).toEqual(directResult.normalizations ?? []);
        tracked.acceptAll();
        const span = replacedSpan(operation, block.text);
        if (span !== null && !redlineKeepsWhatDirectKeeps(span.source, span.replacement)) {
          expect({ operation, text: tracked.presentation().map(({ text }) => text) }).toEqual({
            operation,
            text: direct.presentation().map(({ text }) => text),
          });
          return;
        }
        // Zero-width anchors (a comment's reference mark, a bookmark) inside a
        // stretch the redline rewrites stay by its deletion, while the direct
        // edit keeps the text around them: where such an anchor lands is the
        // planners' alignment, not the formatting rule, and is left out.
        expect({ operation, document: tracked.presentation({ withAnchors: false }) }).toEqual({
          operation,
          document: direct.presentation({ withAnchors: false }),
        });
      },
    );
    await fc.assert(equivalenceProperty, propertyConfig({ numRuns: 150, seed: 1383974287 }));
    await fc.assert(equivalenceProperty, propertyConfig({ numRuns: 150, seed: 1938427620 }));
    await fc.assert(equivalenceProperty, propertyConfig({ numRuns: 150 }));
  }, 300_000);
});
