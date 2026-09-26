/**
 * An offset an operation takes must stand between two characters.
 *
 * Offsets count UTF-16 code units. One between the two halves of a surrogate
 * pair names no character, and text cut there cannot be saved: each half is a
 * lone surrogate, which XML cannot hold, so the save drops both and the
 * character is gone — even for an operation that only formats or comments. An
 * edit of text or a paragraph break inside a grapheme cluster (a letter and
 * its combining marks, a joined emoji, a flag) changes a character nobody
 * named. Such offsets are refused (`splitsCharacter`) and nothing is applied.
 */

import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import { propertyConfig, propertyTestTimeout } from "../../../../test/property-testing";

import {
  blockTexts,
  openReviewer,
  paragraphsDocx,
  reopened,
  reopenedAccepted,
  textRun,
} from "../__tests__/operationBatchDocuments";
import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperation,
} from "../document-operations";
import { splitsGraphemeCluster, splitsSurrogatePair } from "./character-boundaries";
import type { FolioDocxReviewer } from "./headless";
import { createFolioAITextRangeHandle, hashFolioAIBlockText } from "./snapshot";
import type { FolioAITextRangeHandle } from "./types";

setDefaultTimeout(propertyTestTimeout(60_000));

type Mode = "direct" | "tracked-changes";
const MODES: readonly Mode[] = ["direct", "tracked-changes"];

/** A range handle as a caller may build one by hand, bypassing the helper's checks. */
const handle = (
  blockId: string,
  text: string,
  startOffset: number,
  endOffset: number,
): FolioAITextRangeHandle => ({
  type: "textRange",
  story: "main",
  blockId,
  startOffset,
  endOffset,
  selectedTextHash: hashFolioAIBlockText(text.slice(startOffset, endOffset)),
});

const applyOne = (reviewer: FolioDocxReviewer, mode: Mode, operation: FolioDocumentOperation) =>
  reviewer.applyDocumentOperations({
    version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
    mode,
    operations: [operation],
  });

const settled = (reviewer: FolioDocxReviewer, mode: Mode) =>
  mode === "direct" ? reopened(reviewer) : reopenedAccepted(reviewer);

const PARTY = "Party 🧑 agrees.";

describe("an offset inside an emoji", () => {
  test.each(MODES)(
    "refuses a split between its halves and splits beside it (%s, saved and reopened)",
    async (mode) => {
      const buffer = await paragraphsDocx([[textRun(PARTY)]]);
      const inside = await openReviewer(buffer.slice(0));
      const blockId = inside.snapshot().blocks[0]?.id ?? "";
      const refused = applyOne(inside, mode, {
        id: "split",
        type: "splitBlock",
        blockId,
        offset: 7,
      });
      expect(refused.skipped).toEqual([
        {
          id: "split",
          reason: "splitsCharacter",
          message: 'offset 7 falls inside "🧑", which spans offsets 6 to 8; use 6 or 8.',
        },
      ]);
      expect(refused.issues?.[0]).toMatchObject({
        code: "splitsCharacter",
        retryable: true,
        recovery: "changeTarget",
      });
      expect(blockTexts(await settled(inside, mode))).toEqual([PARTY]);

      const beside = await openReviewer(buffer.slice(0));
      expect(
        applyOne(beside, mode, { id: "split", type: "splitBlock", blockId, offset: 8 }).skipped,
      ).toEqual([]);
      expect(blockTexts(await settled(beside, mode))).toEqual(["Party 🧑", " agrees."]);
    },
  );

  test("is no range a handle can name", () => {
    expect(
      createFolioAITextRangeHandle({ blockId: "b", text: PARTY, startOffset: 6, endOffset: 7 }),
    ).toBeNull();
    expect(
      createFolioAITextRangeHandle({ blockId: "b", text: PARTY, startOffset: 7, endOffset: 9 }),
    ).toBeNull();
    expect(
      createFolioAITextRangeHandle({ blockId: "b", text: PARTY, startOffset: 6, endOffset: 8 }),
    ).not.toBeNull();
  });

  for (const mode of MODES) {
    test.each([
      ["formatRange", { formatting: { bold: true } }],
      ["commentOnRange", { comment: { text: "Check the term." } }],
      ["replaceRange", { replace: "X" }],
    ] as const)(
      `refuses %s over half of it and keeps the emoji (${mode}, saved and reopened)`,
      async (type, fields) => {
        const reviewer = await openReviewer(await paragraphsDocx([[textRun(PARTY)]]));
        const blockId = reviewer.snapshot().blocks[0]?.id ?? "";
        const operation = {
          id: type,
          type,
          range: handle(blockId, PARTY, 6, 7),
          ...fields,
        } as FolioDocumentOperation;
        const result = applyOne(reviewer, mode, operation);
        expect(result.skipped.map(({ reason }) => reason)).toEqual(["splitsCharacter"]);
        expect(blockTexts(await settled(reviewer, mode))).toEqual([PARTY]);
      },
    );
  }
});

describe("an offset inside a grapheme cluster", () => {
  const CAFE = "Café au lait"; // é written as e + U+0301

  test.each(MODES)("refuses a replacement of its base letter alone (%s)", async (mode) => {
    const text = CAFE.replace("é", "é");
    const reviewer = await openReviewer(await paragraphsDocx([[textRun(text)]]));
    const blockId = reviewer.snapshot().blocks[0]?.id ?? "";
    const result = applyOne(reviewer, mode, {
      id: "replace",
      type: "replaceInBlock",
      blockId,
      find: "Cafe",
      replace: "Coffee",
    });
    expect(result.skipped.map(({ reason }) => reason)).toEqual(["splitsCharacter"]);
    expect(blockTexts(await settled(reviewer, mode))).toEqual([text]);
  });

  test("still lets a comment start at a combining mark the document formats apart", async () => {
    const text = "Café";
    const reviewer = await openReviewer(await paragraphsDocx([[textRun(text)]]));
    const blockId = reviewer.snapshot().blocks[0]?.id ?? "";
    const result = applyOne(reviewer, "direct", {
      id: "comment",
      type: "commentOnRange",
      range: handle(blockId, text, 4, 5),
      comment: { text: "Accent." },
    });
    expect(result.skipped).toEqual([]);
    expect(blockTexts(await reopened(reviewer))).toEqual([text]);
  });
});

// ---------------------------------------------------------------------------
// The class: every offset-taking operation at every offset of text made of
// astral characters, combining sequences, joined emoji and flags.
// ---------------------------------------------------------------------------

const FRAGMENTS = ["ab", " ", "🧑", "é", "👨‍👩‍👧", "🇨🇿", "x", "𝔸"] as const;

const textArbitrary = fc
  .array(fc.constantFrom(...FRAGMENTS), { minLength: 2, maxLength: 8 })
  .map((parts) => parts.join(""));

type OffsetOperation =
  | "splitBlock"
  | "replaceRange"
  | "replaceInBlock"
  | "formatRange"
  | "commentOnRange"
  | "commentOnBlock";

/** Annotations only add marks; everything else writes text or breaks the paragraph. */
const ANNOTATING: ReadonlySet<OffsetOperation> = new Set([
  "formatRange",
  "commentOnRange",
  "commentOnBlock",
]);

const hasLoneSurrogate = (value: string): boolean => {
  for (let index = 0; index < value.length; index++) {
    const code = value.charCodeAt(index);
    const high = code >= 0xd8_00 && code <= 0xdb_ff;
    const low = code >= 0xdc_00 && code <= 0xdf_ff;
    if (low) {
      return true;
    }
    if (high) {
      const next = value.charCodeAt(index + 1);
      if (!(next >= 0xdc_00 && next <= 0xdf_ff)) {
        return true;
      }
      index++;
    }
  }
  return false;
};

describe("an offset-taking operation at any offset", () => {
  test("changes no text beyond what it names, and cuts no character", async () => {
    await fc.assert(
      fc.asyncProperty(
        textArbitrary,
        fc.constantFrom<OffsetOperation>(
          "splitBlock",
          "replaceRange",
          "replaceInBlock",
          "formatRange",
          "commentOnRange",
          "commentOnBlock",
        ),
        fc.nat(),
        fc.nat(),
        fc.constantFrom(...MODES),
        async (text, type, first, second, mode) => {
          const start = 1 + (first % (text.length - 1));
          const end = Math.min(text.length, start + 1 + (second % 4));
          const reviewer = await openReviewer(
            await paragraphsDocx([[textRun(text)], [textRun("Untouched.")]]),
          );
          const blockId = reviewer.snapshot().blocks[0]?.id ?? "";
          const selected = text.slice(start, end);
          const unique = text.indexOf(selected) === text.lastIndexOf(selected);
          let operation: FolioDocumentOperation;
          let expected: string[];
          let offsets: number[] = [start, end];
          switch (type) {
            case "splitBlock":
              operation = { id: "op", type, blockId, offset: start };
              expected = [text.slice(0, start), text.slice(start), "Untouched."];
              offsets = [start];
              break;
            case "replaceRange":
              operation = {
                id: "op",
                type,
                range: handle(blockId, text, start, end),
                replace: "Q",
              };
              expected = [`${text.slice(0, start)}Q${text.slice(end)}`, "Untouched."];
              break;
            case "replaceInBlock":
              if (!unique || hasLoneSurrogate(selected)) {
                return;
              }
              operation = { id: "op", type, blockId, find: selected, replace: "Q" };
              expected = [`${text.slice(0, start)}Q${text.slice(end)}`, "Untouched."];
              break;
            case "formatRange":
              operation = {
                id: "op",
                type,
                range: handle(blockId, text, start, end),
                formatting: { bold: true },
              };
              expected = [text, "Untouched."];
              break;
            case "commentOnRange":
              operation = {
                id: "op",
                type,
                range: handle(blockId, text, start, end),
                comment: { text: "Note." },
              };
              expected = [text, "Untouched."];
              break;
            case "commentOnBlock":
              if (!unique || hasLoneSurrogate(selected)) {
                return;
              }
              operation = { id: "op", type, blockId, quote: selected, comment: { text: "Note." } };
              expected = [text, "Untouched."];
              break;
          }
          const allowed = offsets.every((offset) =>
            ANNOTATING.has(type)
              ? !splitsSurrogatePair(text, offset)
              : !splitsGraphemeCluster(text, offset),
          );

          const result = applyOne(reviewer, mode, operation);
          // No text node the operation leaves holds half a character.
          reviewer.state.doc.descendants((node) => {
            if (node.isText) {
              expect(hasLoneSurrogate(node.text ?? "")).toBe(false);
            }
            return true;
          });
          const final = blockTexts(await settled(reviewer, mode));
          if (allowed) {
            expect({ text, type, start, end, skipped: result.skipped }).toEqual({
              text,
              type,
              start,
              end,
              skipped: [],
            });
            expect(final).toEqual(expected);
          } else {
            expect({
              text,
              type,
              start,
              end,
              reasons: result.skipped.map((s) => s.reason),
            }).toEqual({ text, type, start, end, reasons: ["splitsCharacter"] });
            expect(final).toEqual([text, "Untouched."]);
          }
        },
      ),
      propertyConfig({ numRuns: 120 }),
    );
  }, 300_000);
});
