import { afterEach, beforeEach, describe, expect, spyOn, test } from "bun:test";

import {
  createFolioAITextRangeHandle,
  FolioDocxReviewer,
  isFolioAIContentBlock,
} from "@stll/folio-core/server";
import { createDocx } from "@stll/folio-core/docx/rezip";
import * as headlessLayout from "@stll/folio-core/headless-layout";
import { resolveFolioAITextRange } from "@stll/folio-core/ai-edits/blockRange";
import { findPageIndexContainingPmPos } from "@stll/folio-core/layout-engine/pmPageIndex";
import {
  getMeasureProvider,
  setMeasureProvider,
} from "@stll/folio-core/layout-engine/measure/measureProvider";
import { getPageTextFromLayout } from "@stll/folio-core/paged-layout/pageText";
import { createFolioReviewerPageSource } from "@stll/folio-core/reviewer-page-source";
import { createEmptyDocument } from "@stll/folio-core/utils/createDocument";

import { installFixedWidthProvider } from "../../../../test/fixedWidthMeasureProvider";
import { executeFolioToolCall, executeFolioToolCallUntyped } from "../execute";
import { FOLIO_AGENT_TOOL_NAMES } from "../types";
import { createReviewerBridge } from "./reviewer";

let originalProvider = getMeasureProvider();
beforeEach(() => {
  originalProvider = getMeasureProvider();
  installFixedWidthProvider();
});
afterEach(() => setMeasureProvider(originalProvider));

const bridgeFor = (reviewer: FolioDocxReviewer) => {
  const pageSource = createFolioReviewerPageSource(reviewer);
  if (!pageSource) throw new Error("Test measurement provider was not installed");
  return createReviewerBridge(reviewer, { pageSource, mode: "direct" });
};

const referenceLayout = async (bytes: ArrayBuffer) => {
  const result = await headlessLayout.layoutDocxHeadless(bytes);
  expect(result.isOk()).toBe(true);
  if (result.isErr()) throw result.error;
  return result.value;
};

const readAllPages = (bridge: ReturnType<typeof bridgeFor>) => {
  const first = executeFolioToolCall(FOLIO_AGENT_TOOL_NAMES.readPage, { page: 1 }, bridge);
  if (!first.ok) throw new Error(first.error);
  const totalPages = first.result.totalPages;
  if (totalPages === undefined) throw new Error("Reviewer bridge omitted totalPages");
  return Array.from({ length: totalPages }, (_, index) => {
    const page = index + 1;
    const result = executeFolioToolCall(FOLIO_AGENT_TOOL_NAMES.readPage, { page }, bridge);
    expect(result.ok).toBe(true);
    if (!result.ok) throw new Error(result.error);
    // Exercise the untyped dispatch used by tool transports as well.
    expect(executeFolioToolCallUntyped("read_page", { page }, bridge)).toEqual(result);
    expect(result.result).toMatchObject({ page, totalPages });
    return result.result;
  });
};

describe("reviewer page tool execution", () => {
  for (const fixture of ["sample.docx", "docx-editor-demo.docx"]) {
    test(`${fixture}: reads and searches the independently laid out pages`, async () => {
      const bytes = await Bun.file(
        new URL(`../../../../tests/visual/fixtures/${fixture}`, import.meta.url),
      ).arrayBuffer();
      const reference = await referenceLayout(bytes);
      const reviewer = await FolioDocxReviewer.fromBuffer(bytes);
      const bridge = bridgeFor(reviewer);
      const pages = readAllPages(bridge);
      expect(pages).toEqual(
        reference.layout.pages.map((_, index) => ({
          page: index + 1,
          totalPages: reference.layout.pages.length,
          text: getPageTextFromLayout({
            layout: reference.layout,
            doc: reference.proseDoc,
            page: index + 1,
          }),
        })),
      );

      // Build the search oracle from the source snapshot and the separate
      // layout, never from the bridge's own page assignments or search result.
      const snapshot = reviewer.snapshot();
      const query = "the";
      const expectedMatches = snapshot.blocks.filter(isFolioAIContentBlock).flatMap((block) =>
        [...block.text.matchAll(/the/giu)].map((match) => {
          const range = createFolioAITextRangeHandle({
            blockId: block.id,
            text: block.text,
            startOffset: match.index,
            endOffset: match.index + query.length,
          });
          if (!range) throw new Error("Fixture match has no text range");
          const resolved = resolveFolioAITextRange({ range, doc: reference.proseDoc, snapshot });
          if (!resolved) throw new Error("Fixture range did not resolve in reference layout");
          const index = findPageIndexContainingPmPos(reference.layout, resolved.from);
          if (index === null) throw new Error("Fixture match was not paginated");
          return { range, page: index + 1 };
        }),
      );
      expect(expectedMatches.length).toBeGreaterThan(0);
      if (fixture === "docx-editor-demo.docx") {
        expect(new Set(expectedMatches.map(({ page }) => page)).size).toBeGreaterThan(1);
      }
      for (const { page } of pages) {
        const args = { query, scope: { type: "page", page } } as const;
        const found = executeFolioToolCall(FOLIO_AGENT_TOOL_NAMES.findText, args, bridge);
        expect(executeFolioToolCallUntyped("find_text", args, bridge)).toEqual(found);
        expect(found.ok).toBe(true);
        if (!found.ok) throw new Error(found.error);
        const expected = expectedMatches.filter((match) => match.page === page);
        expect(found.result.totalMatches).toBe(expected.length);
        expect(found.result.truncated).toBe(false);
        expect(
          found.result.matches.map((match) => {
            if (match.type === "story") throw new Error("Page search returned a secondary story");
            return { range: match.range, page: match.page };
          }),
        ).toEqual(expected);
      }
    });
  }

  test("reuses one layout per state, moves targets after insertion, and survives saving", async () => {
    const bytes = await createDocx(createEmptyDocument({ initialText: "Tail needle" }));
    const reviewer = await FolioDocxReviewer.fromBuffer(bytes);
    const tail = reviewer.snapshot().blocks.find(isFolioAIContentBlock);
    if (!tail) throw new Error("Fixture has no content block");
    const range = createFolioAITextRangeHandle({
      blockId: tail.id,
      text: tail.text,
      startOffset: 5,
      endOffset: 11,
    });
    if (!range) throw new Error("Fixture has no needle range");
    const blockTarget = { type: "block", story: "main", blockId: tail.id } as const;
    const layout = spyOn(headlessLayout, "layoutCurrentProseDocHeadless");
    try {
      const bridge = bridgeFor(reviewer);
      expect(layout).toHaveBeenCalledTimes(0);
      const before = readAllPages(bridge);
      expect(before).toHaveLength(1);
      expect(bridge.getTargetPage?.(blockTarget)).toBe(1);
      expect(bridge.getTargetPage?.(range)).toBe(1);
      expect(
        bridge.getTargetPage?.({ type: "block", story: "main", blockId: "FFFFFFFF" }),
      ).toBeNull();
      expect(readAllPages(bridge)).toEqual(before);
      expect(layout).toHaveBeenCalledTimes(1);

      // This sequence is absent from the original single-page example: insert
      // enough paragraphs before a stable target to change both its page and
      // the page count, then interleave text, count, search and target reads.
      const edit = reviewer.applyOperations(
        [
          {
            id: "push-tail",
            type: "insertBeforeBlock",
            blockId: tail.id,
            text: Array.from({ length: 120 }, (_, index) => `Inserted paragraph ${index}.`).join(
              "\n",
            ),
          },
        ],
        { mode: "direct", undefinedReferences: "refuse" },
      );
      expect(edit.skipped).toHaveLength(0);
      const after = readAllPages(bridge);
      expect(after.length).toBeGreaterThan(before.length);
      expect(after.at(0)?.text).not.toBe(before.at(0)?.text);
      expect(after.at(0)?.text).not.toContain("Tail needle");
      const tailPage = bridge.getTargetPage?.(blockTarget);
      expect(tailPage).toBeGreaterThan(1);
      expect(bridge.getTargetPage?.(range)).toBe(tailPage);
      const tailAnswer = after.find(({ page }) => page === tailPage);
      expect(tailAnswer?.text).toContain("Tail needle");
      expect(readAllPages(bridge)).toEqual(after);
      expect(layout).toHaveBeenCalledTimes(2);

      const saved = await reviewer.save();
      if (saved.type === "repackRefused") throw new Error(saved.reason);
      const reference = await referenceLayout(saved.buffer);
      expect(after.map(({ text }) => text)).toEqual(
        reference.layout.pages.map((_, index) =>
          getPageTextFromLayout({
            layout: reference.layout,
            doc: reference.proseDoc,
            page: index + 1,
          }),
        ),
      );
      const reopened = await FolioDocxReviewer.fromBuffer(saved.buffer);
      const reopenedBridge = bridgeFor(reopened);
      expect(readAllPages(reopenedBridge)).toEqual(after);
      expect(reopenedBridge.getTargetPage?.(blockTarget)).toBe(tailPage);
      expect(reopenedBridge.getTargetPage?.(range)).toBe(tailPage);
      for (const { page } of after) {
        const args = { query: "needle", scope: { type: "page", page } } as const;
        const answer = executeFolioToolCall(FOLIO_AGENT_TOOL_NAMES.findText, args, bridge);
        expect(executeFolioToolCall(FOLIO_AGENT_TOOL_NAMES.findText, args, reopenedBridge)).toEqual(
          answer,
        );
        expect(answer.ok).toBe(true);
        if (!answer.ok) throw new Error(answer.error);
        expect(answer.result.totalMatches).toBe(page === tailPage ? 1 : 0);
        expect(
          answer.result.matches.every((match) => match.type !== "story" && match.page === page),
        ).toBe(true);
      }
      // The original reviewer has not changed again, even after serializing it.
      expect(readAllPages(bridge)).toEqual(after);
      // The bytes oracle uses its own pipeline; the reopened reviewer adds
      // exactly one call to the synchronous reviewer layout entry.
      expect(layout).toHaveBeenCalledTimes(3);
    } finally {
      layout.mockRestore();
    }
  });
});
