import { afterEach, beforeEach, describe, expect, test } from "bun:test";

import { installFixedWidthProvider } from "../../../test/fixedWidthMeasureProvider";

import { FolioDocxReviewer } from "./ai-edits/headless";
import { createDocx } from "./docx/rezip";
import { layoutDocxHeadless } from "./headless-layout";
import {
  getMeasureProvider,
  resetMeasureProvider,
  setMeasureProvider,
} from "./layout-engine/measure/measureProvider";
import { getPageTextFromLayout } from "./paged-layout/pageText";
import { createFolioReviewerPageSource } from "./reviewer-page-source";
import { createEmptyDocument } from "./utils/createDocument";

let originalProvider = getMeasureProvider();
beforeEach(() => {
  originalProvider = getMeasureProvider();
  installFixedWidthProvider();
});

afterEach(() => setMeasureProvider(originalProvider));

describe("createFolioReviewerPageSource", () => {
  test("omits the capability without a MeasureProvider", async () => {
    const bytes = await createDocx(createEmptyDocument({ initialText: "Unavailable" }));
    const reviewer = await FolioDocxReviewer.fromBuffer(bytes);
    resetMeasureProvider();
    expect(createFolioReviewerPageSource(reviewer)).toBeNull();
  });

  test("matches headless layout and invalidates after an edit", async () => {
    const bytes = await createDocx(createEmptyDocument({ initialText: "Alpha beta" }));
    const expected = await layoutDocxHeadless(bytes);
    expect(expected.isOk()).toBe(true);
    if (expected.isErr()) throw expected.error;

    const reviewer = await FolioDocxReviewer.fromBuffer(bytes);
    const source = createFolioReviewerPageSource(reviewer);
    expect(source).not.toBeNull();
    if (!source) throw new Error("Test measurement provider was not installed");
    expect(source.pageText(1)).toBe(
      getPageTextFromLayout({
        layout: expected.value.layout,
        doc: expected.value.proseDoc,
        page: 1,
      }),
    );

    const block = reviewer.snapshot().blocks.at(0);
    expect(block).toBeDefined();
    if (!block) throw new Error("Fixture has no blocks");
    reviewer.applyOperations(
      [{ id: "append", type: "insertAfterBlock", blockId: block.id, text: "Gamma" }],
      { mode: "direct", undefinedReferences: "refuse" },
    );
    expect(source.pageText(1)).toContain("Gamma");
  });
});
