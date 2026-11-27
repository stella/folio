import { buildScrollRootDocument } from "../support/scrollRootDocument";
import {
  ensureLiveView,
  expect,
  forEachAdapter,
  openCanonicalEditor,
} from "./parity-fixture";

const FILLER_PAGES = 80;
/** The selection's block, and at most a block or two the browser keeps near it. */
const MOST_LAID_OUT_BLOCKS = 4;
const EDITOR = ".paged-editor__hidden-pm > .ProseMirror";
const BLOCKS = `${EDITOR} > *`;

declare global {
  // eslint-disable-next-line typescript/consistent-type-definitions
  interface Window {
    __folioSkippedBlocks?: WeakMap<EventTarget, boolean>;
  }
}

forEachAdapter(
  "the hidden editor lays out only the blocks near the selection",
  async (adapter, { page }) => {
    // The browser reports each block it may skip as skipped or laid out,
    // first when it decides and then on every change.
    await page.addInitScript(() => {
      const skipped = new WeakMap<EventTarget, boolean>();
      window.__folioSkippedBlocks = skipped;
      document.addEventListener(
        "contentvisibilityautostatechange",
        (event) => {
          if (event.target !== null)
            skipped.set(
              event.target,
              (event as ContentVisibilityAutoStateChangeEvent).skipped,
            );
        },
        true,
      );
    });
    const bytes = await buildScrollRootDocument({ fillerPages: FILLER_PAGES });
    await page.route("**/fixtures/scroll-root.docx", (route) =>
      route.fulfill({
        body: Buffer.from(bytes),
        contentType:
          "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      }),
    );
    await openCanonicalEditor(page, adapter, "scroll-root.docx");
    await ensureLiveView(page);

    const blocks = page.locator(BLOCKS);
    await expect.poll(() => blocks.count()).toBeGreaterThan(FILLER_PAGES);
    // A block the browser never reported skipped is laid out, its text shaped.
    const laidOut = () =>
      blocks.evaluateAll(
        (elements) =>
          elements.filter(
            (element) => window.__folioSkippedBlocks?.get(element) !== true,
          ).length,
      );
    await expect.poll(laidOut).toBeLessThanOrEqual(MOST_LAID_OUT_BLOCKS);

    // The block that takes the selection is laid out, so the native caret
    // the input method reads has a box.
    const target = Math.floor(FILLER_PAGES / 2);
    await page.locator(EDITOR).first().focus();
    await blocks.nth(target).evaluate((element) => {
      const text = document
        .createTreeWalker(element, NodeFilter.SHOW_TEXT)
        .nextNode();
      if (text === null) throw new Error("the block has no text");
      document.getSelection()?.collapse(text, 1);
    });
    await expect
      .poll(() =>
        blocks
          .nth(target)
          .evaluate(
            (element) => window.__folioSkippedBlocks?.get(element) === false,
          ),
      )
      .toBe(true);
    await expect
      .poll(() =>
        page.evaluate(() => {
          const selection = document.getSelection();
          if (selection === null || selection.rangeCount === 0) return 0;
          return selection.getRangeAt(0).getClientRects().length;
        }),
      )
      .toBeGreaterThan(0);
    await expect.poll(laidOut).toBeLessThanOrEqual(MOST_LAID_OUT_BLOCKS);
  },
);
