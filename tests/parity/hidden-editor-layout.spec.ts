import { buildScrollRootDocument } from "../support/scrollRootDocument";
import {
  ensureLiveView,
  expect,
  forEachAdapter,
  openCanonicalEditor,
} from "./parity-fixture";

const FILLER_PAGES = 80;
/** The blocks the selection reaches, and the block before and after them. */
const MOST_LAID_OUT_BLOCKS = 4;
const EDITOR = ".paged-editor__hidden-pm > .ProseMirror";
const BLOCKS = `${EDITOR} > *`;

declare global {
  // eslint-disable-next-line typescript/consistent-type-definitions
  interface Window {
    /** Whether the browser lays out a hidden editor block. */
    __folioBlockLaidOut?: (block: Element) => boolean;
  }
}

forEachAdapter(
  "the hidden editor lays out only the blocks near the selection",
  async (adapter, { page }) => {
    // The browser reports each block it may skip as skipped or laid out,
    // first when it decides and then on every change. A block it may not
    // skip, or never reported skipped, is laid out, its text shaped.
    await page.addInitScript(() => {
      const skipped = new WeakMap<EventTarget, boolean>();
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
      window.__folioBlockLaidOut = (block) =>
        getComputedStyle(block).contentVisibility !== "auto" ||
        skipped.get(block) !== true;
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
    const laidOut = () =>
      blocks.evaluateAll(
        (elements) =>
          elements.filter((element) => window.__folioBlockLaidOut?.(element))
            .length,
      );
    await expect.poll(laidOut).toBeLessThanOrEqual(MOST_LAID_OUT_BLOCKS);

    // The block that takes the selection, and its neighbours, are laid out
    // in the same update, so the browser edits and composes in laid-out text.
    const target = Math.floor(FILLER_PAGES / 2);
    await page.locator(EDITOR).first().focus();
    const laidOutAfterSelecting = await blocks
      .nth(target)
      .evaluate((element) => {
        const text = document
          .createTreeWalker(element, NodeFilter.SHOW_TEXT)
          .nextNode();
        if (text === null) throw new Error("the block has no text");
        document.getSelection()?.collapse(text, 1);
        return new Promise<boolean[]>((resolve) => {
          // ProseMirror reads the selection on `selectionchange`.
          document.addEventListener(
            "selectionchange",
            () =>
              queueMicrotask(() =>
                resolve(
                  [
                    element.previousElementSibling,
                    element,
                    element.nextElementSibling,
                  ].map(
                    (block) =>
                      block !== null &&
                      getComputedStyle(block).contentVisibility !== "auto",
                  ),
                ),
              ),
            { once: true },
          );
        });
      });
    expect(laidOutAfterSelecting).toEqual([true, true, true]);
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
