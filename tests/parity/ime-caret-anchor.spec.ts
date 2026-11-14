import type { Locator, Page } from "@playwright/test";

import {
  ensureLiveView,
  expect,
  forEachAdapter,
  openCanonicalEditor,
  openEditor,
} from "./parity-fixture";

type CompositionAnchorOptions = {
  expectAligned: () => Promise<void>;
  hiddenHost: Locator;
  input: Locator;
  page: Page;
};

const COMPOSITION_ANCHOR_SENTINEL = "translate3d(1px, 2px, 0px)";

const verifyCompositionAnchorLifecycle = async ({
  expectAligned,
  hiddenHost,
  input,
  page,
}: CompositionAnchorOptions): Promise<void> => {
  await input.dispatchEvent("compositionstart");
  await hiddenHost.evaluate((host, transform) => {
    host.style.transform = transform;
  }, COMPOSITION_ANCHOR_SENTINEL);
  await page.evaluate(() => window.dispatchEvent(new Event("scroll")));
  await expect
    .poll(() => hiddenHost.evaluate((host) => host.style.transform))
    .toBe(COMPOSITION_ANCHOR_SENTINEL);

  await input.dispatchEvent("compositionend");
  await page.waitForTimeout(50);
  await page.evaluate(() => window.dispatchEvent(new Event("scroll")));
  await expectAligned();
};

forEachAdapter(
  "CJK IME input caret follows the painted caret",
  async (adapter, { page }) => {
    await openEditor(page, adapter);
    await ensureLiveView(page);

    const readAlignment = () =>
      page.evaluate(() => {
        const paintedCaret = document.querySelector<HTMLElement>(
          "[data-folio-caret-rect]",
        );
        const selection = document.getSelection();
        if (!paintedCaret || !selection || selection.rangeCount === 0) {
          return {
            aligned: false,
            reason: "missing selection or painted caret",
          };
        }

        const range = selection.getRangeAt(0);
        const hiddenHost = Array.from(
          document.querySelectorAll<HTMLElement>(".paged-editor__hidden-pm"),
        ).find((candidate) => candidate.contains(range.startContainer));
        if (!hiddenHost) {
          return {
            aligned: false,
            reason: "selection is outside the hidden editor",
          };
        }

        // A caret with no box of its own (in an empty paragraph, beside its
        // line break, in an empty inline wrapper) stands at the left edge of
        // the node after it, or of the element holding it, beside that box:
        // only its left edge is exact.
        const caretBox = range.getClientRects().length > 0;
        const holder = range.startContainer;
        const after = holder.childNodes[range.startOffset];
        let nativeCaret = range.getBoundingClientRect();
        if (!caretBox && after instanceof Element)
          nativeCaret = after.getBoundingClientRect();
        else if (!caretBox && holder instanceof Element)
          nativeCaret = holder.getBoundingClientRect();
        const visibleCaret = paintedCaret.getBoundingClientRect();
        const deltaX = Math.abs(nativeCaret.left - visibleCaret.left);
        const deltaY = Math.abs(nativeCaret.top - visibleCaret.top);
        const withinBox =
          visibleCaret.top <= nativeCaret.bottom + 1 &&
          visibleCaret.bottom >= nativeCaret.top - 1;
        return {
          aligned:
            deltaX <= 1 &&
            (caretBox ? deltaY <= 1 : withinBox) &&
            hiddenHost.style.transform.startsWith("translate3d("),
          deltaX,
          deltaY,
          transform: hiddenHost.style.transform,
        };
      });
    const expectAligned = () =>
      expect.poll(readAlignment).toMatchObject({ aligned: true });

    await page
      .locator(".paged-editor__hidden-pm > .ProseMirror")
      .first()
      .focus();
    await expectAligned();

    const firstRun = page
      .locator(".layout-page-content span[data-pm-start][data-pm-end]")
      .first();
    await firstRun.click({ position: { x: 80, y: 8 } });
    await expectAligned();

    await firstRun.click({ position: { x: 120, y: 8 } });
    await expectAligned();

    await page.keyboard.press("End");
    await page.keyboard.press("Enter");
    await expectAligned();

    await verifyCompositionAnchorLifecycle({
      expectAligned,
      hiddenHost: page.locator(".paged-editor__hidden-pm").first(),
      input: page.locator(".paged-editor__hidden-pm > .ProseMirror").first(),
      page,
    });

    const scrollContainer = page.locator(
      adapter.name === "react"
        ? "[data-folio-scroll]"
        : ".docx-editor-vue__editor-scroll",
    );
    await scrollContainer.evaluate((element) => {
      element.scrollTop += 120;
    });
    await expectAligned();
  },
);

forEachAdapter(
  "CJK IME input caret follows the painted header caret",
  async (adapter, { page }) => {
    await openEditor(page, adapter);
    await ensureLiveView(page);

    const header = page.locator(".layout-page-header").first();
    await header.dblclick();
    await expect(page.locator(".hf-inline-editor")).toBeVisible();
    await page
      .locator(".layout-page-header span[data-pm-start][data-pm-end]")
      .first()
      .click({ position: { x: 20, y: 8 } });

    const readAlignment = () =>
      page.evaluate(() => {
        const paintedCaret = document.querySelector<HTMLElement>(
          "[data-testid='hf-caret']",
        );
        const selection = document.getSelection();
        if (!paintedCaret || !selection || selection.rangeCount === 0) {
          return {
            aligned: false,
            reason: "missing selection or painted caret",
            hasPaintedCaret: Boolean(paintedCaret),
            hasSelection: Boolean(selection?.rangeCount),
          };
        }

        const range = selection.getRangeAt(0);
        const hiddenHost = Array.from(
          document.querySelectorAll<HTMLElement>(".paged-editor__hidden-hf-pm"),
        ).find((candidate) => candidate.contains(range.startContainer));
        if (!hiddenHost) {
          return {
            aligned: false,
            reason: "selection is outside the hidden header editor",
          };
        }

        // A caret with no box of its own (in an empty paragraph, beside its
        // line break, in an empty inline wrapper) stands at the left edge of
        // the node after it, or of the element holding it, beside that box:
        // only its left edge is exact.
        const caretBox = range.getClientRects().length > 0;
        const holder = range.startContainer;
        const after = holder.childNodes[range.startOffset];
        let nativeCaret = range.getBoundingClientRect();
        if (!caretBox && after instanceof Element)
          nativeCaret = after.getBoundingClientRect();
        else if (!caretBox && holder instanceof Element)
          nativeCaret = holder.getBoundingClientRect();
        const visibleCaret = paintedCaret.getBoundingClientRect();
        const deltaX = Math.abs(nativeCaret.left - visibleCaret.left);
        const deltaY = Math.abs(nativeCaret.top - visibleCaret.top);
        const withinBox =
          visibleCaret.top <= nativeCaret.bottom + 1 &&
          visibleCaret.bottom >= nativeCaret.top - 1;
        return {
          aligned:
            deltaX <= 1 &&
            (caretBox ? deltaY <= 1 : withinBox) &&
            hiddenHost.style.transform.startsWith("translate3d("),
          deltaX,
          deltaY,
          transform: hiddenHost.style.transform,
        };
      });
    const expectAligned = () =>
      expect.poll(readAlignment).toMatchObject({ aligned: true });

    await expectAligned();

    await verifyCompositionAnchorLifecycle({
      expectAligned,
      hiddenHost: page.locator(".paged-editor__hidden-hf-pm"),
      input: page.locator(".paged-editor__hidden-hf-pm .ProseMirror").first(),
      page,
    });

    const scrollContainer = page.locator(
      adapter.name === "react"
        ? "[data-folio-scroll]"
        : ".docx-editor-vue__editor-scroll",
    );
    await scrollContainer.evaluate((element) => {
      element.scrollTop += 120;
    });
    await expectAligned();
  },
);

forEachAdapter(
  "CJK IME input caret stays beside a composition as it grows",
  async (adapter, { page }) => {
    // The run's East Asian slot names a face the document embeds, so the kana
    // draw with one face whatever fonts the machine has.
    await openCanonicalEditor(page, adapter, "cjk-composition.docx");
    await ensureLiveView(page);
    await page
      .locator(".paged-editor__hidden-pm > .ProseMirror")
      .first()
      .focus();
    await page
      .locator(".layout-page-content span[data-pm-start][data-pm-end]")
      .first()
      .click({ position: { x: 80, y: 8 } });

    // While composing, the anchor stays where the composition began and the
    // hidden editor draws the provisional text with the faces the painter
    // uses, so the native caret, where the candidate window opens, settles
    // beside the painted one, within the rounding of the two measurements.
    const readDistance = () =>
      page.evaluate(() => {
        const selection = document.getSelection();
        const painted = document
          .querySelector<HTMLElement>("[data-folio-caret-rect]")
          ?.getBoundingClientRect();
        if (!selection || selection.rangeCount === 0 || !painted) return null;
        const range = selection.getRangeAt(0);
        if (range.getClientRects().length === 0) return null;
        const native = range.getBoundingClientRect();
        return {
          dx: Math.abs(native.left - painted.left),
          dy: Math.abs(native.top - painted.top),
        };
      });

    // Compose once the anchor has followed the click.
    await expect
      .poll(async () => {
        const distance = await readDistance();
        return distance !== null && distance.dx <= 1 && distance.dy <= 1;
      })
      .toBe(true);

    const cdp = await page.context().newCDPSession(page);
    try {
      const text = "あいうえおかきくけこさしすせそ";
      for (let length = 1; length <= text.length; length += 2) {
        await cdp.send("Input.imeSetComposition", {
          text: text.slice(0, length),
          selectionStart: length,
          selectionEnd: length,
        });
        // The painted caret follows the provisional text after a layout
        // pass; the native one must end up beside it.
        await expect
          .poll(async () => {
            const distance = await readDistance();
            return distance !== null && distance.dx < 12 && distance.dy < 6;
          }, adapter.name)
          .toBe(true);
      }
      await cdp.send("Input.insertText", { text });
    } finally {
      if (!page.isClosed()) await cdp.detach();
    }
  },
);
