import { ensureLiveView, expect, forEachAdapter, openEditor } from "./parity-fixture";

forEachAdapter("CJK IME input caret follows the painted caret", async (adapter, { page }) => {
  await openEditor(page, adapter);
  await ensureLiveView(page);

  const readAlignment = () =>
    page.evaluate(() => {
      const paintedCaret = document.querySelector<HTMLElement>("[data-folio-caret-rect]");
      const selection = document.getSelection();
      if (!paintedCaret || !selection || selection.rangeCount === 0) {
        return { aligned: false, reason: "missing selection or painted caret" };
      }

      const range = selection.getRangeAt(0);
      const hiddenHost = Array.from(
        document.querySelectorAll<HTMLElement>(".paged-editor__hidden-pm"),
      ).find((candidate) => candidate.contains(range.startContainer));
      if (!hiddenHost) {
        return { aligned: false, reason: "selection is outside the hidden editor" };
      }

      const nativeCaret = range.getBoundingClientRect();
      const visibleCaret = paintedCaret.getBoundingClientRect();
      const deltaX = Math.abs(nativeCaret.left - visibleCaret.left);
      const deltaY = Math.abs(nativeCaret.top - visibleCaret.top);
      return {
        aligned:
          deltaX <= 1 && deltaY <= 1 && hiddenHost.style.transform.startsWith("translate3d("),
        deltaX,
        deltaY,
        transform: hiddenHost.style.transform,
      };
    });
  const expectAligned = () => expect.poll(readAlignment).toMatchObject({ aligned: true });

  await page.locator(".paged-editor__hidden-pm > .ProseMirror").first().focus();
  await expectAligned();

  const firstRun = page.locator(".layout-page-content span[data-pm-start][data-pm-end]").first();
  await firstRun.click({ position: { x: 80, y: 8 } });
  await expectAligned();

  await firstRun.click({ position: { x: 120, y: 8 } });
  await expectAligned();

  await page.keyboard.press("End");
  await page.keyboard.press("Enter");
  await expectAligned();

  const scrollContainer = page.locator(
    adapter.name === "react" ? "[data-folio-scroll]" : ".docx-editor-vue__editor-scroll",
  );
  await scrollContainer.evaluate((element) => {
    element.scrollTop += 120;
  });
  await expectAligned();
});

forEachAdapter(
  "CJK IME input caret follows the painted header caret",
  async (adapter, { page }) => {
    await openEditor(page, adapter);

    const header = page.locator(".layout-page-header").first();
    await header.dblclick();
    await expect(page.locator(".hf-inline-editor")).toBeVisible();
    await page
      .locator(".layout-page-header span[data-pm-start][data-pm-end]")
      .first()
      .click({ position: { x: 20, y: 8 } });

    const readAlignment = () =>
      page.evaluate(() => {
        const paintedCaret = document.querySelector<HTMLElement>("[data-testid='hf-caret']");
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
          return { aligned: false, reason: "selection is outside the hidden header editor" };
        }

        const nativeCaret = range.getBoundingClientRect();
        const visibleCaret = paintedCaret.getBoundingClientRect();
        const deltaX = Math.abs(nativeCaret.left - visibleCaret.left);
        const deltaY = Math.abs(nativeCaret.top - visibleCaret.top);
        return {
          aligned:
            deltaX <= 1 && deltaY <= 1 && hiddenHost.style.transform.startsWith("translate3d("),
          deltaX,
          deltaY,
          transform: hiddenHost.style.transform,
        };
      });
    const expectAligned = () => expect.poll(readAlignment).toMatchObject({ aligned: true });

    await expectAligned();

    const scrollContainer = page.locator(
      adapter.name === "react" ? "[data-folio-scroll]" : ".docx-editor-vue__editor-scroll",
    );
    await scrollContainer.evaluate((element) => {
      element.scrollTop += 120;
    });
    await expectAligned();
  },
);
