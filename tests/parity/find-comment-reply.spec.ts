import {
  buildCommentSearchDocument,
  COMMENT_SEARCH_REPLY_LENGTHS,
  REPLY_SEARCH_TARGET,
} from "../support/commentSearchDocument";
import { expect, forEachAdapter, openEditor } from "./parity-fixture";

for (const replyLength of COMMENT_SEARCH_REPLY_LENGTHS) {
  forEachAdapter(
    `find reveals and re-reveals a ${replyLength}-sentence reply beyond the collapsed text clamp`,
    async (adapter, { page }) => {
      const bytes = await buildCommentSearchDocument(replyLength);
      await page.route("**/fixtures/comment-search.docx", (route) =>
        route.fulfill({
          contentType:
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
          body: Buffer.from(bytes),
        }),
      );
      await openEditor(page, adapter, "comment-search.docx");
      const reply = page.locator('[data-find-comment-id="3"]');
      await expect(reply).toBeAttached();
      await expect(page.locator('[data-find-comment-id="2"]')).toHaveCount(0);

      const readTarget = () =>
        reply.evaluate((span, target) => {
          const body = span.parentElement;
          const text = span.firstChild;
          if (!body || !text) throw new TypeError("Reply text is unavailable");
          const offset = text.textContent?.indexOf(target) ?? -1;
          if (offset < 0) throw new TypeError("Reply search target is missing");
          const range = document.createRange();
          range.setStart(text, offset);
          range.setEnd(text, offset + target.length);
          const targetBox = range.getBoundingClientRect();
          const bodyBox = body.getBoundingClientRect();
          const root = span.closest("[data-folio-scroll]");
          if (!root) throw new TypeError("Reply scroll root is unavailable");
          const rootBox = root.getBoundingClientRect();
          return {
            clamp: getComputedStyle(body).webkitLineClamp,
            insideBody:
              targetBox.top >= bodyBox.top &&
              targetBox.bottom <= bodyBox.bottom,
            insideViewport:
              targetBox.top >= Math.max(0, rootBox.top) &&
              targetBox.bottom <= Math.min(innerHeight, rootBox.bottom),
          };
        }, REPLY_SEARCH_TARGET);
      expect(await readTarget()).toMatchObject({
        clamp: "2",
        insideBody: false,
      });

      await page.locator(".layout-page").first().click();
      await page.keyboard.press("ControlOrMeta+f");
      const dialog = page.getByTestId("find-replace-dialog");
      await expect(dialog).toBeVisible();
      await dialog.locator("input").first().fill(REPLY_SEARCH_TARGET);
      await expect(dialog).toContainText("1 / 1");
      await expect
        .poll(readTarget)
        .toEqual({ clamp: "none", insideBody: true, insideViewport: true });
      await expect(page.locator('[data-find-comment-id="2"]')).toBeVisible();
      await expect(
        page
          .locator('[data-find-match-index="0"][data-current="true"]')
          .first(),
      ).toBeVisible();
      if (adapter.name === "vue") {
        // The column begins above the scrolled viewport. Bubble from its real
        // empty background so the mounted sidebar owns the collapse transition.
        await page
          .locator(".unified-sidebar__inner")
          .dispatchEvent("mousedown");
      } else {
        await page
          .locator("[data-folio-scroll]")
          .first()
          .click({ position: { x: 2, y: 2 } });
      }
      await expect
        .poll(readTarget)
        .toMatchObject({ clamp: "2", insideBody: false });
      await dialog.locator("input").first().press("Enter");
      await expect
        .poll(readTarget)
        .toEqual({ clamp: "none", insideBody: true, insideViewport: true });
      await expect(page.locator('[data-find-comment-id="2"]')).toBeVisible();
    },
  );
}
