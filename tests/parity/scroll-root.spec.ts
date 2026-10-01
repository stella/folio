import { ensureLiveView, expect, forEachAdapter, openEditor } from "./parity-fixture";

// The former React-only assertion used toBeVisible(), which accepts elements
// outside an overflow viewport. Assert actual viewport overlap and scrollTop:
// scrolling the non-overflowing pages viewport cannot satisfy this oracle.
for (const handle of ["document", "paged"] as const) {
  forEachAdapter(
    `scrollToPage(3) moves the real scroll root through the ${handle} ref`,
    async (adapter, { page }) => {
      await page.setViewportSize({ width: 1280, height: 600 });
      await openEditor(page, adapter);
      await ensureLiveView(page);
      expect(await page.evaluate(() => window.__folioParity?.getTotalPages())).toBeGreaterThan(2);
      const before = await page.evaluate(() => window.__folioParity?.readScrollViewport(3));
      expect(before).not.toBeNull();
      expect(before?.rootMatches).toBe(true);
      expect(before?.scrollHeight).toBeGreaterThan(before?.clientHeight ?? 0);
      expect(before?.pageTop).toBeGreaterThan(before?.viewportBottom ?? 0);
      expect(
        await page.evaluate(
          (apiHandle) => window.__folioParity?.scrollToPage(3, apiHandle),
          handle,
        ),
      ).toBe(true);
      await expect(async () => {
        const after = await page.evaluate(() => window.__folioParity?.readScrollViewport(3));
        expect(after?.scrollTop).toBeGreaterThan(before?.scrollTop ?? 0);
        expect(after?.pageTop).toBeLessThan(after?.viewportBottom ?? 0);
        expect(after?.pageBottom).toBeGreaterThan(after?.viewportTop ?? 0);
      }).toPass({ timeout: 5_000 });
    },
  );
}
