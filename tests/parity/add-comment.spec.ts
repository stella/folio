import { ensureLiveView, expect, forEachAdapter, openEditor } from "./parity-fixture";

// Both adapters now wire the comment lifecycle: applying a `comment` mark via
// the shared core schema paints a `[data-comment-id]` anchor through the shared
// layout painter, and that anchor is what the sidebar threads a comment card
// onto. Driving the mark through the bridge proves the comment-anchor path is
// present and identical in each adapter (React CommentsSidebar / Vue
// UnifiedSidebar both read the same painter attribute).
forEachAdapter("comment: paints a comment anchor through the bridge", async (adapter, { page }) => {
  await openEditor(page, adapter);
  await ensureLiveView(page);

  const before = await page.evaluate(() => window.__folioParity?.countCommentAnchors() ?? -1);
  const applied = await page.evaluate(() => window.__folioParity?.commentFirstWord() ?? false);

  expect(applied).toBe(true);
  await expect
    .poll(() => page.evaluate(() => window.__folioParity?.countCommentAnchors() ?? -1), {
      timeout: 30_000,
    })
    .toBeGreaterThan(before);
});
