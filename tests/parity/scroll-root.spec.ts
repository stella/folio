import { test } from "@playwright/test";
import type { Page } from "@playwright/test";
import {
  PAGED_SCROLL_NAVIGATION_CASES,
  SCROLL_NAVIGATION_CASES,
} from "../../packages/playground/src/scrollParityBridge";
import { buildScrollRootDocument } from "../support/scrollRootDocument";
import { ensureLiveView, expect, forEachAdapter, openEditor } from "./parity-fixture";
import type { AdapterFixture } from "./parity-fixture";

// The former React-only assertion used toBeVisible(), which accepts elements
// outside an overflow viewport. Assert actual viewport overlap and scrollTop:
// scrolling the non-overflowing pages viewport cannot satisfy this oracle.
for (const handle of ["document", "paged"] as const) {
  forEachAdapter(
    `scrollToPage(3) moves the real scroll root through the ${handle} ref`,
    async (adapter, { page }) => {
      await openScrollFixture(page, adapter);
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

const openScrollFixture = async (page: Page, adapter: AdapterFixture, readyScroll = false) => {
  const bytes = await buildScrollRootDocument();
  await page.route("**/fixtures/scroll-root.docx", (route) =>
    route.fulfill({
      contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      body: Buffer.from(bytes),
    }),
  );
  // Both hosts must be shorter than a page. Outer host scrolling also catches
  // scrollIntoView accidentally propagating to ancestors of the editor root.
  await page.setViewportSize({ width: 1280, height: 600 });
  await page.emulateMedia({ reducedMotion: "reduce" });
  await openEditor(page, adapter, `scroll-root.docx${readyScroll ? "&readyScroll=500" : ""}`);
  await ensureLiveView(page);
  await page.evaluate(() => {
    document.documentElement.style.height = "2200px";
    window.scrollTo({ top: 120, behavior: "instant" });
  });
  expect(await page.evaluate(() => window.scrollY)).toBe(120);
};

const expectTargetInScrollRoot = async (
  page: Page,
  method?: keyof typeof SCROLL_NAVIGATION_CASES,
) => {
  await expect(async () => {
    const state = await page.evaluate((api) => window.__folioScrollParity?.readTarget(api), method);
    expect(state).not.toBeNull();
    expect(state?.scrollTop).toBeGreaterThan(0);
    expect(state?.top).toBeGreaterThanOrEqual(state?.viewportTop ?? Infinity);
    expect(state?.top).toBeLessThan(state?.viewportBottom ?? -Infinity);
    expect(state?.bottom).toBeLessThanOrEqual(state?.viewportBottom ?? -Infinity);
    expect(await page.evaluate(() => window.scrollY)).toBe(120);
  }).toPass({ timeout: 5_000 });
};

for (const { method } of Object.values(SCROLL_NAVIGATION_CASES)) {
  forEachAdapter(
    `public ${method} reveals its target in the scroll root`,
    async (adapter, { page }) => {
      await openScrollFixture(page, adapter);
      if (method === "scrollToSuggestion") {
        expect(await page.evaluate(() => window.__folioScrollParity?.prepareSuggestion())).toBe(
          true,
        );
        await page.evaluate(
          () =>
            new Promise<void>((resolve) =>
              requestAnimationFrame(() => requestAnimationFrame(() => resolve())),
            ),
        );
      }
      const before = await page.evaluate(
        (api) => window.__folioScrollParity?.readTarget(api),
        method,
      );
      expect(before?.scrollTop).toBe(0);
      expect(before?.top).toBeGreaterThan(before?.viewportBottom ?? Infinity);
      expect(await page.evaluate((api) => window.__folioScrollParity?.navigate(api), method)).toBe(
        true,
      );
      await expectTargetInScrollRoot(page, method);
    },
  );
}

for (const { method } of Object.values(PAGED_SCROLL_NAVIGATION_CASES)) {
  forEachAdapter(
    `paged ${method} reveals page-three target in the scroll root`,
    async (adapter, { page }) => {
      await openScrollFixture(page, adapter);
      expect(
        await page.evaluate((api) => window.__folioScrollParity?.navigatePaged(api), method),
      ).toBe(true);
      await expectTargetInScrollRoot(page);
    },
  );
}

forEachAdapter(
  "host scroll in onEditorViewReady survives document readiness",
  async (adapter, { page }) => {
    await openScrollFixture(page, adapter, true);
    const before = await page.evaluate(() => window.__folioScrollParity?.readReady());
    if (adapter.name === "vue") {
      // Vue creates a new view during a buffer reload. Preserved pages let the
      // host's real callback set a nonzero scroll position before readiness.
      expect(await page.evaluate(() => window.__folioScrollParity?.reloadForReady())).toBe(true);
      await ensureLiveView(page);
    }
    // React buffer reload updates the same EditorView and emits no new view
    // lifecycle signal. openScrollFixture creates its first lazy view only
    // after pages are painted, which is the real host callback sequence.
    const ready = await page.evaluate(() => window.__folioScrollParity?.readReady());
    if (adapter.name === "vue") {
      expect(ready?.count).toBeGreaterThan(before?.count ?? 0);
    } else {
      expect(ready?.count).toBe(1);
    }
    expect(ready?.appliedTop).toBe(500);
    expect(ready?.scrollTop).toBe(500);
    // Wait across the paint that previously reset the callback's scroll request.
    await page.evaluate(
      () =>
        new Promise<void>((resolve) =>
          requestAnimationFrame(() => requestAnimationFrame(() => resolve())),
        ),
    );
    expect((await page.evaluate(() => window.__folioScrollParity?.readReady()))?.scrollTop).toBe(
      500,
    );
  },
);

// React exposes onEditorViewReady; it has no Vue-style `ready` event.
test("host scroll in Vue ready event survives the next paint [vue]", async ({ page }) => {
  const adapter = {
    name: "vue",
    baseUrl: `http://localhost:${Number(process.env["FOLIO_PLAYGROUND_VUE_PORT"]) || 4201}`,
  } as const;
  const bytes = await buildScrollRootDocument();
  await page.route("**/fixtures/scroll-root.docx", (route) =>
    route.fulfill({
      contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      body: Buffer.from(bytes),
    }),
  );
  await page.setViewportSize({ width: 1280, height: 600 });
  await openEditor(page, adapter, "scroll-root.docx&readyEventScroll");
  await ensureLiveView(page);
  const before = await page.evaluate(() => window.__folioScrollParity?.readReady());
  expect(await page.evaluate(() => window.__folioScrollParity?.reloadForReady())).toBe(true);
  await ensureLiveView(page);
  await expect(async () => {
    const state = await page.evaluate(() => window.__folioScrollParity?.readReady());
    expect(state?.eventCount).toBeGreaterThan(before?.eventCount ?? 0);
    expect(state?.eventAppliedTop).toBe(500);
    expect(state?.scrollTop).toBe(500);
  }).toPass({ timeout: 5_000 });
  await page.evaluate(
    () =>
      new Promise<void>((resolve) =>
        requestAnimationFrame(() => requestAnimationFrame(() => resolve())),
      ),
  );
  expect((await page.evaluate(() => window.__folioScrollParity?.readReady()))?.scrollTop).toBe(500);
});
