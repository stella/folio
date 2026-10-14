import { test } from "@playwright/test";
import type { FindNotificationBridge } from "../../packages/playground-vue/src/findNotificationBridge";
import { expect, openEditor, VUE_ADAPTER } from "./parity-fixture";

declare global {
  var __folioVueFindNotifications: FindNotificationBridge | undefined;
}

// Composable tests invoke a plain callback; mount the adapter so Vue's event
// listener forwarding participates in every notification.
test("find result listener receives one notification per update [vue]", async ({ page }) => {
  await openEditor(page, VUE_ADAPTER);
  const updates = await page.evaluate(() => {
    const bridge = globalThis.__folioVueFindNotifications;
    if (!bridge) throw new TypeError("Missing Vue find notification bridge");
    bridge.takeNotifications();
    bridge.find("editor");
    const found = bridge.takeNotifications();
    bridge.findNext();
    const next = bridge.takeNotifications();
    bridge.findPrevious();
    const previous = bridge.takeNotifications();
    bridge.find("absent-search-notification-match");
    const absent = bridge.takeNotifications();
    bridge.clearFind();
    const cleared = bridge.takeNotifications();
    return { found, next, previous, absent, cleared };
  });
  expect(updates.found).toHaveLength(1);
  expect(updates.found.at(0)?.totalCount).toBeGreaterThan(0);
  expect(updates.next).toHaveLength(1);
  expect(updates.previous).toHaveLength(1);
  expect(updates.absent).toHaveLength(1);
  expect(updates.absent.at(0)?.totalCount).toBe(0);
  expect(updates.cleared).toEqual([null]);
});
