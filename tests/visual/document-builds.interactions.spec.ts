/**
 * Typing builds no whole document between keys, in either host or engine.
 *
 * Copying a document or converting one from the editor state costs time in
 * proportion to the document, so a host that derives toolbar, ruler or panel
 * state from a freshly built document pays it on every key. Once the edits
 * settle, the default session converts the editor state into the model once,
 * and either session copies the model once for the host's change
 * notification, which never hands out the model itself.
 *
 * The page is wide enough that the outline and the comments open beside the
 * page, so their per-edit updates are exercised too.
 *
 * A slow runner can pause between two keys for longer than the notification
 * delay, and then the notification's build lands between them. Such a build
 * comes at least that delay after the key before it; a build an edit causes
 * comes within the key's own handling.
 */

import type { Page } from "@playwright/test";

import { DOCUMENT_CHANGE_NOTIFY_DELAY_MS } from "../../packages/core/src/controller/documentChangeNotification";
import type {
  DocumentBuildKind,
  LayoutInstrumentation,
} from "../../packages/core/src/layout-engine/layoutInstrumentation";
import { type AdapterFixture, ensureLiveView } from "../parity/parity-fixture";
import { PLAYGROUND_HOSTS } from "../parity/playgroundHosts";
import { waitForPanels } from "../support/panelLayoutAssertions";
import { buildPanelLayoutDocument } from "../support/panelLayoutDocument";
import { expect, test } from "../support/unclassifiedProbe";

type DocumentBuild = { kind: DocumentBuildKind; at: number };

declare global {
  var __folioLayoutInstrumentation: LayoutInstrumentation | undefined;
  var __folioDocumentBuilds: DocumentBuild[] | undefined;
  var __folioKeydowns: number[] | undefined;
}

const KEYS = 20;
const KEY_DELAY_MS = 30;
// Long enough for the change notification after the last edit.
const SETTLE_MS = 1_500;

const HOSTS: readonly AdapterFixture[] = [
  { name: "react", baseUrl: PLAYGROUND_HOSTS.react },
  { name: "vue", baseUrl: PLAYGROUND_HOSTS.vue },
];
const ENGINES = ["default", "canonical"] as const;

test.use({ viewport: { width: 1500, height: 900 } });

const recordBuildsAndKeys = async (page: Page) => {
  await page.addInitScript(() => {
    const builds: DocumentBuild[] = [];
    const keydowns: number[] = [];
    globalThis.__folioDocumentBuilds = builds;
    globalThis.__folioKeydowns = keydowns;
    globalThis.__folioLayoutInstrumentation = {
      ...globalThis.__folioLayoutInstrumentation,
      onDocumentBuilt: ({ kind }) => {
        builds.push({ kind, at: performance.now() });
      },
    };
    window.addEventListener(
      "keydown",
      () => {
        keydowns.push(performance.now());
      },
      { capture: true },
    );
  });
};

const readBuildsAndKeys = (page: Page) =>
  page.evaluate(() => ({
    builds: [...(globalThis.__folioDocumentBuilds ?? [])],
    keydowns: [...(globalThis.__folioKeydowns ?? [])],
  }));

const clearBuildsAndKeys = (page: Page) =>
  page.evaluate(() => {
    globalThis.__folioDocumentBuilds?.splice(0);
    globalThis.__folioKeydowns?.splice(0);
  });

for (const host of HOSTS) {
  for (const engine of ENGINES) {
    test(`typing builds no document between keys [${host.name}, ${engine}]`, async ({
      page,
    }) => {
      const bytes = await buildPanelLayoutDocument("comment-and-changes");
      await page.route("**/fixtures/document-builds.docx", (route) =>
        route.fulfill({ body: Buffer.from(bytes) }),
      );
      await recordBuildsAndKeys(page);
      const session = engine === "canonical" ? "&session=canonical" : "";
      await page.goto(`${host.baseUrl}/?file=document-builds.docx${session}`);
      await page.evaluate(() => document.fonts.ready);
      await ensureLiveView(page);
      await waitForPanels(page, "comment-and-changes");

      // Place the caret inside a clause, then let the click's own work settle.
      await page
        .locator(".layout-page")
        .first()
        .getByText("The Receiving Party shall hold", { exact: false })
        .first()
        .click();
      await expect
        .poll(() =>
          page.evaluate(() =>
            document.activeElement?.closest(".ProseMirror") ? "focused" : "",
          ),
        )
        .toBe("focused");
      await page.waitForTimeout(SETTLE_MS);
      const before = await page.evaluate(
        () => window.__folioParity?.getDocumentText() ?? "",
      );
      await clearBuildsAndKeys(page);

      await page.keyboard.type("x".repeat(KEYS), { delay: KEY_DELAY_MS });
      await page.waitForTimeout(SETTLE_MS);

      const { builds, keydowns } = await readBuildsAndKeys(page);
      const firstKey = keydowns.at(0) ?? Number.NaN;
      const lastKey = keydowns.at(-1) ?? Number.NaN;
      // Builds between keys that no pause long enough for a notification
      // explains.
      const betweenKeys = builds.filter((build) => {
        if (build.at < firstKey || build.at > lastKey) return false;
        const keyBefore = keydowns.findLast((at) => at <= build.at) ?? firstKey;
        return build.at - keyBefore < DOCUMENT_CHANGE_NOTIFY_DELAY_MS;
      });
      const afterKeys = builds.filter((build) => build.at > lastKey);
      const after = await page.evaluate(
        () => window.__folioParity?.getDocumentText() ?? "",
      );

      // Every key reached the document, so the edits really happened.
      expect(keydowns).toHaveLength(KEYS);
      expect(after.length - before.length).toBe(KEYS);
      expect(betweenKeys).toEqual([]);
      // The edits settle once: at most one conversion into the model, and one
      // copy of it for the change notification.
      expect(
        afterKeys.filter(({ kind }) => kind === "conversion").length,
      ).toBeLessThanOrEqual(1);
      expect(
        afterKeys.filter(({ kind }) => kind === "copy").length,
      ).toBeLessThanOrEqual(1);
    });
  }
}
