/**
 * Typing builds no whole document between keys, in either host or engine.
 *
 * Copying a document or converting one from the editor state costs time in
 * proportion to the document, so a host that derives toolbar, ruler or panel
 * state from a freshly built document pays it on every key. The document is
 * built once the edits settle, for the host's change notification.
 *
 * The page is wide enough that the outline and the comments open beside the
 * page, so their per-edit updates are exercised too.
 */

import { expect, test } from "@playwright/test";
import type { Page } from "@playwright/test";

import type {
  DocumentBuildKind,
  LayoutInstrumentation,
} from "../../packages/core/src/layout-engine/layoutInstrumentation";
import { type AdapterFixture, ensureLiveView } from "../parity/parity-fixture";
import { PLAYGROUND_HOSTS } from "../parity/playgroundHosts";
import { waitForPanels } from "../support/panelLayoutAssertions";
import { buildPanelLayoutDocument } from "../support/panelLayoutDocument";

type DocumentBuild = { kind: DocumentBuildKind; at: number };

declare global {
  var __folioLayoutInstrumentation: LayoutInstrumentation | undefined;
  var __folioDocumentBuilds: DocumentBuild[] | undefined;
  var __folioKeydowns: number[] | undefined;
}

const KEYS = 20;
const KEY_DELAY_MS = 30;
// The change notification waits this long after the last edit.
const SETTLE_MS = 1_500;

const HOSTS: readonly AdapterFixture[] = [
  { name: "react", baseUrl: PLAYGROUND_HOSTS.react },
  { name: "vue", baseUrl: PLAYGROUND_HOSTS.vue },
];
const ENGINES = ["default", "canonical"] as const;

test.use({ viewport: { width: 1500, height: 900 } });
test.describe.configure({ timeout: 90_000 });

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

      await page.evaluate(() => {
        const frames: unknown[] = [];
        (globalThis as { __probeFrames?: unknown[] }).__probeFrames = frames;
        new PerformanceObserver((list) => {
          for (const entry of list.getEntries()) {
            const frame = entry as PerformanceEntry & {
              renderStart: number;
              styleAndLayoutStart: number;
              blockingDuration: number;
              scripts: {
                invoker: string;
                sourceFunctionName: string;
                sourceURL: string;
                duration: number;
                forcedStyleAndLayoutDuration: number;
              }[];
            };
            frames.push({
              start: Math.round(frame.startTime),
              duration: Math.round(frame.duration),
              render: Math.round(
                frame.startTime + frame.duration - frame.renderStart,
              ),
              styleLayout: Math.round(
                frame.startTime + frame.duration - frame.styleAndLayoutStart,
              ),
              scripts: frame.scripts.map((script) => ({
                invoker: script.invoker,
                fn: script.sourceFunctionName,
                url: script.sourceURL.split("/").pop(),
                duration: Math.round(script.duration),
                forced: Math.round(script.forcedStyleAndLayoutDuration),
              })),
            });
          }
        }).observe({ type: "long-animation-frame", buffered: false });
      });
      const PROBE_KEYS = 4;
      const startedAt = Date.now();
      await page.keyboard.type("x".repeat(PROBE_KEYS), { delay: KEY_DELAY_MS });
      console.log("PROBE-KEYS", host.name, engine, Date.now() - startedAt);
      await page.waitForTimeout(SETTLE_MS);
      console.log(
        "PROBE-FRAMES",
        host.name,
        engine,
        JSON.stringify(
          await page.evaluate(
            () => (globalThis as { __probeFrames?: unknown[] }).__probeFrames,
          ),
        ).slice(0, 6000),
      );

      const { builds, keydowns } = await readBuildsAndKeys(page);
      const firstKey = keydowns.at(0) ?? Number.NaN;
      const lastKey = keydowns.at(-1) ?? Number.NaN;
      const betweenKeys = builds.filter(
        (build) => build.at >= firstKey && build.at <= lastKey,
      );
      const afterKeys = builds.filter((build) => build.at > lastKey);
      const after = await page.evaluate(
        () => window.__folioParity?.getDocumentText() ?? "",
      );

      // Every key reached the document, so the edits really happened.
      expect(keydowns).toHaveLength(KEYS);
      expect(after.length - before.length).toBe(KEYS);
      expect(betweenKeys).toEqual([]);
      // The edits settle once: one document for the change notification.
      expect(afterKeys.length).toBeLessThanOrEqual(1);
    });
  }
}
