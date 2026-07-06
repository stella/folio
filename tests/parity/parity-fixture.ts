/**
 * Cross-adapter parity harness. Declares each spec body once and forks it per
 * adapter (React on 4200, Vue on 4201), appending `[react]` / `[vue]` to the
 * title. Both playgrounds expose an identical `window.__folioParity` bridge
 * (see packages/playground/src/App.tsx and packages/playground-vue/src/App.vue),
 * so one spec body proves the two editors behave the same.
 *
 * The bridge is built only on `DocxEditorRef` members classified `paired` in
 * scripts/parity/parity.contract.json — the members that are real (not stubbed)
 * in both adapters.
 */

import { test, expect, type Page } from "@playwright/test";

/** Mirror of the playgrounds' `FolioParityBridge`. Kept in sync by hand. */
type FolioParityBridge = {
  getTotalPages: () => number;
  ensureView: () => void;
  hasView: () => boolean;
  getDocumentText: () => string;
  insertText: (text: string) => boolean;
  boldFirstWord: () => boolean;
  insertTable: (rows: number, cols: number) => boolean;
  countTables: () => number;
  commentFirstWord: () => boolean;
  countCommentAnchors: () => number;
  aiSnapshotBlockCount: () => number;
  save: () => Promise<number>;
};

declare global {
  // Window augmentation requires an interface (declaration merging).
  // eslint-disable-next-line typescript/consistent-type-definitions
  interface Window {
    __folioParity?: FolioParityBridge;
  }
}

export type AdapterFixture = {
  /** Stable identifier — appears in the test title. */
  name: "react" | "vue";
  /** Playground origin for this adapter. */
  baseUrl: string;
};

const ADAPTERS: AdapterFixture[] = [
  { name: "react", baseUrl: "http://localhost:4200" },
  { name: "vue", baseUrl: "http://localhost:4201" },
];

/** The pages container both adapters paint into (React + Vue share the class). */
const PAGES_SELECTOR = ".paged-editor__pages";
const DEFAULT_FIXTURE = "sample.docx";

/**
 * Navigate to the adapter playground with `?file=<fixture>`, wait for the pages
 * container, and poll until at least one page has been laid out through the
 * shared bridge.
 */
export async function openEditor(
  page: Page,
  adapter: AdapterFixture,
  fixture = DEFAULT_FIXTURE,
): Promise<void> {
  await page.goto(`${adapter.baseUrl}/?file=${fixture}`);
  await page.waitForSelector(PAGES_SELECTOR, { timeout: 30_000 });
  await expect
    .poll(() => page.evaluate(() => window.__folioParity?.getTotalPages() ?? 0), {
      timeout: 30_000,
    })
    .toBeGreaterThan(0);
}

/**
 * Force-create the deferred editor view and wait until the live ProseMirror
 * view exists — required before any bridge method that touches editor state
 * (insertText, boldFirstWord, getDocumentText).
 */
export async function ensureLiveView(page: Page): Promise<void> {
  await page.evaluate(() => window.__folioParity?.ensureView());
  await expect
    .poll(() => page.evaluate(() => window.__folioParity?.hasView() ?? false), {
      timeout: 30_000,
    })
    .toBe(true);
}

/** Declare a spec once; run it once per adapter with a `[<adapter>]` suffix. */
export function forEachAdapter(
  title: string,
  body: (adapter: AdapterFixture, args: { page: Page }) => Promise<void>,
): void {
  for (const adapter of ADAPTERS) {
    // Playwright requires the first arg to be a literal destructuring pattern.
    test(`${title} [${adapter.name}]`, ({ page }) => body(adapter, { page }));
  }
}

export { expect };
