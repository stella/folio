import { expect, type Page, test } from "@playwright/test";

import { shapeArrayBuffer } from "../../packages/core/src/__tests__/documentShapes";
import type { DocxEditorRef } from "../../packages/react/src/components/DocxEditor.props";
import type {} from "../parity/canonicalFuzzErrors";
import { evaluateCanonicalLoad } from "../parity/canonicalLoadEvaluation";
import { driveBrowserIme } from "./browserImeDriver";
import type { BrowserInputAction } from "./browserInputTrace";

declare global {
  var __folioPlayground:
    { getEditorRef: () => DocxEditorRef | null } | undefined;
}

const COMPOSITIONS = [
  { updates: ["alpha"], completion: "cancel" },
  { updates: ["alpha", "beta gamma"], completion: "cancel" },
  { updates: ["alpha", "beta gamma"], completion: "commit" },
] as const;

const bookmarkBoundaries = (page: Page) =>
  page.evaluate(() => {
    const view = globalThis.__folioPlayground
      ?.getEditorRef()
      ?.getEditorRef()
      ?.getView();
    let count = 0;
    view?.state.doc.descendants((node) => {
      if (node.type.name === "bookmarkBoundary") count += 1;
    });
    return count;
  });

const load = async (page: Page) => {
  const bytes = [
    ...new Uint8Array(await shapeArrayBuffer("fields-links-bookmarks")),
  ];
  await page.goto("/");
  await page.waitForSelector(".layout-page");
  await page.evaluate(() =>
    globalThis.__folioPlayground?.getEditorRef()?.ensureEditorView(),
  );
  await page.waitForFunction(
    () =>
      !!globalThis.__folioPlayground?.getEditorRef()?.getEditorRef()?.getView(),
  );
  await page.evaluate(() => {
    globalThis.__folioCanonicalFuzzErrors = [];
  });
  expect(await evaluateCanonicalLoad(page, bytes)).toBe(true);
  await expect.poll(() => bookmarkBoundaries(page)).toBe(2);
  await page.evaluate(() =>
    globalThis.__folioPlayground
      ?.getEditorRef()
      ?.getEditorRef()
      ?.getView()
      ?.focus(),
  );
};

for (const composition of COMPOSITIONS) {
  test(`${composition.completion} of ${composition.updates.length} composition updates keeps the bookmark beside it (canonical)`, async ({
    page,
  }) => {
    await load(page);
    const action = {
      kind: "imeReplacement",
      updates: [...composition.updates],
      completion: composition.completion,
    } satisfies BrowserInputAction;
    await driveBrowserIme(page, action);
    // A browser that sweeps the element takes the boundary, and then its pair.
    await page.waitForTimeout(100);
    expect(await bookmarkBoundaries(page)).toBe(2);
    expect(
      await page.evaluate(() =>
        globalThis.__folioCanonicalFuzzErrors?.map(({ message }) => message),
      ),
    ).toEqual([]);
  });
}
