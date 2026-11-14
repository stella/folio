import { shapeArrayBuffer } from "../../packages/core/src/__tests__/documentShapes";
import type { DocxEditorRef } from "../../packages/react/src/components/DocxEditor.props";
import type {} from "../parity/canonicalFuzzErrors";
import { evaluateCanonicalLoad } from "../parity/canonicalLoadEvaluation";
import { expect, test } from "../support/unclassifiedProbe";
import { driveBrowserIme } from "./browserImeDriver";

declare global {
  var __folioPlayground:
    { getEditorRef: () => DocxEditorRef | null } | undefined;
}

const MODIFIER = process.platform === "darwin" ? "Meta" : "Control";

test("a composition after a paste starts at the caret, not at the document's start", async ({
  page,
}) => {
  const bytes = [
    ...new Uint8Array(await shapeArrayBuffer("single-bullet-list")),
  ];
  await page.goto("/?session=canonical");
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
  await page.evaluate(() =>
    globalThis.__folioPlayground
      ?.getEditorRef()
      ?.getEditorRef()
      ?.getView()
      ?.focus(),
  );
  await page.context().grantPermissions(["clipboard-read", "clipboard-write"]);
  await page.evaluate(async () => {
    await navigator.clipboard.write([
      new ClipboardItem({
        "text/plain": new Blob(["Title\nBody"], { type: "text/plain" }),
        "text/html": new Blob(["<p>Title</p><p>Body</p>"], {
          type: "text/html",
        }),
      }),
    ]);
  });
  await page.keyboard.press(`${MODIFIER}+v`);
  const paragraphs = () =>
    page.evaluate(() => {
      const texts: string[] = [];
      globalThis.__folioPlayground
        ?.getEditorRef()
        ?.getEditorRef()
        ?.getView()
        ?.state.doc.descendants((node) => {
          if (!node.isTextblock) return true;
          texts.push(node.textContent);
          return false;
        });
      return texts.slice(0, 2);
    });
  await expect.poll(paragraphs).toEqual(["Title", "BodyIntro paragraph."]);

  await driveBrowserIme(page, {
    kind: "imeReplacement",
    updates: ["alpha"],
    completion: "commit",
  });

  await expect.poll(paragraphs).toEqual(["Title", "BodyalphaIntro paragraph."]);
  expect(
    await page.evaluate(() =>
      globalThis.__folioCanonicalFuzzErrors?.map(({ message }) => message),
    ),
  ).toEqual([]);
});
