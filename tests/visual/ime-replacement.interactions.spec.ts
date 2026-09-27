import { expect, test } from "@playwright/test";

import { FolioDocxReviewer } from "../../packages/core/src/ai-edits/headless";
import type { DocxEditorRef } from "../../packages/react/src/components/DocxEditor.props";

declare global {
  var __folioPlayground: { getEditorRef: () => DocxEditorRef | null } | undefined;
}

test("IME replacement preserves the selected text as a deletion after save @browser-input:imeReplacement", async ({
  page,
}) => {
  await page.goto("/?file=sample.docx");
  await page.waitForSelector(".layout-page");
  await page.evaluate(() =>
    globalThis.__folioPlayground?.getEditorRef()?.ensureEditorView({ focus: true }),
  );
  await page.waitForFunction(
    () => !!globalThis.__folioPlayground?.getEditorRef()?.getEditorRef()?.getView(),
  );
  const original = await page.evaluate(async () => {
    const editor = globalThis.__folioPlayground?.getEditorRef();
    const bytes = await editor?.save({ selective: false });
    return bytes ? [...new Uint8Array(bytes)] : null;
  });
  if (!original) throw new Error("editor did not return the original DOCX");

  await page.getByRole("button", { name: "Track Changes", exact: true }).click();
  const selected = await page.evaluate(() => {
    const view = globalThis.__folioPlayground?.getEditorRef()?.getEditorRef()?.getView();
    if (!view) throw new Error("editor view unavailable");
    let from: number | null = null;
    let text = "";
    view.state.doc.descendants((node, pos) => {
      if (from !== null) return false;
      if (
        node.isText &&
        (node.text?.length ?? 0) >= 6 &&
        node.marks.every((mark) => mark.type.name !== "insertion" && mark.type.name !== "deletion")
      ) {
        from = pos;
        text = node.text?.slice(0, 6) ?? "";
        return false;
      }
      return true;
    });
    if (from === null) throw new Error("fixture has no six-character plain text run");
    view.dispatch(
      view.state.tr.setSelection(
        view.state.selection.constructor.create(view.state.doc, from, from + 6),
      ),
    );
    view.focus();
    return text;
  });
  expect(await page.evaluate(() => window.getSelection()?.toString())).toBe(selected);

  const cdp = await page.context().newCDPSession(page);
  await cdp.send("Input.imeSetComposition", { text: "契", selectionStart: 1, selectionEnd: 1 });
  await cdp.send("Input.imeSetComposition", { text: "契約", selectionStart: 2, selectionEnd: 2 });
  await cdp.send("Input.insertText", { text: "契約" });

  await expect
    .poll(() =>
      page.evaluate(() => {
        const view = globalThis.__folioPlayground?.getEditorRef()?.getEditorRef()?.getView();
        const marked = { deletion: "", insertion: "" };
        view?.state.doc.descendants((node) => {
          if (!node.isText) return true;
          if (node.marks.some((mark) => mark.type.name === "deletion"))
            marked.deletion += node.text ?? "";
          if (node.marks.some((mark) => mark.type.name === "insertion"))
            marked.insertion += node.text ?? "";
          return true;
        });
        return marked;
      }),
    )
    .toEqual({ deletion: selected, insertion: "契約" });

  const saved = await page.evaluate(async () => {
    const bytes = await globalThis.__folioPlayground?.getEditorRef()?.save({ selective: false });
    return bytes ? [...new Uint8Array(bytes)] : null;
  });
  if (!saved) throw new Error("editor did not return the revised DOCX");
  const bytes = new Uint8Array(saved);
  const reviewed = await FolioDocxReviewer.fromBuffer(bytes.buffer);
  const baseline = await FolioDocxReviewer.fromBuffer(new Uint8Array(original).buffer);
  const originalBlocks = baseline.snapshot().blocks.map((block) => block.text);
  expect(reviewed.getChanges().map(({ type, text }) => ({ type, text }))).toEqual(
    expect.arrayContaining([
      { type: "deletion", text: selected },
      { type: "insertion", text: "契約" },
    ]),
  );
  const accepting = await FolioDocxReviewer.fromBuffer(bytes.buffer);
  accepting.acceptAll();
  const accepted = await FolioDocxReviewer.fromBuffer(await accepting.toBuffer());
  const changedBlock = originalBlocks.findIndex((text) => text.includes(selected));
  expect(changedBlock).toBeGreaterThanOrEqual(0);
  expect(accepted.snapshot().blocks.map((block) => block.text)).toEqual(
    originalBlocks.map((text, index) =>
      index === changedBlock ? text.replace(selected, "契約") : text,
    ),
  );
  reviewed.rejectAll();
  const reopened = await FolioDocxReviewer.fromBuffer(await reviewed.toBuffer());
  expect(reopened.snapshot().blocks.map((block) => block.text)).toEqual(originalBlocks);
});
