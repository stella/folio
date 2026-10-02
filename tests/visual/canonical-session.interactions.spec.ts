import { expect, test, type Page } from "@playwright/test";

import { parseDocx } from "../../packages/core/src/docx/parser";
import { createDocx } from "../../packages/core/src/docx/rezip";
import { createEmptyDocument } from "../../packages/core/src/utils/createDocument";
import type { buildCanonicalBridge } from "../parity/canonicalBridge";

const MODIFIER = process.platform === "darwin" ? "Meta" : "Control";
const reactPort = Number(process.env["FOLIO_PLAYGROUND_PORT"]) || 4200;
const vuePort = Number(process.env["FOLIO_PLAYGROUND_VUE_PORT"]) || 4201;

const snapshot = (page: Page) => page.evaluate(() => globalThis.__folioCanonical?.snapshot());

type ExpectedProjection = Pick<
  ReturnType<ReturnType<typeof buildCanonicalBridge>["snapshot"]>,
  "text" | "selection"
>;

const expectProjection = async (page: Page, expected: ExpectedProjection) => {
  const current = await snapshot(page);
  expect(current?.active).toBe(true);
  expect(current?.text).toBe(expected.text);
  expect(current?.selection).toEqual(expected.selection);
  if (!current?.document) throw new Error("Canonical document unavailable.");
  expect(current.projectionMatchesCanonical).toBe(true);
  expect(current.projectionJSON).toEqual(current.canonicalProjectionJSON);
  expect(current.provenance.valid).toBe(true);
  expect(current.provenance.capturedParagraphCount).toBeGreaterThan(0);
  return { ...current, document: current.document };
};

const select = async (page: Page, anchor: number, head = anchor) => {
  expect(
    await page.evaluate(({ from, to }) => globalThis.__folioCanonical?.select(from, to), {
      from: anchor,
      to: head,
    }),
  ).toBe(true);
};

test("canonical input, history and saved document agree across both adapters", async ({ page }) => {
  test.setTimeout(60_000);
  const sourceDocument = createEmptyDocument({ initialText: "A😀B" });
  const paragraph = sourceDocument.package.document.content.at(0);
  if (!paragraph || paragraph.type !== "paragraph")
    throw new TypeError("Expected source paragraph.");
  paragraph.formatting = { ...paragraph.formatting, keepNext: true, contextualSpacing: true };
  const source = await createDocx(sourceDocument);
  let firstSaved: Awaited<ReturnType<typeof parseDocx>> | undefined;
  for (const port of [reactPort, vuePort]) {
    await page.goto(`http://localhost:${port}/?session=canonical`);
    await page.waitForSelector(".layout-page");
    expect(
      await page.evaluate(
        async (bytes) => globalThis.__folioCanonical?.load(bytes),
        [...new Uint8Array(source)],
      ),
    ).toBe(true);
    await page.evaluate(
      () => new Promise<void>((resolve) => requestAnimationFrame(() => resolve())),
    );
    await select(page, 5);
    const initial = await expectProjection(page, { text: "A😀B", selection: { from: 5, to: 5 } });
    expect(initial.canUndo).toBe(false);

    await page.keyboard.type("x");
    await expectProjection(page, { text: "A😀Bx", selection: { from: 6, to: 6 } });
    await page.keyboard.press("Backspace");
    await expectProjection(page, { text: "A😀B", selection: { from: 5, to: 5 } });
    await select(page, 2);
    await page.keyboard.press("Delete");
    await expectProjection(page, { text: "AB", selection: { from: 2, to: 2 } });
    await page.keyboard.insertText("é");
    await expectProjection(page, { text: "AéB", selection: { from: 3, to: 3 } });
    await select(page, 1, 3);
    await page.keyboard.press("Backspace");
    const deleted = await expectProjection(page, { text: "B", selection: { from: 1, to: 1 } });
    expect(deleted.canUndo).toBe(true);

    await page.keyboard.press(`${MODIFIER}+z`);
    const undone = await expectProjection(page, { text: "AéB", selection: { from: 1, to: 3 } });
    expect(undone.canRedo).toBe(true);
    await page.keyboard.press(`${MODIFIER}+Shift+z`);
    const redone = await expectProjection(page, { text: "B", selection: { from: 1, to: 1 } });
    expect(redone.document).toEqual(deleted.document);
    expect(redone.canRedo).toBe(false);

    // Refused structural input must leave the canonical model, projection and journal intact.
    await page.keyboard.press("Enter");
    expect(await snapshot(page)).toEqual(redone);

    await select(page, 1, 2);
    const compositionBaseline = await expectProjection(page, {
      text: "B",
      selection: { from: 1, to: 2 },
    });
    const compositionTrace = await page.evaluate(async () => {
      const editor = document.querySelector<HTMLElement>(".ProseMirror");
      const bridge = globalThis.__folioCanonical;
      if (!editor || !bridge) throw new Error("Canonical editor unavailable.");
      const replaceDOMText = (before: string, after: string) => {
        const walker = document.createTreeWalker(editor, NodeFilter.SHOW_TEXT);
        let node = walker.nextNode();
        while (node && node.textContent !== before) node = walker.nextNode();
        if (!node) throw new Error("Composition text node unavailable.");
        node.textContent = after;
        const range = document.createRange();
        range.selectNodeContents(node);
        range.collapse(false);
        const selection = window.getSelection();
        selection?.removeAllRanges();
        selection?.addRange(range);
      };
      editor.dispatchEvent(new CompositionEvent("compositionstart", { bubbles: true }));
      const provisional = new InputEvent("beforeinput", {
        bubbles: true,
        inputType: "insertCompositionText",
        data: "契",
        isComposing: true,
        cancelable: true,
      });
      editor.dispatchEvent(provisional);
      replaceDOMText("B", "契");
      editor.dispatchEvent(
        new InputEvent("input", {
          bubbles: true,
          inputType: "insertCompositionText",
          data: "契",
          isComposing: true,
        }),
      );
      // Let the real DOM observer produce the provisional PM composition transaction.
      await new Promise<void>((resolve) => setTimeout(resolve, 0));
      const provisionalText = editor.textContent;
      let snapshotBlocked = false;
      try {
        bridge.snapshot();
      } catch {
        snapshotBlocked = true;
      }
      const saveBlocked = await bridge.save().then(
        (saved) => saved === null,
        () => true,
      );
      editor.dispatchEvent(new CompositionEvent("compositionend", { bubbles: true, data: "契" }));
      await Promise.resolve();
      await new Promise<void>((resolve) => setTimeout(resolve, 0));
      const final = new InputEvent("beforeinput", {
        bubbles: true,
        inputType: "insertFromComposition",
        data: "契約",
        cancelable: true,
      });
      // Chromium clears inputType values emitted by other engines in synthetic events.
      // Preserve the late-final-input fixture instead of testing an unclassified event.
      Object.defineProperty(final, "inputType", { value: "insertFromComposition" });
      editor.dispatchEvent(final);
      replaceDOMText("契", "契約");
      for (let index = 0; index < 2; index++) {
        editor.dispatchEvent(
          new InputEvent("input", {
            bubbles: true,
            inputType: "insertFromComposition",
            data: "契約",
          }),
        );
        editor.dispatchEvent(
          new CompositionEvent("compositionend", { bubbles: true, data: "契約" }),
        );
      }
      await new Promise<void>((resolve) => setTimeout(resolve, 80));
      return {
        provisionalText,
        snapshotBlocked,
        saveBlocked,
        provisionalPrevented: provisional.defaultPrevented,
        finalPrevented: final.defaultPrevented,
        finalInputType: final.inputType,
      };
    });
    expect(compositionTrace).toEqual({
      provisionalText: "契",
      snapshotBlocked: true,
      saveBlocked: true,
      provisionalPrevented: false,
      finalPrevented: false,
      finalInputType: "insertFromComposition",
    });
    const composed = await expectProjection(page, {
      text: "契約",
      selection: { from: 3, to: 3 },
    });
    await page.keyboard.press(`${MODIFIER}+z`);
    const compositionUndone = await expectProjection(page, {
      text: "B",
      selection: { from: 1, to: 2 },
    });
    expect(compositionUndone.document).toEqual(compositionBaseline.document);
    await page.keyboard.press(`${MODIFIER}+Shift+z`);
    const compositionRedone = await expectProjection(page, {
      text: "契約",
      selection: { from: 3, to: 3 },
    });
    expect(compositionRedone.document).toEqual(composed.document);
    await page.keyboard.press(`${MODIFIER}+z`);
    await expectProjection(page, { text: "B", selection: { from: 1, to: 2 } });

    const saved = await page.evaluate(() => globalThis.__folioCanonical?.save());
    if (!saved) throw new Error("Canonical editor did not save.");
    const reopened = await parseDocx(new Uint8Array(saved), {
      preloadFonts: false,
      detectVariables: false,
    });
    expect(reopened.package.document.content).toEqual(redone.document.package.document.content);
    if (firstSaved) {
      expect(reopened.package.document).toEqual(firstSaved.package.document);
    } else {
      firstSaved = reopened;
    }
    expect(
      await page.evaluate(async (bytes) => globalThis.__folioCanonical?.load(bytes), saved),
    ).toBe(true);
    await page.evaluate(
      () => new Promise<void>((resolve) => requestAnimationFrame(() => resolve())),
    );
    await select(page, 1);
    const reloaded = await expectProjection(page, { text: "B", selection: { from: 1, to: 1 } });
    expect(reloaded.document.package.document.content).toEqual(
      redone.document.package.document.content,
    );
    expect(reloaded.canUndo).toBe(false);
  }
});
