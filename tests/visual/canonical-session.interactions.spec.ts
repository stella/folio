import { expect, test, type Page } from "@playwright/test";

import { parseDocx } from "../../packages/core/src/docx/parser";
import { createDocx } from "../../packages/core/src/docx/rezip";
import { createEmptyDocument } from "../../packages/core/src/utils/createDocument";
import { identityKeysIn, IDENTITY_SPACES } from "../../packages/docx-core/src/ops/ids";
import { normalizeForOps } from "../../packages/docx-core/src/ops/contract";
import type { buildCanonicalBridge } from "../parity/canonicalBridge";

const MODIFIER = process.platform === "darwin" ? "Meta" : "Control";
const reactPort = Number(process.env["FOLIO_PLAYGROUND_PORT"]) || 4200;
const vuePort = Number(process.env["FOLIO_PLAYGROUND_VUE_PORT"]) || 4201;

const snapshot = (page: Page) => page.evaluate(() => globalThis.__folioCanonical?.snapshot());

const clearRefusals = (page: Page) =>
  page.evaluate(() => {
    globalThis.__folioCanonicalFuzzErrors = [];
  });

const expectNoRefusals = async (page: Page) =>
  expect(await page.evaluate(() => globalThis.__folioCanonicalFuzzErrors)).toEqual([]);

test.afterEach(async ({ page }, testInfo) => {
  if (testInfo.status === testInfo.expectedStatus) return;
  console.log(
    "Canonical interaction failure state:",
    JSON.stringify(
      await page.evaluate(() => ({
        activeElement: document.activeElement?.outerHTML,
        refusals: globalThis.__folioCanonicalFuzzErrors,
        snapshot: globalThis.__folioCanonical?.snapshot(),
      })),
    ),
  );
});

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

type LoadedFixture = { bytes: number[]; content: string };

const loadReady = async (page: Page, source: LoadedFixture) => {
  expect(
    await page.evaluate((bytes) => globalThis.__folioCanonical?.load(bytes), source.bytes),
  ).toBe(true);
  // Adapter loading schedules external-document synchronization after parsing.
  // Input starts only once the canonical owner exposes the loaded baseline.
  await page.waitForFunction((content) => {
    const current = globalThis.__folioCanonical?.snapshot();
    return (
      current?.active &&
      current.projectionMatchesCanonical &&
      current.canUndo === false &&
      JSON.stringify(current.document?.package.document.content) === content
    );
  }, source.content);
};

const loadedFixture = async (buffer: ArrayBuffer): Promise<LoadedFixture> => {
  const bytes = new Uint8Array(buffer);
  const parsed = await parseDocx(bytes, { preloadFonts: false, detectVariables: false });
  return {
    bytes: [...bytes],
    content: JSON.stringify(normalizeForOps(parsed).package.document.content),
  };
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

    // Structural input shares the same journal, including the caret before its split.
    await page.keyboard.press("Enter");
    expect((await snapshot(page))?.document?.package.document.content).toHaveLength(2);
    await page.keyboard.press(`${MODIFIER}+z`);
    const splitUndone = await snapshot(page);
    expect(splitUndone?.document).toEqual(redone.document);
    expect(splitUndone?.selection).toEqual(redone.selection);
    expect(splitUndone?.projectionMatchesCanonical).toBe(true);

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
    const reopened = structuredClone(
      await parseDocx(new Uint8Array(saved), {
        preloadFonts: false,
        detectVariables: false,
      }),
    );
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

test("canonical structural and formatting hooks survive save and reopen", async ({ page }) => {
  test.setTimeout(60_000);
  const source = await loadedFixture(await createDocx(createEmptyDocument({ initialText: "ab" })));
  for (const port of [reactPort, vuePort]) {
    await page.goto(`http://localhost:${port}/?session=canonical`);
    await page.waitForSelector(".layout-page");
    await loadReady(page, source);
    await select(page, 2);
    const initial = await expectProjection(page, { text: "ab", selection: { from: 2, to: 2 } });
    const original = initial.document.package.document.content.at(0);
    if (original?.type !== "paragraph") throw new TypeError("Expected paragraph.");
    await page.keyboard.press("Enter");
    const split = await expectProjection(page, { text: "ab", selection: { from: 4, to: 4 } });
    expect(split.document.package.document.content).toHaveLength(2);
    const splitFirst = split.document.package.document.content.at(0);
    const splitSecond = split.document.package.document.content.at(1);
    expect(splitSecond?.type === "paragraph" ? splitSecond.paraId : undefined).toBe(
      original.paraId,
    );
    expect(splitFirst?.type === "paragraph" ? splitFirst.paraId : undefined).not.toBe(
      original.paraId,
    );
    await page.keyboard.press("Backspace");
    const joined = await expectProjection(page, { text: "ab", selection: { from: 2, to: 2 } });
    expect(joined.document.package.document.content).toHaveLength(1);
    const joinedParagraph = joined.document.package.document.content.at(0);
    expect(joinedParagraph?.type === "paragraph" ? joinedParagraph.paraId : undefined).toBe(
      original.paraId,
    );
    await select(page, 1, 3);
    await clearRefusals(page);
    await page.keyboard.press(`${MODIFIER}+b`);
    await expectNoRefusals(page);
    const formatted = await expectProjection(page, { text: "ab", selection: { from: 1, to: 3 } });
    const paragraph = formatted.document.package.document.content.at(0);
    expect(
      paragraph?.type === "paragraph" &&
        paragraph.content.some((run) => run.type === "run" && run.formatting?.bold),
    ).toBe(true);
    await select(page, 3);
    await page.keyboard.press("Enter");
    await page.keyboard.type("- ");
    await page.keyboard.press("Tab");
    const beforeItem = await snapshot(page);
    await page.keyboard.type("item");
    const current = await snapshot(page);
    expect(current?.projectionMatchesCanonical).toBe(true);
    expect(current?.projectionJSON).toEqual(current?.canonicalProjectionJSON);
    const listed = current?.document?.package.document.content.at(1);
    expect(
      listed?.type === "paragraph" &&
        listed.formatting?.numPr?.kind === "reference" &&
        listed.formatting.numPr.ilvl,
    ).toBe(1);
    const beforeUndo = current?.document;
    await page.keyboard.press(`${MODIFIER}+z`);
    const itemUndone = await snapshot(page);
    expect(itemUndone?.document).toEqual(beforeItem?.document);
    expect(itemUndone?.selection).toEqual(beforeItem?.selection);
    expect(itemUndone?.text).toBe("ab");
    expect(itemUndone?.projectionMatchesCanonical).toBe(true);
    await page.keyboard.press(`${MODIFIER}+Shift+z`);
    expect((await snapshot(page))?.document).toEqual(beforeUndo);
    const saved = await page.evaluate(() => globalThis.__folioCanonical?.save());
    if (!saved) throw new TypeError("Canonical save unavailable.");
    const reopened = await parseDocx(new Uint8Array(saved), {
      preloadFonts: false,
      detectVariables: false,
    });
    const first = reopened.package.document.content.at(0);
    const last = reopened.package.document.content.at(1);
    expect(
      first?.type === "paragraph" &&
        first.content.some((run) => run.type === "run" && run.formatting?.bold),
    ).toBe(true);
    expect(
      last?.type === "paragraph" &&
        last.formatting?.numPr?.kind === "reference" &&
        last.formatting.numPr.ilvl,
    ).toBe(1);
    expect(reopened.package.numbering?.nums).toHaveLength(1);
    expect(await page.evaluate((bytes) => globalThis.__folioCanonical?.load(bytes), saved)).toBe(
      true,
    );
    expect((await snapshot(page))?.text).toBe("abitem");
    expect((await snapshot(page))?.projectionMatchesCanonical).toBe(true);
  }
});

test("canonical structural gestures and toolbar operations preserve each intermediate history state", async ({
  page,
}) => {
  test.setTimeout(90_000);
  const source = await loadedFixture(await createDocx(createEmptyDocument({ initialText: "ab" })));
  for (const port of [reactPort, vuePort]) {
    await page.goto(`http://localhost:${port}/?session=canonical`);
    await page.waitForSelector(".layout-page");
    const reload = async () => {
      await loadReady(page, source);
      await select(page, 2);
    };
    const undoExactly = async (before: Awaited<ReturnType<typeof snapshot>>) => {
      const after = await snapshot(page);
      expect(after?.document).not.toEqual(before?.document);
      expect(after?.projectionMatchesCanonical).toBe(true);
      await page.keyboard.press(`${MODIFIER}+z`);
      const undone = await snapshot(page);
      expect(undone?.document).toEqual(before?.document);
      expect(undone?.selection).toEqual(before?.selection);
      expect(undone?.projectionJSON).toEqual(before?.projectionJSON);
      await page.keyboard.press(`${MODIFIER}+Shift+z`);
      const redone = await snapshot(page);
      expect(redone?.document).toEqual(after?.document);
      expect(redone?.selection).toEqual(after?.selection);
      expect(redone?.projectionJSON).toEqual(after?.projectionJSON);
    };

    for (const shortcut of ["Shift+Enter", `${MODIFIER}+Enter`]) {
      await reload();
      const before = await snapshot(page);
      await clearRefusals(page);
      await page.keyboard.press(shortcut);
      await expectNoRefusals(page);
      const inserted = await snapshot(page);
      expect(inserted?.document?.package.document.content).toHaveLength(1);
      const paragraph = inserted?.document?.package.document.content.at(0);
      if (paragraph?.type !== "paragraph") throw new TypeError("Expected break paragraph.");
      const breaks = paragraph.content.flatMap((run) =>
        run.type === "run" ? run.content.filter((leaf) => leaf.type === "break") : [],
      );
      expect(breaks).toHaveLength(1);
      expect(breaks.at(0)).toMatchObject({
        type: "break",
        breakType: shortcut === "Shift+Enter" ? "textWrapping" : "page",
      });
      await undoExactly(before);
    }

    await reload();
    await page.keyboard.press("Enter");
    await select(page, 2);
    const split = await snapshot(page);
    await page.keyboard.press("Delete");
    expect((await snapshot(page))?.document?.package.document.content).toHaveLength(1);
    await undoExactly(split);

    await reload();
    const beforeShiftTab = await snapshot(page);
    await page.keyboard.press("Shift+Tab");
    expect(await snapshot(page)).toEqual(beforeShiftTab);

    await reload();
    await select(page, 1, 3);
    const beforeToolbar = await snapshot(page);
    await page.getByRole("button", { name: "Bold", exact: true }).click();
    const bold = (await snapshot(page))?.document?.package.document.content.at(0);
    expect(
      bold?.type === "paragraph" &&
        bold.content.some((run) => run.type === "run" && run.formatting?.bold),
    ).toBe(true);
    await undoExactly(beforeToolbar);

    await reload();
    await select(page, 1, 3);
    await page.keyboard.press("Backspace");
    await page.keyboard.type("- ");
    await page.keyboard.press("Tab");
    const nested = await snapshot(page);
    await page.keyboard.press("Shift+Tab");
    const outdented = (await snapshot(page))?.document?.package.document.content.at(0);
    expect(
      outdented?.type === "paragraph" &&
        outdented.formatting?.numPr?.kind === "reference" &&
        (outdented.formatting.numPr.ilvl ?? 0),
    ).toBe(0);
    await undoExactly(nested);
    await page.keyboard.type("item");
    await select(page, 1);
    const beforeListBackspace = await snapshot(page);
    await page.keyboard.press("Backspace");
    const removed = (await snapshot(page))?.document?.package.document.content.at(0);
    expect(removed?.type === "paragraph" && removed.formatting?.numPr?.kind === "reference").toBe(
      false,
    );
    await undoExactly(beforeListBackspace);

    await reload();
    await select(page, 1, 3);
    await page.keyboard.press("Backspace");
    const beforeMarker = await snapshot(page);
    await page.keyboard.type("-");
    const beforeRule = await snapshot(page);
    await page.keyboard.type(" ");
    const ruled = (await snapshot(page))?.document?.package.document.content.at(0);
    expect(ruled?.type === "paragraph" && ruled.formatting?.numPr?.kind).toBe("reference");
    await page.keyboard.press("Backspace");
    const ruleUndone = await snapshot(page);
    expect(ruleUndone?.document).toEqual(beforeRule?.document);
    expect(ruleUndone?.selection).toEqual(beforeRule?.selection);
    expect(ruleUndone?.text).toBe("-");
    await page.keyboard.press(`${MODIFIER}+z`);
    expect((await snapshot(page))?.document).toEqual(beforeMarker?.document);
    expect((await snapshot(page))?.selection).toEqual(beforeMarker?.selection);

    await reload();
    await page.keyboard.type("x");
    const beforeFormatting = await snapshot(page);
    await select(page, 1, 4);
    const formattingSelection = await snapshot(page);
    await page.keyboard.press(`${MODIFIER}+i`);
    await select(page, 4);
    const beforeSecondTyping = await snapshot(page);
    await page.keyboard.type("y");
    await page.keyboard.press(`${MODIFIER}+z`);
    expect((await snapshot(page))?.document).toEqual(beforeSecondTyping?.document);
    expect((await snapshot(page))?.selection).toEqual(beforeSecondTyping?.selection);
    await page.keyboard.press(`${MODIFIER}+z`);
    expect((await snapshot(page))?.document).toEqual(beforeFormatting?.document);
    expect((await snapshot(page))?.selection).toEqual(formattingSelection?.selection);
  }
});

test("canonical suggestions survive save and reopen before acceptance or rejection", async ({
  page,
}) => {
  test.setTimeout(60_000);
  const sourceDocument = createEmptyDocument({ initialText: "A😀B" });
  sourceDocument.package.document.content.push(
    ...createEmptyDocument({ initialText: "C" }).package.document.content,
  );
  const source = await createDocx(sourceDocument);
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
    expect(await page.evaluate(() => globalThis.__folioCanonical?.setMode("suggesting"))).toBe(
      true,
    );
    await select(page, 5);
    await page.keyboard.type("x");
    await expectProjection(page, { text: "A😀BxC", selection: { from: 6, to: 6 } });
    await select(page, 2);
    await page.keyboard.press("Delete");
    const typed = await snapshot(page);
    const paragraph = typed?.document?.package.document.content.at(0);
    if (!paragraph || paragraph.type !== "paragraph")
      throw new TypeError("Expected suggested paragraph.");
    expect(paragraph.content.map(({ type }) => type)).toEqual(
      expect.arrayContaining(["insertion", "deletion"]),
    );

    await select(page, 2);
    await page.keyboard.press("Enter");
    const split = await snapshot(page);
    expect(split?.document?.package.document.content).toHaveLength(3);
    expect(split?.projectionMatchesCanonical).toBe(true);
    // Join the original final paragraph to the split tail, a different paragraph mark.
    await select(page, 10);
    await page.keyboard.press("Backspace");
    const suggested = await snapshot(page);
    if (!suggested?.document) throw new Error("Canonical suggestions unavailable.");
    expect(suggested.projectionMatchesCanonical).toBe(true);
    expect(suggested.projectionJSON).toEqual(suggested.canonicalProjectionJSON);
    expect(suggested.document.package.document.content).toHaveLength(3);
    const revisionPrefix = `${IDENTITY_SPACES.REVISION}:`;
    const revisionIds = identityKeysIn(suggested.document.package.document.content).flatMap(
      (key) => (key.startsWith(revisionPrefix) ? [Number(key.slice(revisionPrefix.length))] : []),
    );
    expect(revisionIds).toHaveLength(4);
    expect(new Set(revisionIds).size).toBe(revisionIds.length);

    const saved = await page.evaluate(() => globalThis.__folioCanonical?.save());
    if (!saved) throw new Error("Canonical suggestions did not save.");
    const reopened = structuredClone(
      await parseDocx(new Uint8Array(saved), {
        preloadFonts: false,
        detectVariables: false,
      }),
    );
    expect(reopened.package.document.content).toEqual(suggested.document.package.document.content);

    for (const decision of ["accept", "reject"] as const) {
      expect(
        await page.evaluate(async (bytes) => globalThis.__folioCanonical?.load(bytes), saved),
      ).toBe(true);
      await page.evaluate(
        () => new Promise<void>((resolve) => requestAnimationFrame(() => resolve())),
      );
      const reloaded = await snapshot(page);
      expect(reloaded?.document?.package.document.content).toEqual(
        reopened.package.document.content,
      );
      expect(reloaded?.canUndo).toBe(false);
      expect(
        await page.evaluate(
          ({ ids, resolution }) => globalThis.__folioCanonical?.resolveRevisions(ids, resolution),
          { ids: revisionIds, resolution: decision },
        ),
      ).toBe(true);
      await select(page, 1);
      const resolved = await expectProjection(page, {
        text: decision === "accept" ? "ABxC" : "A😀BC",
        selection: { from: 1, to: 1 },
      });
      expect(identityKeysIn(resolved.document.package.document.content)).toEqual([]);
      expect(resolved.document.package.document.content).toHaveLength(2);
      expect(resolved.document.package.document.content.at(0)).toMatchObject({
        type: "paragraph",
        content: [
          { type: "run", content: [{ type: "text", text: decision === "accept" ? "A" : "A😀B" }] },
        ],
      });
      expect(resolved.canUndo).toBe(true);
      await page.keyboard.press(`${MODIFIER}+z`);
      const undone = await snapshot(page);
      expect(undone?.document?.package.document.content).toEqual(reopened.package.document.content);
      expect(undone?.projectionMatchesCanonical).toBe(true);
      await page.keyboard.press(`${MODIFIER}+Shift+z`);
      const redone = await snapshot(page);
      expect(redone?.document).toEqual(resolved.document);
      expect(redone?.projectionMatchesCanonical).toBe(true);
    }
  }
});
