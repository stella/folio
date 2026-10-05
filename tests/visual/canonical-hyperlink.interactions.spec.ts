import { expect, test, type Page } from "@playwright/test";
import { parseDocx } from "../../packages/core/src/docx/parser";
import { createDocx } from "../../packages/core/src/docx/rezip";
import { createEmptyDocument } from "../../packages/core/src/utils/createDocument";
import type { CanonicalHyperlinkAction } from "../parity/canonicalBridge";
import type {} from "../parity/canonicalFuzzErrors";

const MODIFIER = process.platform === "darwin" ? "Meta" : "Control";
const reactPort = Number(process.env["FOLIO_PLAYGROUND_PORT"]) || 4200;
const vuePort = Number(process.env["FOLIO_PLAYGROUND_VUE_PORT"]) || 4201;

const snapshot = (page: Page) => page.evaluate(() => globalThis.__folioCanonical?.snapshot());

const waitForProjection = async (page: Page) => {
  await page.waitForFunction(() => globalThis.__folioCanonical?.canSnapshot());
  const current = await snapshot(page);
  expect(current?.active).toBe(true);
  expect(current?.projectionMatchesCanonical).toBe(true);
  expect(current?.projectionJSON).toEqual(current?.canonicalProjectionJSON);
  if (!current?.document) throw new TypeError("Canonical document unavailable.");
  return { ...current, document: current.document };
};

type SelectOptions = Readonly<{ page: Page; from: number; to: number }>;

const select = async ({ page, from, to }: SelectOptions) => {
  expect(
    await page.evaluate(({ start, end }) => globalThis.__folioCanonical?.select(start, end), {
      start: from,
      end: to,
    }),
  ).toBe(true);
};

const execute = (page: Page, action: CanonicalHyperlinkAction) =>
  page.evaluate((value) => globalThis.__folioCanonical?.executeHyperlink(value) ?? false, action);

type ExpectCommandUndoRedoOptions = Readonly<{
  page: Page;
  action: CanonicalHyperlinkAction;
  before: Awaited<ReturnType<typeof waitForProjection>>;
}>;

const expectCommandUndoRedo = async ({ page, action, before }: ExpectCommandUndoRedoOptions) => {
  expect(await execute(page, action)).toBe(true);
  const after = await waitForProjection(page);
  await page.keyboard.press(`${MODIFIER}+z`);
  const undone = await waitForProjection(page);
  expect(undone.document).toEqual(before.document);
  expect(undone.selectionJSON).toEqual(before.selectionJSON);
  expect(undone.canRedo).toBe(true);
  await page.keyboard.press(`${MODIFIER}+Shift+z`);
  const redone = await waitForProjection(page);
  expect(redone.document).toEqual(after.document);
  expect(redone.selectionJSON).toEqual(after.selectionJSON);
  return redone;
};

for (const { adapter, port } of [
  { adapter: "React", port: reactPort },
  { adapter: "Vue", port: vuePort },
]) {
  test(`canonical hyperlink commands use public dispatch in ${adapter}`, async ({ page }) => {
    test.setTimeout(90_000);
    const sourceDocument = createEmptyDocument({ initialText: "Alpha😀Beta" });
    const paragraph = sourceDocument.package.document.content.at(0);
    if (!paragraph || paragraph.type !== "paragraph") {
      throw new TypeError("Expected the hyperlink fixture paragraph.");
    }
    paragraph.paraId = "12345678";
    const source = await createDocx(sourceDocument);
    await page.goto(`http://localhost:${port}/?session=canonical`);
    await page.waitForSelector(".layout-page");
    await page.evaluate(() => {
      globalThis.__folioCanonicalFuzzErrors = [];
    });
    expect(
      await page.evaluate(
        (bytes) => globalThis.__folioCanonical?.load(bytes),
        [...new Uint8Array(source)],
      ),
    ).toBe(true);
    await page.waitForFunction(() => globalThis.__folioCanonical?.canSnapshot());
    expect(await page.evaluate(() => globalThis.__folioCanonical?.setMode("editing"))).toBe(true);

    await select({ page, from: 1, to: 6 });
    const beforeSet = await waitForProjection(page);
    const afterSet = await expectCommandUndoRedo({
      page,
      action: {
        type: "setHyperlink",
        from: 1,
        to: 6,
        href: "https://set.example/path",
        tooltip: "Set",
      },
      before: beforeSet,
    });
    const linkedParagraph = afterSet.document.package.document.content.at(0);
    if (!linkedParagraph || linkedParagraph.type !== "paragraph") {
      throw new TypeError("Expected the edited paragraph.");
    }
    expect(linkedParagraph.content).toContainEqual(
      expect.objectContaining({
        type: "hyperlink",
        href: "https://set.example/path",
        tooltip: "Set",
      }),
    );

    const firstSave = await page.evaluate(() => globalThis.__folioCanonical?.save());
    if (!firstSave) throw new TypeError("Expected the set-link document to save.");
    const firstReopen = await parseDocx(new Uint8Array(firstSave), {
      preloadFonts: false,
      detectVariables: false,
    });
    const firstReopenedParagraph = firstReopen.package.document.content.at(0);
    if (!firstReopenedParagraph || firstReopenedParagraph.type !== "paragraph") {
      throw new TypeError("Expected a reopened paragraph.");
    }
    expect(firstReopenedParagraph.content).toContainEqual(
      expect.objectContaining({ type: "hyperlink", href: "https://set.example/path" }),
    );

    await select({ page, from: 1, to: 6 });
    const beforeRemove = await waitForProjection(page);
    const afterRemove = await expectCommandUndoRedo({
      page,
      action: { type: "removeHyperlink", from: 1, to: 6 },
      before: beforeRemove,
    });
    const unlinkedParagraph = afterRemove.document.package.document.content.at(0);
    if (!unlinkedParagraph || unlinkedParagraph.type !== "paragraph") {
      throw new TypeError("Expected the unlinked paragraph.");
    }
    expect(unlinkedParagraph.content).not.toContainEqual(
      expect.objectContaining({ type: "hyperlink", href: "https://set.example/path" }),
    );

    await select({ page, from: 6, to: 8 });
    const beforeInsert = await waitForProjection(page);
    const afterInsert = await expectCommandUndoRedo({
      page,
      action: {
        type: "insertHyperlink",
        from: 6,
        to: 8,
        text: "Link",
        href: "https://insert.example/path",
        tooltip: "Inserted",
      },
      before: beforeInsert,
    });
    expect(afterInsert.text).toBe("AlphaLinkBeta");

    const finalSave = await page.evaluate(() => globalThis.__folioCanonical?.save());
    if (!finalSave) throw new TypeError("Expected the inserted-link document to save.");
    const finalReopen = await parseDocx(new Uint8Array(finalSave), {
      preloadFonts: false,
      detectVariables: false,
    });
    const finalParagraph = finalReopen.package.document.content.at(0);
    if (!finalParagraph || finalParagraph.type !== "paragraph") {
      throw new TypeError("Expected a reopened paragraph.");
    }
    expect(finalParagraph.content).toContainEqual(
      expect.objectContaining({
        type: "hyperlink",
        href: "https://insert.example/path",
        tooltip: "Inserted",
      }),
    );

    expect(await page.evaluate(() => globalThis.__folioCanonical?.setMode("suggesting"))).toBe(
      true,
    );
    const refusalActions = [
      {
        type: "setHyperlink",
        from: 6,
        to: 10,
        href: "https://refused.example/path",
      },
      { type: "removeHyperlink", from: 6, to: 10 },
      {
        type: "insertHyperlink",
        from: 6,
        to: 10,
        text: "Replacement",
        href: "https://refused.example/path",
      },
    ] as const satisfies readonly CanonicalHyperlinkAction[];
    for (const action of refusalActions) {
      await select({ page, from: action.from, to: action.to });
      await page.evaluate(() => {
        globalThis.__folioCanonicalFuzzErrors = [];
      });
      const beforeRefusal = await waitForProjection(page);
      expect(await execute(page, action)).toBe(false);
      const refused = await waitForProjection(page);
      expect(refused.document).toEqual(beforeRefusal.document);
      expect(refused.projectionJSON).toEqual(beforeRefusal.projectionJSON);
      expect(refused.selectionJSON).toEqual(beforeRefusal.selectionJSON);
      expect(refused.canUndo).toBe(beforeRefusal.canUndo);
      expect(refused.canRedo).toBe(beforeRefusal.canRedo);
      if (adapter === "React") {
        expect(await page.evaluate(() => globalThis.__folioCanonicalFuzzErrors?.splice(0))).toEqual(
          [
            expect.objectContaining({
              type: "CanonicalSessionRefusalError",
              gap: "tracked-hyperlink-resolution",
            }),
          ],
        );
      }
    }
  });
}
