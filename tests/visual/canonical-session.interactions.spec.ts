import { expect, test, type Page } from "@playwright/test";

import { parseDocx } from "../../packages/core/src/docx/parser";
import { createDocx } from "../../packages/core/src/docx/rezip";
import { createEmptyDocument } from "../../packages/core/src/utils/createDocument";
import { identityKeysIn, IDENTITY_SPACES } from "../../packages/docx-core/src/ops/ids";
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
    const reopened = await parseDocx(new Uint8Array(saved), {
      preloadFonts: false,
      detectVariables: false,
    });
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
