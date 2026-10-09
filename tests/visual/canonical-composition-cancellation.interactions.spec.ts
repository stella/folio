import { expect, type Page } from "@playwright/test";
import { test, assertCanonicalInputTimersSettled } from "./canonicalTimerProbe";
import { createDocx } from "../../packages/core/src/docx/rezip";
import { createEmptyDocument } from "../../packages/core/src/utils/createDocument";
import { canonicalLoadFixture } from "../parity/canonicalLoadFixture";
import {
  reactPort,
  vuePort,
  snapshot,
  select,
  loadReady,
  clearRefusals,
  expectNoRefusals,
} from "../parity/canonicalInteractionAssertions";

const expectCancelledSelection = async (
  page: Page,
  before: Awaited<ReturnType<typeof snapshot>>,
) => {
  const after = await snapshot(page);
  expect(after?.document).toEqual(before?.document);
  expect(after?.projectionJSON).toEqual(before?.projectionJSON);
  expect(after?.canonicalProjectionJSON).toEqual(before?.canonicalProjectionJSON);
  expect(after?.projectionMatchesCanonical).toBe(true);
  expect(after?.selection).toEqual(before?.selection);
  expect(after?.selectionJSON).toEqual(before?.selectionJSON);
  expect(after?.textSelection).toEqual(before?.textSelection);
  expect(after?.canUndo).toBe(before?.canUndo);
  expect(after?.canRedo).toBe(before?.canRedo);
  expect(after?.textSelection?.selected).toBe("alpha");
  expect(await page.evaluate(() => window.getSelection()?.toString())).toBe("alpha");
};

test("native composition cancellation preserves a selection", async ({ page }) => {
  const source = await canonicalLoadFixture(
    await createDocx(createEmptyDocument({ initialText: "alpha😀café東京" })),
  );
  for (const { adapter, port } of [
    { adapter: "React", port: reactPort },
    { adapter: "Vue", port: vuePort },
  ]) {
    await page.goto(`http://localhost:${port}/?session=canonical`);
    await page.waitForSelector(".layout-page");
    await loadReady(page, source);
    await assertCanonicalInputTimersSettled(page);
    await select(page, 1, 6);
    await clearRefusals(page);

    const before = await snapshot(page);
    expect(before?.textSelection?.selected, adapter).toBe("alpha");
    expect(before?.selection, adapter).toEqual({ from: 1, to: 6 });
    expect(await page.evaluate(() => window.getSelection()?.toString()), adapter).toBe("alpha");
    expect(before?.canUndo, adapter).toBe(false);
    expect(before?.canRedo, adapter).toBe(false);

    const cdp = await page.context().newCDPSession(page);
    try {
      for (const text of ["shall", "café 東京 é"]) {
        await cdp.send("Input.imeSetComposition", {
          text,
          selectionStart: text.length,
          selectionEnd: text.length,
        });
        await page.waitForFunction((expected) => {
          const editor = document.activeElement;
          return (
            (globalThis.__folioCanonicalFuzzErrors?.length ?? 0) > 0 ||
            (globalThis.__folioCanonical?.nativeComposing() === true &&
              editor?.classList.contains("ProseMirror") === true &&
              editor.textContent?.includes(expected) === true)
          );
        }, text);
        await expectNoRefusals(page);
      }
      await cdp.send("Input.imeSetComposition", {
        text: "",
        selectionStart: 0,
        selectionEnd: 0,
      });
    } finally {
      if (!page.isClosed()) await cdp.detach();
    }

    await page.waitForFunction(() => globalThis.__folioCanonical?.nativeComposing() === false);
    await page.waitForFunction(() => globalThis.__folioCanonicalInputTimers?.size === 0);
    await assertCanonicalInputTimersSettled(page);
    await expectNoRefusals(page);
    await expectCancelledSelection(page, before);
  }
});
