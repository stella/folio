import { expect, type Page } from "@playwright/test";

import { mintListInstance } from "../../packages/core/src/docx/listNumberingInstances";
import { NO_NUMBERING_IDS_IN_USE } from "../../packages/core/src/docx/numberingIds";
import { parseDocx } from "../../packages/core/src/docx/parser";
import { createDocx } from "../../packages/core/src/docx/rezip";
import type { Paragraph } from "../../packages/core/src/types/document";
import { createEmptyDocument } from "../../packages/core/src/utils/createDocument";
import {
  clearRefusals,
  expectNoRefusals,
  loadReady,
  reactPort,
  vuePort,
} from "../parity/canonicalInteractionAssertions";
import { canonicalLoadFixture } from "../parity/canonicalLoadFixture";
import { test } from "../parity/canonicalTest";

/**
 * Taking a paragraph out of a list with the toolbar, in both hosts, saves the
 * numbering the paragraph stated before: none for a heading whose style does
 * not number it, and the cancellation for a paragraph its style numbers.
 */

const paragraph = (text: string, styleId: string): Paragraph => ({
  type: "paragraph",
  formatting: { styleId },
  content: [{ type: "run", content: [{ type: "text", text }] }],
});

const savedParagraphs = async (page: Page) => {
  const saved = await page.evaluate(async () => {
    const buffer = await globalThis.__folioCanonical?.save();
    return buffer ? [...new Uint8Array(buffer)] : null;
  });
  if (!saved) throw new TypeError("The canonical document did not save.");
  const document = await parseDocx(new Uint8Array(saved), {
    preloadFonts: false,
    detectVariables: false,
  });
  return document.package.document.content.flatMap((block) =>
    block.type === "paragraph" ? [block] : [],
  );
};

const clickToolbarButton = async (page: Page, name: string) => {
  const button = page
    .getByRole("button", { name, exact: true })
    .filter({ visible: true })
    .first();
  const overflow = (await button.count()) === 0;
  if (overflow)
    await page.getByRole("button", { name: "More formatting" }).click();
  await button.click();
  // An overflow popover left open would take the next click in the page.
  if (overflow) await page.keyboard.press("Escape");
};

const clickText = (page: Page, text: string) =>
  page
    .locator(".layout-page .layout-run-text", { hasText: text })
    .first()
    .click();

test("toggling a list off saves the numbering the paragraph stated before", async ({
  page,
}) => {
  test.setTimeout(90_000);
  const minted = mintListInstance(
    undefined,
    { kind: "numbered" },
    NO_NUMBERING_IDS_IN_USE,
  );
  const document = createEmptyDocument();
  document.package.numbering = minted.definitions;
  document.package.styles = {
    ...document.package.styles,
    styles: [
      ...(document.package.styles?.styles ?? []),
      {
        type: "paragraph",
        styleId: "StyledList",
        name: "Styled List",
        pPr: { numPr: { kind: "reference", numId: minted.numId, ilvl: 0 } },
      },
    ],
  };
  document.package.document.content = [
    paragraph("Heading text", "Heading1"),
    paragraph("Styled item", "StyledList"),
  ];
  const source = await canonicalLoadFixture(await createDocx(document));
  for (const [host, port] of [
    ["React", reactPort],
    ["Vue", vuePort],
  ] as const) {
    await page.goto(`http://localhost:${port}/?session=canonical`);
    await page.waitForSelector(".layout-page");
    await clearRefusals(page);
    await loadReady(page, source);
    const [heading, styled] = await savedParagraphs(page);
    expect(heading?.formatting?.numPr, host).toBeUndefined();
    expect(styled?.listRendering, host).toBeDefined();

    // Bullets, then numbers, then numbers again to take the heading out.
    await clickText(page, "Heading text");
    await clickToolbarButton(page, "Bullet List");
    await expect
      .poll(
        async () => (await savedParagraphs(page))[0]?.listRendering?.isBullet,
      )
      .toBe(true);
    await clickToolbarButton(page, "Numbered List");
    await expect
      .poll(
        async () => (await savedParagraphs(page))[0]?.listRendering?.isBullet,
      )
      .toBe(false);
    await clickToolbarButton(page, "Numbered List");
    await expect
      .poll(async () => (await savedParagraphs(page))[0]?.listRendering)
      .toBeUndefined();

    // The paragraph its style numbers states the cancellation instead.
    await clickText(page, "Styled item");
    await clickToolbarButton(page, "Numbered List");
    await expect
      .poll(async () => (await savedParagraphs(page))[1]?.listRendering)
      .toBeUndefined();

    await expectNoRefusals(page);
    const [savedHeading, savedStyled] = await savedParagraphs(page);
    expect([host, savedHeading?.formatting?.numPr]).toEqual([host, undefined]);
    expect(savedHeading?.formatting?.styleId, host).toBe("Heading1");
    expect([host, savedStyled?.formatting?.numPr]).toEqual([
      host,
      { kind: "none" },
    ]);
    expect(savedStyled?.formatting?.styleId, host).toBe("StyledList");
  }
});
