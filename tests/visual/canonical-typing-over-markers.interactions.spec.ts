/**
 * Typing over a selection that holds comment, bookmark and permission markers
 * keeps the typed characters together and in order, in editing and suggesting
 * mode, in both adapters.
 */
import { expect, type Page } from "@playwright/test";

import { DOCUMENT_SHAPES } from "../../packages/core/src/__tests__/documentShapes";
import {
  clearRefusals,
  expectNoRefusals,
  loadReady,
  MODIFIER,
  reactPort,
  snapshot,
  vuePort,
} from "../parity/canonicalInteractionAssertions";
import { canonicalLoadFixture } from "../parity/canonicalLoadFixture";
import { test } from "../parity/canonicalTest";

const shape = DOCUMENT_SHAPES.find(({ id }) => id === "invisible-markers");
if (shape === undefined) throw new TypeError("Missing the markers shape.");

/** A node's text; with `withDeleted`, tracked deletions included in place. */
const textOf = (node: unknown, withDeleted: boolean): string => {
  if (typeof node !== "object" || node === null) return "";
  const { type, text, content } = node as {
    type?: unknown;
    text?: unknown;
    content?: unknown;
  };
  if (!withDeleted && (type === "deletion" || type === "moveFrom")) return "";
  if (type === "text" && typeof text === "string") return text;
  return Array.isArray(content)
    ? content.map((child) => textOf(child, withDeleted)).join("")
    : "";
};

/** The main story's paragraphs, in order. */
const paragraphs = async (page: Page) => {
  const current = await snapshot(page);
  return (current?.document?.package.document.content ?? []).filter(
    (block) => block.type === "paragraph",
  );
};

for (const [adapter, port] of [
  ["React", reactPort],
  ["Vue", vuePort],
] as const)
  for (const mode of ["editing", "suggesting"] as const)
    test(`${adapter}, ${mode}: text typed over the whole document stays in order`, async ({
      page,
    }) => {
      test.setTimeout(60_000);
      const source = await canonicalLoadFixture(await shape.build());
      await page.goto(`http://localhost:${port}/?session=canonical`);
      await page.waitForSelector(".layout-page");
      await clearRefusals(page);
      await loadReady(page, source);
      if (mode === "suggesting")
        await page
          .getByRole("button", { name: "Toggle Track Changes" })
          .click();
      await page.locator(".layout-page").first().click();
      await page.keyboard.press(`${MODIFIER}+a`);
      await page.keyboard.type("xyz");
      await expectNoRefusals(page);
      const story = await paragraphs(page);
      expect(story.map((block) => textOf(block, false)).join("")).toBe("xyz");
      // Suggesting, the typed run stays whole after the struck text, as a
      // redline reads.
      expect(textOf(story.at(-1), true).endsWith("xyz")).toBe(true);
    });
