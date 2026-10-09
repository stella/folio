import { expect, type Page } from "@playwright/test";
import type { canonicalLoadFixture } from "./canonicalLoadFixture";

export const MODIFIER = process.platform === "darwin" ? "Meta" : "Control";
export const reactPort = Number(process.env["FOLIO_PLAYGROUND_PORT"]) || 4200;
export const vuePort = Number(process.env["FOLIO_PLAYGROUND_VUE_PORT"]) || 4201;

export const snapshot = (page: Page) =>
  page.evaluate(() => globalThis.__folioCanonical?.snapshot());

export const clearRefusals = (page: Page) =>
  page.evaluate(() => {
    globalThis.__folioCanonicalFuzzErrors = [];
  });

export const expectNoRefusals = async (page: Page) =>
  expect(await page.evaluate(() => globalThis.__folioCanonicalFuzzErrors)).toEqual([]);

export const select = async (page: Page, anchor: number, head = anchor) => {
  expect(
    await page.evaluate(({ from, to }) => globalThis.__folioCanonical?.select(from, to), {
      from: anchor,
      to: head,
    }),
  ).toBe(true);
};

type LoadedFixture = Awaited<ReturnType<typeof canonicalLoadFixture>>;

export const loadReady = async (page: Page, source: LoadedFixture) => {
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
