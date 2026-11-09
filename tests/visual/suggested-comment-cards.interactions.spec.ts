import { expect, type Page } from "@playwright/test";

import { createDocx } from "../../packages/core/src/docx/rezip";
import type { Document } from "../../packages/core/src/types/document";
import { reactPort, vuePort } from "../parity/canonicalInteractionAssertions";
import { test } from "../parity/canonicalTest";

/**
 * A comment a suggestion proposes shows as a card while the suggestion is
 * pending; rejecting the suggestion takes the card away at once, and
 * accepting keeps it, in both hosts, in the canonical session.
 */

const source = (): Document => ({
  package: {
    document: {
      content: [
        {
          type: "paragraph",
          paraId: "10000000",
          content: [
            {
              type: "run",
              content: [{ type: "text", text: "Payment is due in ten days." }],
            },
          ],
        },
      ],
    },
  },
});

const HOSTS = [
  { name: "React", port: reactPort, card: ".docx-comment-card" },
  { name: "Vue", port: vuePort, card: ".comment-card" },
] as const;

const load = async (page: Page, bytes: number[]) => {
  expect(
    await page.evaluate(
      (data) => globalThis.__folioCanonical?.loadAnySession(data),
      bytes,
    ),
  ).toBe(true);
  await expect(page.locator(".layout-page").first()).toContainText(
    "Payment is due",
  );
};

const proposeComment = (page: Page) =>
  page.evaluate(() =>
    globalThis.__folioCanonical?.suggestBlockComment(
      "proposal",
      "10000000",
      "Confirm the period.",
    ),
  );

const resolve = (page: Page, resolution: "accept" | "reject") =>
  page.evaluate(
    (decision) =>
      globalThis.__folioCanonical?.resolveSuggestion("proposal", decision),
    resolution,
  );

const suggestionIds = (page: Page) =>
  page.evaluate(() => globalThis.__folioCanonical?.suggestionIds());

test("a proposed comment's card goes with a rejected suggestion and stays with an accepted one", async ({
  page,
}) => {
  test.setTimeout(120_000);
  const bytes = [...new Uint8Array(await createDocx(source()))];
  for (const host of HOSTS)
    for (const resolution of ["reject", "accept"] as const) {
      const label = `${host.name}, ${resolution}`;
      await page.goto(`http://localhost:${host.port}/?session=canonical`);
      await page.waitForSelector(".layout-page");
      await load(page, bytes);
      expect(await proposeComment(page), label).toEqual({
        applied: ["proposal"],
        skipped: [],
      });
      expect(await suggestionIds(page), label).toEqual(["proposal"]);
      const card = page.locator(host.card, {
        hasText: "Confirm the period.",
      });
      await expect(card, label).toHaveCount(1);
      expect(await resolve(page, resolution), label).toBe(true);
      expect(await suggestionIds(page), label).toEqual([]);
      // At once: no save or reload between the decision and the sidebar.
      await expect(card, label).toHaveCount(resolution === "reject" ? 0 : 1, {
        timeout: 2_000,
      });
    }
});
