import { expect } from "@playwright/test";
import type { Page } from "@playwright/test";

import { ZipArchive } from "@stll/docx-core/zip";

import { createDocx } from "../../packages/core/src/docx/rezip";
import type { Document } from "../../packages/core/src/types/document";
import {
  clearRefusals,
  expectNoRefusals,
  loadReady,
  reactPort,
  snapshot,
  vuePort,
} from "../parity/canonicalInteractionAssertions";
import { canonicalLoadFixture } from "../parity/canonicalLoadFixture";
import { test } from "../parity/canonicalTest";

const source = (): Document => ({
  package: {
    document: {
      content: [
        {
          type: "paragraph",
          paraId: "10000000",
          content: [
            { type: "run", content: [{ type: "text", text: "alpha beta" }] },
          ],
        },
      ],
    },
  },
});

/** The saved body's text and the authors of its tracked changes. */
const saved = async (page: Page) => {
  const bytes = await page.evaluate(async () => {
    const buffer = await globalThis.__folioCanonical?.save();
    return buffer ? [...new Uint8Array(buffer)] : null;
  });
  if (!bytes) throw new TypeError("The canonical document did not save.");
  const xml =
    (await (
      await ZipArchive.read(new Uint8Array(bytes))
    )
      .entry("word/document.xml")
      ?.text()) ?? "";
  return {
    text: [...xml.matchAll(/<w:t\b[^>]*>([^<]*)</gu)]
      .map(([, t]) => t)
      .join(""),
    revisionAuthors: [
      ...xml.matchAll(/<w:(?:ins|del)\b[^>]*w:author="([^"]*)"/gu),
    ].map(([, author]) => author),
  };
};

const suggest = (page: Page) =>
  page.evaluate(() =>
    globalThis.__folioCanonical?.suggestReplace(
      "s1",
      "10000000",
      "alpha",
      "gamma",
    ),
  );

test("a suggested batch is a pending suggestion that accepts to a tracked change and rejects away", async ({
  page,
}) => {
  test.setTimeout(90_000);
  const fixture = await canonicalLoadFixture(await createDocx(source()));
  for (const port of [reactPort, vuePort]) {
    for (const resolution of ["accept", "reject"] as const) {
      await page.goto(`http://localhost:${port}/?session=canonical`);
      await page.waitForSelector(".layout-page");
      await clearRefusals(page);
      await loadReady(page, fixture);
      expect(await suggest(page)).toEqual({ applied: ["s1"], skipped: [] });
      expect(
        await page.evaluate(() => globalThis.__folioCanonical?.suggestionIds()),
      ).toEqual(["s1"]);
      expect(
        await page.evaluate(() =>
          globalThis.__folioCanonical?.pendingSuggestionCount(),
        ),
      ).toBe(1);
      // The file holds none of a pending suggestion.
      expect(await saved(page)).toEqual({
        text: "alpha beta",
        revisionAuthors: [],
      });
      expect(
        await page.evaluate(
          (decision) =>
            globalThis.__folioCanonical?.resolveSuggestion("s1", decision),
          resolution,
        ),
      ).toBe(true);
      expect(
        await page.evaluate(() => globalThis.__folioCanonical?.suggestionIds()),
      ).toEqual([]);
      const after = await saved(page);
      if (resolution === "reject") {
        expect(after).toEqual({ text: "alpha beta", revisionAuthors: [] });
        expect((await snapshot(page))?.text).toBe("alpha beta");
      } else {
        // Accepted: an ordinary tracked change by the accepting author.
        expect(after.revisionAuthors.length).toBeGreaterThan(0);
        expect(new Set(after.revisionAuthors)).toEqual(new Set(["Adopter"]));
      }
      expect((await snapshot(page))?.projectionMatchesCanonical).toBe(true);
      await expectNoRefusals(page);
    }
  }
});
