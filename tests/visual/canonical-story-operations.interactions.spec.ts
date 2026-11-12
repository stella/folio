import { expect, type Page, test } from "@playwright/test";

import { parseDocx } from "../../packages/core/src/docx/parser";
import { createDocx } from "../../packages/core/src/docx/rezip";
import type { Document } from "../../packages/core/src/types/document";
import { createEmptyDocument } from "../../packages/core/src/utils/createDocument";
import { createEmptyHeaderFooter } from "../../packages/core/src/utils/headerFooter";
import type {} from "../parity/canonicalBridge";
import { evaluateCanonicalLoad } from "../parity/canonicalLoadEvaluation";

const adapters = [
  { name: "React", port: Number(process.env["FOLIO_PLAYGROUND_PORT"]) || 4200 },
  {
    name: "Vue",
    port: Number(process.env["FOLIO_PLAYGROUND_VUE_PORT"]) || 4201,
  },
] as const;

/** A body with a footnote reference, a header and a footnote, each with text. */
const fixture = async (): Promise<{ bytes: number[]; headerRId: string }> => {
  const source = createEmptyDocument({ initialText: "Body seed" });
  const document = createEmptyHeaderFooter(source, "header", false);
  if (!document) throw new TypeError("Could not create header fixture.");
  const headers = document.package.headers;
  const headerRId = [...(headers?.keys() ?? [])].at(0);
  const header = headerRId === undefined ? undefined : headers?.get(headerRId);
  if (!headers || headerRId === undefined || !header)
    throw new TypeError("Missing header fixture.");
  headers.set(headerRId, {
    ...header,
    content: [
      {
        type: "paragraph",
        content: [
          { type: "run", content: [{ type: "text", text: "Header seed" }] },
        ],
      },
    ],
  });
  const body = document.package.document.content.at(0);
  if (body?.type !== "paragraph") throw new TypeError("Expected a paragraph.");
  body.content.push({ type: "run", content: [{ type: "footnoteRef", id: 1 }] });
  document.package.footnotes = [
    {
      type: "footnote",
      id: 1,
      content: [
        {
          type: "paragraph",
          content: [
            { type: "run", content: [{ type: "text", text: "Footnote seed" }] },
          ],
        },
      ],
    },
  ];
  return { bytes: [...new Uint8Array(await createDocx(document))], headerRId };
};

const canonicalDocument = async (page: Page): Promise<Document> => {
  await page.waitForFunction(() => globalThis.__folioCanonical?.canSnapshot());
  const current = await page.evaluate(() =>
    globalThis.__folioCanonical?.snapshot(),
  );
  expect(current?.projectionMatchesCanonical).toBe(true);
  if (!current?.document)
    throw new TypeError("Canonical document unavailable.");
  return current.document;
};

const textOf = (blocks: readonly unknown[] | undefined): string =>
  JSON.stringify(blocks ?? [])
    .match(/"text":"[^"]*"/gu)
    ?.map((entry) => entry.slice(8, -1))
    .join("") ?? "";

// A header part map does not cross into the test, so it is read in the page.
const headerText = async (page: Page, rId: string) =>
  textOf(
    await page.evaluate(
      (id) =>
        globalThis.__folioCanonical
          ?.snapshot()
          .document?.package.headers?.get(id)?.content,
      rId,
    ),
  );
const footnoteText = async (page: Page) =>
  textOf(
    (await canonicalDocument(page)).package.footnotes?.find(
      ({ id }) => id === 1,
    )?.content,
  );

for (const adapter of adapters) {
  test(`public operations edit a header and a footnote from any focus, undo per story and save in ${adapter.name}`, async ({
    page,
  }) => {
    test.setTimeout(90_000);
    const { bytes, headerRId } = await fixture();
    await page.goto(`http://localhost:${adapter.port}/?session=canonical`);
    await page.waitForSelector(".layout-page");
    expect(await evaluateCanonicalLoad(page, bytes)).toBe(true);
    const header = { kind: "header", rId: headerRId } as const;
    const footnote = { kind: "footnote", id: 1 } as const;

    // A header edit through the API while the body has focus.
    await page.locator(".layout-paragraph").first().click();
    expect(
      await page.evaluate(
        (story) =>
          globalThis.__folioCanonical?.replaceStoryBlock(
            story,
            "Header by API",
          ),
        header,
      ),
    ).toEqual({ applied: 1, skipped: [] });
    expect(await headerText(page, headerRId)).toBe("Header by API");
    const paintedHeader = page.locator(".layout-page-header").first();
    await expect(paintedHeader).toContainText("Header by API");
    expect(
      textOf((await canonicalDocument(page)).package.document.content),
    ).toBe("Body seed");

    // The header editor opens on the edited header; an API edit to it shows there.
    await paintedHeader.dblclick();
    await expect(page.locator(".hf-inline-editor")).toBeVisible();
    const headerEditor = page.locator(
      `[data-hf-r-id="${headerRId}"] .ProseMirror`,
    );
    await expect(headerEditor).toContainText("Header by API");
    expect(
      await page.evaluate(
        (story) =>
          globalThis.__folioCanonical?.replaceStoryBlock(story, "Header twice"),
        header,
      ),
    ).toEqual({ applied: 1, skipped: [] });
    await expect(headerEditor).toContainText("Header twice");

    // A footnote edit through the API while the header editor is open.
    expect(
      await page.evaluate(
        (story) =>
          globalThis.__folioCanonical?.replaceStoryBlock(
            story,
            "Footnote by API",
          ),
        footnote,
      ),
    ).toEqual({ applied: 1, skipped: [] });
    expect(await footnoteText(page)).toBe("Footnote by API");
    await expect(headerEditor).toContainText("Header twice");

    // Undo takes back the footnote edit alone, then the header edit alone.
    expect(
      await page.evaluate(() => globalThis.__folioCanonical?.undoStoryBlock()),
    ).toBe("undone");
    expect(await footnoteText(page)).toBe("Footnote seed");
    expect(await headerText(page, headerRId)).toBe("Header twice");
    expect(
      await page.evaluate(
        (story) =>
          globalThis.__folioCanonical?.replaceStoryBlock(
            story,
            "Footnote by API",
          ),
        footnote,
      ),
    ).toEqual({ applied: 1, skipped: [] });

    // The saved file carries both edits.
    const saved = await page.evaluate(() =>
      globalThis.__folioCanonical?.save(),
    );
    if (!saved) throw new TypeError("Canonical save unavailable.");
    const reopened = await parseDocx(new Uint8Array(saved).buffer, {
      preloadFonts: false,
      detectVariables: false,
    });
    expect(textOf(reopened.package.headers?.get(headerRId)?.content)).toBe(
      "Header twice",
    );
    expect(
      textOf(reopened.package.footnotes?.find(({ id }) => id === 1)?.content),
    ).toBe("Footnote by API");
  });
}
