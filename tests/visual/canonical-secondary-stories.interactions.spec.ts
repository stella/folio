import { expect, test, type Page } from "@playwright/test";
import { createDocx } from "../../packages/core/src/docx/rezip";
import { parseDocx } from "../../packages/core/src/docx/parser";
import { createEmptyDocument } from "../../packages/core/src/utils/createDocument";
import { reviewDifferences } from "../../test/reviewDifferences";
import type {} from "../parity/canonicalBridge";

const adapters = [
  { name: "React", port: Number(process.env["FOLIO_PLAYGROUND_PORT"]) || 4200 },
  { name: "Vue", port: Number(process.env["FOLIO_PLAYGROUND_VUE_PORT"]) || 4201 },
] as const;

const canonicalSnapshot = async (page: Page) => {
  await page.waitForFunction(() => globalThis.__folioCanonical?.canSnapshot());
  const current = await page.evaluate(() => globalThis.__folioCanonical?.snapshot());
  expect(current?.active).toBe(true);
  expect(current?.projectionMatchesCanonical).toBe(true);
  if (!current?.document) throw new TypeError("Canonical document unavailable.");
  return { ...current, document: current.document };
};

const hasInsertionBy = (value: unknown, author: string): boolean => {
  if (Array.isArray(value)) return value.some((entry) => hasInsertionBy(entry, author));
  if (value === null || typeof value !== "object") return false;
  if (
    "type" in value &&
    value.type === "insertion" &&
    "info" in value &&
    value.info !== null &&
    typeof value.info === "object" &&
    "author" in value.info &&
    value.info.author === author
  ) {
    return true;
  }
  return Object.values(value).some((entry) => hasInsertionBy(entry, author));
};

const makeNoteDocument = async () => {
  const document = createEmptyDocument({ initialText: "Body seed" });
  const body = document.package.document.content.at(0);
  if (body?.type !== "paragraph") throw new TypeError("Expected body paragraph.");
  body.content.push(
    { type: "run", content: [{ type: "footnoteRef", id: 1 }] },
    {
      type: "run",
      content: [
        { type: "text", text: " " },
        { type: "endnoteRef", id: 1 },
      ],
    },
  );
  document.package.footnotes = [
    {
      type: "footnote",
      id: 1,
      content: [
        {
          type: "paragraph",
          content: [{ type: "run", content: [{ type: "text", text: "Footnote seed" }] }],
        },
      ],
    },
  ];
  document.package.endnotes = [
    {
      type: "endnote",
      id: 1,
      content: [
        {
          type: "paragraph",
          content: [{ type: "run", content: [{ type: "text", text: "Endnote seed" }] }],
        },
      ],
    },
  ];
  return createDocx(document);
};

for (const adapter of adapters) {
  test(`canonical note stories support native editing, suggestions, history and save in ${adapter.name}`, async ({
    page,
  }) => {
    test.setTimeout(90_000);
    const bytes = [...new Uint8Array(await makeNoteDocument())];
    await page.goto(`http://localhost:${adapter.port}/?session=canonical`);
    await page.waitForSelector(".layout-page");
    expect(await page.evaluate((input) => globalThis.__folioCanonical?.load(input), bytes)).toBe(
      true,
    );
    const initial = await canonicalSnapshot(page);
    const initialBody = initial.document.package.document.content;
    const initialFootnote = initial.document.package.footnotes?.find(({ id }) => id === 1)?.content;
    const initialEndnote = initial.document.package.endnotes?.find(({ id }) => id === 1)?.content;
    if (!initialFootnote || !initialEndnote) throw new TypeError("Note fixture did not load.");

    const body = page.locator(".layout-paragraph").first();
    await body.click();
    await page.keyboard.press("End");
    await page.keyboard.type(" edited");
    const bodyEdited = (await canonicalSnapshot(page)).document.package.document.content;
    expect(bodyEdited).not.toEqual(initialBody);
    await page.getByRole("button", { name: "Undo", exact: true }).click();
    expect((await canonicalSnapshot(page)).document.package.document.content).toEqual(initialBody);
    await page.getByRole("button", { name: "Redo", exact: true }).click();
    expect((await canonicalSnapshot(page)).document.package.document.content).toEqual(bodyEdited);

    const footnoteRef = page.locator('[data-note-kind="footnote"][data-note-id="1"]').first();
    await expect(footnoteRef).toBeVisible();
    await footnoteRef.dblclick();
    const noteEditor = page.locator(
      'aside[aria-label="Footnotes"], aside[aria-label="Endnotes"], .docx-editor-vue__note-editor',
    );
    await expect(noteEditor).toBeVisible();
    const noteProseMirror = noteEditor.locator(
      '[data-note-kind="footnote"][data-note-id="1"] .ProseMirror',
    );
    await expect(noteProseMirror).toBeVisible();
    await noteProseMirror.click();
    await page.keyboard.press("End");
    await page.keyboard.type(" edited");
    await expect(noteProseMirror).toContainText("Footnote seed edited");
    const footnoteEdited = (await canonicalSnapshot(page)).document.package.footnotes?.find(
      ({ id }) => id === 1,
    )?.content;
    expect(footnoteEdited).not.toEqual(initialFootnote);
    await page.getByRole("button", { name: "Undo", exact: true }).click();
    expect(
      (await canonicalSnapshot(page)).document.package.footnotes?.find(({ id }) => id === 1)
        ?.content,
    ).toEqual(initialFootnote);
    await page.getByRole("button", { name: "Redo", exact: true }).click();
    expect(
      (await canonicalSnapshot(page)).document.package.footnotes?.find(({ id }) => id === 1)
        ?.content,
    ).toEqual(footnoteEdited);
    expect((await canonicalSnapshot(page)).document.package.document.content).toEqual(bodyEdited);
    await noteEditor.getByRole("button", { name: "Close", exact: true }).click();

    const trackChanges = page.getByRole("button", {
      name: "Toggle Track Changes",
      exact: true,
    });
    await trackChanges.click();
    await expect(trackChanges).toHaveAttribute("aria-pressed", "true");
    const endnoteRef = page.locator('[data-note-kind="endnote"][data-note-id="1"]').first();
    await expect(endnoteRef).toBeVisible();
    await endnoteRef.dblclick();
    await expect(noteEditor).toBeVisible();
    const endnoteEditor = noteEditor.locator(
      '[data-note-kind="endnote"][data-note-id="1"] .ProseMirror',
    );
    await endnoteEditor.click();
    await page.keyboard.press("End");
    await page.keyboard.type(" suggested");
    await expect(endnoteEditor).toContainText("Endnote seed suggested");
    const endnoteSuggested = (await canonicalSnapshot(page)).document.package.endnotes?.find(
      ({ id }) => id === 1,
    )?.content;
    expect(hasInsertionBy(endnoteSuggested, "Folio User")).toBe(true);
    await page.getByRole("button", { name: "Undo", exact: true }).click();
    expect(
      (await canonicalSnapshot(page)).document.package.endnotes?.find(({ id }) => id === 1)
        ?.content,
    ).toEqual(initialEndnote);
    await page.getByRole("button", { name: "Redo", exact: true }).click();
    expect(
      (await canonicalSnapshot(page)).document.package.endnotes?.find(({ id }) => id === 1)
        ?.content,
    ).toEqual(endnoteSuggested);
    const beforeSave = await canonicalSnapshot(page);
    expect(beforeSave.document.package.document.content).toEqual(bodyEdited);
    expect(beforeSave.document.package.footnotes?.find(({ id }) => id === 1)?.content).toEqual(
      footnoteEdited,
    );

    const saved = await page.evaluate(() => globalThis.__folioCanonical?.save());
    if (!saved) throw new TypeError("Canonical save unavailable.");
    const reopened = await parseDocx(new Uint8Array(saved).buffer, {
      preloadFonts: false,
      detectVariables: false,
    });
    expect(reviewDifferences(beforeSave.document, reopened)).toEqual({ messages: [], omitted: 0 });
    expect(
      hasInsertionBy(reopened.package.endnotes?.find(({ id }) => id === 1)?.content, "Folio User"),
    ).toBe(true);
  });
}
