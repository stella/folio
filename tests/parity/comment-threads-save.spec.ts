import type { Page } from "@playwright/test";

import { ZipArchive } from "@stll/docx-core/zip";

import {
  type AdapterFixture,
  ensureLiveView,
  expect,
  forEachAdapter,
} from "./parity-fixture";

/**
 * Comment threads edited in the sidebar are what both hosts save: a comment
 * added, a reply, a resolution and a deleted thread each survive a save and
 * a reopen of the saved package.
 */

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W14 = "http://schemas.microsoft.com/office/word/2010/wordml";
const W15 = "http://schemas.microsoft.com/office/word/2012/wordml";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

/** Two threads: the first with a reply, the second on its own. */
const commentedPackage = async (): Promise<number[]> => {
  const zip = new ZipArchive();
  zip.set(
    "[Content_Types].xml",
    `${HEAD}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/><Override PartName="/word/commentsExtended.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml"/></Types>`,
  );
  zip.set(
    "_rels/.rels",
    `${HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`,
  );
  zip.set(
    "word/_rels/document.xml.rels",
    `${HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId9" Type="${R}/comments" Target="comments.xml"/><Relationship Id="rId10" Type="http://schemas.microsoft.com/office/2011/relationships/commentsExtended" Target="commentsExtended.xml"/></Relationships>`,
  );
  zip.set(
    "word/document.xml",
    `${HEAD}<w:document xmlns:w="${W}"><w:body>` +
      "<w:p><w:r><w:t>Opening words</w:t></w:r></w:p>" +
      '<w:p><w:commentRangeStart w:id="1"/><w:commentRangeStart w:id="2"/><w:r><w:t>Please remit funds.</w:t></w:r><w:commentRangeEnd w:id="1"/><w:r><w:commentReference w:id="1"/></w:r><w:commentRangeEnd w:id="2"/><w:r><w:commentReference w:id="2"/></w:r></w:p>' +
      '<w:p><w:commentRangeStart w:id="3"/><w:r><w:t>Closing words</w:t></w:r><w:commentRangeEnd w:id="3"/><w:r><w:commentReference w:id="3"/></w:r></w:p>' +
      "</w:body></w:document>",
  );
  const comment = (id: number, paraId: string, text: string) =>
    `<w:comment w:id="${id}" w:author="Reviewer" w:date="2024-01-01T00:00:00Z"><w:p w14:paraId="${paraId}"><w:r><w:t>${text}</w:t></w:r></w:p></w:comment>`;
  zip.set(
    "word/comments.xml",
    `${HEAD}<w:comments xmlns:w="${W}" xmlns:w14="${W14}">` +
      comment(1, "10000001", "Check the account") +
      comment(2, "10000002", "Agreed reply text") +
      comment(3, "10000003", "Second thread") +
      "</w:comments>",
  );
  zip.set(
    "word/commentsExtended.xml",
    `${HEAD}<w15:commentsEx xmlns:w15="${W15}"><w15:commentEx w15:paraId="10000001" w15:done="0"/><w15:commentEx w15:paraId="10000002" w15:paraIdParent="10000001" w15:done="0"/><w15:commentEx w15:paraId="10000003" w15:done="0"/></w15:commentsEx>`,
  );
  return [...new Uint8Array(await zip.generate())];
};

type SavedComments = { comments: string; extended: string };

const savedComments = async (bytes: number[]): Promise<SavedComments> => {
  const zip = await ZipArchive.read(Uint8Array.from(bytes).buffer);
  return {
    comments: (await zip.entry("word/comments.xml")?.text()) ?? "",
    extended: (await zip.entry("word/commentsExtended.xml")?.text()) ?? "",
  };
};

const CARD = { react: ".docx-comment-card", vue: ".comment-card" } as const;

const card = (page: Page, adapter: AdapterFixture, id: number | string) =>
  page.locator(`${CARD[adapter.name]}[data-comment-id="${id}"]`);

const load = async (page: Page, bytes: number[]): Promise<void> => {
  await page.evaluate((data) => window.__folioParity?.loadDocx(data), bytes);
  await ensureLiveView(page);
  await expect(page.locator(".layout-page").first()).toContainText(
    "Closing words",
  );
};

const save = async (page: Page): Promise<number[]> => {
  expect(
    await page.evaluate(() => globalThis.__folioCanonicalFuzzErrors ?? []),
  ).toEqual([]);
  return page.evaluate(
    () => window.__folioParity?.saveDocx() ?? Promise.resolve([]),
  );
};

/**
 * Once a suggestion has added a paragraph, a body that holds comment ranges
 * is written whole rather than spliced, even after the suggestion is
 * rejected: the only record such a save may leave is that fallback.
 */
const expectOnlyBodyRewrites = async (page: Page): Promise<void> => {
  const records = await page.evaluate(() =>
    (globalThis.__folioCanonicalFuzzErrors ?? []).splice(0),
  );
  expect(
    records.filter(
      (entry) =>
        !(
          entry.status === "saveDiagnostic" &&
          "diagnostic" in entry &&
          entry.diagnostic.type === "selectiveSaveRefused" &&
          entry.diagnostic.part === "word/document.xml"
        ),
    ),
  ).toEqual([]);
};

/** Save, reopen what was saved, and return the saved comment parts. */
const saveAndReopen = async (page: Page): Promise<SavedComments> => {
  const bytes = await save(page);
  await load(page, bytes);
  return savedComments(bytes);
};

/** Each flow runs in the default session and in the canonical one. */
const SESSIONS = [
  { title: "", query: "" },
  { title: " in the canonical session", query: "/?session=canonical" },
] as const;

const open = async (
  page: Page,
  adapter: AdapterFixture,
  query: string,
): Promise<void> => {
  // A blank editor, so the sidebar opens for the commented package.
  await page.goto(`${adapter.baseUrl}${query}`);
  // Refusals are recorded from here on; each save checks there were none.
  await page.evaluate(() => {
    globalThis.__folioCanonicalFuzzErrors = [];
  });
  await expect
    .poll(() => page.evaluate(() => window.__folioParity !== undefined), {
      timeout: 30_000,
    })
    .toBe(true);
  await load(page, await commentedPackage());
  expect(
    await page.evaluate(
      () => globalThis.__folioCanonical?.snapshot()?.active ?? false,
    ),
  ).toBe(query !== "");
  // Vue's sidebar opens by itself only for the first commented document an
  // editor shows; the blank editor's own document may have used that.
  const opened = await card(page, adapter, 3)
    .waitFor({ state: "visible", timeout: 5000 })
    .then(() => true)
    .catch(() => false);
  if (!opened && adapter.name === "vue")
    await page
      .getByRole("button", { name: "Comments & Changes", exact: true })
      .click();
  await expect(card(page, adapter, 3)).toBeVisible();
};

/** Undo in the document, as a user does after editing a thread. */
const undoInDocument = async (page: Page): Promise<void> => {
  await page.locator(".layout-page").first().getByText("Opening words").click();
  await page.keyboard.press("ControlOrMeta+z");
};

for (const session of SESSIONS) {
  forEachAdapter(
    `comment threads: a comment added in the sidebar is saved${session.title}`,
    async (adapter, { page }) => {
      await open(page, adapter, session.query);
      await page
        .locator(".layout-page")
        .first()
        .getByText("Opening words")
        .dblclick();
      await page
        .getByRole("button", { name: "Add comment" })
        .dispatchEvent("mousedown");
      const field = page.getByPlaceholder(/^Add a comment/u);
      await field.fill("Brand new note");
      if (adapter.name === "react") await field.press("Enter");
      else
        await page
          .getByRole("button", { name: "Comment", exact: true })
          .click();
      await expect(
        page.locator(CARD[adapter.name]).filter({ hasText: "Brand new note" }),
      ).toHaveCount(1);

      const saved = await saveAndReopen(page);
      expect(saved.comments).toContain("Brand new note");
      await expect(
        page.locator(CARD[adapter.name]).filter({ hasText: "Brand new note" }),
      ).toHaveCount(1);
    },
  );

  forEachAdapter(
    `comment threads: a deleted thread and its reply stay deleted after a save${session.title}`,
    async (adapter, { page }) => {
      await open(page, adapter, session.query);
      const thread = card(page, adapter, 1);
      await thread.click();
      await thread.getByTitle("More options").click();
      await page.getByRole("button", { name: "Delete", exact: true }).click();
      await expect(thread).toHaveCount(0);

      const saved = await saveAndReopen(page);
      expect(saved.comments).not.toContain("Check the account");
      expect(saved.comments).not.toContain("Agreed reply text");
      expect(saved.comments).toContain("Second thread");
      await expect(card(page, adapter, 1)).toHaveCount(0);
      await expect(card(page, adapter, 3)).toBeVisible();
    },
  );

  forEachAdapter(
    `comment threads: a reply is saved under its thread${session.title}`,
    async (adapter, { page }) => {
      await open(page, adapter, session.query);
      const thread = card(page, adapter, 3);
      await thread.click();
      await thread
        .getByPlaceholder(/^Reply/u)
        .first()
        .click();
      await thread
        .getByPlaceholder(/^Reply/u)
        .last()
        .fill("A fresh reply");
      await thread.getByRole("button", { name: "Reply", exact: true }).click();
      await expect(thread).toContainText("A fresh reply");

      const saved = await saveAndReopen(page);
      expect(saved.comments).toContain("A fresh reply");
      // The reply is threaded under the second thread's paragraph.
      expect(saved.extended).toMatch(/w15:paraIdParent="10000003"/u);
    },
  );

  forEachAdapter(
    `comment threads: a resolved thread is saved resolved${session.title}`,
    async (adapter, { page }) => {
      await open(page, adapter, session.query);
      const thread = card(page, adapter, 3);
      await thread.click();
      await thread.getByTitle("Resolve").click();

      const saved = await saveAndReopen(page);
      expect(saved.extended).toMatch(
        /w15:paraId="10000003"[^>]*w15:done="1"|w15:done="1"[^>]*w15:paraId="10000003"/u,
      );
    },
  );

  forEachAdapter(
    `comment threads: undo brings back a deleted thread and its reply, and the save keeps them${session.title}`,
    async (adapter, { page }) => {
      await open(page, adapter, session.query);
      const thread = card(page, adapter, 1);
      await thread.click();
      await thread.getByTitle("More options").click();
      await page.getByRole("button", { name: "Delete", exact: true }).click();
      await expect(thread).toHaveCount(0);

      await undoInDocument(page);
      await expect(card(page, adapter, 1)).toBeVisible();
      await expect(card(page, adapter, 1)).toContainText("Agreed reply text");

      const saved = await saveAndReopen(page);
      expect(saved.comments).toContain("Check the account");
      expect(saved.comments).toContain("Agreed reply text");
      expect(saved.extended).toMatch(/w15:paraIdParent="10000001"/u);
      await expect(card(page, adapter, 1)).toBeVisible();
      await expect(card(page, adapter, 1)).toContainText("Agreed reply text");
    },
  );

  forEachAdapter(
    `comment threads: undo reopens a resolved thread, and the save keeps it open${session.title}`,
    async (adapter, { page }) => {
      await open(page, adapter, session.query);
      const thread = card(page, adapter, 3);
      await thread.click();
      await thread.getByTitle("Resolve").click();

      await undoInDocument(page);
      const saved = await saveAndReopen(page);
      expect(saved.extended).toMatch(
        /w15:paraId="10000003"[^>]*w15:done="0"|w15:done="0"[^>]*w15:paraId="10000003"/u,
      );
      await expect(card(page, adapter, 3)).toBeVisible();
    },
  );

  forEachAdapter(
    `comment threads: rejecting a suggested paragraph takes its comment with it, and undo brings both back${session.title}`,
    async (adapter, { page }) => {
      await open(page, adapter, session.query);
      const suggestionId = await page.evaluate(
        () =>
          window.__folioParity?.suggestCommentedInsertion(
            "Closing",
            "A suggested last line.",
            "Suggested wording note",
          ) ?? null,
      );
      expect(suggestionId).not.toBeNull();
      const suggested = page
        .locator(CARD[adapter.name])
        .filter({ hasText: "Suggested wording note" });
      await expect(suggested).toHaveCount(1);
      // A pending suggestion's comment is held back with it.
      const pending = await savedComments(await save(page));
      await expectOnlyBodyRewrites(page);
      expect(pending.comments).not.toContain("Suggested wording note");
      expect(pending.comments).toContain("Second thread");

      await page.evaluate(
        (id) => window.__folioParity?.rejectSuggestion(id ?? ""),
        suggestionId,
      );
      await expect(suggested).toHaveCount(0);
      const rejected = await savedComments(await save(page));
      await expectOnlyBodyRewrites(page);
      expect(rejected.comments).not.toContain("Suggested wording note");
      expect(rejected.comments).toContain("Second thread");

      // A pending suggestion is not part of the saved package; undo brings the
      // suggestion back with its comment.
      await undoInDocument(page);
      await expect(suggested).toHaveCount(1);
      expect(
        await page.evaluate(() => globalThis.__folioCanonicalFuzzErrors ?? []),
      ).toEqual([]);
    },
  );
}
