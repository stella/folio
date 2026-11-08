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

const save = (page: Page): Promise<number[]> =>
  page.evaluate(() => window.__folioParity?.saveDocx() ?? Promise.resolve([]));

/** Save, reopen what was saved, and return the saved comment parts. */
const saveAndReopen = async (page: Page): Promise<SavedComments> => {
  const bytes = await save(page);
  await load(page, bytes);
  return savedComments(bytes);
};

const open = async (page: Page, adapter: AdapterFixture): Promise<void> => {
  // A blank editor, so the sidebar opens for the commented package.
  await page.goto(adapter.baseUrl);
  await expect
    .poll(() => page.evaluate(() => window.__folioParity !== undefined), {
      timeout: 30_000,
    })
    .toBe(true);
  await load(page, await commentedPackage());
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

forEachAdapter(
  "comment threads: a comment added in the sidebar is saved",
  async (adapter, { page }) => {
    await open(page, adapter);
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
      await page.getByRole("button", { name: "Comment", exact: true }).click();
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
  "comment threads: a deleted thread and its reply stay deleted after a save",
  async (adapter, { page }) => {
    await open(page, adapter);
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
  "comment threads: a reply is saved under its thread",
  async (adapter, { page }) => {
    await open(page, adapter);
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
  "comment threads: a resolved thread is saved resolved",
  async (adapter, { page }) => {
    await open(page, adapter);
    const thread = card(page, adapter, 3);
    await thread.click();
    await thread.getByTitle("Resolve").click();

    const saved = await saveAndReopen(page);
    expect(saved.extended).toMatch(
      /w15:paraId="10000003"[^>]*w15:done="1"|w15:done="1"[^>]*w15:paraId="10000003"/u,
    );
  },
);
