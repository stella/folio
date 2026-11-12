import type { Page } from "@playwright/test";

import { ZipArchive } from "@stll/docx-core/zip";

import {
  type AdapterFixture,
  ensureLiveView,
  expect,
  forEachAdapter,
} from "./parity-fixture";
import { withLegacyEngine } from "./playgroundHosts";

/**
 * Typing over exactly the words a comment covers keeps the typed words inside
 * the comment in both hosts, typing directly and tracking changes.
 */

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const COMMENTED = "Please remit funds.";

const commentedPackage = async (): Promise<number[]> => {
  const zip = new ZipArchive();
  zip.set(
    "[Content_Types].xml",
    `${HEAD}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/></Types>`,
  );
  zip.set(
    "_rels/.rels",
    `${HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`,
  );
  zip.set(
    "word/_rels/document.xml.rels",
    `${HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId9" Type="${R}/comments" Target="comments.xml"/></Relationships>`,
  );
  zip.set(
    "word/document.xml",
    `${HEAD}<w:document xmlns:w="${W}"><w:body>` +
      "<w:p><w:r><w:t>Opening words</w:t></w:r></w:p>" +
      `<w:p><w:commentRangeStart w:id="1"/><w:r><w:t>${COMMENTED}</w:t></w:r><w:commentRangeEnd w:id="1"/><w:r><w:commentReference w:id="1"/></w:r></w:p>` +
      "<w:p><w:r><w:t>Closing words</w:t></w:r></w:p>" +
      "</w:body></w:document>",
  );
  zip.set(
    "word/comments.xml",
    `${HEAD}<w:comments xmlns:w="${W}"><w:comment w:id="1" w:author="Reviewer" w:date="2024-01-01T00:00:00Z"><w:p><w:r><w:t>Check the account</w:t></w:r></w:p></w:comment></w:comments>`,
  );
  return [...new Uint8Array(await zip.generate())];
};

const CARD = { react: ".docx-comment-card", vue: ".comment-card" } as const;

const card = (page: Page, adapter: AdapterFixture) =>
  page.locator(`${CARD[adapter.name]}[data-comment-id="1"]`);

const load = async (page: Page, bytes: number[]): Promise<void> => {
  await page.evaluate((data) => window.__folioParity?.loadDocx(data), bytes);
  await ensureLiveView(page);
  await expect(page.locator(".layout-page").first()).toContainText(
    "Closing words",
  );
};

const open = async (page: Page, adapter: AdapterFixture): Promise<void> => {
  await page.goto(withLegacyEngine(`${adapter.baseUrl}/`));
  await expect
    .poll(() => page.evaluate(() => window.__folioParity !== undefined), {
      timeout: 30_000,
    })
    .toBe(true);
  await load(page, await commentedPackage());
};

/** The saved body, and the text between comment 1's range markers. */
const saved = async (page: Page) => {
  const bytes = await page.evaluate(
    () => window.__folioParity?.saveDocx() ?? Promise.resolve([]),
  );
  const zip = await ZipArchive.read(Uint8Array.from(bytes).buffer);
  const body = (await zip.entry("word/document.xml")?.text()) ?? "";
  const comments = (await zip.entry("word/comments.xml")?.text()) ?? "";
  const inside =
    /<w:commentRangeStart w:id="1"\/>(.*?)<w:commentRangeEnd w:id="1"\/>/u.exec(
      body,
    )?.[1] ?? "";
  return { bytes, body, comments, inside };
};

/** Select exactly the commented words, as a drag over them does. */
const selectCommentedWords = async (page: Page): Promise<void> => {
  // "Opening words" is the first paragraph; the commented words open the second.
  const from = 1 + "Opening words".length + 2;
  await page.evaluate(
    ([start, end]) => globalThis.__folioCanonical?.select(start, end),
    [from, from + COMMENTED.length],
  );
  await expect
    .poll(() => page.evaluate(() => window.__folioParity?.getSelectedText()))
    .toBe(COMMENTED);
};

forEachAdapter(
  "typing over the words a comment covers keeps them in the comment",
  async (adapter, { page }) => {
    await open(page, adapter);
    await selectCommentedWords(page);
    await page.keyboard.type("New words");
    await expect(page.locator(".layout-page").first()).toContainText(
      "New words",
    );

    const after = await saved(page);
    expect(after.comments).toContain("Check the account");
    expect(after.inside.replace(/<[^>]+>/gu, "")).toBe("New words");
    await load(page, after.bytes);
    await expect(card(page, adapter)).toBeVisible();
  },
);

forEachAdapter(
  "typing over the words a comment covers while tracking keeps both versions in the comment",
  async (adapter, { page }) => {
    await open(page, adapter);
    // React's toolbar names the toggle by its mode, Vue's by its action.
    await page
      .getByRole("button", { name: /^(?:Toggle )?Track Changes$/u })
      .first()
      .click();
    await selectCommentedWords(page);
    await page.keyboard.type("New words");
    await expect(page.locator(".layout-page").first()).toContainText(
      "New words",
    );

    const after = await saved(page);
    expect(after.comments).toContain("Check the account");
    // The struck words and the typed ones both sit inside the comment, so
    // accepting leaves it on the new words and rejecting on the old.
    expect(after.inside).toMatch(
      /<w:delText>Please remit funds\.<\/w:delText>.*<w:ins\b[^>]*>.*New words.*<\/w:ins>/u,
    );
    await load(page, after.bytes);
    await expect(card(page, adapter)).toBeVisible();
  },
);
