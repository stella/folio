/**
 * Accepting or rejecting a tracked change inside a content control whose
 * content is locked is refused in the editors' own review controls, with the
 * reason, while a change outside it still resolves.
 */
import { expect, type Page, test } from "@playwright/test";

import { ZipArchive } from "@stll/docx-core/zip";

import {
  ensureLiveView,
  openCanonicalEditor,
  VUE_ADAPTER,
} from "../parity/parity-fixture";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const FIXTURE = "content-control-lock-review.docx";
const HELD =
  "A locked content control prevents accepting or rejecting this change.";

const run = (text: string) =>
  `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const inserted = (id: number, author: string, text: string) =>
  `<w:ins w:id="${id}" w:author="${author}" w:date="2024-01-01T00:00:00Z">${run(text)}</w:ins>`;

const fixture = async (): Promise<Buffer> => {
  const zip = new ZipArchive();
  zip.set(
    "[Content_Types].xml",
    `${HEAD}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`,
  );
  zip.set(
    "_rels/.rels",
    `${HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`,
  );
  zip.set(
    "word/document.xml",
    `${HEAD}<w:document xmlns:w="${W}"><w:body>` +
      `<w:p>${run("Guarded ")}<w:sdt><w:sdtPr><w:alias w:val="Guarded"/><w:lock w:val="contentLocked"/></w:sdtPr>` +
      `<w:sdtContent>${run("clause")}${inserted(1, "Locked Author", " heldword")}</w:sdtContent></w:sdt></w:p>` +
      `<w:p>${run("Free ")}${inserted(2, "Free Author", "freeword")}</w:p>` +
      `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`,
  );
  return Buffer.from(await zip.generate());
};

const serveFixture = async (page: Page) => {
  const body = await fixture();
  await page.route(`**/fixtures/${FIXTURE}`, (route) =>
    route.fulfill({
      body,
      contentType:
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    }),
  );
};

const insertion = (page: Page, text: string) =>
  page
    .locator(".paged-editor__pages .docx-insertion")
    .filter({ hasText: text })
    .first();

const savedDocument = async (page: Page) => {
  const saved = await page.evaluate(async () => {
    const buffer = await globalThis.__folioPlayground?.getEditorRef()?.save();
    return buffer ? [...new Uint8Array(buffer)] : null;
  });
  expect(saved).not.toBeNull();
  const zip = await ZipArchive.read(Uint8Array.from(saved ?? []).buffer);
  return (await zip.entry("word/document.xml")?.text()) ?? "";
};

test("React: the toolbar and context menu refuse a change a lock holds", async ({
  page,
}) => {
  await serveFixture(page);
  await page.goto(`/?session=canonical&file=${FIXTURE}`);
  await expect(insertion(page, "heldword")).toBeVisible({ timeout: 30_000 });

  const accept = page.getByTestId("toolbar-accept-change");
  const reject = page.getByTestId("toolbar-reject-change");
  // The review buttons sit past the toolbar's fold.
  const showReviewButtons = async () => {
    if (!(await accept.isVisible()))
      await page.getByRole("button", { name: "More formatting" }).click();
  };

  await insertion(page, "heldword").click();
  await showReviewButtons();
  await expect(accept).toBeDisabled();
  await expect(reject).toBeDisabled();
  await expect(accept).toHaveAttribute("title", HELD);
  await expect(reject).toHaveAttribute("title", HELD);
  // Close the toolbar overflow before reaching the text again.
  await page.keyboard.press("Escape");

  await insertion(page, "heldword").click({ button: "right" });
  const menuAccept = page.getByRole("menuitem", { name: "Accept Change" });
  await expect(menuAccept).toHaveAttribute("aria-disabled", "true");
  await expect(menuAccept).toHaveAttribute("title", HELD);
  await expect(
    page.getByRole("menuitem", { name: "Reject Change" }),
  ).toHaveAttribute("title", HELD);
  await page.keyboard.press("Escape");

  await insertion(page, "freeword").click();
  await showReviewButtons();
  await expect(accept).toBeEnabled();
  await accept.click();
  await page.keyboard.press("Escape");
  await expect(insertion(page, "freeword")).toHaveCount(0);
  await expect(insertion(page, "heldword")).toBeVisible();

  const xml = await savedDocument(page);
  expect(xml).toContain("freeword");
  expect(xml).not.toContain('w:author="Free Author"');
  expect(xml).toContain('w:author="Locked Author"');
  expect(xml).toContain("heldword");
});

test("Vue: a change card's buttons refuse a change a lock holds", async ({
  page,
}) => {
  await serveFixture(page);
  await openCanonicalEditor(page, VUE_ADAPTER, FIXTURE);
  await ensureLiveView(page);

  // The sidebar opens by itself only for comment threads.
  await page.getByTestId("toolbar-comments-toggle").click();
  const cards = page.locator(".tc-card");
  await expect(cards).toHaveCount(2, { timeout: 30_000 });

  const held = cards.filter({ hasText: "heldword" });
  await held.click();
  const heldAccept = held.getByRole("button", { name: "Accept Change" });
  const heldReject = held.getByRole("button", { name: "Reject Change" });
  await expect(heldAccept).toBeDisabled();
  await expect(heldReject).toBeDisabled();
  await expect(heldAccept).toHaveAttribute("title", HELD);
  await expect(heldReject).toHaveAttribute("aria-description", HELD);

  const free = cards.filter({ hasText: "freeword" });
  await free.click();
  const freeAccept = free.getByRole("button", { name: "Accept Change" });
  await expect(freeAccept).toBeEnabled();
  await freeAccept.click();
  await expect(cards).toHaveCount(1);
  await expect(cards.first()).toContainText("heldword");
});
