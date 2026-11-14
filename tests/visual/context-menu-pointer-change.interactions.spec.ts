import { ZipArchive } from "@stll/docx-core/zip";

/**
 * A context menu's Accept and Reject items act on the tracked change under
 * the pointer.
 *
 * A right-click inside a range selection keeps the range, which can start
 * outside the change the pointer is on. The menu read the change at the
 * selection's start, so it offered no review items for the change the user
 * right-clicked.
 */
import { expect, type Page, test } from "../support/unclassifiedProbe";
import type {} from "./interleavingBridge";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const FIXTURE = "context-menu-pointer-change.docx";
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
      `<w:p>${run("Free text ")}${inserted(2, "Free Author", "freeword")}</w:p>` +
      `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`,
  );
  return Buffer.from(await zip.generate());
};

const open = async (page: Page) => {
  const body = await fixture();
  await page.route(`**/fixtures/${FIXTURE}`, (route) =>
    route.fulfill({
      body,
      contentType:
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    }),
  );
  await page.goto(`/?file=${FIXTURE}`);
  await expect(insertion(page, "freeword")).toBeVisible({ timeout: 30_000 });
  // The editing view mounts on first use; the range is set on it directly.
  await page.evaluate(() =>
    globalThis.__folioPlayground?.getEditorRef()?.ensureEditorView(),
  );
  await page.waitForFunction(
    () =>
      globalThis.__folioPlayground?.getEditorRef()?.getEditorRef()?.getView() !=
      null,
  );
};

const insertion = (page: Page, text: string) =>
  page
    .locator(".paged-editor__pages .docx-insertion")
    .filter({ hasText: text })
    .first();

/**
 * Select from the start of the paragraph that reads `paragraph` to the end of
 * its text: a range that starts outside the change and holds the pointer.
 */
const selectParagraph = (page: Page, paragraph: string) =>
  page.evaluate((wanted) => {
    const editor = globalThis.__folioPlayground?.getEditorRef()?.getEditorRef();
    const view = editor?.getView();
    if (!editor || !view) throw new Error("browser editor unavailable");
    const ranges: { from: number; to: number }[] = [];
    view.state.doc.descendants((node, pos) => {
      if (ranges.length > 0) return false;
      if (node.type.name !== "paragraph") return true;
      if (node.textContent === wanted)
        ranges.push({ from: pos + 1, to: pos + node.nodeSize - 1 });
      return false;
    });
    const range = ranges.at(0);
    if (!range) throw new Error(`no paragraph reads "${wanted}"`);
    const { from, to } = range;
    editor.setSelection(from, to);
    const selection = view.state.selection;
    return { from: selection.from, to: selection.to };
  }, paragraph);

test("a right-click on a change inside a range selection resolves that change", async ({
  page,
}) => {
  await open(page);
  const selected = await selectParagraph(page, "Free text freeword");
  expect(selected.to).toBeGreaterThan(selected.from);

  await insertion(page, "freeword").click({ button: "right" });
  const menuAccept = page.getByRole("menuitem", { name: "Accept Change" });
  await expect(menuAccept).toBeVisible();
  await expect(menuAccept).not.toHaveAttribute("aria-disabled", "true");
  await expect(
    page.getByRole("menuitem", { name: "Reject Change" }),
  ).toBeVisible();
  await menuAccept.click();

  // The change under the pointer resolved, and only that one.
  await expect(insertion(page, "freeword")).toHaveCount(0);
  await expect(page.locator(".paged-editor__pages")).toContainText("freeword");
  await expect(insertion(page, "heldword")).toBeVisible();
});

test("a right-click on a locked change inside a range selection shows the lock", async ({
  page,
}) => {
  await open(page);
  const selected = await selectParagraph(page, "Guarded clause heldword");
  expect(selected.to).toBeGreaterThan(selected.from);

  await insertion(page, "heldword").click({ button: "right" });
  const menuAccept = page.getByRole("menuitem", { name: "Accept Change" });
  await expect(menuAccept).toHaveAttribute("aria-disabled", "true");
  await expect(menuAccept).toHaveAttribute("title", HELD);
  await expect(
    page.getByRole("menuitem", { name: "Reject Change" }),
  ).toHaveAttribute("title", HELD);
});
