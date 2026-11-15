import type { Page } from "@playwright/test";
import { readFileSync } from "node:fs";

import { ZipArchive } from "@stll/docx-core/zip";

import {
  type AdapterFixture,
  ensureLiveView,
  expect,
  forEachAdapter,
  openCanonicalEditor,
  openEditor,
} from "./parity-fixture";

forEachAdapter(
  "find: the dialog highlights matches without changing selection",
  async (adapter, { page }) => {
    await openEditor(page, adapter);
    await ensureLiveView(page);
    await page.locator(".layout-page").first().click();
    expect(
      await page.evaluate(() => window.__folioParity?.selectFirstWord()),
    ).toBe(true);
    const before = await page.evaluate(() => ({
      selection: window.__folioParity?.getSelectedText(),
      text: window.__folioParity?.getDocumentText(),
    }));
    expect(before.selection).toBeTruthy();
    expect(before.selection?.toLowerCase()).not.toBe("editor");

    await page.keyboard.press("ControlOrMeta+f");
    const dialog = page.getByTestId("find-replace-dialog");
    await expect(dialog).toBeVisible();
    await dialog.locator("input").first().fill("editor");

    const current = page
      .locator('[data-find-match-index="0"][data-current="true"]')
      .first();
    await expect(current).toBeVisible();
    await expect(dialog).toContainText("1 / 3");
    await dialog.locator("input").first().press("Enter");
    await expect(
      page.locator('[data-find-match-index="1"][data-current="true"]').first(),
    ).toBeVisible();
    await expect(
      page.locator('[data-find-match-index="0"][data-current="true"]'),
    ).toHaveCount(0);
    await expect(dialog).toContainText("2 / 3");
    expect(
      await page.evaluate(() => ({
        selection: window.__folioParity?.getSelectedText(),
        text: window.__folioParity?.getDocumentText(),
      })),
    ).toEqual(before);
  },
);

const REPLACEMENT = "folioReplacement";

const paragraph = (paraId: string, runs: string) =>
  `<w:p w14:paraId="${paraId}" w14:textId="77777777">${runs}</w:p>`;
const run = (text: string, bold = false) =>
  `<w:r>${bold ? "<w:rPr><w:b/></w:rPr>" : ""}<w:t xml:space="preserve">${text}</w:t></w:r>`;

/**
 * The sample package with three matches of "quokka": one in a plain run, one
 * split across a bold run and a plain run, and one in a table cell.
 */
const findReplaceDocument = async (): Promise<number[]> => {
  const zip = await ZipArchive.read(
    readFileSync(new URL("../visual/fixtures/sample.docx", import.meta.url)),
  );
  const xml = (await zip.entry("word/document.xml")?.text()) ?? "";
  const bodyStart = xml.indexOf("<w:body>") + "<w:body>".length;
  const sectionStart = xml.lastIndexOf("<w:sectPr");
  if (bodyStart < "<w:body>".length || sectionStart < bodyStart)
    throw new TypeError("The sample body is missing");
  const body = [
    paragraph("3FA00001", run("alpha quokka beta")),
    paragraph("3FA00002", `${run("quo", true)}${run("kka gamma")}`),
    '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>',
    paragraph("3FA00003", run("cell quokka")),
    "</w:tc></w:tr></w:tbl>",
    paragraph("3FA00004", run("closing")),
  ].join("");
  zip.set(
    "word/document.xml",
    `${xml.slice(0, bodyStart)}${body}${xml.slice(sectionStart)}`,
  );
  return [...(await zip.generate())];
};

const openFindReplaceDocument = async (
  page: Page,
  adapter: AdapterFixture,
  mode: "editing" | "suggesting",
) => {
  await openCanonicalEditor(page, adapter, undefined, mode);
  const bytes = await findReplaceDocument();
  await page.evaluate((data) => window.__folioParity?.loadDocx(data), bytes);
  await ensureLiveView(page);
  await expect
    .poll(() =>
      page.evaluate(() => window.__folioParity?.getDocumentText() ?? ""),
    )
    .toContain("cell quokka");
};

/** Open the replace row from the public ref and search for "quokka". */
const searchForQuokka = async (page: Page, adapter: AdapterFixture) => {
  await page.evaluate(() => window.__folioParity?.openReplace());
  const dialog = page.getByTestId("find-replace-dialog");
  await expect(dialog).toBeVisible();
  const replaceInput = dialog.getByLabel("Replacement text");
  if (adapter.name === "vue" && !(await replaceInput.isVisible()))
    await dialog.getByTestId("find-replace-toggle").click();
  await dialog.getByLabel("Find text").fill("quokka");
  await expect(
    page.locator('[data-find-match-index="0"][data-current="true"]').first(),
  ).toBeVisible();
  await expect(dialog).toContainText("1 / 3");
  await replaceInput.fill(REPLACEMENT);
  return {
    dialog,
    replaceCurrent:
      adapter.name === "vue"
        ? dialog.getByTestId("replace-current")
        : dialog.getByRole("button", { name: "Replace current match" }),
    replaceAll:
      adapter.name === "vue"
        ? dialog.getByTestId("replace-all")
        : dialog.getByRole("button", { name: "Replace all" }),
  };
};

const TEXT = /<w:(t|delText)(?:\s[^>]*)?>([^<]*)<\/w:\1>/gu;

/** The saved body's text, read with every change accepted or rejected. */
const savedText = async (page: Page, decision: "accept" | "reject") => {
  const bytes = await page.evaluate(
    async () => (await window.__folioParity?.saveDocx()) ?? [],
  );
  const zip = await ZipArchive.read(new Uint8Array(bytes));
  const xml = (await zip.entry("word/document.xml")?.text()) ?? "";
  const resolved =
    decision === "accept"
      ? xml.replaceAll(/<w:del\b[\s\S]*?<\/w:del>/gu, "")
      : xml.replaceAll(/<w:ins\b[\s\S]*?<\/w:ins>/gu, "");
  return {
    xml,
    text: [...resolved.matchAll(TEXT)].map((match) => match[2]).join(""),
  };
};

const SOURCE_TEXT = "alpha quokka betaquokka gammacell quokkaclosing";
const REPLACED_TEXT = `alpha ${REPLACEMENT} beta${REPLACEMENT} gammacell ${REPLACEMENT}closing`;

forEachAdapter(
  "find/replace: replace current writes the highlighted match on the canonical session",
  async (adapter, { page }) => {
    await openFindReplaceDocument(page, adapter, "editing");
    const { dialog, replaceCurrent } = await searchForQuokka(page, adapter);
    await replaceCurrent.click();
    await expect
      .poll(() =>
        page.evaluate(() => window.__folioParity?.getDocumentText() ?? ""),
      )
      .toContain(`alpha ${REPLACEMENT} beta`);
    await expect(dialog).toContainText("1 / 2");
    const saved = await savedText(page, "accept");
    expect(saved.text).toContain(`alpha ${REPLACEMENT} beta`);
    expect(saved.text).toContain("quokka gamma");
    expect(saved.text).toContain("cell quokka");
  },
);

forEachAdapter(
  "find/replace: replace all writes a split match and a table cell on the canonical session",
  async (adapter, { page }) => {
    await openFindReplaceDocument(page, adapter, "editing");
    const { replaceAll } = await searchForQuokka(page, adapter);
    await replaceAll.click();
    await expect
      .poll(() =>
        page.evaluate(() => window.__folioParity?.getDocumentText() ?? ""),
      )
      .toContain(`cell ${REPLACEMENT}`);
    const saved = await savedText(page, "accept");
    expect(saved.text).toContain(REPLACED_TEXT);
    expect(saved.xml).not.toContain("<w:ins ");
    expect(saved.xml).not.toContain("<w:del ");
  },
);

forEachAdapter(
  "find/replace: replace all while suggesting records tracked changes on the canonical session",
  async (adapter, { page }) => {
    await openFindReplaceDocument(page, adapter, "suggesting");
    const { replaceAll } = await searchForQuokka(page, adapter);
    await replaceAll.click();
    await expect
      .poll(() =>
        page.evaluate(() => window.__folioParity?.getDocumentText() ?? ""),
      )
      .toContain(REPLACEMENT);
    await expect
      .poll(
        async () =>
          (
            await page.evaluate(
              () => window.__folioParity?.getDocumentText() ?? "",
            )
          ).split(REPLACEMENT).length,
      )
      .toBe(4);
    const accepted = await savedText(page, "accept");
    expect(accepted.xml).toContain("<w:ins ");
    expect(accepted.xml).toContain("<w:del ");
    expect(accepted.text).toContain(REPLACED_TEXT);
    expect((await savedText(page, "reject")).text).toContain(SOURCE_TEXT);
  },
);
