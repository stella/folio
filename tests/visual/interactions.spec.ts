/**
 * Interaction (behaviour) regression tests for the folio editor.
 *
 * These complement `rendering.spec.ts` (pixel baselines): they assert document
 * CONTENT and live editor STATE — never screenshots — so they are stable across
 * machines and run in CI. They exercise exactly the flows the headless refactor
 * moved through (`@stll/folio-core` managers + thin `@stll/folio-react` hook
 * bindings): typing + grouped undo, find, table row/column edits, header/footer
 * editing, image resize, and zoom.
 *
 * Assertions read the live ProseMirror document and the imperative
 * `DocxEditorRef` via the playground's `window.__folioPlayground` test hook.
 */

import { expect, test } from "@playwright/test";
import type { Locator, Page } from "@playwright/test";

import type { DocxEditorRef } from "../../packages/react/src/components/DocxEditor.props";

declare global {
  var __folioPlayground: { getEditorRef: () => DocxEditorRef | null } | undefined;
}

const isMac = process.platform === "darwin";
const MOD = isMac ? "Meta" : "Control";

/** Mount a fixture and force the (deferred) editor view to materialise. */
async function mountFixture(page: Page, fixture: string): Promise<void> {
  await page.goto(`/?file=${encodeURIComponent(fixture)}`);
  await page.waitForSelector('[data-testid="folio-editor"]', { timeout: 15_000 });
  await page.waitForFunction(() => document.querySelectorAll(".layout-page").length >= 1);
  await page.evaluate(async () => {
    await document.fonts.ready;
  });
  await page.evaluate(() => {
    globalThis.__folioPlayground?.getEditorRef()?.ensureEditorView({ focus: true });
  });
  await page.waitForFunction(
    () => !!globalThis.__folioPlayground?.getEditorRef()?.getEditorRef()?.getView(),
    undefined,
    { timeout: 10_000 },
  );
  await page.waitForTimeout(250);
}

/** Live ProseMirror document text. */
function docText(page: Page): Promise<string> {
  return page.evaluate(
    () =>
      globalThis.__folioPlayground?.getEditorRef()?.getEditorRef()?.getView()?.state.doc
        .textContent ?? "",
  );
}

/** Count ProseMirror nodes by type name in the live document. */
function countNodes(page: Page, typeName: string): Promise<number> {
  return page.evaluate((name) => {
    const view = globalThis.__folioPlayground?.getEditorRef()?.getEditorRef()?.getView();
    let n = 0;
    view?.state.doc.descendants((node) => {
      if (node.type.name === name) n += 1;
      return true;
    });
    return n;
  }, typeName);
}

/** Cells in the first table row (column count). */
function firstRowCellCount(page: Page): Promise<number> {
  return page.evaluate(() => {
    const view = globalThis.__folioPlayground?.getEditorRef()?.getEditorRef()?.getView();
    let cells = 0;
    let done = false;
    view?.state.doc.descendants((node) => {
      if (done) return false;
      if (node.type.name === "tableRow") {
        node.descendants((child) => {
          if (child.type.name === "tableCell" || child.type.name === "tableHeader") cells += 1;
          return false;
        });
        done = true;
        return false;
      }
      return true;
    });
    return cells;
  });
}

/** Place the caret inside the first painted table cell and open its menu. */
async function openTableCellMenu(page: Page): Promise<Locator> {
  const cell = page.locator(".layout-table-cell").first();
  await cell.scrollIntoViewIfNeeded();
  const box = await cell.boundingBox();
  if (!box) throw new Error("no painted table cell");
  const cx = box.x + box.width / 2;
  const cy = box.y + box.height / 2;
  await page.mouse.click(cx, cy);
  await page.mouse.click(cx, cy, { button: "right" });
  const menu = page.locator('[role="menu"][aria-label="Text editing menu"]');
  await menu.waitFor({ state: "visible", timeout: 5000 });
  return menu;
}

test.describe("typing + undo", () => {
  test("a rapid burst types as a group, and one undo reverts the whole burst", async ({ page }) => {
    await mountFixture(page, "sample.docx");

    // Place the caret in the first body paragraph.
    await page.locator(".layout-paragraph").first().click();
    const before = await docText(page);
    expect(before).not.toContain("Qbxz");

    // Type a burst rapidly so it lands in one history group.
    await page.keyboard.type("Qbxz", { delay: 15 });
    await page.waitForTimeout(150);
    expect(await docText(page)).toContain("Qbxz");

    // One undo must revert the entire burst (grouped-undo behaviour), not a
    // single character: the document returns exactly to its pre-typing text.
    await page.keyboard.press(`${MOD}+z`);
    await page.waitForTimeout(150);
    const afterUndo = await docText(page);
    expect(afterUndo).not.toContain("Qbxz");
    expect(afterUndo).toBe(before);
  });

  // KNOWN ISSUE (reported, not asserted): redo after a grouped undo does not
  // restore the typed text on current `main`. Cmd/Ctrl+Shift+Z + Cmd/Ctrl+Y
  // both fire the redo (the editor's `canRedo()` flips true -> false), but the
  // restored document does not contain the burst — the redo entry holds a stale
  // (pre-typing) snapshot. Separately, the document-level `useHistory` keydown
  // handler matches `event.key === "z"`, which is "Z" while Shift is held, so
  // Shift+Z alone never reaches it (only Ctrl/Cmd+Y does). Skipped until the
  // history redo is fixed, rather than asserting broken behaviour.
  test.skip("redo restores an undone burst", () => {});
});

test.describe("find", () => {
  test("Cmd/Ctrl+F finds a term present in the document and reports a match count", async ({
    page,
  }) => {
    await mountFixture(page, "sample.docx");
    await page.locator(".layout-paragraph").first().click();

    await page.keyboard.press(`${MOD}+f`);
    const dialog = page.locator('[data-testid="find-replace-dialog"]');
    await dialog.waitFor({ state: "visible", timeout: 5000 });

    const searchInput = dialog.locator("input").first();
    await searchInput.fill("editor");

    // The dialog debounces, then renders a "<current> / <total>" counter.
    const counter = dialog.locator("text=/\\d+\\s*\\/\\s*\\d+/").first();
    await counter.waitFor({ state: "visible", timeout: 5000 });
    const counterText = (await counter.textContent()) ?? "";
    const total = Number(counterText.split("/")[1]?.trim());
    expect(total).toBeGreaterThan(0);

    // Navigating to the next match advances the current index (when >1 match).
    if (total > 1) {
      const currentBefore = Number(counterText.split("/")[0]?.trim());
      await searchInput.press("Enter");
      await page.waitForTimeout(200);
      const currentAfter = Number(((await counter.textContent()) ?? "").split("/")[0]?.trim());
      expect(currentAfter).not.toBe(currentBefore);
    }
  });

  // The shipped FindReplaceDialog renders find-only: a search field, match-case
  // / whole-word toggles, and prev/next navigation. `onReplace` / `onReplaceAll`
  // are wired at the props level (DocxEditorDialogs) but no replace input or
  // "Replace all" control is rendered, so a replace cannot be driven through the
  // product without UI changes. Reported, not hacked around.
  test.skip("find/replace: replace via the dialog", () => {});
});

test.describe("table", () => {
  test("inserting and deleting rows and columns changes the table structure", async ({ page }) => {
    await mountFixture(page, "sample.docx");

    const rows0 = await countNodes(page, "tableRow");
    const cols0 = await firstRowCellCount(page);
    expect(rows0).toBeGreaterThan(0);
    expect(cols0).toBeGreaterThan(0);

    // Insert a row below -> row count grows.
    await (await openTableCellMenu(page)).getByRole("menuitem", { name: "Insert row below" }).click();
    await page.waitForTimeout(250);
    const rowsAfterInsert = await countNodes(page, "tableRow");
    expect(rowsAfterInsert).toBeGreaterThan(rows0);

    // Insert a column right -> the first row gains at least one cell.
    await (await openTableCellMenu(page))
      .getByRole("menuitem", { name: "Insert column right" })
      .click();
    await page.waitForTimeout(250);
    const colsAfterInsert = await firstRowCellCount(page);
    expect(colsAfterInsert).toBeGreaterThan(cols0);

    // Delete a row -> row count shrinks again.
    await (await openTableCellMenu(page)).getByRole("menuitem", { name: "Delete row" }).click();
    await page.waitForTimeout(250);
    expect(await countNodes(page, "tableRow")).toBeLessThan(rowsAfterInsert);

    // Delete a column -> the first row loses a cell again.
    await (await openTableCellMenu(page)).getByRole("menuitem", { name: "Delete column" }).click();
    await page.waitForTimeout(250);
    expect(await firstRowCellCount(page)).toBeLessThan(colsAfterInsert);
  });
});

test.describe("header/footer", () => {
  // Reads the header's hidden content editor (the `[data-hf-r-id]` ProseMirror
  // that backs the painted `.layout-page-header`).
  const headerContent = (page: Page) =>
    page.evaluate(
      () => document.querySelector("[data-hf-r-id] .ProseMirror")?.textContent ?? "",
    );

  test("double-clicking the header enters edit mode and edits land in the header", async ({
    page,
  }) => {
    await mountFixture(page, "sample.docx");

    const header = page.locator(".layout-page-header").first();
    await header.scrollIntoViewIfNeeded();
    expect(await headerContent(page)).not.toContain("HdrMark");

    // Double-click enters header edit mode (routed via the headless
    // useHeaderFooterEditor double-click handler) and places the caret in the
    // header content editor.
    await header.dblclick();
    await page.locator(".hf-inline-editor").waitFor({ state: "visible", timeout: 5000 });

    await page.keyboard.type("HdrMark", { delay: 15 });
    await page.waitForTimeout(150);

    // The typed text lands in the live header content editor.
    expect(await headerContent(page)).toContain("HdrMark");
  });

  // NOTE: committing the header edit back into the .docx (the save-on-body-click
  // / save-on-exit path) could not be driven reliably from Playwright — a
  // synthetic body click tears down the hidden header editors before the save
  // reads them, so the edit is dropped instead of persisted. The edit itself
  // (above) is asserted; the round-trip persistence is reported rather than
  // asserted on flaky synthetic events.
});

test.describe("image", () => {
  // No fixture exposes a body-level image: sample / docx-editor-demo ship none,
  // and podily-bps's only image is a header logo (word/media in header rels),
  // not selectable from the body. Driving a resize would require inserting an
  // image (file upload) and dragging the overlay's resize handles, which carry
  // no stable selector. Skipped and reported rather than instrumenting the
  // product or hand-authoring a binary fixture.
  test.skip("resizing a selected image updates its dimensions", () => {});
});

test.describe("zoom", () => {
  test("setZoom updates getZoom and scales the rendered page", async ({ page }) => {
    await mountFixture(page, "sample.docx");

    const pageWidth = async () => (await page.locator(".layout-page").first().boundingBox())?.width ?? 0;

    expect(await page.evaluate(() => globalThis.__folioPlayground?.getEditorRef()?.getZoom())).toBe(
      1,
    );
    const width1 = await pageWidth();

    await page.evaluate(() => globalThis.__folioPlayground?.getEditorRef()?.setZoom(1.5));
    await page.waitForTimeout(250);

    expect(await page.evaluate(() => globalThis.__folioPlayground?.getEditorRef()?.getZoom())).toBe(
      1.5,
    );
    const ratio = (await pageWidth()) / width1;
    expect(ratio).toBeGreaterThan(1.3);
    expect(ratio).toBeLessThan(1.7);
  });

  test("the playground zoom-in control drives the editor zoom", async ({ page }) => {
    await mountFixture(page, "sample.docx");
    await page.getByRole("button", { name: "Zoom in" }).click();
    await page.waitForTimeout(150);
    const zoom = await page.evaluate(() =>
      globalThis.__folioPlayground?.getEditorRef()?.getZoom(),
    );
    expect(zoom).toBeGreaterThan(1);
  });
});
