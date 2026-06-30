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

  test("redo after a grouped undo restores the typed burst (Shift+Z and Ctrl/Cmd+Y)", async ({
    page,
  }) => {
    await mountFixture(page, "sample.docx");
    await page.locator(".layout-paragraph").first().click();
    const before = await docText(page);
    expect(before).not.toContain("Qbxz");

    await page.keyboard.type("Qbxz", { delay: 15 });
    await page.waitForTimeout(150);
    expect(await docText(page)).toContain("Qbxz");

    // Grouped undo reverts the whole burst back to the pre-typing text.
    await page.keyboard.press(`${MOD}+z`);
    await page.waitForTimeout(150);
    expect(await docText(page)).toBe(before);

    // The standard redo chord (Shift held) restores the entire burst. With
    // Shift down the browser reports the uppercase "Z"; the history shortcut is
    // matched case-insensitively so this chord triggers redo.
    await page.keyboard.press(`${MOD}+Shift+z`);
    await page.waitForTimeout(150);
    expect(await docText(page)).toContain("Qbxz");

    // Undo once more, then redo via the alternate Ctrl/Cmd+Y chord also restores
    // the burst (not a stale, pre-typing snapshot).
    await page.keyboard.press(`${MOD}+z`);
    await page.waitForTimeout(150);
    expect(await docText(page)).toBe(before);

    await page.keyboard.press(`${MOD}+y`);
    await page.waitForTimeout(150);
    expect(await docText(page)).toContain("Qbxz");
  });
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

    // Insert a column right -> the first row gains exactly one cell (a collapsed
    // caret must add a single column, uniformly, not one per existing column).
    await (await openTableCellMenu(page))
      .getByRole("menuitem", { name: "Insert column right" })
      .click();
    await page.waitForTimeout(250);
    const colsAfterInsert = await firstRowCellCount(page);
    expect(colsAfterInsert).toBe(cols0 + 1);

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

test.describe("insert toolbar", () => {
  test("the opt-in Insert Table control adds a table to the document model", async ({ page }) => {
    await mountFixture(page, "sample.docx");

    // Place the caret in a body paragraph (outside any existing table) so the
    // insert lands in the flow, then count the table nodes before inserting.
    await page.locator(".layout-paragraph").first().click();
    const tablesBefore = await countNodes(page, "table");

    // The Insert group renders only because the playground passes the opt-in
    // `onInsert*` handlers; clicking it runs the real `insertTableInView` op.
    const insertTable = page.locator('[data-testid="toolbar-insert-table"]');
    await expect(insertTable).toBeVisible();
    await insertTable.click();
    await page.waitForTimeout(250);

    // A new table node appears in the live ProseMirror document.
    expect(await countNodes(page, "table")).toBe(tablesBefore + 1);
  });
});

test.describe("header/footer", () => {
  // Reads the header's hidden content editor (the `[data-hf-r-id]` ProseMirror
  // that backs the painted `.layout-page-header`).
  const headerContent = (page: Page) =>
    page.evaluate(
      () => document.querySelector("[data-hf-r-id] .ProseMirror")?.textContent ?? "",
    );

  // Text of every header in the *saved* document model. `getDocument()`
  // (buildCurrentDocument) flushes the persistent hidden header/footer views
  // into the returned Document, so this is exactly what a "Save As .docx" would
  // serialise — the round-trip persistence path, independent of any UI commit.
  const savedHeaderText = (page: Page) =>
    page.evaluate(() => {
      const doc = globalThis.__folioPlayground?.getEditorRef()?.getDocument() as
        | { package?: { headers?: unknown } }
        | null;
      const collect = (node: unknown): string => {
        if (typeof node === "string") return node;
        if (node === null || typeof node !== "object") return "";
        const record = node as Record<string, unknown>;
        if (typeof record["text"] === "string") return record["text"];
        const content = record["content"];
        if (Array.isArray(content)) return content.map(collect).join("");
        return collect(content);
      };
      const headers = doc?.package?.headers;
      const values = headers instanceof Map ? [...headers.values()] : Object.values(headers ?? {});
      return values.map((value) => collect(value)).join("");
    });

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

    // Round-trip persistence: the edit is already in the saved document model
    // (getDocument flushes the persistent header/footer views), so a save would
    // write it into the .docx without depending on any UI commit step.
    expect(await savedHeaderText(page)).toContain("HdrMark");
  });

  test("double-clicking an image-only header still routes edits into the header", async ({
    page,
  }) => {
    await mountFixture(page, "podily-bps.docx");

    const header = page.locator(".layout-page-header").first();
    await header.scrollIntoViewIfNeeded();

    // This first-page header holds only a floating logo image plus an empty
    // paragraph — there is no body-flow text under the cursor. A double-click in
    // the header band must still enter header edit mode AND land the caret in the
    // header content editor, rather than falling through to the body (the nearest
    // editable text). Regression guard for the title-page / image-only header.
    await header.dblclick();
    await page.locator(".hf-inline-editor").waitFor({ state: "visible", timeout: 5000 });

    // Pause before typing: a real user does not type instantly, so focus must
    // stay in the header view past the brief entry-focus window, not snap back
    // to the body once the window elapses.
    await page.waitForTimeout(1200);
    await page.keyboard.type("HdrMark", { delay: 15 });
    await page.waitForTimeout(150);

    const result = await page.evaluate(() => {
      const top = (
        globalThis as unknown as {
          __folioPlayground: {
            getEditorRef: () => {
              getDocument?: () => { package?: { headers?: unknown } };
              getEditorRef: () => {
                getView?: () => { state: { doc: { textContent: string } } } | null;
              };
            };
          };
        }
      ).__folioPlayground.getEditorRef();
      const headers = top.getDocument?.()?.package?.headers;
      const modelJson = JSON.stringify(
        headers instanceof Map ? [...headers.values()] : (headers ?? {}),
      );
      return {
        painted:
          document.querySelector('[data-page-number="1"] .layout-page-header')?.textContent ?? "",
        modelHasEdit: modelJson.includes("HdrMark"),
        body: top.getEditorRef().getView?.()?.state.doc.textContent ?? "",
      };
    });

    // The edit is visible in the painted header (the user sees what they type)...
    expect(result.painted).toContain("HdrMark");
    // ...is persisted into the saved document model (survives save)...
    expect(result.modelHasEdit).toBe(true);
    // ...and never leaks into the body (the nearest editable text).
    expect(result.body).not.toContain("HdrMark");
  });

  // NOTE: the *UI* save-on-body-click / save-on-exit commit path is not driven
  // here. While the inline header overlay is open the painted body paragraph is
  // not hittable, so a synthetic body click times out (a Playwright limitation,
  // not a product bug). Persistence itself is asserted above via getDocument():
  // the persistent hidden HF view holds the edit and is flushed on save, so the
  // edit is never dropped.
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
