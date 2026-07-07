/**
 * Core editing-flow regression tests for the folio editor.
 *
 * Companion to `interactions.spec.ts`: same playground test hook
 * (`window.__folioPlayground`), same "assert document CONTENT and live editor
 * STATE, never screenshots" discipline, so these run in CI alongside the other
 * interaction specs and stay stable across machines. They cover flows the
 * sibling spec does not: inline mark shortcuts, footer (not header) editing,
 * list toggles, and structural undo.
 *
 * Assertions read the live ProseMirror document and the saved document model
 * via the imperative `DocxEditorRef`; they never depend on pixels.
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

/**
 * Click a toolbar button by its accessible name. The font/color/alignment/
 * list group collapses behind the "More formatting" overflow trigger
 * whenever the bar is too narrow to show it inline (measured live against
 * the toolbar's actual content, not a fixed viewport breakpoint) — open that
 * popover first when the button is not already on screen.
 */
async function clickToolbarButton(page: Page, name: string): Promise<void> {
  const button = page.getByRole("button", { name });
  const visible = await button.isVisible().catch(() => false);
  if (!visible) {
    await page.getByRole("button", { name: "More formatting" }).click();
  }
  await button.click();
}

/** Live body ProseMirror document text. */
function bodyText(page: Page): Promise<string> {
  return page.evaluate(
    () =>
      globalThis.__folioPlayground?.getEditorRef()?.getEditorRef()?.getView()?.state.doc
        .textContent ?? "",
  );
}

/** Count body ProseMirror nodes by type name in the live document. */
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

type SelectionMarkState = {
  text: string;
  /** Whether EVERY text node inside the selection carries the mark. */
  bold: boolean;
  italic: boolean;
  underline: boolean;
  /** Whether the single character immediately before the selection has the mark. */
  beforeBold: boolean;
  beforeItalic: boolean;
  beforeUnderline: boolean;
  /** Whether the single character immediately after the selection has the mark. */
  afterBold: boolean;
  afterItalic: boolean;
  afterUnderline: boolean;
};

/**
 * Mark coverage around the live selection. For each formatting mark the
 * `bold` / `italic` / `underline` fields report FULL coverage: every text node
 * inside `from..to` must carry the mark, so a command that formatted only part
 * of the selection (or none of it) fails the assertion. The `before*` / `after*`
 * fields report the mark on the single adjacent character on each side, so a
 * leak of formatting onto neighbouring text is caught too. The selected text is
 * returned so the caller can prove the selection is the span it meant to format.
 */
function selectionMarkState(page: Page): Promise<SelectionMarkState> {
  return page.evaluate(() => {
    const empty = {
      text: "",
      bold: false,
      italic: false,
      underline: false,
      beforeBold: false,
      beforeItalic: false,
      beforeUnderline: false,
      afterBold: false,
      afterItalic: false,
      afterUnderline: false,
    };
    const view = globalThis.__folioPlayground?.getEditorRef()?.getEditorRef()?.getView();
    if (!view) return empty;
    const { from, to } = view.state.selection;
    const { schema, doc } = view.state;

    // True only when the mark is present on EVERY text node touched by the
    // selection — partial or absent coverage returns false.
    const wholeSelectionHasMark = (name: string): boolean => {
      const type = schema.marks[name];
      if (!type) return false;
      let sawText = false;
      let covered = true;
      doc.nodesBetween(from, to, (node) => {
        if (node.isText) {
          sawText = true;
          if (!type.isInSet(node.marks)) covered = false;
        }
        return true;
      });
      return sawText && covered;
    };

    // Whether the single character at [pos, pos + 1] carries the mark; used to
    // confirm a toggle did not bleed onto the chars adjacent to the selection.
    const charHasMark = (pos: number, name: string): boolean => {
      const type = schema.marks[name];
      if (!type || pos < 0 || pos + 1 > doc.content.size) return false;
      return doc.rangeHasMark(pos, pos + 1, type);
    };

    return {
      text: doc.textBetween(from, to),
      bold: wholeSelectionHasMark("bold"),
      italic: wholeSelectionHasMark("italic"),
      underline: wholeSelectionHasMark("underline"),
      beforeBold: charHasMark(from - 1, "bold"),
      beforeItalic: charHasMark(from - 1, "italic"),
      beforeUnderline: charHasMark(from - 1, "underline"),
      afterBold: charHasMark(to, "bold"),
      afterItalic: charHasMark(to, "italic"),
      afterUnderline: charHasMark(to, "underline"),
    };
  });
}

/**
 * The numbering id of the paragraph holding the caret, or null when it is not a
 * list item. folio models lists as paragraph `numPr` attrs (numId 1 = bullets),
 * not wrapper nodes, so the toggle's effect is visible here.
 */
function caretParagraphNumId(page: Page): Promise<number | null> {
  return page.evaluate(() => {
    const view = globalThis.__folioPlayground?.getEditorRef()?.getEditorRef()?.getView();
    if (!view) return null;
    const { $from } = view.state.selection;
    for (let depth = $from.depth; depth >= 0; depth -= 1) {
      const node = $from.node(depth);
      if (node.type.name === "paragraph") {
        const numPr = node.attrs["numPr"] as { numId?: number } | null | undefined;
        return numPr?.numId ?? null;
      }
    }
    return null;
  });
}

/**
 * Text of every header / footer in the *saved* document model. `getDocument()`
 * (buildCurrentDocument) flushes the persistent hidden HF views into the
 * returned Document, so this is exactly what a "Save As .docx" would serialise:
 * the round-trip persistence path, independent of any UI commit step.
 */
function savedHfText(page: Page, slot: "headers" | "footers"): Promise<string> {
  return page.evaluate((key) => {
    const doc = globalThis.__folioPlayground?.getEditorRef()?.getDocument() as {
      package?: Record<string, unknown>;
    } | null;
    const collect = (node: unknown): string => {
      if (typeof node === "string") return node;
      if (node === null || typeof node !== "object") return "";
      const record = node as Record<string, unknown>;
      if (typeof record["text"] === "string") return record["text"];
      const content = record["content"];
      if (Array.isArray(content)) return content.map(collect).join("");
      return collect(content);
    };
    const slotMap = doc?.package?.[key];
    const values = slotMap instanceof Map ? [...slotMap.values()] : Object.values(slotMap ?? {});
    return values.map((value) => collect(value)).join("");
  }, slot);
}

/** Place the caret inside the first painted table cell and open its menu. */
async function openTableCellMenu(page: Page): Promise<Locator> {
  const cell = page.locator(".layout-table-cell").first();
  await cell.scrollIntoViewIfNeeded();
  await cell.click();
  await cell.click({ button: "right" });
  const menu = page.locator('[role="menu"][aria-label="Text editing menu"]');
  await menu.waitFor({ state: "visible", timeout: 5000 });
  return menu;
}

test.describe("inline formatting", () => {
  test("bold / italic / underline shortcuts each toggle their own mark over the selection", async ({
    page,
  }) => {
    await mountFixture(page, "sample.docx");
    await page.locator(".layout-paragraph").first().click();

    // Type a unique marker, then select exactly it (one ArrowLeft per char).
    // Working off a known selection makes the assertions independent of the
    // fixture's pre-existing formatting.
    const MARKER = "Fmtzz";
    await page.keyboard.type(MARKER, { delay: 15 });
    await page.waitForTimeout(200); // let the insertion + re-layout settle
    for (let i = 0; i < MARKER.length; i += 1) {
      await page.keyboard.press("Shift+ArrowLeft");
    }
    await expect.poll(async () => (await selectionMarkState(page)).text).toBe(MARKER);
    const before = await selectionMarkState(page);

    // The chars flanking the selection must keep their formatting through every
    // toggle: a shortcut that bled onto adjacent text would change these. Held
    // constant in each expectation below.
    const neighbors = {
      beforeBold: before.beforeBold,
      beforeItalic: before.beforeItalic,
      beforeUnderline: before.beforeUnderline,
      afterBold: before.afterBold,
      afterItalic: before.afterItalic,
      afterUnderline: before.afterUnderline,
    };

    // Each shortcut flips ONLY its own mark across the WHOLE selection; the other
    // two marks and both neighbours are untouched. This proves the keymap wires
    // Mod-b/i/u to the matching schema mark and applies it to the full selection
    // (not just one char, not the whole document, not the adjacent text).
    await page.keyboard.press(`${MOD}+b`);
    await expect
      .poll(() => selectionMarkState(page))
      .toEqual({
        text: MARKER,
        bold: !before.bold,
        italic: before.italic,
        underline: before.underline,
        ...neighbors,
      });
    const afterBold = await selectionMarkState(page);

    await page.keyboard.press(`${MOD}+i`);
    await expect
      .poll(() => selectionMarkState(page))
      .toEqual({
        text: MARKER,
        bold: afterBold.bold,
        italic: !afterBold.italic,
        underline: afterBold.underline,
        ...neighbors,
      });
    const afterItalic = await selectionMarkState(page);

    await page.keyboard.press(`${MOD}+u`);
    await expect
      .poll(() => selectionMarkState(page))
      .toEqual({
        text: MARKER,
        bold: afterItalic.bold,
        italic: afterItalic.italic,
        underline: !afterItalic.underline,
        ...neighbors,
      });
    const afterUnderline = await selectionMarkState(page);

    // Toggling bold a second time removes only bold; italic + underline persist,
    // confirming the three marks are independent on the same range.
    await page.keyboard.press(`${MOD}+b`);
    await expect
      .poll(() => selectionMarkState(page))
      .toEqual({
        text: MARKER,
        bold: before.bold,
        italic: afterUnderline.italic,
        underline: afterUnderline.underline,
        ...neighbors,
      });
  });
});

test.describe("footer", () => {
  test("double-clicking the footer enters edit mode; edits land in the footer, not the body", async ({
    page,
  }) => {
    await mountFixture(page, "sample.docx");

    const footer = page.locator(".layout-page-footer").first();
    await footer.scrollIntoViewIfNeeded();
    expect(await savedHfText(page, "footers")).not.toContain("FtrMark");

    // Double-click enters footer edit mode (routed via the same headless
    // useHeaderFooterEditor double-click handler as the header) and places the
    // caret in the footer content editor.
    await footer.dblclick();
    await page.locator(".hf-inline-editor").waitFor({ state: "visible", timeout: 5000 });

    // Pause before typing so focus must persist in the footer view past the
    // brief entry-focus window, rather than snapping back to the body.
    await page.waitForTimeout(1200);
    await page.keyboard.type("FtrMark", { delay: 15 });
    await page.waitForTimeout(150);

    // The user sees the edit in the painted footer...
    expect(await footer.textContent()).toContain("FtrMark");
    // ...it is persisted into the saved footers model (survives save)...
    expect(await savedHfText(page, "footers")).toContain("FtrMark");
    // ...and never leaks into the body or the header.
    expect(await bodyText(page)).not.toContain("FtrMark");
    expect(await savedHfText(page, "headers")).not.toContain("FtrMark");
  });
});

test.describe("lists", () => {
  test("the toolbar list buttons toggle bullet and numbered lists on the caret paragraph", async ({
    page,
  }) => {
    await mountFixture(page, "sample.docx");
    await page.locator(".layout-paragraph").first().click();

    // Bullet list applies numbering id 1 to the caret paragraph.
    await clickToolbarButton(page, "Bullet List");
    await expect.poll(() => caretParagraphNumId(page)).toBe(1);

    // Switching to a numbered list moves it to a distinct numbering id (2).
    await clickToolbarButton(page, "Numbered List");
    await expect.poll(() => caretParagraphNumId(page)).toBe(2);

    // Clicking the already-active numbered-list button clears list formatting.
    await clickToolbarButton(page, "Numbered List");
    await expect.poll(() => caretParagraphNumId(page)).toBeNull();
  });
});

test.describe("table", () => {
  test("a single undo reverts a row insertion back to the original structure", async ({ page }) => {
    await mountFixture(page, "sample.docx");

    const rows0 = await countNodes(page, "tableRow");
    expect(rows0).toBeGreaterThan(0);

    const menu = await openTableCellMenu(page);
    await menu.getByRole("menuitem", { name: "Insert row below" }).click();
    await expect.poll(() => countNodes(page, "tableRow")).toBeGreaterThan(rows0);

    // Re-place the caret in the table body so the history shortcut targets the
    // body view, then ONE undo must restore the original row count: a structural
    // insert participates in history as a single revertible step.
    await page.locator(".layout-table-cell").first().click();
    await page.keyboard.press(`${MOD}+z`);
    await expect.poll(() => countNodes(page, "tableRow")).toBe(rows0);
  });
});
