import { assertExactModel } from "../../test/exactModel";
import { expect, test, type Page } from "@playwright/test";

import { FolioDocxReviewer } from "../../packages/core/src/ai-edits/headless";
import type { DocxEditorRef } from "../../packages/react/src/components/DocxEditor.props";
import { paragraphLogicalText } from "../../packages/docx-core/src/ops/offsets";
import type { Document } from "../../packages/core/src/types/document";
import { createDocx } from "../../packages/core/src/docx/rezip";
import { parseDocx } from "../../packages/core/src/docx/parser";
import { createEmptyDocument } from "../../packages/core/src/utils/createDocument";
import type { buildCanonicalBridge } from "../parity/canonicalBridge";

declare global {
  var __folioPlayground: { getEditorRef: () => DocxEditorRef | null } | undefined;
}

const cases = [
  {
    kind: "pasteListHtml",
    name: "nested HTML bullet list",
    html: "<ul><li>Alpha<ul><li>Nested</li></ul></li><li>Omega</li></ul>",
    plain: "Alpha\nNested\nOmega",
    selection: "caret",
  },
  {
    kind: "pasteWordHtml",
    name: "Word-like paragraphs and table",
    html: '<p class="MsoNormal">Opening</p><table border="1"><tbody><tr><td><p class="MsoNormal">Cell A</p></td><td><p class="MsoNormal">Cell B</p></td></tr></tbody></table><p class="MsoNormal">Closing</p>',
    plain: "Opening\nCell A\tCell B\nClosing",
    selection: "replace",
  },
  {
    kind: "pasteHtml",
    name: "ordinary HTML paragraphs",
    html: "<p>First <strong>bold</strong></p><p>Second</p>",
    plain: "First bold\nSecond",
    selection: "replace",
  },
  {
    kind: "pastePlain",
    name: "plain multiline text replacing a selection",
    html: "",
    plain: "First\nSecond\nThird",
    selection: "replace",
  },
  {
    kind: "pastePlain",
    name: "plain multiline text at a caret",
    html: "",
    plain: "First\nSecond\nThird",
    selection: "caret",
  },
] as const;

type PasteCase = (typeof cases)[number];

const project = (reviewer: FolioDocxReviewer) =>
  reviewer.snapshot().blocks.map(({ text, displayLabel, listLevel, table }) => ({
    text,
    displayLabel,
    listLevel,
    table,
  }));

const paste = async (page: Page, input: PasteCase, suggesting: boolean) => {
  await page.goto("/?file=sample.docx");
  await page.waitForSelector(".layout-page");
  await page.evaluate(() =>
    globalThis.__folioPlayground?.getEditorRef()?.ensureEditorView({ focus: true }),
  );
  await page.waitForFunction(
    () => !!globalThis.__folioPlayground?.getEditorRef()?.getEditorRef()?.getView(),
  );
  const baselineBytes = await page.evaluate(async () => {
    const bytes = await globalThis.__folioPlayground?.getEditorRef()?.save({ selective: false });
    return bytes ? [...new Uint8Array(bytes)] : null;
  });
  if (!baselineBytes) throw new Error("editor did not return the original DOCX");
  const baseline = await FolioDocxReviewer.fromBuffer(new Uint8Array(baselineBytes).buffer);

  if (suggesting) await page.getByRole("button", { name: "Track Changes", exact: true }).click();
  const selected = await page.evaluate((selection) => {
    const view = globalThis.__folioPlayground?.getEditorRef()?.getEditorRef()?.getView();
    if (!view) throw new Error("editor view unavailable");
    let from: number | null = null;
    let text = "";
    view.state.doc.descendants((node, pos) => {
      if (from !== null) return false;
      if (node.isText && (node.text?.length ?? 0) >= 6) {
        from = pos;
        text = node.text?.slice(0, 6) ?? "";
        return false;
      }
      return true;
    });
    if (from === null) throw new Error("fixture has no six-character text run");
    const start = selection === "caret" ? from + 6 : from;
    view.dispatch(
      view.state.tr.setSelection(
        view.state.selection.constructor.create(view.state.doc, start, from + 6),
      ),
    );
    view.focus();
    return selection === "caret" ? "" : text;
  }, input.selection);
  expect(await page.evaluate(() => window.getSelection()?.toString())).toBe(selected);

  await page.context().grantPermissions(["clipboard-read", "clipboard-write"]);
  await page.evaluate(async ({ html, plain }) => {
    const items: Record<string, Blob> = {
      "text/plain": new Blob([plain], { type: "text/plain" }),
    };
    if (html) items["text/html"] = new Blob([html], { type: "text/html" });
    await navigator.clipboard.write([new ClipboardItem(items)]);
    document.addEventListener(
      "paste",
      (event) => {
        document.documentElement.dataset["receivedPasteHtml"] =
          event.clipboardData?.getData("text/html") ?? "";
      },
      { capture: true, once: true },
    );
  }, input);
  await page.keyboard.press(process.platform === "darwin" ? "Meta+v" : "Control+v");
  await expect
    .poll(() => page.locator("html").getAttribute("data-received-paste-html"))
    .toContain(input.html);

  const savedBytes = await page.evaluate(async () => {
    const bytes = await globalThis.__folioPlayground?.getEditorRef()?.save({ selective: false });
    return bytes ? [...new Uint8Array(bytes)] : null;
  });
  if (!savedBytes) {
    const status = await page.locator('.pg-status[data-status="error"]').textContent();
    throw new Error(`editor did not return the pasted DOCX: ${status}`);
  }
  const buffer = new Uint8Array(savedBytes).buffer;
  const reviewer = await FolioDocxReviewer.fromBuffer(buffer);
  return { buffer, baseline: project(baseline), selected, pending: reviewer.getChanges() };
};

for (const input of cases) {
  const tag = input.kind === "pastePlain" ? "" : ` @browser-input:${input.kind}`;
  test(`${input.name} paste accepts to editing mode and rejects to original${tag}`, async ({
    page,
  }) => {
    const edited = await paste(page, input, false);
    const suggested = await paste(page, input, true);
    expect(suggested.baseline).toEqual(edited.baseline);
    expect(suggested.pending.length).toBeGreaterThan(0);

    const accepting = await FolioDocxReviewer.fromBuffer(suggested.buffer);
    accepting.acceptAll();
    const accepted = await FolioDocxReviewer.fromBuffer(await accepting.toBuffer());
    const editedSaved = await FolioDocxReviewer.fromBuffer(edited.buffer);
    expect(project(accepted)).toEqual(project(editedSaved));

    const rejecting = await FolioDocxReviewer.fromBuffer(suggested.buffer);
    rejecting.rejectAll();
    const rejected = await FolioDocxReviewer.fromBuffer(await rejecting.toBuffer());
    expect(project(rejected)).toEqual(suggested.baseline);

    if (input.selection === "replace") {
      expect(suggested.pending).toEqual(
        expect.arrayContaining([
          expect.objectContaining({ type: "deletion", text: suggested.selected }),
        ]),
      );
    }
    if (input.kind === "pasteListHtml") {
      const blocks = project(editedSaved);
      const list = ["Alpha", "Nested", "Omega"].map((item) =>
        blocks.find(({ text }) => text.includes(item)),
      );
      expect(list.map((block) => block?.listLevel)).toEqual([0, 1, 0]);
      expect(list.every((block) => !!block?.displayLabel)).toBe(true);
    }
  });
}

test("native cut copies the selection and resolves like editing mode", async ({ page }) => {
  const run = async (suggesting: boolean) => {
    await page.goto("/?file=sample.docx");
    await page.waitForSelector(".layout-page");
    await page.evaluate(() =>
      globalThis.__folioPlayground?.getEditorRef()?.ensureEditorView({ focus: true }),
    );
    await page.waitForFunction(
      () => !!globalThis.__folioPlayground?.getEditorRef()?.getEditorRef()?.getView(),
    );
    const original = await page.evaluate(async () => {
      const bytes = await globalThis.__folioPlayground?.getEditorRef()?.save({ selective: false });
      return bytes ? [...new Uint8Array(bytes)] : null;
    });
    if (!original) throw new Error("editor did not return the original DOCX");
    if (suggesting) await page.getByRole("button", { name: "Track Changes", exact: true }).click();
    const selected = await page.evaluate(() => {
      const view = globalThis.__folioPlayground?.getEditorRef()?.getEditorRef()?.getView();
      if (!view) throw new Error("editor view unavailable");
      let from: number | null = null;
      let text = "";
      view.state.doc.descendants((node, pos) => {
        if (from !== null) return false;
        if (node.isText && (node.text?.length ?? 0) >= 6) {
          from = pos;
          text = node.text?.slice(0, 6) ?? "";
          return false;
        }
        return true;
      });
      if (from === null) throw new Error("fixture has no six-character text run");
      view.dispatch(
        view.state.tr.setSelection(
          view.state.selection.constructor.create(view.state.doc, from, from + 6),
        ),
      );
      view.focus();
      return text;
    });
    await page.context().grantPermissions(["clipboard-read", "clipboard-write"]);
    await page.keyboard.press(process.platform === "darwin" ? "Meta+x" : "Control+x");
    await expect.poll(() => page.evaluate(() => navigator.clipboard.readText())).toBe(selected);
    const saved = await page.evaluate(async () => {
      const bytes = await globalThis.__folioPlayground?.getEditorRef()?.save({ selective: false });
      return bytes ? [...new Uint8Array(bytes)] : null;
    });
    if (!saved) throw new Error("editor did not return the cut DOCX");
    return {
      buffer: new Uint8Array(saved).buffer,
      original: new Uint8Array(original).buffer,
      selected,
    };
  };

  const edited = await run(false);
  const suggested = await run(true);
  expect(suggested.selected).toBe(edited.selected);
  const editedSaved = await FolioDocxReviewer.fromBuffer(edited.buffer);
  const pending = await FolioDocxReviewer.fromBuffer(suggested.buffer);
  expect(pending.getChanges()).toEqual(
    expect.arrayContaining([
      expect.objectContaining({ type: "deletion", text: suggested.selected }),
    ]),
  );
  const accepting = await FolioDocxReviewer.fromBuffer(suggested.buffer);
  accepting.acceptAll();
  const accepted = await FolioDocxReviewer.fromBuffer(await accepting.toBuffer());
  expect(project(accepted)).toEqual(project(editedSaved));
  const rejecting = await FolioDocxReviewer.fromBuffer(suggested.buffer);
  rejecting.rejectAll();
  const rejected = await FolioDocxReviewer.fromBuffer(await rejecting.toBuffer());
  const baseline = await FolioDocxReviewer.fromBuffer(suggested.original);
  expect(project(rejected)).toEqual(project(baseline));
});

type CanonicalSnapshot = ReturnType<ReturnType<typeof buildCanonicalBridge>["snapshot"]>;

const canonicalSnapshot = async (page: Page) => {
  const current = await page.evaluate(() => globalThis.__folioCanonical?.snapshot());
  if (!current?.document) throw new TypeError("Canonical clipboard document unavailable.");
  expect(current.active).toBe(true);
  expect(current.projectionMatchesCanonical).toBe(true);
  expect(current.projectionJSON).toEqual(current.canonicalProjectionJSON);
  expect(current.provenance.valid).toBe(true);
  return { ...current, document: current.document };
};

const authoredClipboardContent = (document: Document) =>
  document.package.document.content.map((paragraph) => {
    if (paragraph.type !== "paragraph") throw new TypeError("Expected pasted paragraph.");
    return {
      paraId: paragraph.paraId,
      text: paragraphLogicalText(paragraph),
      alignment: paragraph.formatting?.alignment,
      keepNext: paragraph.formatting?.keepNext ?? false,
      runs: paragraph.content.map((run) => {
        if (run.type !== "run") throw new TypeError("Expected authored clipboard run.");
        return {
          text: run.content.map((leaf) => (leaf.type === "text" ? leaf.text : leaf.type)).join(""),
          bold: run.formatting?.bold ?? false,
          italic: run.formatting?.italic ?? false,
        };
      }),
    };
  });

const expectCanonicalHistoryState = (actual: CanonicalSnapshot, expected: CanonicalSnapshot) => {
  assertExactModel(actual.document, expected.document);
  expect(actual.projectionJSON).toEqual(expected.projectionJSON);
  assertExactModel(actual.selection, expected.selection);
};

test("canonical rich clipboard preserves generated marks, open ends and exact history in both adapters", async ({
  page,
}) => {
  test.setTimeout(120_000);
  const modifier = process.platform === "darwin" ? "Meta" : "Control";
  const ports = [
    Number(process.env["FOLIO_PLAYGROUND_PORT"]) || 4200,
    Number(process.env["FOLIO_PLAYGROUND_VUE_PORT"]) || 4201,
  ];
  const sourceDocument = createEmptyDocument({ initialText: "leftRIGHT" });
  const destination = sourceDocument.package.document.content.at(0);
  if (destination?.type !== "paragraph") throw new TypeError("Expected destination paragraph.");
  destination.formatting = { alignment: "center", keepNext: true };
  const source = [...new Uint8Array(await createDocx(sourceDocument))];
  await page.context().grantPermissions(["clipboard-read", "clipboard-write"]);

  for (const [mask, openStart, openEnd] of [
    [1, 0, 0],
    [2, 0, 1],
    [3, 1, 0],
    [3, 1, 1],
  ] as const) {
    const tokens = Array.from({ length: mask }, (_, index) => `Paste${mask}${index}😀`);
    const importedIds = tokens.map((_, index) => `4AFE000${index + 1}`);
    const html = tokens
      .map((token, index) => {
        let marked = token;
        if (mask & 1) marked = `<strong>${marked}</strong>`;
        if (mask & 2) marked = `<em>${marked}</em>`;
        const slice = index === 0 ? ` data-pm-slice="${openStart} ${openEnd} []"` : "";
        return `<p data-para-id="${importedIds.at(index)}"${slice}>${marked}</p>`;
      })
      .join("");

    for (const replace of [false, true]) {
      let firstAdapterContent: ReturnType<typeof authoredClipboardContent> | undefined;
      for (const port of ports) {
        await page.goto(`http://localhost:${port}/?session=canonical`);
        await page.waitForSelector(".layout-page");
        expect(
          await page.evaluate((bytes) => globalThis.__folioCanonical?.load(bytes), source),
        ).toBe(true);
        expect(
          await page.evaluate(
            (head) => globalThis.__folioCanonical?.select(3, head),
            replace ? 5 : 3,
          ),
        ).toBe(true);
        const before = await canonicalSnapshot(page);
        await page.evaluate(
          async ({ rich, plain }) => {
            await navigator.clipboard.write([
              new ClipboardItem({
                "text/html": new Blob([rich], { type: "text/html" }),
                "text/plain": new Blob([plain], { type: "text/plain" }),
              }),
            ]);
            document.addEventListener(
              "paste",
              (event) => {
                document.documentElement.dataset["canonicalPasteHtml"] =
                  event.clipboardData?.getData("text/html") ?? "";
              },
              { once: true, capture: true },
            );
          },
          { rich: html, plain: tokens.join("\n") },
        );
        await page.keyboard.press(`${modifier}+v`);
        await expect
          .poll(() => page.locator("html").getAttribute("data-canonical-paste-html"))
          .toContain("data-pm-slice");
        const after = await canonicalSnapshot(page);
        const expectedTexts = [...tokens];
        if (openStart === 1) expectedTexts[0] = `le${expectedTexts[0]}`;
        else expectedTexts.unshift("le");
        const suffix = replace ? "RIGHT" : "ftRIGHT";
        if (openEnd === 1) expectedTexts[expectedTexts.length - 1] += suffix;
        else expectedTexts.push(suffix);
        const authored = authoredClipboardContent(after.document);
        expect(authored.map(({ text }) => text)).toEqual(expectedTexts);
        expect(authored.at(0)?.alignment).toBe("center");
        expect(authored.at(0)?.keepNext).toBe(true);
        const identities = authored.map(({ paraId }) => paraId);
        expect(identities.every((id) => typeof id === "string" && /^[0-9A-F]{8}$/i.test(id))).toBe(
          true,
        );
        expect(new Set(identities).size).toBe(identities.length);
        expect(identities.every((id) => !importedIds.includes(id ?? ""))).toBe(true);
        for (const token of tokens) {
          const pastedRuns = authored
            .flatMap(({ runs }) => runs)
            .filter(({ text }) => text.includes(token));
          expect(pastedRuns).toHaveLength(1);
          expect(pastedRuns.at(0)?.bold).toBe(Boolean(mask & 1));
          expect(pastedRuns.at(0)?.italic).toBe(Boolean(mask & 2));
        }
        expect(after.canUndo).toBe(true);
        await page.keyboard.press(`${modifier}+z`);
        const undone = await canonicalSnapshot(page);
        expectCanonicalHistoryState(undone, before);
        expect(undone.canUndo).toBe(false);
        expect(undone.canRedo).toBe(true);
        await page.keyboard.press(`${modifier}+Shift+z`);
        const redone = await canonicalSnapshot(page);
        expectCanonicalHistoryState(redone, after);
        expect(redone.canRedo).toBe(false);

        const saved = await page.evaluate(() => globalThis.__folioCanonical?.save());
        if (!saved) throw new TypeError("Canonical clipboard document did not save.");
        const reopened = await parseDocx(new Uint8Array(saved), {
          preloadFonts: false,
          detectVariables: false,
        });
        const reopenedContent = authoredClipboardContent(reopened);
        expect(reopenedContent).toEqual(authored);
        if (firstAdapterContent) expect(reopenedContent).toEqual(firstAdapterContent);
        else firstAdapterContent = reopenedContent;
        expect(
          await page.evaluate((bytes) => globalThis.__folioCanonical?.load(bytes), saved),
        ).toBe(true);
        const reloaded = await canonicalSnapshot(page);
        expect(authoredClipboardContent(reloaded.document)).toEqual(authored);
        expect(reloaded.canUndo).toBe(false);
      }
    }
  }
});
