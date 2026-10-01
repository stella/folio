import { expect, test, type Page } from "@playwright/test";
import fc from "fast-check";

import { SUGGESTION_INPUT_KINDS } from "../../packages/core/src/__tests__/suggestionInputKinds";
import { shapeArrayBuffer } from "../../packages/core/src/__tests__/documentShapes";
import { FolioDocxReviewer } from "../../packages/core/src/ai-edits/headless";
import { withFakeTextMeasure } from "../../packages/core/src/layout-engine/measure/__tests__/fakeTextMeasure";
import {
  createFreshRenderRig,
  docFromParagraphs,
} from "../../packages/core/src/controller/__tests__/freshRenderHarness";
import type { DocxEditorRef } from "../../packages/react/src/components/DocxEditor.props";
import {
  browserInputTraceArbitrary,
  BROWSER_PASTE_PAYLOADS,
  browserSuggestionActionKinds,
  parseBrowserInputTraceConfig,
  type BrowserInputAction,
  type BrowserInputTrace,
  type BrowserDragTarget,
} from "./browserInputTrace";
import {
  failureMarker,
  failureRecord,
  logFailureMarker,
  writeFailureRecord,
} from "../../test/consumer-scenarios/support/failure-fingerprints";
import { clipboardHtmlProjection } from "./clipboardHtmlProjection";

declare global {
  var __folioPlayground: { getEditorRef: () => DocxEditorRef | null } | undefined;
}

const MODIFIER = process.platform === "darwin" ? "Meta" : "Control";
const INPUT_TRACE = browserInputTraceArbitrary();

type Block = {
  kind: string;
  text: string;
  displayLabel?: string;
  listLevel?: number | null;
  table?: unknown;
};
type LiveBlock = Pick<Block, "kind" | "text" | "table">;
const project = (reviewer: FolioDocxReviewer): Block[] =>
  reviewer.snapshot().blocks.map(({ kind, text, displayLabel, listLevel, table }) => ({
    kind,
    text,
    displayLabel,
    listLevel,
    table,
  }));
const projectLive = (blocks: readonly Block[]): LiveBlock[] =>
  blocks.map(({ kind, text, table }) => ({ kind, text, table }));

const liveBlocks = (page: Page) =>
  page.evaluate(() => {
    const snapshot = globalThis.__folioPlayground?.getEditorRef()?.createAIEditSnapshot();
    if (!snapshot) throw new Error("live reader unavailable");
    return snapshot.blocks.map(({ kind, text, table }) => ({
      kind,
      text,
      table,
    }));
  });

const editor = (page: Page) =>
  page.evaluate(() => {
    const ref = globalThis.__folioPlayground?.getEditorRef();
    if (!ref?.getEditorRef()?.getView()) throw new Error("browser editor unavailable");
    return true;
  });

const load = async (page: Page, bytes: ArrayBuffer, baseline: Block[], suggesting: boolean) => {
  await page.goto("/");
  await page.waitForSelector(".layout-page");
  await page.evaluate(() => globalThis.__folioPlayground?.getEditorRef()?.ensureEditorView());
  await page.waitForFunction(
    () => !!globalThis.__folioPlayground?.getEditorRef()?.getEditorRef()?.getView(),
  );
  await page.evaluate(
    async (source) => {
      const ref = globalThis.__folioPlayground?.getEditorRef();
      if (!ref) throw new Error("browser editor unavailable");
      await ref.loadDocumentBuffer(new Uint8Array(source));
    },
    [...new Uint8Array(bytes)],
  );
  await expect.poll(() => liveBlocks(page)).toEqual(projectLive(baseline));
  await editor(page);
  if (suggesting) await page.getByRole("button", { name: "Track Changes", exact: true }).click();
  await page.evaluate(() =>
    globalThis.__folioPlayground?.getEditorRef()?.getEditorRef()?.getView()?.focus(),
  );
};

const selectTableTarget = async (page: Page) => {
  const cells = await page.evaluate(() => {
    const view = globalThis.__folioPlayground?.getEditorRef()?.getEditorRef()?.getView();
    if (!view) throw new Error("browser editor unavailable");
    const positions: { pos: number; paragraphPos: number }[] = [];
    view.state.doc.descendants((node, pos) => {
      if (positions.length >= 2) return false;
      if (node.type.name !== "tableCell") return true;
      let paragraphPos: number | null = null;
      node.descendants((child, offset) => {
        if (paragraphPos !== null) return false;
        if (child.isTextblock) paragraphPos = pos + 1 + offset;
        return !child.isTextblock;
      });
      if (paragraphPos !== null) positions.push({ pos, paragraphPos });
      return false;
    });
    return positions;
  });
  const [anchor, head] = cells;
  if (!anchor || !head) return false;

  // The PM view is off-screen. Pointer actions must use the painted body cells,
  // whose paragraph positions also let us wait for layout after preceding edits.
  const paintedCell = (paragraphPos: number) =>
    page
      .locator(`.layout-page-content .layout-table-cell[data-pm-start="${paragraphPos}"]`)
      .first();
  const anchorCell = paintedCell(anchor.paragraphPos);
  const headCell = paintedCell(head.paragraphPos);
  await anchorCell.scrollIntoViewIfNeeded();
  await headCell.scrollIntoViewIfNeeded();
  const [from, to] = await Promise.all([anchorCell.boundingBox(), headCell.boundingBox()]);
  if (!from || !to) throw new Error("table drag targets are not painted");
  await page.mouse.move(from.x + from.width / 2, from.y + from.height / 2);
  await page.mouse.down();
  await page.mouse.move(to.x + to.width / 2, to.y + to.height / 2, { steps: 4 });
  await page.mouse.up();

  // A missed drag must fail at the input boundary, before paste/delete can edit
  // whichever caret the previous action happened to leave in each mode.
  const selected = await page.evaluate(() => {
    const selection = globalThis.__folioPlayground?.getEditorRef()?.getEditorRef()?.getView()
      ?.state.selection;
    if (!selection || !("$anchorCell" in selection) || !("$headCell" in selection)) return null;
    const selectedAnchor = selection.$anchorCell;
    const selectedHead = selection.$headCell;
    if (
      !selectedAnchor ||
      typeof selectedAnchor !== "object" ||
      !("pos" in selectedAnchor) ||
      !selectedHead ||
      typeof selectedHead !== "object" ||
      !("pos" in selectedHead)
    )
      return null;
    return { anchor: selectedAnchor.pos, head: selectedHead.pos };
  });
  expect(selected).toEqual({ anchor: anchor.pos, head: head.pos });
  return true;
};

const selectTarget = async (page: Page, target: BrowserDragTarget) => {
  if (target === "table") return selectTableTarget(page);
  const coords = await page.evaluate((wanted) => {
    const view = globalThis.__folioPlayground?.getEditorRef()?.getEditorRef()?.getView();
    if (!view) throw new Error("browser editor unavailable");
    const positions: number[] = [];
    const targetName = {
      table: "tableCell",
      list: "paragraph",
      note: "footnoteRef",
      field: "field",
      inlineObject: "image",
    }[wanted];
    view.state.doc.descendants((node, pos) => {
      const matches =
        node.type.name === targetName &&
        (wanted !== "list" || (node.attrs["numPr"] !== null && node.attrs["numPr"] !== undefined));
      if (matches && positions.length < 2) positions.push(pos);
      if (wanted === "note" && node.marks.some((mark) => mark.type.name === targetName)) {
        positions.push(pos);
      }
      return true;
    });
    if (positions.length === 0) return null;
    const from = Math.max(1, positions.at(0) ?? 1);
    const to = Math.min(view.state.doc.content.size - 1, (positions.at(1) ?? from + 2) + 1);
    return { from: view.coordsAtPos(from), to: view.coordsAtPos(to) };
  }, target);
  if (!coords) return false;
  await page.mouse.move(coords.from.left, coords.from.top + 3);
  await page.mouse.down();
  await page.mouse.move(coords.to.left, coords.to.top + 3, { steps: 4 });
  await page.mouse.up();
  return true;
};

const paste = async (page: Page, action: Extract<BrowserInputAction, { html: string }>) => {
  const expectedHtml = await page.evaluate(clipboardHtmlProjection, action.html);
  await page.context().grantPermissions(["clipboard-read", "clipboard-write"]);
  await page.evaluate(async ({ html, plain }) => {
    const parts: Record<string, Blob> = { "text/plain": new Blob([plain], { type: "text/plain" }) };
    if (html) parts["text/html"] = new Blob([html], { type: "text/html" });
    await navigator.clipboard.write([new ClipboardItem(parts)]);
    document.documentElement.dataset["fuzzPasteReceived"] = "unseen";
    document.addEventListener(
      "paste",
      (event) => {
        document.documentElement.dataset["fuzzPasteReceived"] = "seen";
        document.documentElement.dataset["fuzzPasteHtml"] =
          event.clipboardData?.getData("text/html") ?? "";
        document.documentElement.dataset["fuzzPastePlain"] =
          event.clipboardData?.getData("text/plain") ?? "";
      },
      { capture: true, once: true },
    );
  }, action);
  await page.keyboard.press(`${MODIFIER}+v`);
  await expect
    .poll(() => page.locator("html").getAttribute("data-fuzz-paste-received"))
    .toBe("seen");
  await expect
    .poll(async () =>
      page.evaluate(
        clipboardHtmlProjection,
        (await page.locator("html").getAttribute("data-fuzz-paste-html")) ?? "",
      ),
    )
    .toBe(expectedHtml);
  await expect(page.locator("html")).toHaveAttribute("data-fuzz-paste-plain", action.plain);
};

const drive = async (page: Page, action: BrowserInputAction) => {
  switch (action.kind) {
    case "typing":
      await page.keyboard.type(action.text);
      return;
    case "enter":
      await page.keyboard.press("Enter");
      return;
    case "backspace":
      await page.keyboard.press("Backspace");
      return;
    case "delete":
      await page.keyboard.press("Delete");
      return;
    case "pastePlain":
    case "pasteHtml":
    case "pasteWordHtml":
    case "pasteListHtml":
    case "pasteTable":
    case "pasteMultiBlock":
      await paste(page, action);
      return;
    case "imeReplacement": {
      const cdp = await page.context().newCDPSession(page);
      try {
        for (const text of action.updates) {
          await cdp.send("Input.imeSetComposition", {
            text,
            // CDP offsets count UTF-16 code units, including surrogate pairs.
            selectionStart: text.length,
            selectionEnd: text.length,
          });
        }
        if (action.completion === "commit") {
          const text = action.updates.at(-1);
          if (text === undefined) throw new Error("IME lifecycle has no updates");
          await cdp.send("Input.insertText", { text });
        } else {
          await cdp.send("Input.imeSetComposition", {
            text: "",
            selectionStart: 0,
            selectionEnd: 0,
          });
        }
      } finally {
        await cdp.detach();
      }
      return;
    }
    case "cut":
      await page.keyboard.press(`${MODIFIER}+x`);
      return;
    case "dragCellDelete":
      if (!(await selectTarget(page, "table")))
        throw new Error("table fixture has no draggable cells");
      await page.keyboard.press("Delete");
      return;
    case "undo":
      await page.keyboard.press(`${MODIFIER}+z`);
      return;
    case "redo":
      await page.keyboard.press(`${MODIFIER}+Shift+z`);
      return;
    case "historyBurst":
      for (const key of action.keys) {
        await page.keyboard.press(key === "undo" ? `${MODIFIER}+z` : `${MODIFIER}+Shift+z`);
      }
      return;
    case "selectionDrag":
      if (!(await selectTarget(page, action.target))) {
        throw new Error(`fixture has no ${action.target} drag target`);
      }
      return;
    default: {
      const unreachable: never = action;
      return unreachable;
    }
  }
};

const save = async (page: Page) => {
  const bytes = await page.evaluate(async () => {
    const saved = await globalThis.__folioPlayground?.getEditorRef()?.save({ selective: false });
    return saved ? [...new Uint8Array(saved)] : null;
  });
  if (!bytes) throw new Error("browser editor did not save");
  return new Uint8Array(bytes).buffer;
};

const runMode = async (
  page: Page,
  source: ArrayBuffer,
  baseline: Block[],
  trace: BrowserInputTrace,
  suggesting: boolean,
) => {
  await load(page, source, baseline, suggesting);
  for (const action of trace.actions) await drive(page, action);
  await page.waitForTimeout(350);
  const live = await liveBlocks(page);
  const painted = await page.locator(".layout-page-content").allTextContents();
  const buffer = await save(page);
  const reopened = await FolioDocxReviewer.fromBuffer(buffer);
  expect(projectLive(project(reopened))).toEqual(live);
  await page.evaluate(
    async (saved) => {
      await globalThis.__folioPlayground?.getEditorRef()?.loadDocumentBuffer(new Uint8Array(saved));
    },
    [...new Uint8Array(buffer)],
  );
  await expect.poll(() => page.locator(".layout-page-content").allTextContents()).toEqual(painted);
  return { buffer, blocks: project(reopened), changes: reopened.getChanges() };
};

/** Replays the same action schedule through #1148's fresh-render scheduler oracle. */
const checkFreshRender = (trace: BrowserInputTrace) => {
  withFakeTextMeasure(() => {
    const rig = createFreshRenderRig({
      initialDoc: docFromParagraphs(["alpha", "beta"]),
      leadingFrame: true,
      ext: null,
    });
    for (const action of trace.actions) {
      if (action.kind === "typing") {
        rig.edit((state) => state.tr.insertText(action.text, 2));
      } else if (action.kind === "backspace" || action.kind === "delete") {
        rig.edit((state) =>
          state.doc.firstChild?.textContent.length ? state.tr.delete(1, 2) : null,
        );
      } else {
        rig.tick(16);
      }
    }
    rig.settle();
    expect(rig.staleCommits).toEqual([]);
    expect(rig.committed()).toEqual(rig.fresh());
  });
};

const config = parseBrowserInputTraceConfig(
  process.env,
  process.env["FOLIO_FUZZ_LANE"] === "nightly" ? "nightly" : "pullRequest",
);
test.setTimeout(600_000);

test("browser generator covers every declared suggestion input kind", () => {
  expect(new Set(browserSuggestionActionKinds)).toEqual(new Set(SUGGESTION_INPUT_KINDS));
});

test("clipboard oracle accepts HTML serialization and detects text, attribute and markup changes", async ({
  page,
}) => {
  await page.goto("/");
  for (const { html } of Object.values(BROWSER_PASTE_PAYLOADS)) {
    const expected = await page.evaluate(clipboardHtmlProjection, html);
    // Clipboard writers parse document wrappers before delivering a paste.
    // Chromium adds the omitted head in the full-document fixture (seed 11).
    expect(await page.evaluate(clipboardHtmlProjection, expected)).toBe(expected);
    expect(await page.evaluate(clipboardHtmlProjection, "")).not.toBe(expected);
    const mutations = await page.evaluate((source) => {
      const textDocument = new DOMParser().parseFromString(source, "text/html");
      const walker = textDocument.createTreeWalker(textDocument.body, NodeFilter.SHOW_TEXT);
      const firstText = walker.nextNode();
      if (!firstText) throw new Error("clipboard fixture has no text to mutate");
      firstText.textContent = "lost clipboard text";
      const attributeDocument = new DOMParser().parseFromString(source, "text/html");
      const element = attributeDocument.body.firstElementChild;
      if (!element) throw new Error("clipboard fixture has no element to mutate");
      element.setAttribute("data-fuzz-mutated", "true");
      const markupDocument = new DOMParser().parseFromString(source, "text/html");
      const wrapper = markupDocument.body.firstElementChild;
      if (!wrapper) throw new Error("clipboard fixture has no markup to mutate");
      wrapper.replaceWith(...wrapper.childNodes);
      return [
        textDocument.documentElement.outerHTML,
        attributeDocument.documentElement.outerHTML,
        markupDocument.documentElement.outerHTML,
      ];
    }, html);
    for (const mutation of mutations) {
      expect(await page.evaluate(clipboardHtmlProjection, mutation)).not.toBe(expected);
    }
  }
  const source = BROWSER_PASTE_PAYLOADS.pasteWordHtml.html;
  const normalized = source.replace("<body>", "<head></head><body>");
  expect(await page.evaluate(clipboardHtmlProjection, source)).toBe(
    await page.evaluate(clipboardHtmlProjection, normalized),
  );
});

for (const seed of config.seeds) {
  test(`seed ${seed}: browser input preserves readers, fresh render, and suggesting equivalence`, async ({
    page,
  }) => {
    const verdict = await fc.check(
      fc.asyncProperty(INPUT_TRACE, async (trace) => {
        checkFreshRender(trace);
        const source = await shapeArrayBuffer(trace.shape);
        const baseline = project(await FolioDocxReviewer.fromBuffer(source));
        const edited = await runMode(page, source, baseline, trace, false);
        const suggested = await runMode(page, source, baseline, trace, true);
        const accepting = await FolioDocxReviewer.fromBuffer(suggested.buffer);
        accepting.acceptAll();
        const accepted = await FolioDocxReviewer.fromBuffer(await accepting.toBuffer());
        expect(project(accepted)).toEqual(edited.blocks);
        const rejecting = await FolioDocxReviewer.fromBuffer(suggested.buffer);
        rejecting.rejectAll();
        const rejected = await FolioDocxReviewer.fromBuffer(await rejecting.toBuffer());
        expect(project(rejected)).toEqual(baseline);
        if (JSON.stringify(edited.blocks) !== JSON.stringify(baseline)) {
          expect(suggested.changes.length).toBeGreaterThan(0);
        }
      }),
      { seed, numRuns: config.runs, endOnFailure: false },
    );
    if (verdict.failed) {
      const failure = verdict.errorInstance;
      const trace = verdict.counterexample?.at(0);
      if (trace !== undefined && failure !== undefined && failure !== null) {
        const marker = failureMarker({
          test: "browser input preserves readers, fresh render, and suggesting equivalence",
          seed,
          path: verdict.counterexamplePath,
          repro: `FOLIO_FUZZ_SEEDS=${seed} FOLIO_FUZZ_RUNS=${config.runs} bunx playwright test --project=browser-fuzzer tests/visual/browser-input-fuzz.interactions.spec.ts --workers=1`,
          failure,
          flow: `${trace.shape}: ${trace.actions.map(({ kind }) => kind).join(" → ")}`,
        });
        logFailureMarker(marker);
        writeFailureRecord(
          "test-results/browser-fuzz-failures",
          failureRecord(marker, failure, { flow: trace }),
        );
      }
      const detail =
        failure instanceof Error ? (failure.stack ?? failure.message) : fc.stringify(failure);
      throw new Error(
        `seed=${seed} path=${verdict.counterexamplePath} trace=${JSON.stringify(verdict.counterexample?.at(0))}\n${detail}`,
        { cause: failure },
      );
    }
  });
}
