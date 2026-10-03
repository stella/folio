import type { BrowserDragTarget } from "./browserDragTarget";
import { driveBrowserIme } from "./browserImeDriver";
import { expect, test, type Page } from "@playwright/test";
import fc from "fast-check";
import { validateDocxPackage } from "../../packages/docx-core/src/validate/docx";

import knownFailures from "../../test/known-failure-fingerprints.json" with { type: "json" };

import { classifyFuzzRun } from "../../test/fuzz-health";
import { reportFuzzHealth } from "../../test/consumer-scenarios/support/fuzz-health";

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
  BROWSER_SHAPE_TARGETS,
  browserInputTraceArbitrary,
  BROWSER_PASTE_PAYLOADS,
  browserSuggestionActionKinds,
  parseBrowserInputTraceConfig,
  type BrowserInputAction,
  type BrowserInputTrace,
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
const PAINTED_TARGET_MODULE = `/@fs${new URL("./browserPaintedTargets.ts", import.meta.url).pathname}`;

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
  const target = await page.evaluate(async (moduleUrl) => {
    const { resolvePaintedTableTarget }: typeof import("./browserPaintedTargets") = await import(
      moduleUrl
    );
    const ref = globalThis.__folioPlayground?.getEditorRef();
    if (!ref) throw new Error("browser editor unavailable");
    return resolvePaintedTableTarget(ref);
  }, PAINTED_TARGET_MODULE);
  if (target.type === "absent") return false;
  await page.mouse.move(target.from.x, target.from.y);
  await page.mouse.down();
  await page.mouse.move(target.to.x, target.to.y, { steps: 4 });
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
  expect(selected).toEqual({ anchor: target.anchor, head: target.head });
  return true;
};

const selectTarget = async (page: Page, target: BrowserDragTarget) => {
  if (target === "table") return selectTableTarget(page);
  const painted = await page.evaluate(
    async ({ moduleUrl, wanted }) => {
      const { resolvePaintedDragTarget }: typeof import("./browserPaintedTargets") = await import(
        moduleUrl
      );
      const ref = globalThis.__folioPlayground?.getEditorRef();
      if (!ref) throw new Error("browser editor unavailable");
      return resolvePaintedDragTarget(ref, wanted);
    },
    { moduleUrl: PAINTED_TARGET_MODULE, wanted: target },
  );
  if (painted.type === "absent") return false;
  const { positions } = painted;
  expect(positions.to).toBeGreaterThan(positions.from);
  await page.mouse.move(painted.from.x, painted.from.y);
  await page.mouse.down();
  await page.mouse.move(painted.to.x, painted.to.y, { steps: 4 });
  await page.mouse.up();
  const selected = await page.evaluate(() => {
    const selection = globalThis.__folioPlayground?.getEditorRef()?.getEditorRef()?.getView()
      ?.state.selection;
    return selection ? { from: selection.from, to: selection.to } : null;
  });
  expect(selected).toEqual({ from: positions.from, to: positions.to });
  if (positions.targetType === "inline") {
    expect(selected?.from).toBeLessThanOrEqual(positions.targetFrom);
    expect(selected?.to).toBeGreaterThanOrEqual(positions.targetTo);
  }
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
    case "imeReplacement":
      await driveBrowserIme(page, action);
      return;
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

const reopenSaved = async (buffer: ArrayBuffer) => {
  const validity = await validateDocxPackage(buffer);
  expect(validity.valid, validity.valid ? "" : validity.error).toBe(true);
  return FolioDocxReviewer.fromBuffer(buffer);
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
  const reopened = await reopenSaved(buffer);
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

const paintedTargetReplays = [
  {
    shape: "tables",
    actions: [
      { kind: "imeReplacement", updates: ["alpha"], completion: "commit" },
      { kind: "selectionDrag", target: "table" },
    ],
  },
  {
    shape: "notes",
    actions: [
      { kind: "delete" },
      { kind: "pasteTable", ...BROWSER_PASTE_PAYLOADS.pasteTable },
      { kind: "selectionDrag", target: "note" },
    ],
  },
] satisfies readonly BrowserInputTrace[];

for (const [index, trace] of paintedTargetReplays.entries()) {
  test(`painted target replay ${index}`, async ({ page }) => {
    const source = await shapeArrayBuffer(trace.shape);
    const baseline = project(await FolioDocxReviewer.fromBuffer(source));
    for (const suggesting of [false, true]) {
      await load(page, source, baseline, suggesting);
      for (const action of trace.actions) await drive(page, action);
    }
  });
}

const isBrowserShape = (shape: string): shape is keyof typeof BROWSER_SHAPE_TARGETS =>
  Object.hasOwn(BROWSER_SHAPE_TARGETS, shape);

const browserAcceptances = knownFailures.known.flatMap(({ fingerprint, acceptance }) => {
  if (acceptance?.type !== "browser") return [];
  const trace = acceptance.trace;
  if (!trace || !isBrowserShape(trace.shape)) throw new Error("Invalid browser acceptance shape");
  const actions = trace.actions.map((action) => {
    switch (action.kind) {
      case "dragCellDelete":
        return { kind: "dragCellDelete" } as const;
      case "typing":
        if (typeof action.text !== "string") throw new Error("Invalid acceptance typing action");
        return { kind: "typing", text: action.text } as const;
      default:
        throw new Error(`Unsupported acceptance action ${action.kind}`);
    }
  });
  return [
    {
      ...acceptance,
      fingerprint,
      trace: { shape: trace.shape, actions } satisfies BrowserInputTrace,
    },
  ];
});

// A changed symptom or a passing replay requires reviewing and removing its acceptance entry.
for (const acceptance of browserAcceptances) {
  test(`known failure #${acceptance.issue} / seed ${acceptance.reportSeed} / ${acceptance.fingerprint}: ${acceptance.title}`, async ({
    page,
  }) => {
    const { trace } = acceptance;
    const source = await shapeArrayBuffer(trace.shape);
    const baseline = project(await FolioDocxReviewer.fromBuffer(source));
    const edited = await runMode(page, source, baseline, trace, false);
    const suggested = await runMode(page, source, baseline, trace, true);
    const accepting = await FolioDocxReviewer.fromBuffer(suggested.buffer);
    accepting.acceptAll();
    const accepted = await FolioDocxReviewer.fromBuffer(await accepting.toBuffer());
    const editedCellIndex = edited.blocks.findIndex(
      ({ text, table }) => text === "alpha" && table !== undefined,
    );
    if (editedCellIndex < 0) throw new Error("#1341 trace did not edit a table cell to alpha");
    const editedCell = edited.blocks[editedCellIndex];
    const acceptedCell = project(accepted).at(editedCellIndex);
    if (editedCell === undefined || acceptedCell === undefined) {
      throw new Error("#1341 trace lost its edited table cell while accepting");
    }
    if (JSON.stringify(acceptedCell.table) !== JSON.stringify(editedCell.table)) {
      throw new Error("#1341 trace accepted a different table cell");
    }
    if (editedCell.text !== "alpha" || acceptedCell.text !== "a") {
      throw new Error(
        `#1341 expected editing alpha and accepting a; got ${JSON.stringify(editedCell.text)} and ${JSON.stringify(acceptedCell.text)}`,
      );
    }

    // Mark the test expected-to-fail only after setup and the exact symptom pass.
    // A setup regression fails normally; a fix fails at the final equality.
    test.fail(true, "#1341: accepting tracked table input truncates alpha to a");
    expect(acceptedCell.text).toBe(editedCell.text);
  });
}

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

// Derive deterministic drag fixtures from the generator's shape map, so every
// shape with a target gets this gesture invariant in both editor modes.
const browserDragShapes = Object.keys(BROWSER_SHAPE_TARGETS).filter(
  (shape): shape is keyof typeof BROWSER_SHAPE_TARGETS =>
    isBrowserShape(shape) && BROWSER_SHAPE_TARGETS[shape] !== null,
);
for (const shape of browserDragShapes) {
  const target = BROWSER_SHAPE_TARGETS[shape];
  if (target === null) throw new Error(`browser drag shape ${shape} has no target`);
  for (const suggesting of [false, true]) {
    test(`painted ${target} drag in ${shape} after paste selects its planned range (${suggesting ? "suggesting" : "editing"})`, async ({
      page,
    }) => {
      const source = await shapeArrayBuffer(shape);
      const baseline = project(await FolioDocxReviewer.fromBuffer(source));
      await load(page, source, baseline, suggesting);
      await paste(page, {
        kind: "pasteHtml",
        plain: "First bold\nSecond",
        html: "<p>First <strong>bold</strong></p><p>Second</p>",
      });
      expect(await selectTarget(page, target)).toBe(true);
    });
  }
}

const config = parseBrowserInputTraceConfig(
  process.env,
  process.env["FOLIO_FUZZ_LANE"] === "nightly" ? "nightly" : "pullRequest",
);
const replayPath = process.env["PROPERTY_TEST_PATH"];
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
    reportFuzzHealth({ status: "started", completed: 0 });
    const verdict = await fc.check(
      fc.asyncProperty(INPUT_TRACE, async (trace) => {
        checkFreshRender(trace);
        const source = await shapeArrayBuffer(trace.shape);
        const baseline = project(await FolioDocxReviewer.fromBuffer(source));
        const edited = await runMode(page, source, baseline, trace, false);
        const suggested = await runMode(page, source, baseline, trace, true);
        const accepting = await FolioDocxReviewer.fromBuffer(suggested.buffer);
        accepting.acceptAll();
        const accepted = await reopenSaved(await accepting.toBuffer());
        expect(project(accepted)).toEqual(edited.blocks);
        const rejecting = await FolioDocxReviewer.fromBuffer(suggested.buffer);
        rejecting.rejectAll();
        const rejected = await reopenSaved(await rejecting.toBuffer());
        expect(project(rejected)).toEqual(baseline);
        if (JSON.stringify(edited.blocks) !== JSON.stringify(baseline)) {
          expect(suggested.changes.length).toBeGreaterThan(0);
        }
      }),
      {
        seed,
        numRuns: config.runs,
        endOnFailure: false,
        ...(replayPath ? { path: replayPath } : {}),
      },
    );
    const health = classifyFuzzRun(verdict);
    reportFuzzHealth(health);
    if (health.status === "infrastructure") {
      throw new Error(`Fuzz infrastructure: ${health.detail}`);
    }
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
          process.env["FOLIO_FUZZ_FAILURES_DIR"] ?? "fuzz-artifacts/browser/findings",
          failureRecord(marker, failure, { flow: trace }),
        );
        if (
          browserAcceptances.some(
            (acceptance) =>
              seed === acceptance.reportSeed &&
              marker.fingerprint === acceptance.fingerprint &&
              marker.primary === acceptance.primary,
          )
        ) {
          test.fail(true, "#1341 expected browser-fuzz failure");
        }
      }
      const detail =
        failure instanceof Error ? (failure.stack ?? failure.message) : fc.stringify(failure);
      logFailureMarker(
        failureMarker({
          test: "browser input preserves readers, fresh render, and suggesting equivalence",
          seed,
          path: verdict.counterexamplePath,
          repro: `FOLIO_FUZZ_SEEDS=${seed} FOLIO_FUZZ_RUNS=${config.runs} PROPERTY_TEST_PATH=${verdict.counterexamplePath} bunx playwright test --project=browser-fuzzer --workers=1`,
          failure,
          flow: fc.stringify(verdict.counterexample?.at(0)),
        }),
      );
      throw new Error(
        `seed=${seed} path=${verdict.counterexamplePath} trace=${JSON.stringify(verdict.counterexample?.at(0))}\n${detail}`,
        { cause: failure },
      );
    }
  });
}
