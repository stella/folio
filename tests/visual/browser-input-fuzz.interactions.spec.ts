import { expect, type Page, test, type TestInfo } from "@playwright/test";
import fc from "fast-check";

import { shapeArrayBuffer } from "../../packages/core/src/__tests__/documentShapes";
import { SUGGESTION_INPUT_KINDS } from "../../packages/core/src/__tests__/suggestionInputKinds";
import { FolioDocxReviewer } from "../../packages/core/src/ai-edits/headless";
import {
  createFreshRenderRig,
  docFromParagraphs,
} from "../../packages/core/src/controller/__tests__/freshRenderHarness";
import { parseDocx } from "../../packages/core/src/docx/parser";
import { withFakeTextMeasure } from "../../packages/core/src/layout-engine/measure/__tests__/fakeTextMeasure";
import { validateDocxPackage } from "../../packages/docx-core/src/validate/docx";
import type { DocxEditorRef } from "../../packages/react/src/components/DocxEditor.props";
import {
  BROWSER_FUZZ_BUDGET,
  checkWithBoundedShrink,
  DEFAULT_EDITOR_FUZZ_BUDGET,
} from "../../test/bounded-async-fuzz";
import { createMissingOpBurndown } from "../../test/canonical-missing-ops";
import {
  canonicalActivationRefusalRow,
  canonicalBrowserRefusalRows,
  matchCanonicalRefusalRow,
  validateHarnessRefusalRows,
} from "../../test/canonical-refusal-rows";
import {
  failureMarker,
  failureRecord,
  logFailureMarker,
  writeFailureRecord,
} from "../../test/consumer-scenarios/support/failure-fingerprints";
import { reportFuzzHealth } from "../../test/consumer-scenarios/support/fuzz-health";
import { classifyFuzzRun } from "../../test/fuzz-health";
import type {} from "../parity/canonicalFuzzErrors";
import { evaluateCanonicalLoad } from "../parity/canonicalLoadEvaluation";
import type { BrowserDragTarget } from "./browserDragTarget";
import { driveBrowserIme } from "./browserImeDriver";
import {
  BROWSER_INPUT_REGRESSIONS,
  PINNED_BROWSER_INPUT_TAG,
} from "./browserInputRegressions";
import {
  BROWSER_PARAGRAPH_DRAG_SHAPES,
  BROWSER_PASTE_PAYLOADS,
  BROWSER_SHAPE_TARGETS,
  type BrowserInputAction,
  type BrowserInputTrace,
  browserInputTraceArbitrary,
  browserSuggestionActionKinds,
  parseBrowserInputTraceConfig,
} from "./browserInputTrace";
import type {} from "./browserTestBridge";
import { clipboardHtmlProjection } from "./clipboardHtmlProjection";
import { PINNED_REPLAY_FONT_DELAYS_MS } from "./pinnedReplayRepeats";

declare global {
  var __folioPlayground:
    { getEditorRef: () => DocxEditorRef | null } | undefined;
}

const MODIFIER = process.platform === "darwin" ? "Meta" : "Control";

/**
 * What the running pinned replay observed: each drag's selection and each
 * mode's saved outcome. Repeated under other font timings, it must observe the
 * same. Generated runs record nothing.
 */
let replayObservations: unknown[] | null = null;
const INPUT_TRACE = browserInputTraceArbitrary();
const missingCanonical = createMissingOpBurndown();
test.afterAll(() => console.info(missingCanonical.markdown()));

const drainCanonicalErrors = (page: Page) =>
  page.evaluate(() => {
    if (!globalThis.__folioCanonicalFuzzErrors)
      throw new TypeError("Canonical refusal sink unavailable.");
    return globalThis.__folioCanonicalFuzzErrors.splice(0);
  });

type Block = {
  kind: string;
  text: string;
  displayLabel?: string;
  effectiveLevel?: number;
  table?: unknown;
};
type LiveBlock = Pick<Block, "kind" | "text" | "table">;
const project = (reviewer: FolioDocxReviewer): Block[] =>
  reviewer.snapshot().blocks.map((block) => {
    const projection: Block = { kind: block.kind, text: block.text };
    if (block.table !== undefined) projection.table = block.table;
    if (block.kind === "diagnostic") return projection;
    if (block.displayLabel !== undefined)
      projection.displayLabel = block.displayLabel;
    if (block.listReference !== undefined)
      projection.effectiveLevel = block.listReference.level;
    return projection;
  });
const projectLive = (blocks: readonly Block[]): LiveBlock[] =>
  blocks.map(({ kind, text, table }) => ({ kind, text, table }));

const liveBlocks = (page: Page) =>
  page.evaluate(() => {
    const snapshot = globalThis.__folioPlayground
      ?.getEditorRef()
      ?.createAIEditSnapshot();
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
    if (!ref?.getEditorRef()?.getView())
      throw new Error("browser editor unavailable");
    return true;
  });

type FuzzAuthority = "canonical" | "prosemirror";
type LoadFuzzOptions = {
  page: Page;
  bytes: ArrayBuffer;
  baseline: Block[];
  suggesting: boolean;
  authority: FuzzAuthority;
};
const LOAD_TIMEOUT_MS = 60_000;

const load = async ({
  page,
  bytes,
  baseline,
  suggesting,
  authority,
}: LoadFuzzOptions) => {
  const activationRow =
    authority === "canonical"
      ? canonicalActivationRefusalRow(
          await parseDocx(bytes, {
            preloadFonts: false,
            detectVariables: false,
          }),
        )
      : undefined;
  if (activationRow) validateHarnessRefusalRows([activationRow]);
  await page.goto(authority === "canonical" ? "/?session=canonical" : "/");
  // A page that cannot mount fails here, not at the test's long fuzz budget.
  await page.waitForSelector(".layout-page", { timeout: LOAD_TIMEOUT_MS });
  await page.evaluate(() =>
    globalThis.__folioPlayground?.getEditorRef()?.ensureEditorView(),
  );
  await page.waitForFunction(
    () =>
      !!globalThis.__folioPlayground?.getEditorRef()?.getEditorRef()?.getView(),
    undefined,
    { timeout: LOAD_TIMEOUT_MS },
  );
  if (authority === "canonical")
    await page.evaluate(() => {
      globalThis.__folioCanonicalFuzzErrors = [];
    });
  if (authority === "canonical" && !activationRow) {
    if (!(await evaluateCanonicalLoad(page, [...new Uint8Array(bytes)])))
      throw new TypeError("Canonical fuzz load unavailable.");
  } else {
    await page.evaluate(
      async (source) => {
        const ref = globalThis.__folioPlayground?.getEditorRef();
        if (!ref) throw new Error("browser editor unavailable");
        await ref.loadDocumentBuffer(new Uint8Array(source));
      },
      [...new Uint8Array(bytes)],
    );
  }
  if (activationRow) {
    await page.waitForFunction(
      () => (globalThis.__folioCanonicalFuzzErrors?.length ?? 0) > 0,
    );
    const errors = await drainCanonicalErrors(page);
    expect(errors).toHaveLength(1);
    for (const error of errors) {
      expect(error.status).toBe("refusal");
      if (error.status !== "refusal") throw new TypeError(error.message);
      expect(
        matchCanonicalRefusalRow({ rows: [activationRow], refusal: error }),
      ).toEqual(activationRow);
    }
    expect(
      await page.evaluate(() => {
        const ref = globalThis.__folioPlayground?.getEditorRef();
        const core = ref?.getEditor();
        return {
          view: ref?.getEditorRef()?.getView() != null,
          owner: core?.getCanonicalDocument() != null,
          undo: core?.canUndo() ?? false,
          redo: core?.canRedo() ?? false,
        };
      }),
    ).toEqual({ view: false, owner: false, undo: false, redo: false });
    missingCanonical.record(activationRow.gap);
    return { type: "activationRefused", row: activationRow } as const;
  }
  await expect.poll(() => liveBlocks(page)).toEqual(projectLive(baseline));
  await editor(page);
  if (authority === "canonical") {
    expect(
      await page.evaluate((tracked) => {
        const core = globalThis.__folioPlayground?.getEditorRef()?.getEditor();
        if (!core?.getCanonicalDocument())
          throw new TypeError("Canonical fuzz authority unavailable.");
        return core.setCanonicalMode(
          tracked
            ? { type: "suggesting", author: "Browser conformance" }
            : { type: "editing" },
        );
      }, suggesting),
    ).toBe(true);
  } else if (suggesting)
    await page
      .getByRole("button", { name: "Track Changes", exact: true })
      .click();
  await page.evaluate(() =>
    globalThis.__folioPlayground
      ?.getEditorRef()
      ?.getEditorRef()
      ?.getView()
      ?.focus(),
  );
  if (authority === "canonical")
    expect(await drainCanonicalErrors(page)).toEqual([]);
  return { type: "ready" } as const;
};

const selectTableTarget = async (page: Page) => {
  const target = await page.evaluate(() => {
    const bridge = globalThis.__folioBrowserTestBridge;
    const ref = globalThis.__folioPlayground?.getEditorRef();
    if (!bridge || !ref)
      throw new Error("bundled browser editor bridge unavailable");
    return bridge.resolvePaintedTableTarget(ref);
  });
  if (target.type === "absent") return false;
  await page.mouse.move(target.from.x, target.from.y);
  await page.mouse.down();
  await page.mouse.move(target.to.x, target.to.y, { steps: 4 });
  await page.mouse.up();

  // A missed drag must fail at the input boundary, before paste/delete can edit
  // whichever caret the previous action happened to leave in each mode.
  const selected = await page.evaluate(() => {
    const selection = globalThis.__folioPlayground
      ?.getEditorRef()
      ?.getEditorRef()
      ?.getView()?.state.selection;
    if (
      !selection ||
      !("$anchorCell" in selection) ||
      !("$headCell" in selection)
    )
      return null;
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
  replayObservations?.push({ drag: "table", selected });
  expect(selected).toEqual({ anchor: target.anchor, head: target.head });
  return true;
};

const selectTarget = async (page: Page, target: BrowserDragTarget) => {
  if (target === "table") return selectTableTarget(page);
  const painted = await page.evaluate((wanted) => {
    const bridge = globalThis.__folioBrowserTestBridge;
    const ref = globalThis.__folioPlayground?.getEditorRef();
    if (!bridge || !ref)
      throw new Error("bundled browser editor bridge unavailable");
    return bridge.resolvePaintedDragTarget(ref, wanted);
  }, target);
  if (painted.type === "absent") return false;
  const { positions } = painted;
  expect(positions.to).toBeGreaterThan(positions.from);
  await page.mouse.move(painted.from.x, painted.from.y);
  await page.mouse.down();
  await page.mouse.move(painted.to.x, painted.to.y, { steps: 4 });
  await page.mouse.up();
  const selected = await page.evaluate(() => {
    const selection = globalThis.__folioPlayground
      ?.getEditorRef()
      ?.getEditorRef()
      ?.getView()?.state.selection;
    return selection ? { from: selection.from, to: selection.to } : null;
  });
  replayObservations?.push({ drag: target, selected });
  expect(selected).toEqual({ from: positions.from, to: positions.to });
  if (positions.targetType === "inline") {
    expect(selected?.from).toBeLessThanOrEqual(positions.targetFrom);
    expect(selected?.to).toBeGreaterThanOrEqual(positions.targetTo);
  }
  return true;
};

const paste = async (
  page: Page,
  action: Extract<BrowserInputAction, { html: string }>,
) => {
  const expectedHtml = await page.evaluate(
    clipboardHtmlProjection,
    action.html,
  );
  await page.context().grantPermissions(["clipboard-read", "clipboard-write"]);
  await page.evaluate(async ({ html, plain }) => {
    const parts: Record<string, Blob> = {
      "text/plain": new Blob([plain], { type: "text/plain" }),
    };
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
  await expect(page.locator("html")).toHaveAttribute(
    "data-fuzz-paste-plain",
    action.plain,
  );
};

const PROVISIONAL_TEXT_AT_CARET =
  "a composition shows its text in progress at the caret";

/**
 * Drives one action. With `readsProvisional`, a composition's text in
 * progress must show where its caret was after every update.
 */
const drive = async (
  page: Page,
  action: BrowserInputAction,
  { readsProvisional = false }: { readsProvisional?: boolean } = {},
) => {
  switch (action.kind) {
    case "typing":
      await page.keyboard.type(action.text);
      return;
    case "enter":
      await page.keyboard.press("Enter");
      return;
    case "keys":
      for (const key of action.keys) await page.keyboard.press(key);
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
      const readings = await driveBrowserIme(page, action, {
        readProvisional: readsProvisional,
      });
      // The browser may keep part of an earlier update beside the new one;
      // the text around the caret stays as it was.
      for (const { update, before, after, actual } of readings) {
        const between =
          actual !== null &&
          actual.length >= before.length + after.length &&
          actual.startsWith(before) &&
          actual.endsWith(after)
            ? actual.slice(before.length, actual.length - after.length)
            : null;
        expect(
          between?.includes(update) ?? false,
          `${PROVISIONAL_TEXT_AT_CARET}: ${JSON.stringify(update)} between ${JSON.stringify(before)} and ${JSON.stringify(after)}, shown ${JSON.stringify(actual)}`,
        ).toBe(true);
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
        await page.keyboard.press(
          key === "undo" ? `${MODIFIER}+z` : `${MODIFIER}+Shift+z`,
        );
      }
      return;
    case "selectionDrag":
      // An earlier action can remove the target, in one mode and not the other.
      // Such a trace states no edit to compare.
      fc.pre(await selectTarget(page, action.target));
      return;
    case "selectAll":
      await page.keyboard.press(`${MODIFIER}+a`);
      return;
    default: {
      const unreachable: never = action;
      return unreachable;
    }
  }
};

const save = async (page: Page) => {
  const bytes = await page.evaluate(async () => {
    const saved = await globalThis.__folioPlayground
      ?.getEditorRef()
      ?.save({ selective: false });
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

const canonicalSnapshot = async (page: Page) => {
  await page.waitForFunction(() => globalThis.__folioCanonical?.canSnapshot());
  return page.evaluate(() => {
    const ref = globalThis.__folioPlayground?.getEditorRef();
    const core = ref?.getEditor();
    const view = ref?.getEditorRef()?.getView();
    const document = core?.getCanonicalDocument();
    if (!document || !view || !core)
      throw new TypeError("Canonical action owner unavailable.");
    return {
      document,
      projection: view.state.doc.toJSON(),
      selection: view.state.selection.toJSON(),
      undo: core.canUndo(),
      redo: core.canRedo(),
    };
  });
};

type DriveFuzzOptions = {
  page: Page;
  action: BrowserInputAction;
  authority: FuzzAuthority;
};
const driveWithAuthority = async ({
  page,
  action,
  authority,
}: DriveFuzzOptions) => {
  if (authority === "prosemirror") return drive(page, action);
  expect(await drainCanonicalErrors(page)).toEqual([]);
  const before = await canonicalSnapshot(page);
  const rows = canonicalBrowserRefusalRows(action, {
    selection: before.selection["type"] === "cell" ? "cells" : "other",
  });
  validateHarnessRefusalRows(rows);
  await drive(page, action, { readsProvisional: true });
  const after = await canonicalSnapshot(page);
  const errors = await drainCanonicalErrors(page);
  if (rows.length > 0) expect(errors.length).toBeGreaterThan(0);
  for (const error of errors) {
    if (error.status !== "refusal") throw new TypeError(error.message);
    const row = matchCanonicalRefusalRow({ rows, refusal: error });
    expect(
      row,
      `Unregistered canonical refusal: ${error.gap}: ${error.message}`,
    ).toBeDefined();
    if (!row) throw new TypeError("Unregistered canonical refusal.");
    missingCanonical.record(row.gap);
  }
  if (errors.length > 0) expect(after).toEqual(before);
};

type RunFuzzOptions = {
  page: Page;
  source: ArrayBuffer;
  baseline: Block[];
  trace: BrowserInputTrace;
  suggesting: boolean;
  authority: FuzzAuthority;
};
const runMode = async ({
  page,
  source,
  baseline,
  trace,
  suggesting,
  authority,
}: RunFuzzOptions) => {
  const loaded = await load({
    page,
    bytes: source,
    baseline,
    suggesting,
    authority,
  });
  if (loaded.type === "activationRefused") {
    replayObservations?.push({ authority, suggesting, refused: loaded.row });
    return { type: "activationRefused", row: loaded.row } as const;
  }
  for (const action of trace.actions)
    await driveWithAuthority({ page, action, authority });
  await page.waitForTimeout(350);
  const live = await liveBlocks(page);
  const painted = await page.locator(".layout-page-content").allTextContents();
  const buffer = await save(page);
  const reopened = await reopenSaved(buffer);
  expect(
    projectLive(project(reopened)),
    "the saved package reads as the live editor did",
  ).toEqual(live);
  if (authority === "canonical") {
    if (!(await evaluateCanonicalLoad(page, [...new Uint8Array(buffer)])))
      throw new TypeError("Canonical saved load unavailable.");
  } else {
    await page.evaluate(
      async (saved) => {
        await globalThis.__folioPlayground
          ?.getEditorRef()
          ?.loadDocumentBuffer(new Uint8Array(saved));
      },
      [...new Uint8Array(buffer)],
    );
  }
  await expect
    .poll(() => page.locator(".layout-page-content").allTextContents(), {
      message: "a fresh load paints what the page showed before it",
    })
    .toEqual(painted);
  const changes = reopened.getChanges();
  replayObservations?.push({
    authority,
    suggesting,
    blocks: project(reopened),
    changes: changes.map(({ type, text }) => ({ type, text })),
  });
  return {
    type: "ready",
    buffer,
    blocks: project(reopened),
    changes,
  } as const;
};

type ResolveFuzzOptions = {
  page: Page;
  source: ArrayBuffer;
  revisionIds: readonly number[];
  decision: "accept" | "reject";
};
const resolveCanonicalSaved = async ({
  page,
  source,
  revisionIds,
  decision,
}: ResolveFuzzOptions) => {
  if (!(await evaluateCanonicalLoad(page, [...new Uint8Array(source)])))
    throw new TypeError("Canonical fuzz load unavailable.");
  await page.evaluate(
    ({ ids, resolution }) => {
      const ref = globalThis.__folioPlayground?.getEditorRef();
      if (!ref) throw new TypeError("Canonical fuzz editor unavailable.");
      const core = ref.getEditor();
      if (!core?.getCanonicalDocument())
        throw new TypeError("Canonical review authority unavailable.");
      if (ids.length > 0 && !core.resolveCanonicalRevisions(ids, resolution))
        throw new TypeError(
          "Canonical fuzz review refused existing revisions.",
        );
    },
    { ids: revisionIds, resolution: decision },
  );
  return reopenSaved(await save(page));
};

// Each oracle names itself, so failures of two of them are never one class.
const ACCEPTED_EQUALS_EDITED = "accepting tracked input equals the direct edit";
const REJECTED_EQUALS_SOURCE = "rejecting tracked input restores the source";
const TRACKED_INPUT_LEAVES_A_CHANGE = "tracked input leaves a revision";

type CanonicalTraceOptions = {
  page: Page;
  source: ArrayBuffer;
  baseline: Block[];
  trace: BrowserInputTrace;
};
/** One trace through both canonical modes, then its review outcomes against each. */
const checkCanonicalTrace = async ({
  page,
  source,
  baseline,
  trace,
}: CanonicalTraceOptions) => {
  const edited = await runMode({
    page,
    source,
    baseline,
    trace,
    suggesting: false,
    authority: "canonical",
  });
  const suggested = await runMode({
    page,
    source,
    baseline,
    trace,
    suggesting: true,
    authority: "canonical",
  });
  if (
    edited.type === "activationRefused" ||
    suggested.type === "activationRefused"
  ) {
    expect(edited.type).toBe("activationRefused");
    expect(suggested.type).toBe("activationRefused");
    if (
      edited.type !== "activationRefused" ||
      suggested.type !== "activationRefused"
    )
      throw new TypeError("Canonical activation differs between modes.");
    expect(edited.row).toEqual(suggested.row);
    return;
  }
  const accepted = await resolveCanonicalSaved({
    page,
    source: suggested.buffer,
    revisionIds: suggested.changes.map(({ id }) => id),
    decision: "accept",
  });
  expect(project(accepted), ACCEPTED_EQUALS_EDITED).toEqual(edited.blocks);
  const rejected = await resolveCanonicalSaved({
    page,
    source: suggested.buffer,
    revisionIds: suggested.changes.map(({ id }) => id),
    decision: "reject",
  });
  expect(project(rejected), REJECTED_EQUALS_SOURCE).toEqual(baseline);
  if (JSON.stringify(edited.blocks) !== JSON.stringify(baseline)) {
    expect(
      suggested.changes.length,
      TRACKED_INPUT_LEAVES_A_CHANGE,
    ).toBeGreaterThan(0);
  }
};

type DefaultEditorTraceOptions = CanonicalTraceOptions;
/** One trace through the default editor's two modes, then its review outcomes. */
const checkDefaultEditorTrace = async ({
  page,
  source,
  baseline,
  trace,
}: DefaultEditorTraceOptions) => {
  const edited = await runMode({
    page,
    source,
    baseline,
    trace,
    suggesting: false,
    authority: "prosemirror",
  });
  const suggested = await runMode({
    page,
    source,
    baseline,
    trace,
    suggesting: true,
    authority: "prosemirror",
  });
  if (edited.type !== "ready" || suggested.type !== "ready")
    throw new TypeError("The default editor unexpectedly refused activation.");
  const accepting = await FolioDocxReviewer.fromBuffer(suggested.buffer);
  accepting.acceptAll();
  expect(
    project(await reopenSaved(await accepting.toBuffer())),
    ACCEPTED_EQUALS_EDITED,
  ).toEqual(edited.blocks);
  const rejecting = await FolioDocxReviewer.fromBuffer(suggested.buffer);
  rejecting.rejectAll();
  expect(
    project(await reopenSaved(await rejecting.toBuffer())),
    REJECTED_EQUALS_SOURCE,
  ).toEqual(baseline);
};

/** Both modes commit through the canonical session what the default editor commits. */
const checkCommittedAsDefaultEditor = async ({
  page,
  source,
  baseline,
  trace,
}: CanonicalTraceOptions) => {
  for (const suggesting of [false, true]) {
    const canonical = await runMode({
      page,
      source,
      baseline,
      trace,
      suggesting,
      authority: "canonical",
    });
    const viaDefault = await runMode({
      page,
      source,
      baseline,
      trace,
      suggesting,
      authority: "prosemirror",
    });
    if (canonical.type !== "ready" || viaDefault.type !== "ready")
      throw new TypeError("Both editors must open the compared trace.");
    expect(canonical.blocks).toEqual(viaDefault.blocks);
    expect(canonical.changes.map(({ type, text }) => ({ type, text }))).toEqual(
      viaDefault.changes.map(({ type, text }) => ({ type, text })),
    );
  }
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
      const loaded = await load({
        page,
        bytes: source,
        baseline,
        suggesting,
        authority: "canonical",
      });
      if (loaded.type === "activationRefused") continue;
      for (const action of trace.actions)
        await driveWithAuthority({ page, action, authority: "canonical" });
    }
  });
}

const isBrowserShape = (
  shape: string,
): shape is keyof typeof BROWSER_SHAPE_TARGETS =>
  Object.hasOwn(BROWSER_SHAPE_TARGETS, shape);

test.describe("pinned browser input", () => {
  test.setTimeout(60_000);
  for (const regression of BROWSER_INPUT_REGRESSIONS) {
    const { seed, path, fingerprint, trace } = regression;
    test(
      `regression seed ${seed} path ${path} / ${fingerprint}`,
      { tag: PINNED_BROWSER_INPUT_TAG },
      async ({ page }) => {
        // The replay runs once per font timing, the fonts arriving as served
        // and then after the editor has painted, and must observe the same
        // each time: a target read from paint a late face moves is read
        // differently. Routing also keeps the cache from serving the fonts.
        let fontDelay = 0;
        await page.route(/\.woff2(?:\?|$)/u, async (route) => {
          if (fontDelay > 0)
            await new Promise((resolve) => setTimeout(resolve, fontDelay));
          await route.continue();
        });
        const replay = async () => {
          const source = await shapeArrayBuffer(trace.shape);
          const baseline = project(await FolioDocxReviewer.fromBuffer(source));
          if ("authority" in regression) {
            await checkCanonicalTrace({ page, source, baseline, trace });
            if ("comparesWithDefaultEditor" in regression)
              await checkCommittedAsDefaultEditor({
                page,
                source,
                baseline,
                trace,
              });
            return;
          }
          await checkDefaultEditorTrace({ page, source, baseline, trace });
        };
        const observed: { delay: number; observations: unknown[] }[] = [];
        for (const delay of PINNED_REPLAY_FONT_DELAYS_MS) {
          fontDelay = delay;
          const observations: unknown[] = [];
          replayObservations = observations;
          await test.step(`fonts delivered after ${String(delay)} ms`, replay);
          replayObservations = null;
          observed.push({ delay, observations });
        }
        const [first] = observed;
        for (const { delay, observations } of observed)
          expect({ delay, observations }).toEqual({
            delay,
            observations: first?.observations,
          });
      },
    );
  }
});

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
          state.doc.firstChild?.textContent.length
            ? state.tr.delete(1, 2)
            : null,
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
// A drag across paragraphs, in list, note and right-to-left paint.
const paragraphDragShapes = [
  "plain-markdown",
  "single-decimal-list",
  "notes",
  "rtl-cjk",
] as const satisfies readonly (keyof typeof BROWSER_SHAPE_TARGETS)[];
const paintedDrags = [
  ...browserDragShapes.map(
    (dragShape) => [dragShape, BROWSER_SHAPE_TARGETS[dragShape]] as const,
  ),
  ...paragraphDragShapes.map((dragShape) => [dragShape, "paragraphs"] as const),
];
for (const [shape, target] of paintedDrags) {
  if (
    target === null ||
    (target === "paragraphs" && !BROWSER_PARAGRAPH_DRAG_SHAPES.has(shape))
  )
    throw new Error(`browser drag shape ${shape} has no ${target} target`);
  for (const suggesting of [false, true]) {
    test(`painted ${target} drag in ${shape} after paste selects its planned range (${suggesting ? "suggesting" : "editing"})`, async ({
      page,
    }) => {
      const source = await shapeArrayBuffer(shape);
      const baseline = project(await FolioDocxReviewer.fromBuffer(source));
      const loaded = await load({
        page,
        bytes: source,
        baseline,
        suggesting,
        authority: "canonical",
      });
      if (loaded.type === "activationRefused") return;
      await driveWithAuthority({
        page,
        authority: "canonical",
        action: {
          kind: "pasteHtml",
          plain: "First bold\nSecond",
          html: "<p>First <strong>bold</strong></p><p>Second</p>",
        },
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
test.setTimeout(BROWSER_FUZZ_BUDGET.testMs);

test("browser generator covers every declared suggestion input kind", () => {
  expect(new Set(browserSuggestionActionKinds)).toEqual(
    new Set(SUGGESTION_INPUT_KINDS),
  );
});

test("clipboard oracle accepts HTML serialization and detects text, attribute and markup changes", async ({
  page,
}) => {
  await page.goto("/");
  for (const { html } of Object.values(BROWSER_PASTE_PAYLOADS)) {
    const expected = await page.evaluate(clipboardHtmlProjection, html);
    // Clipboard writers parse document wrappers before delivering a paste.
    // Chromium adds the omitted head in the full-document fixture (seed 11).
    expect(await page.evaluate(clipboardHtmlProjection, expected)).toBe(
      expected,
    );
    expect(await page.evaluate(clipboardHtmlProjection, "")).not.toBe(expected);
    const mutations = await page.evaluate((source) => {
      const textDocument = new DOMParser().parseFromString(source, "text/html");
      const walker = textDocument.createTreeWalker(
        textDocument.body,
        NodeFilter.SHOW_TEXT,
      );
      const firstText = walker.nextNode();
      if (!firstText)
        throw new Error("clipboard fixture has no text to mutate");
      firstText.textContent = "lost clipboard text";
      const attributeDocument = new DOMParser().parseFromString(
        source,
        "text/html",
      );
      const element = attributeDocument.body.firstElementChild;
      if (!element)
        throw new Error("clipboard fixture has no element to mutate");
      element.setAttribute("data-fuzz-mutated", "true");
      const markupDocument = new DOMParser().parseFromString(
        source,
        "text/html",
      );
      const wrapper = markupDocument.body.firstElementChild;
      if (!wrapper)
        throw new Error("clipboard fixture has no markup to mutate");
      wrapper.replaceWith(...wrapper.childNodes);
      return [
        textDocument.documentElement.outerHTML,
        attributeDocument.documentElement.outerHTML,
        markupDocument.documentElement.outerHTML,
      ];
    }, html);
    for (const mutation of mutations) {
      expect(await page.evaluate(clipboardHtmlProjection, mutation)).not.toBe(
        expected,
      );
    }
  }
  const source = BROWSER_PASTE_PAYLOADS.pasteWordHtml.html;
  const normalized = source.replace("<body>", "<head></head><body>");
  expect(await page.evaluate(clipboardHtmlProjection, source)).toBe(
    await page.evaluate(clipboardHtmlProjection, normalized),
  );
});

const INPUT_TEST_NAME =
  "browser input preserves readers, fresh render, and suggesting equivalence";
const DEFAULT_EDITOR_TEST_NAME = `default editor ${INPUT_TEST_NAME}`;

type GeneratedLaneOptions = {
  info: TestInfo;
  seed: number;
  testName: string;
  /** The environment a replay of this lane needs, before its path. */
  replayEnv: string;
  budget: { discoveryMs: number; shrinkMs: number };
  check: (trace: BrowserInputTrace) => Promise<void>;
};
/** One seed's generated traces, with failures reported the same way for every lane. */
const runGeneratedLane = async ({
  info,
  seed,
  testName,
  replayEnv,
  budget,
  check,
}: GeneratedLaneOptions) => {
  reportFuzzHealth({ status: "started", completed: 0 });
  const { verdict } = await checkWithBoundedShrink({
    arbitrary: INPUT_TRACE,
    evaluate: async (trace) => {
      checkFreshRender(trace);
      await check(trace);
    },
    seed,
    numRuns: config.runs,
    discoveryMs: budget.discoveryMs,
    shrinkMs: budget.shrinkMs,
    ...(replayPath ? { path: replayPath } : {}),
    onFirstFailure: async ({
      seed: originalSeed,
      path,
      value: trace,
      error: failure,
    }) => {
      const marker = failureMarker({
        test: testName,
        seed: originalSeed,
        path,
        repro: `${replayEnv}=${originalSeed} FOLIO_FUZZ_RUNS=${config.runs} PROPERTY_TEST_PATH=${path} bunx playwright test --project=browser-fuzzer tests/visual/browser-input-fuzz.interactions.spec.ts --workers=1`,
        failure,
        flow: `${trace.shape}: ${trace.actions.map(({ kind }) => kind).join(" → ")}`,
      });
      logFailureMarker(marker);
      const artifact = writeFailureRecord(
        "fuzz-artifacts/browser/original",
        failureRecord(marker, failure, { flow: trace }),
      );
      await info.attach("browser-original-failure", {
        path: artifact,
        contentType: "application/json",
      });
    },
  });
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
        test: testName,
        seed,
        path: verdict.counterexamplePath,
        repro: `${replayEnv}=${seed} FOLIO_FUZZ_RUNS=${config.runs} bunx playwright test --project=browser-fuzzer tests/visual/browser-input-fuzz.interactions.spec.ts --workers=1`,
        failure,
        flow: `${trace.shape}: ${trace.actions.map(({ kind }) => kind).join(" → ")}`,
      });
      logFailureMarker(marker);
      writeFailureRecord(
        process.env["FOLIO_FUZZ_FAILURES_DIR"] ??
          "fuzz-artifacts/browser/findings",
        failureRecord(marker, failure, { flow: trace }),
      );
    }
    const detail =
      failure instanceof Error
        ? (failure.stack ?? failure.message)
        : fc.stringify(failure);
    logFailureMarker(
      failureMarker({
        test: testName,
        seed,
        path: verdict.counterexamplePath,
        repro: `${replayEnv}=${seed} FOLIO_FUZZ_RUNS=${config.runs} PROPERTY_TEST_PATH=${verdict.counterexamplePath} bunx playwright test --project=browser-fuzzer --workers=1`,
        failure,
        flow: fc.stringify(verdict.counterexample?.at(0)),
      }),
    );
    throw new Error(
      `seed=${seed} path=${verdict.counterexamplePath} trace=${JSON.stringify(verdict.counterexample?.at(0))}\n${detail}`,
      { cause: failure },
    );
  }
};

// First in the file: when a slow night runs out of step time, the editor people
// use keeps its seeds, and its smaller budget leaves the rest of the step.
for (const seed of config.defaultEditorSeeds) {
  test(`default editor seed ${seed}: ${INPUT_TEST_NAME}`, async ({
    page,
  }, info) => {
    await runGeneratedLane({
      info,
      seed,
      testName: DEFAULT_EDITOR_TEST_NAME,
      replayEnv: "FOLIO_FUZZ_DEFAULT_EDITOR_SEEDS",
      budget: DEFAULT_EDITOR_FUZZ_BUDGET,
      check: async (trace) => {
        const source = await shapeArrayBuffer(trace.shape);
        const baseline = project(await FolioDocxReviewer.fromBuffer(source));
        await checkDefaultEditorTrace({ page, source, baseline, trace });
      },
    });
  });
}

for (const seed of config.seeds) {
  test(`seed ${seed}: ${INPUT_TEST_NAME}`, async ({ page }, info) => {
    await runGeneratedLane({
      info,
      seed,
      testName: INPUT_TEST_NAME,
      replayEnv: "FOLIO_FUZZ_SEEDS",
      budget: BROWSER_FUZZ_BUDGET,
      check: async (trace) => {
        const source = await shapeArrayBuffer(trace.shape);
        const baseline = project(await FolioDocxReviewer.fromBuffer(source));
        await checkCanonicalTrace({ page, source, baseline, trace });
      },
    });
  });
}
