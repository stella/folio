import { afterAll, beforeAll, expect, test } from "bun:test";
import { GlobalRegistrator } from "@happy-dom/global-registrator";
import { EditorState } from "prosemirror-state";
import type { Transaction } from "prosemirror-state";
import { applyFolioAIEditOperations } from "../packages/core/src/ai-edits/apply";
import { getSuggestions } from "../packages/core/src/prosemirror/commands/comments";
import {
  resolveFolioAIBlockRange,
  resolvePassageRange,
} from "../packages/core/src/ai-edits/blockRange";
import { getTrackedChangesFromDoc } from "../packages/core/src/ai-edits/read";
import { createFolioAIEditSnapshot } from "../packages/core/src/ai-edits/snapshot";
import { parseDocx } from "../packages/core/src/docx/parser";
import { findBlockSdtMatches } from "../packages/core/src/prosemirror/commands/contentControls";
import { toProseDoc } from "../packages/core/src/prosemirror/conversion/toProseDoc";
import {
  SCROLL_CONTROL_TAG,
  SCROLL_REVISION_ID,
  SCROLL_TARGET_PARA_ID,
  SCROLL_TARGET_TEXT,
  SCROLL_TARGET_SUGGESTION,
  findScrollSuggestionTarget,
} from "../packages/playground/src/scrollParityBridge";
import { buildScrollRootDocument } from "../tests/support/scrollRootDocument";

let ownsDomGlobals = false;
beforeAll(() => {
  if (GlobalRegistrator.isRegistered) return;
  GlobalRegistrator.register();
  ownsDomGlobals = true;
});
afterAll(() => {
  if (ownsDomGlobals) return GlobalRegistrator.unregister();
});

test("scroll browser fixture preserves every navigation target through the real DOCX parser", async () => {
  const document = await parseDocx(await buildScrollRootDocument());
  const doc = toProseDoc(document);
  const snapshot = createFolioAIEditSnapshot(doc);
  const target = snapshot.blocks.find(({ id }) => id === SCROLL_TARGET_PARA_ID);
  expect(target?.text).toContain(SCROLL_TARGET_TEXT);
  expect(snapshot.anchors[SCROLL_TARGET_PARA_ID]?.from).toBeGreaterThan(1);
  expect(
    resolveFolioAIBlockRange({ blockId: SCROLL_TARGET_PARA_ID, doc, snapshot }),
  ).not.toBeNull();
  expect(
    resolvePassageRange({
      blockId: SCROLL_TARGET_PARA_ID,
      text: SCROLL_TARGET_TEXT,
      doc,
      snapshot,
    }),
  ).not.toBeNull();
  expect(getTrackedChangesFromDoc(doc)).toContainEqual(
    expect.objectContaining({
      id: SCROLL_REVISION_ID,
      type: "insertion",
      blockId: SCROLL_TARGET_PARA_ID,
    }),
  );
  expect(findBlockSdtMatches(doc, { tag: SCROLL_CONTROL_TAG })).toHaveLength(1);
  const boundaries: string[] = [];
  doc.descendants((node) => {
    if (node.type.name === "paragraph" && node.attrs["pageBreakBefore"] === true) {
      boundaries.push(String(node.attrs["paraId"]));
    }
  });
  expect(boundaries).toEqual(["13300200", SCROLL_TARGET_PARA_ID, "13300400"]);
});

test("suggestion scroll oracle measures the inserted block instead of its source", async () => {
  const parsed = await parseDocx(await buildScrollRootDocument());
  const doc = toProseDoc(parsed);
  const snapshot = createFolioAIEditSnapshot(doc);
  const view = {
    state: EditorState.create({ schema: doc.type.schema, doc }),
    dispatch(transaction: Transaction) {
      view.state = view.state.apply(transaction);
    },
  };
  const result = applyFolioAIEditOperations({
    view,
    snapshot,
    mode: "suggested",
    operations: [SCROLL_TARGET_SUGGESTION],
  });
  const suggestionId = result.applied.at(0)?.suggestionId;
  expect(suggestionId).toBeDefined();
  const suggestion = getSuggestions(view.state).find(
    (candidate) => candidate.suggestionId === suggestionId,
  );
  expect(suggestion).toBeDefined();
  if (!suggestion) return;
  const liveSnapshot = createFolioAIEditSnapshot(view.state.doc);
  const root = document.createElement("div");
  for (const block of liveSnapshot.blocks) {
    const anchor = liveSnapshot.anchors[block.id];
    if (!anchor) continue;
    const paragraph = document.createElement("div");
    paragraph.className = "layout-paragraph";
    paragraph.dataset["pmStart"] = String(anchor.from);
    paragraph.textContent = block.text;
    root.append(paragraph);
  }
  const target = findScrollSuggestionTarget({ root, snapshot: liveSnapshot, suggestion });
  expect(target?.textContent).toBe(SCROLL_TARGET_SUGGESTION.text);
  expect(target?.dataset["pmStart"]).not.toBe(
    String(liveSnapshot.anchors[SCROLL_TARGET_PARA_ID]?.from),
  );
  target?.remove();
  expect(findScrollSuggestionTarget({ root, snapshot: liveSnapshot, suggestion })).toBeNull();
});

// Each host owns its public-package imports; their browser operations must stay identical.
test("scroll bridges stay identical across adapters", async () => {
  const react = await Bun.file(
    new URL("../packages/playground/src/scrollParityBridge.ts", import.meta.url),
  ).text();
  const vue = await Bun.file(
    new URL("../packages/playground-vue/src/scrollParityBridge.ts", import.meta.url),
  ).text();
  expect(vue.replaceAll("@stll/folio-vue", "@stll/folio-react")).toBe(react);
});
