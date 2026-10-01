import { expect, test } from "bun:test";
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
} from "../packages/playground-vue/src/scrollParityBridge";
import { buildScrollRootDocument } from "../tests/support/scrollRootDocument";

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
