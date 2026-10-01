import { expect, test } from "bun:test";
import path from "node:path";
import { EditorState } from "prosemirror-state";
import { createFolioAIEditSnapshot } from "../../packages/core/src/ai-edits/snapshot";
import { applyFolioDocumentOperations } from "../../packages/core/src/document-operations";
import { getTrackedChangesFromDoc } from "../../packages/core/src/ai-edits/read";
import { parseDocx } from "../../packages/core/src/docx/parser";
import { toProseDoc } from "../../packages/core/src/prosemirror/conversion/toProseDoc";
import { shapeArrayBuffer } from "../../packages/core/src/__tests__/documentShapes";
import { createInterleavingSuggest } from "./interleavingBridge";

// Compile the actual browser entry, following every runtime import. A future
// server-barrel or node: import must fail before the browser waits for a global.
test("interleaving bridge bundles without Node-only dependencies", async () => {
  const result = await Bun.build({
    entrypoints: [path.join(import.meta.dir, "interleavingBridge.ts")],
    target: "browser",
    write: false,
  });
  expect(result.logs.map(String)).toEqual([]);
  expect(result.success).toBe(true);
  expect(result.outputs.length).toBeGreaterThan(0);
});

test("live interleaving edits retain tracked mode, guarded targets and distinct operation ids", async () => {
  const doc = await parseDocx(await shapeArrayBuffer("plain-markdown"));
  const view = {
    state: EditorState.create({ doc: toProseDoc(doc) }),
    dispatch: (tr: Parameters<EditorState["apply"]>[0]) => {
      view.state = view.state.apply(tr);
    },
  };
  const ids: string[] = [];
  const suggest = createInterleavingSuggest({
    createAIEditSnapshot: () => createFolioAIEditSnapshot(view.state.doc),
    applyDocumentOperations: ({ snapshot, batch, mode, author }) => {
      const operation = batch.operations.at(0);
      if (!operation || operation.type !== "insertAfterBlock") {
        throw new Error("Missing interleaving insertion");
      }
      ids.push(operation.id);
      expect(batch.version).toBe(1);
      expect(batch.mode).toBe("tracked-changes");
      expect(mode).toBe("tracked-changes");
      expect(operation.precondition?.blockTextHash).toBe(
        snapshot.anchors[operation.blockId]?.textHash,
      );
      return applyFolioDocumentOperations({ view, snapshot, batch, author });
    },
    getTrackedChanges: () => getTrackedChangesFromDoc(view.state.doc),
  });
  expect(suggest("Model authored insertion")).toBeGreaterThan(0);
  expect(suggest("Another insertion")).toBeGreaterThan(1);
  expect(new Set(ids).size).toBe(2);
  expect(view.state.doc.textContent).toContain("Model authored insertion");
  expect(view.state.doc.textContent).toContain("Another insertion");
});
