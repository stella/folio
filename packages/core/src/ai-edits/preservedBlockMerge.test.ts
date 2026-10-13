import { expect, test } from "bun:test";
import { EditorState } from "prosemirror-state";
import { schema } from "../prosemirror/schema";
import { createFolioAIEditSnapshot } from "./snapshot";
import { applyFolioAIEditOperations } from "./apply";

// Structural-edit fixtures had only joinable paragraphs, never opaque siblings.
for (const xml of [
  '<w:moveFromRangeEnd w:id="1"/>',
  '<w:permStart w:id="7" w:edGrp="everyone"/>',
  '<w:altChunk r:id="rIdChunk"/>',
])
  for (const mode of ["direct", "tracked-changes", "suggested"] as const) {
    for (const separator of ["", " "])
      for (const properties of [undefined, { styleId: null }]) {
        test(`merge preserves opaque sibling ${xml} ${mode} ${separator.length} ${properties !== undefined}`, () => {
          const doc = schema.node("doc", null, [
            schema.node("paragraph", { paraId: "11000001" }, [schema.text("Left")]),
            schema.node("preservedBlock", { xml }),
            schema.node("paragraph", { paraId: "11000002" }, [schema.text("Right")]),
          ]);
          const view = {
            state: EditorState.create({ doc }),
            dispatch: (tr: Parameters<EditorState["apply"]>[0]) => {
              view.state = view.state.apply(tr);
            },
          };
          const snapshot = createFolioAIEditSnapshot(doc);
          const first = snapshot.blocks.at(0);
          if (!first) throw new Error("Missing first paragraph");
          const before = view.state.doc.toJSON();
          const result = applyFolioAIEditOperations({
            view,
            snapshot,
            mode,
            operations: [
              {
                id: "merge",
                type: "mergeBlockWithNext",
                blockId: first.id,
                separator,
                ...(properties === undefined ? {} : { mergedParagraphProperties: properties }),
              },
            ],
          });
          expect(result.applied).toEqual([]);
          expect(result.skipped).toMatchObject([{ id: "merge", reason: "unsupportedBlock" }]);
          expect(view.state.doc.toJSON()).toEqual(before);
        });
      }
  }
