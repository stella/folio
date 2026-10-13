import { expect, test } from "bun:test";
import { EditorState } from "prosemirror-state";
import { acceptAllChanges } from "../prosemirror/commands/comments";
import { FOLIO_DOCUMENT_OPERATION_MODES_BY_TYPE } from "../document-operations";
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

// Typed bookmarks have a lossless inline representation when their break goes.
for (const mode of FOLIO_DOCUMENT_OPERATION_MODES_BY_TYPE.mergeBlockWithNext) {
  for (const count of [0, 1, 2, 4]) {
    for (const pendingInsertion of [false, true]) {
      for (const separator of ["", " "]) {
        test(`merge carries typed bookmarks ${mode} ${count} ${pendingInsertion} ${separator.length}`, () => {
          const boundaries = Array.from({ length: count }, (_, index) =>
            schema.node("blockBookmarkBoundary", {
              type: index % 2 === 0 ? "start" : "end",
              id: index,
              ...(index % 2 === 0 ? { name: `bookmark${index}` } : {}),
            }),
          );
          const doc = schema.node("doc", null, [
            schema.node(
              "paragraph",
              {
                paraId: "11000001",
                ...(pendingInsertion
                  ? {
                      pPrMark: {
                        kind: "ins",
                        info: { id: 900, author: "Reviewer", date: "2000-01-01" },
                      },
                    }
                  : {}),
              },
              [schema.text("Left")],
            ),
            ...boundaries,
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
          const result = applyFolioAIEditOperations({
            view,
            snapshot,
            mode,
            operations: [{ id: "merge", type: "mergeBlockWithNext", blockId: first.id, separator }],
          });
          expect(result.skipped).toEqual([]);
          expect(result.applied.map(({ id }) => id)).toEqual(["merge"]);
          acceptAllChanges()(view.state, view.dispatch);
          expect(() => view.state.doc.check()).not.toThrow();
          expect(view.state.doc.childCount).toBe(1);
          expect(view.state.doc.textContent).toBe(`Left${separator}Right`);
          expect(
            view.state.doc.firstChild?.content.content
              .filter((node) => node.type.name === "bookmarkBoundary")
              .map((node) => node.attrs["id"]),
          ).toEqual(boundaries.map((node) => node.attrs["id"]));
        });
      }
    }
  }
}
