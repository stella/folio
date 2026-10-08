import assert from "node:assert/strict";
import { EditorState, TextSelection } from "prosemirror-state";
import { test } from "bun:test";
import { schema, singletonManager } from "../../schema";

test("legacy collapsed removal retires all formatting segments of one hyperlink", () => {
  const link = schema.mark("hyperlink", { href: "https://example.com/" });
  const doc = schema.node("doc", undefined, [
    schema.node("paragraph", { paraId: "12345678" }, [
      schema.text("first", [link]),
      schema.text("second", [link, schema.mark("bold")]),
    ]),
  ]);
  let state = EditorState.create({ schema, doc, selection: TextSelection.create(doc, 2) });
  assert.equal(
    singletonManager.requireCommand("removeHyperlink")()(state, (transaction) => {
      state = state.apply(transaction);
    }),
    true,
  );
  let linked = false;
  state.doc.descendants((node) => {
    if (node.marks.some((mark) => mark.type.name === "hyperlink")) linked = true;
  });
  assert.equal(linked, false, "legacy collapsed removal leaves part of the same hyperlink linked");
});
