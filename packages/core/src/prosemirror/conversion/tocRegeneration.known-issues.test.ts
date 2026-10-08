import assert from "node:assert/strict";
import { EditorState, TextSelection } from "prosemirror-state";
import {
  expectedFailure,
  FINDING_SYMPTOMS,
} from "../../../../../test/consumer-scenarios/support/known-issues";
import { schema, singletonManager } from "../schema";
import { createDocumentStylesPlugin } from "../plugins/documentStyles";

expectedFailure(
  "LEGACY_TOC_REGENERATION_DANGLING_REFERENCES",
  "legacy TOC regeneration retains existing link and PAGEREF targets",
  FINDING_SYMPTOMS.LEGACY_TOC_REGENERATION_DANGLING_REFERENCES,
  () => {
    const oldName = "_TocExisting";
    let state = EditorState.create({
      schema,
      plugins: [
        createDocumentStylesPlugin({
          styles: [{ type: "paragraph", styleId: "Heading", name: "heading 1" }],
        }),
      ],
      doc: schema.node("doc", null, [
        schema.node("paragraph", { styleId: "Heading", bookmarks: [{ id: 7, name: oldName }] }, [
          schema.text("Heading"),
        ]),
        schema.node("paragraph", null, [
          schema.text("Old entry", [schema.mark("hyperlink", { href: `#${oldName}` })]),
          schema.node("field", {
            fieldKind: "complex",
            fieldType: "PAGEREF",
            instruction: `PAGEREF ${oldName} \\h`,
            displayText: "1",
          }),
        ]),
      ]),
    });
    state = state.apply(state.tr.setSelection(TextSelection.atEnd(state.doc)));
    assert.equal(
      singletonManager.requireCommand("generateTOC")({ title: "Contents" })(
        state,
        (transaction) => {
          state = state.apply(transaction);
        },
      ),
      true,
    );
    const names = new Set<string>();
    let oldLink = false;
    let oldField = false;
    state.doc.descendants((node) => {
      if (node.type.name === "paragraph" && Array.isArray(node.attrs["bookmarks"])) {
        for (const bookmark of node.attrs["bookmarks"])
          if (typeof bookmark.name === "string") names.add(bookmark.name);
      }
      if (
        node.marks.some(
          (mark) => mark.type.name === "hyperlink" && mark.attrs["href"] === `#${oldName}`,
        )
      )
        oldLink = true;
      if (node.type.name === "field" && node.attrs["instruction"] === `PAGEREF ${oldName} \\h`)
        oldField = true;
    });
    assert.equal(oldLink && oldField, true);
    assert.equal(
      names.has(oldName),
      true,
      "legacy TOC regeneration preserves previously referenced bookmark names",
    );
  },
);
