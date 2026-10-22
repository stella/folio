import { describe, expect, spyOn, test } from "bun:test";
import { Node as PMNode } from "prosemirror-model";
import { EditorState } from "prosemirror-state";

import { paragraphNumberingFromSlots } from "@stll/docx-core/model";
import { schema } from "../schema";
import { createDocumentNumberingPlugin, getDocumentNumbering } from "./documentNumbering";

const paragraph = (text: string, attrs: Record<string, unknown> = {}) =>
  schema.node("paragraph", attrs, [schema.text(text)]);

describe("document numbering invalidation", () => {
  test("plain text and mark changes do not traverse unrelated paragraphs", () => {
    const doc = schema.node(
      "doc",
      null,
      Array.from({ length: 200 }, (_, index) => paragraph(`Paragraph ${index}`)),
    );
    const state = EditorState.create({
      doc,
      plugins: [createDocumentNumberingPlugin(null)],
    });
    const original = PMNode.prototype.descendants;
    let paragraphVisits = 0;
    const traversal = spyOn(PMNode.prototype, "descendants").mockImplementation(function (
      this: PMNode,
      callback,
    ) {
      return original.call(this, (node, pos, parent, index) => {
        if (node.type.name === "paragraph") paragraphVisits += 1;
        return callback(node, pos, parent, index);
      });
    });

    try {
      const inserted = state.apply(state.tr.insertText("x", 2));
      expect(paragraphVisits).toBe(0);

      inserted.apply(inserted.tr.addMark(1, 2, schema.mark("bold")));
      expect(paragraphVisits).toBe(0);
    } finally {
      traversal.mockRestore();
    }
  });

  test("adding paragraph numbering recomputes document numbering", () => {
    const doc = schema.node("doc", null, [paragraph("Item"), paragraph("Remaining")]);
    const state = EditorState.create({
      doc,
      plugins: [createDocumentNumberingPlugin(null)],
    });
    const first = state.doc.firstChild;
    if (!first) throw new TypeError("The test document has no paragraph");
    const numbered = state.apply(
      state.tr.setNodeMarkup(0, undefined, {
        ...first.attrs,
        numPr: paragraphNumberingFromSlots({ numId: 42, ilvl: 0 }),
        listMarker: "1.",
        listMarkerTemplate: "%1.",
        listNumFmt: "decimal",
        listAbstractNumId: 42,
      }),
    );

    expect(getDocumentNumbering(numbered)?.hasNumbering(42)).toBe(true);

    const withoutNumberedParagraph = numbered.apply(
      numbered.tr.delete(0, numbered.doc.child(0).nodeSize),
    );
    expect(getDocumentNumbering(withoutNumberedParagraph)?.hasNumbering(42) ?? false).toBe(false);
  });
});
