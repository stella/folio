import { describe, expect, setDefaultTimeout, spyOn, test } from "bun:test";
import fc from "fast-check";
import { Node as PMNode } from "prosemirror-model";
import { EditorState } from "prosemirror-state";

import { paragraphNumberingFromSlots } from "@stll/docx-core/model";

import {
  propertyConfig,
  propertyTestTimeout,
} from "../../../../../test/property-testing";
import type { NumberingDefinitions } from "../../types/document";
import { schema } from "../schema";
import {
  createDocumentNumberingPlugin,
  getDocumentNumbering,
} from "./documentNumbering";

// Synchronous generated transaction chains use a 15-second base budget.
setDefaultTimeout(propertyTestTimeout(15_000));

const paragraph = (text: string, attrs: Record<string, unknown> = {}) =>
  schema.node("paragraph", attrs, [schema.text(text)]);

/** The package defines instance 5; 7 and 8 only paragraphs name. */
const DEFINITIONS: NumberingDefinitions = {
  abstractNums: [
    {
      abstractNumId: 5,
      multiLevelType: "singleLevel",
      levels: [
        {
          ilvl: 0,
          start: 1,
          numFmt: "decimal",
          lvlText: "%1.",
          suffix: "space",
        },
      ],
    },
  ],
  nums: [{ numId: 5, abstractNumId: 5 }],
};

const listItem = (text: string, numId: number | null) =>
  paragraph(
    text,
    numId === null
      ? {}
      : {
          numPr: paragraphNumberingFromSlots({ numId, ilvl: 0 }),
          listMarker: "1.",
          listMarkerTemplate: "%1.",
          listNumFmt: "decimal",
          listAbstractNumId: 5,
        },
  );

const countParagraphVisits = (run: () => void): number => {
  const original = PMNode.prototype.descendants;
  let visits = 0;
  const traversal = spyOn(PMNode.prototype, "descendants").mockImplementation(
    function (this: PMNode, callback) {
      return original.call(this, (node, pos, parent, index) => {
        if (node.type.name === "paragraph") visits += 1;
        return callback(node, pos, parent, index);
      });
    },
  );
  try {
    run();
  } finally {
    traversal.mockRestore();
  }
  return visits;
};

describe("document numbering invalidation", () => {
  test("plain text and mark changes do not traverse unrelated paragraphs", () => {
    const doc = schema.node(
      "doc",
      null,
      Array.from({ length: 200 }, (_, index) =>
        paragraph(`Paragraph ${index}`),
      ),
    );
    const state = EditorState.create({
      doc,
      plugins: [createDocumentNumberingPlugin(null)],
    });
    const original = PMNode.prototype.descendants;
    let paragraphVisits = 0;
    const traversal = spyOn(PMNode.prototype, "descendants").mockImplementation(
      function (this: PMNode, callback) {
        return original.call(this, (node, pos, parent, index) => {
          if (node.type.name === "paragraph") paragraphVisits += 1;
          return callback(node, pos, parent, index);
        });
      },
    );

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
    const doc = schema.node("doc", null, [
      paragraph("Item"),
      paragraph("Remaining"),
    ]);
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
    expect(
      getDocumentNumbering(withoutNumberedParagraph)?.hasNumbering(42) ?? false,
    ).toBe(false);
  });

  test("replacing a list item whose instance the package defines leaves the numbering", () => {
    const doc = schema.node(
      "doc",
      null,
      Array.from({ length: 200 }, (_, index) => listItem(`Item ${index}`, 5)),
    );
    const state = EditorState.create({
      doc,
      plugins: [createDocumentNumberingPlugin(DEFINITIONS)],
    });
    const first = state.doc.child(0);
    let replaced = state;
    const visits = countParagraphVisits(() => {
      replaced = state.apply(
        state.tr.replaceWith(0, first.nodeSize, listItem("Item 0x", 5)),
      );
    });
    expect(visits).toBe(0);
    expect(getDocumentNumbering(replaced)).toBe(getDocumentNumbering(state));
  });

  test("keeps the numbering a fresh state computes through any edit", () => {
    const numIds = fc.constantFrom<number | null>(null, 5, 7, 8);
    const edit = fc.oneof(
      fc.record({
        type: fc.constant("replace" as const),
        at: fc.nat(),
        numId: numIds,
      }),
      fc.record({
        type: fc.constant("insert" as const),
        at: fc.nat(),
        numId: numIds,
      }),
      fc.record({ type: fc.constant("delete" as const), at: fc.nat() }),
      fc.record({ type: fc.constant("type" as const), at: fc.nat() }),
      fc.record({ type: fc.constant("join" as const), at: fc.nat() }),
    );
    fc.assert(
      fc.property(
        fc.array(numIds, { minLength: 1, maxLength: 5 }),
        fc.array(edit, { minLength: 1, maxLength: 8 }),
        (initial, edits) => {
          let state = EditorState.create({
            doc: schema.node(
              "doc",
              null,
              initial.map((numId, index) => listItem(`P${index}`, numId)),
            ),
            plugins: [createDocumentNumberingPlugin(DEFINITIONS)],
          });
          for (const step of edits) {
            const count = state.doc.childCount;
            const index = step.at % count;
            let pos = 0;
            for (let at = 0; at < index; at += 1)
              pos += state.doc.child(at).nodeSize;
            const node = state.doc.child(index);
            const tr = state.tr;
            if (step.type === "replace")
              tr.replaceWith(
                pos,
                pos + node.nodeSize,
                listItem("R", step.numId),
              );
            else if (step.type === "insert")
              tr.insert(pos, listItem("I", step.numId));
            else if (step.type === "delete" && count > 1)
              tr.delete(pos, pos + node.nodeSize);
            else if (step.type === "type") tr.insertText("t", pos + 1);
            else if (step.type === "join" && index + 1 < count)
              tr.join(pos + node.nodeSize);
            state = state.apply(tr);
            const fresh = EditorState.create({
              doc: state.doc,
              plugins: [createDocumentNumberingPlugin(DEFINITIONS)],
            });
            expect(getDocumentNumbering(state)?.definitions).toEqual(
              getDocumentNumbering(fresh)?.definitions,
            );
          }
        },
      ),
      propertyConfig({ numRuns: 300 }),
    );
  });
});
