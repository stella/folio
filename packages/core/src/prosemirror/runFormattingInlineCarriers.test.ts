import { describe, expect, test } from "bun:test";

import { schema } from "./schema";
import {
  expandRunFormattingCarrier,
  RUN_FORMATTING_INLINE_ATOM_DISPOSITIONS,
  selectRunFormattingCarrierRepresentations,
} from "./runFormattingInlineCarriers";

describe("run-formatting inline carrier contract", () => {
  test("classifies every inline atom exposed by the editor schema", () => {
    const schemaAtoms = Object.values(schema.nodes)
      .filter((node) => node.isInline && node.isAtom)
      .map(({ name }) => name)
      .toSorted();

    expect(Object.keys(RUN_FORMATTING_INLINE_ATOM_DISPOSITIONS).toSorted()).toEqual(schemaAtoms);
    expect(RUN_FORMATTING_INLINE_ATOM_DISPOSITIONS).toEqual({
      bookmarkBoundary: "not-a-run",
      field: "field-run",
      hardBreak: "break-run",
      image: "not-a-run",
      math: "not-a-run",
      pageBreakRun: "page-break-carrier",
      renderedPageBreak: "not-a-run",
      shape: "not-a-run",
      structuredField: "structured-field",
      symbol: "symbol-run",
      tab: "tab-run",
      text: "text-run",
      textBoxAnchor: "not-a-run",
    });
  });

  test.each([
    ["text-run", () => schema.text("x")],
    ["tab-run", () => schema.node("tab")],
    ["break-run", () => schema.node("hardBreak")],
    ["page-break-carrier", () => schema.node("pageBreakRun")],
    ["symbol-run", () => schema.node("symbol", { font: "Wingdings", char: "F06F" })],
    [
      "field-run",
      () =>
        schema.node("field", {
          fieldType: "PAGE",
          instruction: " PAGE ",
          displayText: "1",
          fieldKind: "simple",
          fldLock: false,
          dirty: false,
        }),
    ],
  ] as const)("expands a %s into exactly one owning representation", (disposition, makeNode) => {
    const node = makeNode();
    expect(expandRunFormattingCarrier(node, 7)).toEqual({
      disposition,
      node,
      position: 7,
      representations: [{ node, position: 7, role: "owner" }],
    });
  });

  test("does not expand an inline atom whose serialization has no run properties", () => {
    expect(expandRunFormattingCarrier(schema.node("renderedPageBreak"), 7)).toBeNull();
  });

  test("keeps a page-break carrier structural during generic formatting selection", () => {
    const pageBreak = schema.node("pageBreakRun");
    const doc = schema.node("doc", null, [schema.node("paragraph", null, [pageBreak])]);

    expect(expandRunFormattingCarrier(pageBreak, 1)).toEqual({
      disposition: "page-break-carrier",
      node: pageBreak,
      position: 1,
      representations: [{ node: pageBreak, position: 1, role: "owner" }],
    });
    expect(selectRunFormattingCarrierRepresentations({ doc, from: 1, to: 2 })).toEqual([]);
  });

  test("expands a structured field once into only its serialized run representations", () => {
    const field = schema.node(
      "structuredField",
      {
        fieldType: "REF",
        instruction: " REF carrier ",
        displayText: "A\tB",
        fieldKind: "simple",
        fldLock: false,
        dirty: false,
      },
      [schema.text("A"), schema.node("tab"), schema.node("renderedPageBreak"), schema.text("B")],
    );
    const carrier = expandRunFormattingCarrier(field, 7);

    expect(carrier?.node).toBe(field);
    expect(
      carrier?.representations.map(({ node, position, role }) => ({
        name: node.type.name,
        position,
        role,
      })),
    ).toEqual([
      { name: "structuredField", position: 7, role: "owner" },
      { name: "text", position: 8, role: "serialized-result" },
      { name: "tab", position: 9, role: "serialized-result" },
      { name: "text", position: 11, role: "serialized-result" },
    ]);
  });

  test("selects the same physical carriers for a complete structured field", () => {
    const field = schema.node(
      "structuredField",
      {
        fieldType: "REF",
        instruction: " REF carrier ",
        displayText: "A\tB",
        fieldKind: "simple",
        fldLock: false,
        dirty: false,
      },
      [schema.text("A"), schema.node("tab"), schema.node("renderedPageBreak"), schema.text("B")],
    );
    const doc = schema.node("doc", null, [schema.node("paragraph", null, [field])]);

    expect(
      selectRunFormattingCarrierRepresentations({ doc, from: 1, to: 7 }).map(
        ({ node, position, role, from, to }) => ({
          name: node.type.name,
          position,
          role,
          from,
          to,
        }),
      ),
    ).toEqual([
      { name: "structuredField", position: 1, role: "owner", from: 1, to: 7 },
      { name: "text", position: 2, role: "serialized-result", from: 2, to: 3 },
      { name: "tab", position: 3, role: "serialized-result", from: 3, to: 4 },
      { name: "text", position: 5, role: "serialized-result", from: 5, to: 6 },
    ]);
  });

  test("selects only result runs when a range covers part of a structured field", () => {
    const field = schema.node(
      "structuredField",
      {
        fieldType: "REF",
        instruction: " REF carrier ",
        displayText: "AB",
        fieldKind: "simple",
        fldLock: false,
        dirty: false,
      },
      [schema.text("AB")],
    );
    const doc = schema.node("doc", null, [schema.node("paragraph", null, [field])]);

    expect(
      selectRunFormattingCarrierRepresentations({ doc, from: 2, to: 3 }).map(
        ({ node, position, role, from, to }) => ({
          name: node.type.name,
          position,
          role,
          from,
          to,
        }),
      ),
    ).toEqual([{ name: "text", position: 2, role: "serialized-result", from: 2, to: 3 }]);
  });
});
