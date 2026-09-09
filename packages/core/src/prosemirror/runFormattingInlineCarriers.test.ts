import { describe, expect, test } from "bun:test";

import { schema } from "./schema";
import {
  expandRunFormattingCarrier,
  RUN_FORMATTING_INLINE_ATOM_DISPOSITIONS,
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
});
