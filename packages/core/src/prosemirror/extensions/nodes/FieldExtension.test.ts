import { describe, expect, test } from "bun:test";

import { schema } from "../../schema";

describe("FieldExtension", () => {
  test("uses DOCX field-instruction parsing for quoted MERGEFIELD names", () => {
    const field = schema.node("field", {
      fieldType: "MERGEFIELD",
      instruction: ' MERGEFIELD "Client Name" \\* MERGEFORMAT ',
      displayText: "",
      fieldKind: "simple",
    });

    const toDOM = field.type.spec.toDOM;
    if (!toDOM) {
      throw new Error("Expected field node to provide toDOM");
    }

    const domSpec = toDOM(field);

    expect(domSpec).toEqual([
      "span",
      expect.objectContaining({ "data-field-type": "MERGEFIELD" }),
      "«Client Name»",
    ]);
  });

  test("keeps cached display text for fields the layout path does not recompute", () => {
    const field = schema.node("field", {
      fieldType: "REF",
      instruction: " REF _Ref123 \\h ",
      displayText: "Clause 4.2",
      fieldKind: "complex",
    });

    const toDOM = field.type.spec.toDOM;
    if (!toDOM) {
      throw new Error("Expected field node to provide toDOM");
    }

    const domSpec = toDOM(field);

    expect(domSpec).toEqual([
      "span",
      expect.objectContaining({ "data-field-type": "REF" }),
      "Clause 4.2",
    ]);
    expect(field.textContent).toBe("Clause 4.2");
    expect(schema.node("paragraph", null, [field]).textBetween(0, field.nodeSize)).toBe(
      "Clause 4.2",
    );
  });

  test("uses one fallback for an empty complex PAGE field in DOM and text semantics", () => {
    const field = schema.node("field", {
      fieldType: "PAGE",
      instruction: " PAGE ",
      displayText: "",
      fieldKind: "complex",
    });
    const toDOM = field.type.spec.toDOM;
    if (!toDOM) {
      throw new Error("Expected field node to provide toDOM");
    }

    expect(toDOM(field)).toEqual([
      "span",
      expect.objectContaining({ "data-field-type": "PAGE" }),
      "{page}",
    ]);
    expect(field.textContent).toBe("{page}");
    expect(schema.node("paragraph", null, [field]).textBetween(0, field.nodeSize)).toBe("{page}");
  });

  test("preserves hyperlink bookmark boundaries through DOM serialization", () => {
    const hyperlink = schema.mark("hyperlink", {
      href: "https://example.test/field",
      _docxHyperlinkIndex: 1,
    });
    const field = schema.node(
      "structuredField",
      {
        fieldType: "REF",
        instruction: " REF target \\h ",
        displayText: "Target",
        fieldKind: "simple",
      },
      [
        schema.node("bookmarkBoundary", { type: "start", id: 31, name: "target" }, null, [
          hyperlink,
        ]),
        schema.text("Target", [hyperlink]),
        schema.node("bookmarkBoundary", { type: "end", id: 31 }, null, [hyperlink]),
      ],
    );
    const toDOM = field.type.spec.toDOM;
    if (!toDOM) {
      throw new Error("FieldExtension must define toDOM");
    }

    expect(toDOM(field)).toEqual([
      "span",
      expect.objectContaining({
        class: "docx-field docx-field-ref",
        "data-field-type": "REF",
        "data-field-structured": "true",
        "data-instruction": " REF target \\h ",
        "data-docx-internal-clipboard": expect.any(String),
      }),
      0,
    ]);
    expect(JSON.stringify(toDOM(field))).not.toContain("aria-hidden");
    expect(JSON.stringify(toDOM(field))).not.toContain("Target");
    expect(field.textContent).toBe("Target");
    expect(field.content.toJSON().map((child) => child.type)).toEqual([
      "bookmarkBoundary",
      "text",
      "bookmarkBoundary",
    ]);
  });

  test("keeps ordinary and structured DOM parse rules disjoint", () => {
    const parseRules = schema.nodes.field.spec.parseDOM ?? [];
    const structuredRules = schema.nodes.structuredField.spec.parseDOM ?? [];
    const leafRule = parseRules.at(0);
    const structuredRule = structuredRules.at(0);

    expect(structuredRule?.tag).toBe('span.docx-field[data-field-structured="true"]');
    expect(leafRule?.tag).toBe("span.docx-field:not([data-field-structured])");
    expect(schema.nodes.field.isLeaf).toBe(true);
    expect(schema.nodes.structuredField.isLeaf).toBe(false);
  });

  test.each([
    { fieldKind: "complex", hasHyperlink: true },
    { fieldKind: "simple", hasHyperlink: false },
  ])("rejects malformed structured field DOM ($fieldKind, link=$hasHyperlink)", (input) => {
    const structuredRule = schema.nodes.structuredField.spec.parseDOM?.at(0);
    const dom = Object.assign(Object.create(null), {
      dataset: { fieldKind: input.fieldKind },
      textContent: "forged",
      querySelector: () => (input.hasHyperlink ? Object.create(null) : null),
    });

    expect(structuredRule?.getAttrs?.(dom)).toBe(false);
  });
});
