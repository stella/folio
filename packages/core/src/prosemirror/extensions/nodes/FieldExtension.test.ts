import { describe, expect, test } from "bun:test";
import { Window } from "happy-dom";
import { DOMParser, DOMSerializer, type Node as PMNode } from "prosemirror-model";

import type { Document, Paragraph } from "../../../types/document";
import { fromProseDoc } from "../../conversion/fromProseDoc";
import { toProseDoc } from "../../conversion/toProseDoc";
import { schema } from "../../schema";

const roundTripThroughDom = (field: PMNode) => {
  const window = new Window();
  const document = window.document as unknown as globalThis.Document;
  const fragment = DOMSerializer.fromSchema(schema).serializeFragment(
    schema.node("paragraph", null, [field]).content,
    { document },
  );
  const host = document.createElement("div");
  host.append(fragment);
  return { host, parsed: DOMParser.fromSchema(schema).parse(host) };
};

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

  test("carries empty result runs through field DOM attributes", () => {
    const paragraph = {
      type: "paragraph",
      content: [
        {
          type: "complexField",
          instruction: " PAGE ",
          fieldType: "PAGE",
          fieldCode: [],
          fieldResult: [
            { type: "run", formatting: { bold: true }, content: [{ type: "text", text: "" }] },
            { type: "run", formatting: { italic: true }, content: [{ type: "text", text: "" }] },
          ],
        },
      ],
    } as const satisfies Paragraph;
    const source: Document = { package: { document: { content: [paragraph] } } };
    const field = toProseDoc(source).firstChild?.firstChild;
    if (!field) throw new Error("Expected the projected PAGE field");
    const runs = paragraph.content[0].fieldResult;
    const { host, parsed } = roundTripThroughDom(field);
    const span = host.querySelector("span.docx-field");
    expect(span?.textContent).toBe("{page}");
    expect(span?.getAttribute("data-display-text")).toBe("");
    expect(span?.getAttribute("data-empty-result-runs")).toBe(JSON.stringify(runs));
    expect(parsed.firstChild?.firstChild?.attrs).toMatchObject({
      displayText: "",
      _docxEmptyResultRuns: runs,
    });
    expect(fromProseDoc(parsed).package.document.content).toEqual([paragraph]);
  });

  test.each([
    ["invalid JSON", "{"],
    ["scalar", "42"],
    ["empty array", "[]"],
    ["visible result", '[{"type":"run","content":[{"type":"text","text":"7"}]}]'],
    ["invalid formatting", '[{"type":"run","content":[],"formatting":{"bold":"yes"}}]'],
  ])("discards %s from pasted empty-result metadata", (_name, value) => {
    const window = new Window();
    const document = window.document as unknown as globalThis.Document;
    const host = document.createElement("div");
    const field = document.createElement("span");
    field.className = "docx-field";
    field.setAttribute("data-field-type", "PAGE");
    field.setAttribute("data-instruction", " PAGE ");
    field.setAttribute("data-field-kind", "complex");
    field.setAttribute("data-empty-result-runs", value);
    field.setAttribute("data-display-text", "");
    field.textContent = "{page}";
    host.append(field);

    const parsed = DOMParser.fromSchema(schema).parse(host);
    expect(parsed.firstChild?.firstChild?.attrs["_docxEmptyResultRuns"]).toBeNull();
    expect(parsed.firstChild?.firstChild?.attrs["displayText"]).toBe("{page}");
    expect(() => fromProseDoc(parsed)).not.toThrow();
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
