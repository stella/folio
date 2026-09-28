import { describe, expect, test } from "bun:test";
import { Window } from "happy-dom";

import { DOMParser, DOMSerializer, Fragment, Slice, type Node as PMNode } from "prosemirror-model";

import type { ComplexField, Document } from "../../types/document";
import { schema } from "../schema";
import { readFieldAttrs } from "../attrs";
import { fromProseDoc } from "./fromProseDoc";
import { toProseDoc } from "./toProseDoc";

const fieldCode = [
  {
    type: "run",
    content: [{ type: "instrText", text: " REF " }],
    preservedAttributes: [{ name: "rsidR", value: "00A1B2C3" }],
  },
  {
    type: "run",
    formatting: { bold: true },
    content: [{ type: "instrText", text: "target \\h" }],
    preservedAttributes: [{ name: "rsidRPr", value: "00D4E5F6" }],
  },
] satisfies ComplexField["fieldCode"];

const documentWithField = (): Document => ({
  package: {
    document: {
      content: [
        {
          type: "table",
          rows: [
            {
              cells: [
                {
                  type: "tableCell",
                  content: [
                    {
                      type: "paragraph",
                      content: [
                        {
                          type: "complexField",
                          instruction: " REF target \\h",
                          fieldType: "REF",
                          fieldCode,
                          fieldResult: [
                            { type: "run", content: [{ type: "text", text: "Target" }] },
                          ],
                        },
                      ],
                    },
                  ],
                },
              ],
            },
          ],
        },
      ],
    },
  },
});

const findComplexField = (document: Document) => {
  const table = document.package.document.content.at(0);
  if (table?.type !== "table") throw new Error("Expected table");
  const cell = table.rows.at(0)?.cells.at(0);
  const paragraph = cell?.content.at(0);
  if (paragraph?.type !== "paragraph") throw new Error("Expected cell paragraph");
  const field = paragraph.content.at(0);
  if (field?.type !== "complexField") throw new Error("Expected complex field");
  return field;
};

describe("complex field code metadata", () => {
  test("rejects retained code runs on a simple field", () => {
    const field = schema.nodes.field.create({
      fieldType: "REF",
      instruction: " REF target \\h",
      displayText: "Target",
      fieldKind: "simple",
      _docxFieldCode: { instruction: " REF target \\h", runs: fieldCode },
    });
    const result = readFieldAttrs(field);
    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.issues.map(({ path }) => path)).toContain("field.attrs._docxFieldCode");
    }
  });

  test("preserves all authored code runs inside a table cell", () => {
    const result = fromProseDoc(toProseDoc(documentWithField()));

    expect(findComplexField(result).fieldCode).toEqual(fieldCode);
  });

  test("keeps code runs through field DOM serialization", () => {
    const pmDoc = toProseDoc(documentWithField());
    let field: PMNode | undefined;
    pmDoc.descendants((node) => {
      if (node.type.name !== "field") return true;
      field = node;
      return false;
    });
    if (!field) throw new Error("Expected ProseMirror field");
    const window = new Window();
    const document = window.document as unknown as globalThis.Document;
    const fragment = DOMSerializer.fromSchema(schema).serializeFragment(Fragment.from(field), {
      document,
    });
    const host = document.createElement("div");
    host.append(fragment);
    const span = host.querySelector("span.docx-field");
    expect(span?.getAttribute("data-field-code")).toBe(
      JSON.stringify({
        instruction: " REF target \\h",
        runs: fieldCode,
      }),
    );

    const parsed = DOMParser.fromSchema(schema).parse(host);
    const paragraph = fromProseDoc(parsed).package.document.content.at(0);
    if (paragraph?.type !== "paragraph") throw new Error("Expected parsed paragraph");
    const restored = paragraph.content.at(0);
    if (restored?.type !== "complexField") throw new Error("Expected parsed complex field");
    expect(restored.fieldCode).toEqual(fieldCode);
  });

  test("drops authored code runs when the ProseMirror instruction changes", () => {
    const pmDoc = toProseDoc(documentWithField());
    let fieldPosition: number | undefined;
    let fieldNode: PMNode | undefined;
    pmDoc.descendants((node, pos) => {
      if (node.type.name === "field") {
        fieldPosition = pos;
        fieldNode = node;
        return false;
      }
      return true;
    });
    if (fieldPosition === undefined || !fieldNode) throw new Error("Expected ProseMirror field");

    const editedField = schema.nodes.field.create(
      { ...fieldNode.attrs, instruction: " REF other \\h" },
      null,
      fieldNode.marks,
    );
    const editedDoc = pmDoc.replace(
      fieldPosition,
      fieldPosition + fieldNode.nodeSize,
      new Slice(Fragment.from(editedField), 0, 0),
    );
    const editedModel = fromProseDoc(editedDoc);

    expect(findComplexField(editedModel).instruction).toBe(" REF other \\h");
    expect(findComplexField(editedModel).fieldCode).toEqual([]);
  });
});
