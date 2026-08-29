/**
 * Field extensions — atomic inline fields (PAGE, NUMPAGES, REF, etc.).
 *
 * Ordinary fields are true ProseMirror leaves, so `textContent`, `textBetween`,
 * and the default plain-text clipboard serializer can use `leafText`.
 * Simple fields that must preserve hyperlink structure use a distinct atomic
 * node with inline children; their visible text comes from those children.
 */

import type { Node as PMNode } from "prosemirror-model";

import { parseFieldInstruction } from "../../../docx/fieldParser";
import { expectFieldAttrs } from "../../attrs";
import { createNodeExtension } from "../create";

type StructuredFieldOptions = {
  getInternalClipboardToken?: () => string;
};

const createFieldAttrs = () => ({
  fieldType: { default: "UNKNOWN" },
  instruction: { default: "" },
  displayText: { default: "" },
  _numberedRefBaseline: { default: undefined },
  fieldKind: { default: "simple" },
  fldLock: { default: false },
  dirty: { default: false },
});

const readFieldDomAttrs = (dom: HTMLElement) => ({
  fieldType: dom.dataset["fieldType"] ?? "UNKNOWN",
  instruction: dom.dataset["instruction"] ?? "",
  displayText: dom.textContent ?? "",
  fieldKind: dom.dataset["fieldKind"] ?? "simple",
  fldLock: dom.dataset["fldLock"] === "true",
  dirty: dom.dataset["dirty"] === "true",
});

const getFieldDomAttrs = (node: PMNode) => {
  const { fieldType, instruction, fieldKind, fldLock, dirty } = expectFieldAttrs(node);
  return {
    class: `docx-field docx-field-${fieldType.toLowerCase()}`,
    "data-field-type": fieldType,
    "data-instruction": instruction,
    "data-field-kind": fieldKind,
    ...(fldLock ? { "data-fld-lock": "true" } : {}),
    ...(dirty ? { "data-dirty": "true" } : {}),
    style:
      "outline: 1px solid var(--doc-field-outline, rgba(200,200,200,0.4)); padding: 0 1px; border-radius: 2px;",
  };
};

const getFieldVisibleText = (node: PMNode): string => {
  const { fieldType, instruction, displayText } = expectFieldAttrs(node);
  if (displayText) {
    return displayText;
  }
  switch (fieldType) {
    case "PAGE":
      return "{page}";
    case "NUMPAGES":
      return "{pages}";
    case "DATE":
    case "TIME":
    case "CREATEDATE":
    case "SAVEDATE":
      return new Date().toLocaleDateString();
    case "MERGEFIELD":
      return `«${getMergeFieldName(instruction)}»`;
    default:
      return `{${fieldType}}`;
  }
};

export const FieldExtension = createNodeExtension({
  name: "field",
  schemaNodeName: "field",
  nodeSpec: {
    inline: true,
    group: "inline",
    atom: true,
    selectable: true,
    attrs: createFieldAttrs(),
    leafText: getFieldVisibleText,
    parseDOM: [
      {
        tag: "span.docx-field:not([data-field-structured])",
        getAttrs: readFieldDomAttrs,
      },
    ],
    toDOM(node) {
      return ["span", getFieldDomAttrs(node), getFieldVisibleText(node)];
    },
  },
});

export const StructuredFieldExtension = createNodeExtension<StructuredFieldOptions>({
  name: "structuredField",
  schemaNodeName: "structuredField",
  nodeSpec: (options) => ({
    inline: true,
    group: "inline",
    content:
      "(text | bookmarkBoundary | tab | symbol | hardBreak | image | shape | renderedPageBreak | textBoxAnchor)+",
    atom: true,
    selectable: true,
    attrs: createFieldAttrs(),
    parseDOM: [
      {
        tag: 'span.docx-field[data-field-structured="true"]',
        getAttrs(dom) {
          if (dom.dataset["fieldKind"] !== "simple" || !dom.querySelector("a[href]")) {
            return false;
          }
          return readFieldDomAttrs(dom);
        },
      },
    ],
    toDOM(node) {
      return [
        "span",
        {
          ...getFieldDomAttrs(node),
          "data-field-structured": "true",
          ...(options.getInternalClipboardToken
            ? { "data-docx-internal-clipboard": options.getInternalClipboardToken() }
            : {}),
        },
        0,
      ];
    },
  }),
});

function getMergeFieldName(instruction: string): string {
  return parseFieldInstruction(instruction).argument ?? "";
}
