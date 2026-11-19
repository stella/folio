import { readXml, type XmlElement } from "@stll/docx-core/xml";

import type { BlockView } from "./readers.ts";

const HTML_CELL_INLINE_TAGS = new Set([
  "a",
  "strong",
  "em",
  "u",
  "sup",
  "sub",
  "ins",
  "del",
  "span",
  "p",
]);

/** HTML's void elements this oracle meets, written without an end tag. */
const VOID_ELEMENT = /<(br|img)(\s[^<>]*?)?\s*\/?>/gu;

/** The five XML entities; every other reference stays as written. */
const decodeBuiltInEntities = (text: string): string =>
  text.replace(/&(amp|lt|gt|quot|apos);/gu, (_, name: string) => {
    switch (name) {
      case "amp":
        return "&";
      case "lt":
        return "<";
      case "gt":
        return ">";
      case "quot":
        return '"';
      default:
        return "'";
    }
  });

/**
 * A cell's own text. A table nested in the cell is read as the blocks it
 * holds, after the cell's text, into `nested`, as a reader lists them.
 */
const htmlCellText = (
  nodes: readonly XmlElement[],
  nested: BlockView[],
): string => {
  const rendered: string[] = [];
  for (const node of nodes) {
    if (node.type === "text") {
      rendered.push(decodeBuiltInEntities(String(node.text ?? "")));
      continue;
    }
    const tag = node.name ?? "";
    if (tag === "img") continue;
    if (tag === "br") {
      rendered.push("\n");
      continue;
    }
    if (tag === "table") {
      nested.push(...tableViews(node.elements ?? []));
      continue;
    }
    if (!HTML_CELL_INLINE_TAGS.has(tag)) {
      throw new Error(`Unsupported HTML table cell tag: ${tag}`);
    }
    rendered.push(htmlCellText(node.elements ?? [], nested));
  }
  return rendered.join("");
};

/** The blocks a table's rows hold, cell by cell. */
const tableViews = (nodes: readonly XmlElement[]): BlockView[] => {
  const views: BlockView[] = [];
  for (const node of nodes) {
    if (node.type === "text") {
      if (String(node.text ?? "").trim().length > 0) {
        throw new Error("Unexpected text outside an HTML table cell");
      }
      continue;
    }
    const tag = node.name ?? "";
    if (tag === "table" || tag === "thead" || tag === "tbody" || tag === "tr") {
      views.push(...tableViews(node.elements ?? []));
      continue;
    }
    if (tag === "th" || tag === "td") {
      const nested: BlockView[] = [];
      const text = htmlCellText(node.elements ?? [], nested).trim();
      if (text.length > 0) views.push({ text, kind: "paragraph" });
      views.push(...nested);
      continue;
    }
    throw new Error(`Unsupported HTML table tag: ${tag}`);
  }
  return views;
};

export const htmlTableViews = (html: string): BlockView[] => {
  // Raw references: the oracle decodes only XML's five, so an HTML name such
  // as `&nbsp;` stays literal text rather than becoming a character.
  const read = readXml(html.replace(VOID_ELEMENT, "<$1$2/>"), {
    conformance: "tolerant",
    references: "raw",
  });
  if (read.status === "rejected") {
    throw new Error(`Malformed HTML table: ${read.reason}`);
  }
  const elements = read.value.elements ?? [];
  if (!elements.some((node) => node.type !== "text" && node.name === "table"))
    throw new Error("Markdown HTML block has no table");
  return tableViews(elements);
};
