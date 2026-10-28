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

const htmlCellText = (nodes: readonly XmlElement[]): string => {
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
    if (!HTML_CELL_INLINE_TAGS.has(tag)) {
      throw new Error(`Unsupported HTML table cell tag: ${tag}`);
    }
    rendered.push(htmlCellText(node.elements ?? []));
  }
  return rendered.join("");
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
  const views: BlockView[] = [];
  let sawTable = false;
  const walk = (nodes: readonly XmlElement[]): void => {
    for (const node of nodes) {
      if (node.type === "text") {
        if (String(node.text ?? "").trim().length > 0) {
          throw new Error("Unexpected text outside an HTML table cell");
        }
        continue;
      }
      const tag = node.name ?? "";
      if (
        tag === "table" ||
        tag === "thead" ||
        tag === "tbody" ||
        tag === "tr"
      ) {
        if (tag === "table") sawTable = true;
        walk(node.elements ?? []);
        continue;
      }
      if (tag === "th" || tag === "td") {
        const text = htmlCellText(node.elements ?? []).trim();
        if (text.length > 0) views.push({ text, kind: "paragraph" });
        continue;
      }
      throw new Error(`Unsupported HTML table tag: ${tag}`);
    }
  };
  walk(read.value.elements ?? []);
  if (!sawTable) throw new Error("Markdown HTML block has no table");
  return views;
};
