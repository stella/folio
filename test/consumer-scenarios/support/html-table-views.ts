import { XMLParser } from "fast-xml-parser";
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

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);

const htmlCellText = (nodes: readonly unknown[]): string => {
  const rendered: string[] = [];
  for (const node of nodes) {
    if (!isRecord(node)) throw new Error("Malformed HTML table node");
    for (const [tag, children] of Object.entries(node)) {
      if (tag === ":@") continue;
      if (tag === "#text") {
        if (typeof children !== "string") throw new Error("Malformed HTML table text");
        rendered.push(children);
        continue;
      }
      if (tag === "img") continue;
      if (tag === "br") {
        rendered.push("\n");
        continue;
      }
      if (!HTML_CELL_INLINE_TAGS.has(tag)) {
        throw new Error(`Unsupported HTML table cell tag: ${tag}`);
      }
      if (!Array.isArray(children)) throw new Error(`Malformed HTML table ${tag}`);
      rendered.push(htmlCellText(children));
    }
  }
  return rendered.join("");
};

export const htmlTableViews = (html: string): BlockView[] => {
  const parser = new XMLParser({
    ignoreAttributes: false,
    preserveOrder: true,
    trimValues: false,
    parseTagValue: false,
    unpairedTags: ["br", "img"],
  });
  const parsed: unknown = parser.parse(html);
  if (!Array.isArray(parsed)) throw new Error("Malformed HTML table");
  const views: BlockView[] = [];
  let sawTable = false;
  const walk = (nodes: readonly unknown[]): void => {
    for (const node of nodes) {
      if (!isRecord(node)) throw new Error("Malformed HTML table node");
      for (const [tag, children] of Object.entries(node)) {
        if (tag === ":@") continue;
        if (tag === "#text") {
          if (typeof children !== "string" || children.trim().length > 0) {
            throw new Error("Unexpected text outside an HTML table cell");
          }
          continue;
        }
        if (tag === "table" || tag === "thead" || tag === "tbody" || tag === "tr") {
          if (!Array.isArray(children)) throw new Error(`Malformed HTML table ${tag}`);
          if (tag === "table") sawTable = true;
          walk(children);
          continue;
        }
        if (tag === "th" || tag === "td") {
          if (!Array.isArray(children)) throw new Error(`Malformed HTML table ${tag}`);
          const text = htmlCellText(children).trim();
          if (text.length > 0) views.push({ text, kind: "paragraph" });
          continue;
        }
        throw new Error(`Unsupported HTML table tag: ${tag}`);
      }
    }
  };
  walk(parsed);
  if (!sawTable) throw new Error("Markdown HTML block has no table");
  return views;
};
