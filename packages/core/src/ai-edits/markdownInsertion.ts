import { Result } from "better-result";

import { compileMarkdownToContent } from "@stll/docx-core";

import type {
  BlockContent,
  Paragraph,
  Run,
  RunContent,
} from "../types/document";
import type { FolioAIListKind } from "./types";

type MarkdownInline = Extract<RunContent, { type: "text" | "break" }>;
type MarkdownRun = Omit<Run, "content"> & { content: MarkdownInline[] };

export type MarkdownParagraph = Omit<Paragraph, "content"> & {
  content: MarkdownRun[];
};

const isMarkdownParagraph = (block: BlockContent): block is MarkdownParagraph =>
  block.type === "paragraph" &&
  block.content.every(
    (item) =>
      item.type === "run" &&
      item.content.every(
        (child) =>
          child.type === "text" ||
          (child.type === "break" &&
            child.breakType === undefined &&
            child.sourceElement === undefined),
      ),
  );

type MarkdownInsertion =
  | { type: "unsupported"; message: string }
  | {
      type: "paragraphs";
      paragraphs: MarkdownParagraph[];
      listKind?: FolioAIListKind;
    };

const unsupported = (message: string) =>
  ({ type: "unsupported", message }) as const;

export const compileMarkdownInsertion = (text: string): MarkdownInsertion => {
  const compiled = Result.try({
    try: () => compileMarkdownToContent(text, { editProfile: "paragraphs" }),
    catch: () => "Markdown could not be parsed.",
  });
  if (compiled.isErr()) return unsupported(compiled.error);
  const { content, warnings, numbering } = compiled.value;
  const warning = warnings?.at(0);
  if (warning !== undefined) return unsupported(warning);
  if (content.length === 0) return unsupported("Markdown contains no blocks.");
  if (!content.every(isMarkdownParagraph))
    return unsupported(
      "Markdown must contain supported paragraphs or a flat list.",
    );

  if (numbering !== undefined) {
    const instance = numbering.nums.at(0);
    if (
      numbering.nums.length !== 1 ||
      instance === undefined ||
      content.some(
        ({ listRendering }) =>
          listRendering === undefined ||
          listRendering.numId !== instance.numId ||
          listRendering.level !== 0,
      )
    ) {
      return unsupported(
        "Markdown list input must be one standalone flat list.",
      );
    }
    const listKind = content.at(0)?.listRendering?.isBullet
      ? "bullet"
      : "numbered";
    return { type: "paragraphs", paragraphs: content, listKind };
  }

  // A direct outline level makes headings independent of a template's styles.
  const paragraphs: MarkdownParagraph[] = [];
  for (const paragraph of content) {
    const outlineLevel = paragraph.formatting?.outlineLevel;
    paragraphs.push({
      type: "paragraph",
      content: paragraph.content,
      ...(outlineLevel !== undefined && { formatting: { outlineLevel } }),
    });
  }
  return { type: "paragraphs", paragraphs };
};

export const markdownRunText = (run: MarkdownRun): string =>
  run.content.map((item) => (item.type === "text" ? item.text : "\n")).join("");

/** Plain Markdown table rows share the ordinary insertTable execution path. */
export const compileMarkdownTableRows = (
  text: string,
): Result<string[][], string> => {
  const compiled = Result.try({
    try: () => compileMarkdownToContent(text, { editProfile: "table" }),
    catch: () => "Markdown could not be parsed.",
  });
  if (compiled.isErr()) return Result.err(compiled.error);
  const warning = compiled.value.warnings?.at(0);
  if (warning !== undefined) return Result.err(warning);
  const table = compiled.value.content.at(0);
  if (table?.type !== "table" || compiled.value.content.length !== 1)
    return Result.err("Expected one standalone Markdown table.");
  const rows: string[][] = [];
  for (const row of table.rows) {
    const cells: string[] = [];
    for (const cell of row.cells) {
      const paragraph = cell.content.at(0);
      if (
        cell.content.length !== 1 ||
        paragraph === undefined ||
        !isMarkdownParagraph(paragraph)
      )
        return Result.err("Markdown table cells must contain plain text.");
      cells.push(paragraph.content.map(markdownRunText).join(""));
    }
    rows.push(cells);
  }
  return Result.ok(rows);
};
