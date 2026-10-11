import { expect, test } from "bun:test";

import { MARKDOWN_BLOCK_BOUNDARY_FIXTURES } from "../../typecheck/markdown/blockBoundaries.typecheck";

import type { BlockContent, ListRendering, Paragraph } from "../types/document";
import { fromMarkdown } from "./fromMarkdown";
import { toMarkdown } from "./index";

const paragraph = (text: string, listRendering?: ListRendering): Paragraph => ({
  type: "paragraph",
  content: [{ type: "run", content: [{ type: "text", text }] }],
  ...(listRendering === undefined ? {} : { listRendering }),
});

const render = (content: BlockContent[]): string =>
  toMarkdown({ package: { document: { content } } });

const listParagraph = (text: string): Paragraph =>
  paragraph(text, { marker: "-", level: 0, numId: 1, isBullet: true });

const importedTexts = (markdown: string): string[] =>
  fromMarkdown(markdown).package.document.content.map((block) => {
    expect(block.type).toBe("paragraph");
    if (block.type !== "paragraph") throw new Error(`Unexpected block ${block.type}`);
    expect(block.formatting?.outlineLevel).toBeUndefined();
    return block.content
      .flatMap((item) =>
        item.type === "run"
          ? item.content.flatMap((part) => (part.type === "text" ? [part.text] : []))
          : [],
      )
      .join("");
  });

test("each text-emitting block stays separated from preceding paragraph and list text", () => {
  const suffix = paragraph("SUFFIX");
  const leads = [paragraph("PREFIX"), listParagraph("PREFIX")];

  for (const fixture of Object.values(MARKDOWN_BLOCK_BOUNDARY_FIXTURES)) {
    if (fixture.marker === null) {
      for (const lead of leads) {
        expect(render([lead, fixture.block, suffix])).toBe(render([lead, suffix]));
      }
      continue;
    }

    for (const lead of leads) {
      const actual = render([lead, fixture.block, suffix]);
      const expected = [render([lead]), render([fixture.block]), render([suffix])].join("\n\n");
      expect(actual).toBe(expected);
      expect(actual).toContain(fixture.marker);
    }
  }
});

test("block syntax-like text after paragraph text keeps a blank block boundary", () => {
  for (const text of ["-", "=", "===", "---", "-----"]) {
    for (const prefix of [paragraph("PREFIX"), listParagraph("PREFIX")]) {
      const markdown = render([prefix, paragraph(text)]);
      expect(markdown).toContain("PREFIX\n\n");
      expect(importedTexts(markdown)).toEqual(["PREFIX", text]);
    }
    const custom = paragraph("", { marker: text, level: 1, numId: 2, isBullet: false });
    expect(importedTexts(render([listParagraph("PREFIX"), custom]))).toEqual(["PREFIX", text]);
  }
});

test("list paragraph text cannot introduce nested Markdown block syntax", () => {
  for (const text of ["# Literal", "=", "===", "---", "-----", "> Literal", "1. Literal"]) {
    expect(importedTexts(render([listParagraph(text)]))).toEqual([text]);
  }
});

test("custom list labels remain separate paragraph text in Markdown", () => {
  for (const marker of ["=", "---", "a.", "1.2."]) {
    const custom = paragraph("Value", { marker, level: 1, numId: 2, isBullet: false });
    expect(importedTexts(render([listParagraph("PREFIX"), custom]))).toEqual([
      "PREFIX",
      `${marker} Value`,
    ]);
  }
});

test("bookmark delimiters do not consume the separator between text blocks", () => {
  const prefix = listParagraph("PREFIX");
  const suffix = paragraph("SUFFIX");
  const withMarkers = render([
    prefix,
    MARKDOWN_BLOCK_BOUNDARY_FIXTURES.bookmarkStart.block,
    MARKDOWN_BLOCK_BOUNDARY_FIXTURES.bookmarkEnd.block,
    suffix,
  ]);

  expect(withMarkers).toBe(render([prefix, suffix]));
});
