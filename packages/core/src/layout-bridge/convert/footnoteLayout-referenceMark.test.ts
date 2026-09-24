import { describe, expect, test } from "bun:test";
import { panic } from "better-result";

import { parseFootnotes } from "../../docx/footnoteParser";
import type { FlowBlock, Run } from "../../layout-engine/types";
import { convertFootnoteToContent } from "./footnoteLayout";

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

function footnoteFromParagraphXml(paragraphXml: string) {
  const footnote = parseFootnotes(
    `<w:footnotes xmlns:w="${W_NS}"><w:footnote w:id="4">${paragraphXml}</w:footnote></w:footnotes>`,
  ).byId.get(4);
  if (!footnote) {
    panic("Expected the synthetic footnote to parse");
  }
  return footnote;
}

function convert(paragraphXml: string, displayNumber: number): FlowBlock[] {
  return convertFootnoteToContent(footnoteFromParagraphXml(paragraphXml), displayNumber, 400, {
    measureBlocks: (blocks) =>
      blocks.map(() => ({ kind: "paragraph" as const, lines: [], totalHeight: 12 })),
  }).blocks;
}

function firstParagraphRuns(blocks: FlowBlock[]): Run[] {
  const first = blocks.at(0);
  if (first?.kind !== "paragraph") {
    panic("Expected a paragraph block");
  }
  return first.runs;
}

function visibleText(runs: Run[]): string {
  return runs
    .map((run) => {
      if (run.kind === "text") {
        return run.text;
      }
      return run.kind === "tab" ? "\t" : "";
    })
    .join("");
}

const SUPERSCRIPT_MARK_RUN = `<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteRef/></w:r>`;

describe("note reference mark in the note story", () => {
  test("shows the number where w:footnoteRef sits among the runs", () => {
    const runs = firstParagraphRuns(
      convert(
        `<w:p><w:r><w:t>(</w:t></w:r>${SUPERSCRIPT_MARK_RUN}<w:r><w:t xml:space="preserve">) </w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>Note body</w:t></w:r></w:p>`,
        3,
      ),
    );

    expect(visibleText(runs)).toBe("(3) \tNote body");
    expect(runs.find((run) => run.kind === "text" && run.text === "3")).toMatchObject({
      superscript: true,
    });
  });

  test("shows the number once when the mark leads the paragraph", () => {
    const runs = firstParagraphRuns(
      convert(
        `<w:p>${SUPERSCRIPT_MARK_RUN}<w:r><w:t xml:space="preserve"> Note body</w:t></w:r></w:p>`,
        2,
      ),
    );

    expect(visibleText(runs)).toBe("2 Note body");
  });

  test("a note story without the mark still leads with its number", () => {
    const runs = firstParagraphRuns(convert(`<w:p><w:r><w:t>Note body</w:t></w:r></w:p>`, 5));

    expect(visibleText(runs)).toBe("5 Note body");
  });
});
