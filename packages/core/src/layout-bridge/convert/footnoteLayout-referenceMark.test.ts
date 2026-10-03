import { describe, expect, test } from "bun:test";
import { panic } from "better-result";

import { parseEndnotes, parseFootnotes } from "../../docx/footnoteParser";
import type { FlowBlock, Run } from "../../layout-engine/types";
import { convertFootnoteToContent, convertNoteStoryToFlowBlocks } from "./footnoteLayout";

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

// The parser owns namespace recognition; layout consumes the same typed marker for both note kinds.
for (const kind of ["footnote", "endnote"] as const) {
  for (const namespace of [W_NS, "http://purl.oclc.org/ooxml/wordprocessingml/main"]) {
    for (const prefix of ["producer:", ""]) {
      for (const placement of ["leading", "middle", "trailing"] as const) {
        test(`typed ${kind} marker retains ${placement} placement in ${namespace} (${prefix || "default"})`, () => {
          const binding = prefix
            ? `xmlns:producer="${namespace}"`
            : `xmlns="${namespace}" xmlns:producer="${namespace}"`;
          const before = placement === "leading" ? "" : "before";
          const after = placement === "trailing" ? "" : "after";
          const xml =
            `<${prefix}${kind}s ${binding}><${prefix}${kind} producer:id="4"><${prefix}p>` +
            `<${prefix}r><${prefix}t>${before}</${prefix}t></${prefix}r>` +
            `<${prefix}r><${prefix}rPr><${prefix}vertAlign producer:val="superscript"/></${prefix}rPr><${prefix}${kind}Ref/></${prefix}r>` +
            `<${prefix}r><${prefix}t>${after}</${prefix}t></${prefix}r>` +
            `</${prefix}p></${prefix}${kind}></${prefix}${kind}s>`;
          const note = (kind === "footnote" ? parseFootnotes(xml) : parseEndnotes(xml)).byId.get(4);
          if (!note) panic("Expected parsed note");
          const { flowBlocks, hasReferenceMark } = convertNoteStoryToFlowBlocks(
            note.content,
            {},
            "iv",
          );
          const runs = firstParagraphRuns(flowBlocks);
          expect(hasReferenceMark).toBe(true);
          expect(visibleText(runs)).toBe(`${before}iv${after}`);
          expect(runs.find((run) => run.kind === "text" && run.text === "iv")).toMatchObject({
            superscript: true,
          });
        });
      }
    }
  }
}
