import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import { propertyConfig, propertyTestTimeout } from "../../../../test/property-testing";
import type { BlockContent, Paragraph, ParagraphContent, Run } from "../model/document";
import { compileMarkdownToContent } from "./content";
import { sanitizeMarkdownHref } from "./href";
import { inlineMarkdownToRuns } from "./inline";
import { paragraphNumberingLevel, paragraphNumberingReferenceId } from "@stll/docx-core/model";

setDefaultTimeout(propertyTestTimeout(30_000));

const paragraphs = (content: BlockContent[]): Paragraph[] =>
  content.flatMap((block) => (block.type === "paragraph" ? [block] : []));

const runText = (run: Run): string =>
  run.content.map((node) => (node.type === "text" ? node.text : "")).join("");

/** Every run's text in a paragraph's content, concatenated. */
const paragraphText = (paragraph: Paragraph): string =>
  paragraph.content.flatMap((node) => (node.type === "run" ? [runText(node)] : [])).join("");

const flattenRuns = (content: ParagraphContent[]): Run[] =>
  content.flatMap((node) => {
    if (node.type === "run") {
      return [node];
    }
    if (node.type === "hyperlink") {
      return node.children.flatMap((child) => (child.type === "run" ? [child] : []));
    }
    return [];
  });

/** Text plus formatting of each run, for `toContainEqual` assertions. */
const describeRuns = (content: ParagraphContent[]) =>
  flattenRuns(content).map((run) => ({ text: runText(run), formatting: run.formatting }));

describe("inlineMarkdownToRuns", () => {
  test("renders emphasis, code spans, and links as formatted runs", () => {
    const runs = inlineMarkdownToRuns(
      "The **Buyer** pays *promptly*, see [the schedule](https://example.com/s) and `Section 4`.",
    );
    const hyperlink = runs.find((node) => node.type === "hyperlink");
    expect(hyperlink?.type === "hyperlink" ? hyperlink.href : undefined).toBe(
      "https://example.com/s",
    );
    const formatted = describeRuns(runs);
    expect(formatted).toContainEqual({ text: "Buyer", formatting: { bold: true } });
    expect(formatted).toContainEqual({ text: "promptly", formatting: { italic: true } });
    expect(formatted).toContainEqual({
      text: "Section 4",
      formatting: { fontFamily: { ascii: "Courier New", hAnsi: "Courier New" } },
    });
  });

  test("highlights [[placeholders]] only when asked, inheriting the surrounding emphasis", () => {
    const highlighted = describeRuns(
      inlineMarkdownToRuns("**[[Party]]** shall pay [[Amount]].", { placeholders: true }),
    );
    expect(highlighted).toContainEqual({
      text: "Party",
      formatting: { bold: true, highlight: "yellow" },
    });
    expect(highlighted).toContainEqual({ text: "Amount", formatting: { highlight: "yellow" } });
    expect(highlighted.some((run) => run.text.includes("[["))).toBe(false);

    const literal = flattenRuns(inlineMarkdownToRuns("[[Party]] shall pay."));
    expect(literal.map(runText).join("")).toBe("[[Party]] shall pay.");
  });

  test("drops executable link targets but keeps their text", () => {
    const runs = inlineMarkdownToRuns("[bad](javascript:alert(1)) and [data](data:text/html,x)");
    expect(runs.every((node) => node.type === "run")).toBe(true);
    expect(flattenRuns(runs).map(runText).join("")).toBe("bad and data");
  });

  test("the rendered text equals the plain text for any plain sentence", () => {
    // Words of letters only never form markdown syntax, so the runs must
    // reproduce the sentence verbatim (no lost or doubled characters).
    const word = fc.stringMatching(/^[a-z]{1,8}$/u);
    const sentence = fc.array(word, { minLength: 1, maxLength: 6 }).map((words) => words.join(" "));
    fc.assert(
      fc.property(sentence, (text) => {
        const runs = flattenRuns(inlineMarkdownToRuns(text));
        return runs.map(runText).join("") === text;
      }),
      propertyConfig({ numRuns: 100 }),
    );
  });
});

describe("sanitizeMarkdownHref", () => {
  test("keeps document anchors and safe external targets", () => {
    expect(sanitizeMarkdownHref("#intro")).toBe("#intro");
    expect(sanitizeMarkdownHref("https://example.com/a")).toBe("https://example.com/a");
    expect(sanitizeMarkdownHref("mailto:legal@example.com")).toBe("mailto:legal@example.com");
  });

  test("rejects empty, whitespace-bearing, and executable targets", () => {
    expect(sanitizeMarkdownHref("")).toBeUndefined();
    expect(sanitizeMarkdownHref("#with space")).toBeUndefined();
    expect(sanitizeMarkdownHref("javascript:alert(1)")).toBeUndefined();
    expect(sanitizeMarkdownHref("mailto:")).toBeUndefined();
  });
});

describe("compileMarkdownToContent", () => {
  test("maps headings to Heading styles and quotes to the Quote style", () => {
    const { content } = compileMarkdownToContent("# Title\n\n### Deep\n\n> cited");
    expect(paragraphs(content).map((paragraph) => paragraph.formatting?.styleId)).toEqual([
      "Heading1",
      "Heading3",
      "Quote",
    ]);
  });

  test("a nested ordered list does not inherit a sibling's bullet level", () => {
    const { content, numbering } = compileMarkdownToContent("- a\n  - x\n- b\n  1. y");
    const nested = paragraphs(content).filter(
      (paragraph) => paragraphNumberingLevel(paragraph.formatting?.numPr) === 1,
    );
    expect(nested).toHaveLength(2);
    const [bulletItem, orderedItem] = nested;
    const bulletNumId = paragraphNumberingReferenceId(bulletItem?.formatting?.numPr);
    const orderedNumId = paragraphNumberingReferenceId(orderedItem?.formatting?.numPr);
    expect(orderedNumId).not.toBe(bulletNumId);
    const orderedLevel = numbering?.abstractNums
      .find((abstract) => abstract.abstractNumId === orderedNumId)
      ?.levels.find((level) => level.ilvl === 1);
    expect(orderedLevel?.numFmt).toBe("decimal");
  });

  test("carries no numbering without lists", () => {
    expect(compileMarkdownToContent("Just prose.").numbering).toBeUndefined();
  });

  test("a nested list stays nested, with no warning", () => {
    // The one case the block model can actually express as nesting: a deeper
    // `ilvl` under the parent item. No content was flattened, so no warning.
    const { content, warnings } = compileMarkdownToContent("1. Parent\n   - Child\n2. Next");
    expect(content.map((block) => block.type)).toEqual(["paragraph", "paragraph", "paragraph"]);
    expect(warnings).toBeUndefined();
  });

  test("a table inside a list item becomes a following table block, with a warning", () => {
    // OOXML lists are numbered paragraphs, not containers: a table cannot
    // nest inside a list item, so it must survive as a sibling block instead
    // of being silently dropped (the block model has no other way to keep
    // it).
    const source = "- Parent\n\n  | A | B |\n  | --- | --- |\n  | X | Y |\n\n- Next";
    const { content, warnings } = compileMarkdownToContent(source);
    expect(content.map((block) => block.type)).toEqual(["paragraph", "table", "paragraph"]);
    const [parentPara, table, nextPara] = content;
    expect(parentPara?.type === "paragraph" ? paragraphText(parentPara) : undefined).toBe("Parent");
    expect(nextPara?.type === "paragraph" ? paragraphText(nextPara) : undefined).toBe("Next");
    if (table?.type !== "table") {
      throw new Error("expected a table block");
    }
    const cellText = (rowIndex: number, cellIndex: number): string | undefined => {
      const cell = table.rows[rowIndex]?.cells[cellIndex];
      const cellPara = cell?.content[0];
      return cellPara?.type === "paragraph" ? paragraphText(cellPara) : undefined;
    };
    expect(cellText(0, 0)).toBe("A");
    expect(cellText(0, 1)).toBe("B");
    expect(cellText(1, 0)).toBe("X");
    expect(cellText(1, 1)).toBe("Y");
    expect(warnings).toBeDefined();
    expect(warnings?.some((warning) => warning.includes("table"))).toBe(true);
  });

  test("a code block inside a list item becomes following paragraphs, with a warning", () => {
    const source = "- Parent\n\n  ```\n  const sentinel = 1;\n  ```\n\n- Next";
    const { content, warnings } = compileMarkdownToContent(source);
    expect(content.map((block) => block.type)).toEqual(["paragraph", "paragraph", "paragraph"]);
    const [, codeLine] = content;
    expect(codeLine?.type === "paragraph" ? paragraphText(codeLine) : undefined).toBe(
      "const sentinel = 1;",
    );
    expect(warnings?.some((warning) => warning.includes("code"))).toBe(true);
  });

  test("a blockquote in a list item becomes a following Quote paragraph, with a warning", () => {
    const source = "- Parent\n\n  > cited sentinel\n\n- Next";
    const { content, warnings } = compileMarkdownToContent(source);
    expect(content.map((block) => block.type)).toEqual(["paragraph", "paragraph", "paragraph"]);
    const [, quote] = content;
    expect(quote?.type === "paragraph" ? quote.formatting?.styleId : undefined).toBe("Quote");
    expect(quote?.type === "paragraph" ? paragraphText(quote) : undefined).toBe("cited sentinel");
    expect(warnings?.some((warning) => warning.includes("blockquote"))).toBe(true);
  });

  test("every list paragraph references a numbering instance it synthesized", () => {
    const item = fc.stringMatching(/^[a-z]{1,6}$/u);
    const list = (marker: string, indent: string) =>
      fc
        .array(item, { minLength: 1, maxLength: 3 })
        .map((items) => items.map((text) => `${indent}${marker} ${text}`).join("\n"));
    const markdown = fc
      .tuple(list("-", ""), list("1.", "  "), list("-", ""))
      .map(([first, nested, second]) => `${first}\n${nested}\n\n${second}`);
    fc.assert(
      fc.property(markdown, (source) => {
        const { content, numbering } = compileMarkdownToContent(source);
        const defined = new Set(numbering?.nums.map((num) => num.numId) ?? []);
        return paragraphs(content).every(
          (paragraph) =>
            paragraph.formatting?.numPr === undefined ||
            defined.has(paragraphNumberingReferenceId(paragraph.formatting.numPr) ?? -1),
        );
      }),
      propertyConfig({ numRuns: 50 }),
    );
  });
});
