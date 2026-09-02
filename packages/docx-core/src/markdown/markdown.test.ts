import { describe, expect, test } from "bun:test";
import fc from "fast-check";

import { propertyConfig } from "../../../../test/property-testing";
import type { BlockContent, Paragraph, ParagraphContent, Run } from "../model/document";
import { compileMarkdownToContent } from "./content";
import { sanitizeMarkdownHref } from "./href";
import { inlineMarkdownToRuns } from "./inline";

const paragraphs = (content: BlockContent[]): Paragraph[] =>
  content.flatMap((block) => (block.type === "paragraph" ? [block] : []));

const runText = (run: Run): string =>
  run.content.map((node) => (node.type === "text" ? node.text : "")).join("");

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
      (paragraph) => paragraph.formatting?.numPr?.ilvl === 1,
    );
    expect(nested).toHaveLength(2);
    const [bulletItem, orderedItem] = nested;
    const bulletNumId = bulletItem?.formatting?.numPr?.numId;
    const orderedNumId = orderedItem?.formatting?.numPr?.numId;
    expect(orderedNumId).not.toBe(bulletNumId);
    const orderedLevel = numbering?.abstractNums
      .find((abstract) => abstract.abstractNumId === orderedNumId)
      ?.levels.find((level) => level.ilvl === 1);
    expect(orderedLevel?.numFmt).toBe("decimal");
  });

  test("carries no numbering without lists", () => {
    expect(compileMarkdownToContent("Just prose.").numbering).toBeUndefined();
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
            defined.has(paragraph.formatting.numPr.numId),
        );
      }),
      propertyConfig({ numRuns: 50 }),
    );
  });
});
