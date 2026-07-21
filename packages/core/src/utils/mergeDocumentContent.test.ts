// mergeDocumentContent is the general replacement for hand-merging a parsed
// document's content blocks onto another document (e.g. `fromMarkdown`'s
// output onto a styled preset): appending `source.package.document.content`
// directly is unsafe once `source` carries its own numbering, because
// `fromMarkdown` mints small numIds (1, 2, …) that collide with a preset's
// reserved numbering (see stellaStyle.ts's `createLegalNumbering`, which
// reserves numId 1-5 for clause/definitions/recitals/parties/bullet). A
// colliding markdown list silently renders with the preset's clause/
// definition marker instead of a plain bullet/number.

import { describe, expect, test } from "bun:test";

import { docxToMarkdown } from "../docx/server/docxToMarkdown";
import { createDocx } from "../docx/rezip";
import { fromMarkdown } from "../markdown/fromMarkdown";
import { createStellaStyleDocumentPreset } from "../style-sets/stellaStyle";
import type { Document, Paragraph } from "../types/document";
import { createEmptyDocument } from "./createDocument";
import { mergeDocumentContent } from "./mergeDocumentContent";

const CLEAN = {
  annotations: "strip",
  trackedChanges: "clean",
  comments: "strip",
  hyperlinks: "inline",
  footnotes: "strip",
} as const;

const clauseParagraph = (text: string): Paragraph => ({
  type: "paragraph",
  formatting: { numPr: { numId: 1, ilvl: 0 }, styleId: "ClauseParagraph1" },
  content: [{ type: "run", formatting: {}, content: [{ type: "text", text }] }],
});

const stellaTargetWithClause = (): Document => {
  const target = createEmptyDocument({ preset: createStellaStyleDocumentPreset() });
  target.package.document.content = [clauseParagraph("Clause text.")];
  return target;
};

describe("mergeDocumentContent", () => {
  test("appends source content and merges numbering definitions", () => {
    const target = stellaTargetWithClause();
    const merged = mergeDocumentContent(target, fromMarkdown("1. first\n2. second"));

    expect(merged.package.document.content).toHaveLength(3);
    // Target's own reserved numbering (numId 1-5) is untouched.
    expect(merged.package.numbering?.nums.filter((num) => num.numId <= 5)).toEqual(
      target.package.numbering?.nums,
    );
  });

  test("does not mutate either input document", () => {
    const target = stellaTargetWithClause();
    const source = fromMarkdown("1. first\n2. second");
    const targetContentLength = target.package.document.content.length;
    const sourceNumIds = source.package.numbering?.nums.map((num) => num.numId);

    mergeDocumentContent(target, source);

    expect(target.package.document.content).toHaveLength(targetContentLength);
    expect(source.package.numbering?.nums.map((num) => num.numId)).toEqual(sourceNumIds);
  });

  test("remaps a colliding source numId above the target's existing range", () => {
    const target = stellaTargetWithClause();
    const source = fromMarkdown("1. first\n2. second");
    // fromMarkdown mints numId 1 for its first list — the same id Stella's
    // preset reserves for clause numbering.
    expect(source.package.numbering?.nums.map((num) => num.numId)).toEqual([1]);

    const merged = mergeDocumentContent(target, source);
    const mergedListParagraph = merged.package.document.content.find(
      (block) => block.type === "paragraph" && block.listRendering,
    );
    expect(mergedListParagraph?.type).toBe("paragraph");
    const remappedNumId =
      mergedListParagraph?.type === "paragraph"
        ? mergedListParagraph.formatting?.numPr?.numId
        : undefined;
    expect(remappedNumId).toBeGreaterThan(5);
  });

  test("createDocx(merged) succeeds and round-trips both lists correctly", async () => {
    const target = stellaTargetWithClause();
    const merged = mergeDocumentContent(target, fromMarkdown("1. first\n2. second\n\n- bullet"));

    const bytes = await createDocx(merged);
    const markdown = await docxToMarkdown(bytes, CLEAN);

    // The markdown list keeps its own identity (plain "1."/"- "), not
    // Stella's clause "(a)"/definitions marker.
    expect(markdown).toContain("1. first\n2. second");
    expect(markdown).toContain("- bullet");

    // The preset's clause paragraph still renders with clause numbering
    // (decimal at level 0), not a plain unnumbered paragraph.
    const clauseLine = markdown.split("\n").find((line) => line.includes("Clause text."));
    expect(clauseLine).toMatch(/^1[.)]?\s+Clause text\.$/);
  });
});
