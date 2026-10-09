import { describe, expect, test } from "bun:test";

import { createBuiltInStyleIndex } from "../docx/builtInStyles";
import { parseDocx } from "../docx/parser";
import { createDocx } from "../docx/rezip";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import { createHeadingCollectorOoxmlFixture } from "./fixtures/headingCollectorOoxml";
import {
  HEADING_COLLECTOR_DOCUMENT,
  HEADING_COLLECTOR_RUN_IN_DOCUMENT,
  HEADING_COLLECTOR_STYLES,
} from "./fixtures/headingCollector.synthetic";
import { collectHeadings } from "./headingCollector";

describe("document outline collection", () => {
  test("uses the first formatted run as the title and handles empty paragraphs", () => {
    const styles = createBuiltInStyleIndex(HEADING_COLLECTOR_STYLES);
    const headings = collectHeadings(HEADING_COLLECTOR_RUN_IN_DOCUMENT, styles);

    expect(headings.map(({ text, level }) => [text, level])).toEqual([
      ["Article 7. Scope", 0],
      ["Single run remains the fallback title", 1],
      ["Visible fallback after an empty formatted span", 0],
      ["Underlined title and more", 1],
      ["Bold run", 0],
    ]);
  });

  test("collects heading semantics without inferring headings from numbering or typography", () => {
    const styles = createBuiltInStyleIndex(HEADING_COLLECTOR_STYLES);
    const headings = collectHeadings(HEADING_COLLECTOR_DOCUMENT, styles);

    expect(headings).toHaveLength(66);
    expect(headings.slice(0, 64).map(({ text, level }) => [text, level])).toEqual(
      Array.from({ length: 64 }, (_, index) => [
        `Section ${index + 1}: General provisions governing synthetic agreements and related obligations`,
        index % 2 === 0 ? 0 : 1,
      ]),
    );
    expect(headings.slice(64).map(({ text, level }) => [text, level])).toEqual([
      ["Inherited clause heading", 1],
      ["Explicit outline", 2],
    ]);
    const headingTexts = headings.map(({ text }) => text);
    expect(headingTexts).not.toContain("Numbered list item at level 0");
    expect(headingTexts).not.toContain("Numbered list item at level 1");
    expect(headingTexts).not.toContain("Numbered list item at level 2");
    expect(headingTexts).not.toContain(
      '"Synthetic Term" means a term used only by this synthetic document fixture.',
    );
    expect(headingTexts).not.toContain("PLAIN BOLD CAPS");
    expect(headingTexts).not.toContain(
      "The parties agree that this ordinary clause remains body text.",
    );
    expect(headingTexts).not.toContain("Body override");
    expect(headingTexts).not.toContain("Inherited heading overridden to body text");
  });
});

test("a numbering-level outline is preserved without promoting its paragraph", async () => {
  const document = await parseDocx(await createHeadingCollectorOoxmlFixture(), {
    preloadFonts: false,
  });
  const numberingOutline = document.package.numbering?.abstractNums
    .flatMap(({ levels }) => levels)
    .find(({ ilvl }) => ilvl === 0)?.pPr?.outlineLevel;
  expect(numberingOutline).toEqual({ kind: "heading", level: 2 });

  const paragraphs = document.package.document.content.filter(
    (block) => block.type === "paragraph",
  );
  const numberedBody = paragraphs.find(({ content }) =>
    content.some(
      (item) =>
        item.type === "run" &&
        item.content.some(
          (part) =>
            part.type === "text" && part.text === "Numbered body with numbering-only outline level",
        ),
    ),
  );
  expect(numberedBody?.formatting?.outlineLevel).toBeUndefined();

  const styles = createBuiltInStyleIndex(document.package.styles?.styles ?? []);
  const headings = collectHeadings(toProseDoc(document), styles);
  expect(headings.map(({ text, level }) => [text, level])).toEqual([
    ["Real Heading 1", 0],
    ["Real Heading 2", 1],
    ["Inherited Heading", 1],
    ["Direct paragraph outline", 3],
  ]);

  const reopened = await parseDocx(await createDocx(document), { preloadFonts: false });
  const reopenedOutline = reopened.package.numbering?.abstractNums
    .flatMap(({ levels }) => levels)
    .find(({ ilvl }) => ilvl === 0)?.pPr?.outlineLevel;
  expect(reopenedOutline).toEqual(numberingOutline);
});
