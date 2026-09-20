/**
 * Every consumer that asks "is this paragraph a heading, and of what level"
 * must give the same answer, and none of them may change its answer when the
 * document's style ids are renamed.
 *
 * The first property is what makes `builtInStyles.ts` the single owner: four
 * independent call sites (the outline/TOC collector, the AI snapshot, the
 * markdown exporter, the bilingual builder) are compared against the
 * classifier and against each other over generated style packages. The second
 * is the localisation invariant stated directly: `w:styleId` is an opaque,
 * document-local token, so pushing every id through a bijection is a no-op on
 * anything folio decides.
 */

import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import { propertyConfig, propertyTestTimeout } from "../../../../test/property-testing";

import { createBuiltInStyleIndex, resolveHeadingLevel } from "./builtInStyles";
import { createBilingualDocument } from "./server/createBilingualDocument";
import { createFolioAIEditSnapshotWithStyleResolver } from "../ai-edits/snapshot";
import { toMarkdown } from "../markdown";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import { createStyleEngine } from "../style-engine/styleEngine";
import type { Document, Paragraph, Style, StyleDefinitions } from "../types/document";
import { collectHeadings } from "../utils/headingCollector";

setDefaultTimeout(propertyTestTimeout(30_000));

/** Markdown caps at six levels; Word goes to nine. */
const MAX_MARKDOWN_HEADING_LEVEL = 6;

/**
 * Names drawn from the built-ins plus noise that merely looks like one: the
 * localized display names ODF exports write, and custom style names.
 */
const STYLE_NAMES = [
  "heading 1",
  "heading 2",
  "heading 3",
  "heading 4",
  "heading 5",
  "heading 6",
  "heading 7",
  "heading 8",
  "heading 9",
  "Heading 1",
  "Heading2",
  "HEADING  3",
  "Normal",
  "Title",
  "Subtitle",
  "Quote",
  "Intense Quote",
  "List Paragraph",
  "TOC Heading",
  "toc 1",
  // Noise: localized display names and custom styles, none of them built-ins.
  "Überschrift 1",
  "Nadpis 2",
  "Nagłówek 3",
  "Címsor 4",
  "Encabezado 5",
  "Titre 6",
  "Clause Heading",
  "Body Text",
  "heading",
  "heading 10",
  "subheading 1",
  "MyStyle",
] as const;

/**
 * Ids with no relationship to the style's meaning: what a localized Word, an
 * ODF export and a generator actually write.
 */
const STYLE_IDS = [
  "Heading1",
  "Heading2",
  "Nadpis1",
  "berschrift2",
  "Titre3",
  "Nagwek4",
  "Cmsor5",
  "Ttulo6",
  "Overskrift7",
  "style7",
  "1",
  "a",
  "Überschrift1",
  "标题2",
  "x-y_z",
] as const;

const styleArbitrary = fc.record({
  styleId: fc.constantFrom(...STYLE_IDS),
  name: fc.option(fc.constantFrom(...STYLE_NAMES), { nil: undefined }),
  outlineLevel: fc.option(fc.integer({ min: 0, max: 9 }), { nil: undefined }),
  basedOn: fc.option(fc.constantFrom(...STYLE_IDS), { nil: undefined }),
});

type GeneratedStyle = {
  styleId: string;
  name: string | undefined;
  outlineLevel: number | undefined;
  basedOn: string | undefined;
};

const toStyle = ({ styleId, name, outlineLevel, basedOn }: GeneratedStyle): Style => ({
  styleId,
  type: "paragraph",
  ...(name === undefined ? {} : { name }),
  ...(basedOn === undefined ? {} : { basedOn }),
  ...(outlineLevel === undefined ? {} : { pPr: { outlineLevel } }),
});

/** Distinct ids only: a package with two styles sharing an id is malformed. */
const stylesArbitrary = fc
  .uniqueArray(styleArbitrary, { minLength: 1, maxLength: 8, selector: (s) => s.styleId })
  .map((styles) => styles.map(toStyle));

const paragraphArbitrary = fc.record({
  styleId: fc.option(fc.constantFrom(...STYLE_IDS), { nil: undefined }),
  outlineLevel: fc.option(fc.integer({ min: 0, max: 9 }), { nil: undefined }),
});

const documentArbitrary = fc
  .tuple(stylesArbitrary, fc.array(paragraphArbitrary, { minLength: 1, maxLength: 6 }))
  .map(([styles, paragraphs]) => ({ styles: { styles } as StyleDefinitions, paragraphs }));

type GeneratedParagraph = { styleId: string | undefined; outlineLevel: number | undefined };

/**
 * Text that survives every exporter unchanged: no markdown block syntax to
 * escape, one block per paragraph, and unique so blocks stay addressable.
 */
const paragraphText = (index: number): string => `para${index}`;

const buildDocument = (styles: StyleDefinitions, paragraphs: GeneratedParagraph[]): Document => ({
  package: {
    document: {
      content: paragraphs.map(
        (paragraph, index): Paragraph => ({
          type: "paragraph",
          paraId: (index + 1).toString(16).padStart(8, "0").toUpperCase(),
          content: [{ type: "run", content: [{ type: "text", text: paragraphText(index) }] }],
          formatting: {
            ...(paragraph.styleId === undefined ? {} : { styleId: paragraph.styleId }),
            ...(paragraph.outlineLevel === undefined
              ? {}
              : { outlineLevel: paragraph.outlineLevel }),
          },
        }),
      ),
    },
    styles,
  },
});

/** The answer under test: `undefined` for body text, else the zero-based level. */
type Verdict = number | undefined;

const classifierVerdicts = (
  styles: StyleDefinitions,
  paragraphs: GeneratedParagraph[],
): Verdict[] => {
  const index = createBuiltInStyleIndex(styles.styles);
  return paragraphs.map((paragraph) => resolveHeadingLevel(paragraph, index));
};

const collectorVerdicts = (document: Document): Verdict[] => {
  const pmDoc = toProseDoc(document);
  const headings = collectHeadings(
    pmDoc,
    createBuiltInStyleIndex(document.package.styles?.styles ?? []),
  );
  const byText = new Map(headings.map((heading) => [heading.text, heading.level]));
  return document.package.document.content.map((_block, index) => byText.get(paragraphText(index)));
};

const snapshotVerdicts = (document: Document): Verdict[] => {
  const snapshot = createFolioAIEditSnapshotWithStyleResolver(
    toProseDoc(document),
    createStyleEngine(document.package.styles),
  );
  const byText = new Map(
    snapshot.blocks.map((block) => [
      block.text,
      block.kind === "heading" && block.headingLevel !== undefined
        ? block.headingLevel - 1
        : undefined,
    ]),
  );
  return document.package.document.content.map((_block, index) => byText.get(paragraphText(index)));
};

const markdownVerdicts = (document: Document): Verdict[] => {
  const blocks = toMarkdown(document).split("\n\n");
  const byText = new Map<string, Verdict>();
  for (const block of blocks) {
    const heading = /^(?<hashes>#{1,6}) (?<text>.*)$/u.exec(block);
    if (heading?.groups) {
      byText.set(heading.groups["text"] ?? "", (heading.groups["hashes"] ?? "").length - 1);
      continue;
    }
    // A quote block keeps its `> ` prefix; strip it so the text still matches.
    byText.set(block.replace(/^> /u, ""), undefined);
  }
  return document.package.document.content.map((_block, index) => byText.get(paragraphText(index)));
};

/**
 * The bilingual builder answers heading-or-not, with no level, so its verdict
 * is compared against the classifier's collapsed to a boolean.
 */
const bilingualIsHeading = (document: Document): boolean[] => {
  const paraIds = new Set(
    document.package.document.content.flatMap((block) =>
      block.type === "paragraph" && block.paraId !== undefined ? [block.paraId] : [],
    ),
  );
  const { rows } = createBilingualDocument(document, {
    targetStyleSuffix: "EN",
    editableParagraphIds: paraIds,
  });
  const byText = new Map(
    rows.flatMap((row) => (row.kind === "table" ? [] : [[row.sourceText, row.kind === "heading"]])),
  );
  return document.package.document.content.map(
    (_block, index) => byText.get(paragraphText(index)) ?? false,
  );
};

/**
 * Markdown cannot express levels seven to nine, so it clamps. Every other
 * consumer carries the full range; clamp the reference the same way rather
 * than weakening the comparison for all of them.
 */
const clampToMarkdown = (verdict: Verdict): Verdict =>
  verdict === undefined ? undefined : Math.min(verdict, MAX_MARKDOWN_HEADING_LEVEL - 1);

describe("every heading consumer agrees with the classifier", () => {
  test("the collector, the AI snapshot, markdown and the bilingual builder give one answer", () => {
    fc.assert(
      fc.property(documentArbitrary, ({ styles, paragraphs }) => {
        const document = buildDocument(styles, paragraphs);
        const expected = classifierVerdicts(styles, paragraphs);

        expect(collectorVerdicts(document)).toEqual(expected);
        expect(snapshotVerdicts(document)).toEqual(expected);
        expect(markdownVerdicts(document)).toEqual(expected.map(clampToMarkdown));
        expect(bilingualIsHeading(document)).toEqual(
          expected.map((verdict) => verdict !== undefined),
        );
      }),
      propertyConfig({ numRuns: 200 }),
    );
  });
});

/** Map every id to a fresh one, keeping the mapping injective. */
const renameStyleIds = (
  styles: StyleDefinitions,
  paragraphs: GeneratedParagraph[],
): { styles: StyleDefinitions; paragraphs: GeneratedParagraph[] } => {
  const renamed = new Map<string, string>();
  for (const [index, style] of styles.styles.entries()) {
    renamed.set(style.styleId, `renamed${index}`);
  }
  // The bijection is over the ids the package *defines*. A `w:pStyle` naming a
  // style the package never defines is not one of them: it is content, and
  // `resolveHeadingLevel`'s last tier deliberately reads such an id as the
  // English built-in it names. Leave it alone.
  const rename = (styleId: string): string => renamed.get(styleId) ?? styleId;
  return {
    styles: {
      ...styles,
      styles: styles.styles.map((style) => ({
        ...style,
        styleId: rename(style.styleId),
        ...(style.basedOn === undefined ? {} : { basedOn: rename(style.basedOn) }),
      })),
    },
    paragraphs: paragraphs.map((paragraph) => ({
      ...paragraph,
      ...(paragraph.styleId === undefined ? {} : { styleId: rename(paragraph.styleId) }),
    })),
  };
};

describe("style ids are opaque", () => {
  test("renaming every style id through a bijection changes no consumer's answer", () => {
    fc.assert(
      fc.property(documentArbitrary, ({ styles, paragraphs }) => {
        const before = buildDocument(styles, paragraphs);
        const renamed = renameStyleIds(styles, paragraphs);
        const after = buildDocument(renamed.styles, renamed.paragraphs);

        expect(classifierVerdicts(renamed.styles, renamed.paragraphs)).toEqual(
          classifierVerdicts(styles, paragraphs),
        );
        expect(collectorVerdicts(after)).toEqual(collectorVerdicts(before));
        expect(snapshotVerdicts(after)).toEqual(snapshotVerdicts(before));
        expect(markdownVerdicts(after)).toEqual(markdownVerdicts(before));
        expect(bilingualIsHeading(after)).toEqual(bilingualIsHeading(before));
        // Markdown carries no style ids at all, so the rendered text is
        // byte-identical, not merely equivalent.
        expect(toMarkdown(after)).toEqual(toMarkdown(before));
      }),
      propertyConfig({ numRuns: 200 }),
    );
  });
});
