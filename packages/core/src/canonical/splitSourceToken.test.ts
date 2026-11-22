/**
 * A split's new half derives its properties from the paragraph it was split
 * from, and a durable source token names one paragraph: the paragraph keeps
 * its token and the half takes none, whichever of the two the edit rebuilt.
 */
import { panic } from "better-result";
import { expect, test } from "bun:test";

import {
  assignDocumentParagraphPropertySourceContract,
  assignParagraphPropertySource,
  getParagraphPropertySource,
  getParagraphPropertySourceToken,
  inspectParagraphSourceIdentity,
} from "../docx/paragraphPropertySource";
import type { BlockContent, Document, Paragraph } from "../types/document";
import {
  preservePropertySources,
  preservePropertySourcesWholeStories,
} from "./propertySourcePreservation";

const SOURCE = '<w:pPr><w:jc w:val="center"/></w:pPr>';

const paragraph = (paraId: string, text: string): Paragraph => ({
  type: "paragraph",
  paraId,
  formatting: { alignment: "center" },
  content:
    text === "" ? [] : [{ type: "run", content: [{ type: "text", text }] }],
});

const documentOf = (content: BlockContent[]): Document => ({
  package: { document: { content } },
});

const source = (): Document => {
  const captured = paragraph("00000001", "abc");
  assignParagraphPropertySource(captured, SOURCE);
  const document = documentOf([captured, paragraph("00000002", "d")]);
  assignDocumentParagraphPropertySourceContract(document, "b".repeat(64));
  return document;
};

const paragraphs = (document: Document): Paragraph[] =>
  document.package.document.content.filter(
    (block): block is Paragraph => block.type === "paragraph",
  );

const tokens = (document: Document): (string | undefined)[] =>
  paragraphs(document).map(getParagraphPropertySourceToken);

for (const [name, preserve] of Object.entries({
  changed: preservePropertySources,
  whole: preservePropertySourcesWholeStories,
}))
  for (const [first, second] of [
    ["abc", ""],
    ["a", "bc"],
    ["", "abc"],
  ] as const)
    test(`${name} stories: a half "${second}" split from "${first}" takes no token`, () => {
      const before = source();
      const [token, other] = tokens(before);
      const after = preserve({
        target: documentOf([
          paragraph("00000001", first),
          paragraph("00000003", second),
          paragraphs(before).at(1) ?? panic("The fixture has two paragraphs."),
        ]),
        source: before,
      });
      expect(tokens(after)).toEqual([token, undefined, other]);
      const half =
        paragraphs(after).at(1) ?? panic("The split leaves two halves.");
      // The half still reads its properties from the same source.
      expect(getParagraphPropertySource(half)).toBeDefined();
      expect(inspectParagraphSourceIdentity(half, after).type).toBe(
        "capture-only",
      );
    });
