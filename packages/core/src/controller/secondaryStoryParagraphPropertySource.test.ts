import { panic } from "better-result";
import { describe, expect, test } from "bun:test";

import {
  assignParagraphPropertySource,
  cloneDocumentWithParagraphPropertySources,
} from "../docx/paragraphPropertySource";
import { serializeParagraph } from "../docx/serializer/paragraphSerializer";
import { NAMESPACES } from "../docx/xmlParser";
import { proseDocToBlocks } from "../prosemirror/conversion/fromProseDoc";
import {
  footnoteToProseDoc,
  headerFooterToProseDoc,
  toProseDoc,
} from "../prosemirror/conversion/toProseDoc";
import type { Document, Paragraph } from "../types/document";

const assignOwnerSource = (paragraph: Paragraph, owner: string): void => {
  assignParagraphPropertySource(
    paragraph,
    `<w:pPr xmlns:w="${NAMESPACES.w}" xmlns:x="urn:folio:test">` +
      `<x:property x:owner="${owner}"/></w:pPr>`,
  );
};

const paragraphWithSource = (
  text: string,
  owner: string,
  paraId?: string,
): Paragraph => {
  const paragraph: Paragraph = {
    type: "paragraph",
    ...(paraId ? { paraId } : {}),
    content: [{ type: "run", content: [{ type: "text", text }] }],
  };
  assignOwnerSource(paragraph, owner);
  return paragraph;
};

const expectOwner = (paragraph: Paragraph, owner: string): void => {
  const xml = serializeParagraph(paragraph);
  expect(xml).toContain(`x:owner="${owner}"`);
  expect(xml.match(/x:owner=/gu)).toHaveLength(1);
};

describe("secondary-story paragraph property ownership", () => {
  test("document cloning retains captures across every serialised story", () => {
    const body = paragraphWithSource("body source", "body", "11111111");
    // A bare body conversion materialises the built-in Normal spacing.
    body.formatting = { spaceAfter: 160 };
    assignOwnerSource(body, "body");
    const header = paragraphWithSource("header source", "header", "22222222");
    const footer = paragraphWithSource("footer source", "footer", "33333333");
    const footnote = paragraphWithSource(
      "footnote source",
      "footnote",
      "44444444",
    );
    const endnote = paragraphWithSource(
      "endnote source",
      "endnote",
      "55555555",
    );
    const document: Document = {
      package: {
        document: { content: [body] },
        headers: new Map([
          [
            "rId-header",
            { type: "header", hdrFtrType: "default", content: [header] },
          ],
        ]),
        footers: new Map([
          [
            "rId-footer",
            { type: "footer", hdrFtrType: "default", content: [footer] },
          ],
        ]),
        footnotes: [{ type: "footnote", id: 1, content: [footnote] }],
        endnotes: [{ type: "endnote", id: 1, content: [endnote] }],
      },
    };

    const cloned = cloneDocumentWithParagraphPropertySources(document);
    const clonedStories = [
      cloned.package.document.content.at(0),
      cloned.package.headers?.get("rId-header")?.content.at(0),
      cloned.package.footers?.get("rId-footer")?.content.at(0),
      cloned.package.footnotes?.at(0)?.content.at(0),
      cloned.package.endnotes?.at(0)?.content.at(0),
    ];
    const originals = [body, header, footer, footnote, endnote];
    const owners = ["body", "header", "footer", "footnote", "endnote"];

    for (const [index, owner] of owners.entries()) {
      const clonedStory = clonedStories.at(index);
      if (clonedStory?.type !== "paragraph") {
        panic(`The cloned document lost its ${owner} paragraph.`);
      }
      expect(clonedStory).not.toBe(originals.at(index));
      expectOwner(clonedStory, owner);
    }

    const bodyProseDocument = toProseDoc({
      package: { document: { content: [body] } },
    });
    const proseDocuments = [
      bodyProseDocument,
      headerFooterToProseDoc([header]),
      headerFooterToProseDoc([footer]),
      footnoteToProseDoc([footnote]),
      footnoteToProseDoc([endnote]),
    ];

    for (const [index, owner] of owners.entries()) {
      const proseDocument = proseDocuments.at(index);
      const source = clonedStories.at(index);
      if (!proseDocument || source?.type !== "paragraph") {
        panic(`The ${owner} conversion fixture is incomplete.`);
      }
      const rebuiltProseDocument = proseDocument.type.schema.nodeFromJSON(
        proseDocument.toJSON(),
      );
      const converted = proseDocToBlocks(rebuiltProseDocument, [source]).at(0);
      if (converted?.type !== "paragraph") {
        panic(`The ${owner} conversion lost its paragraph.`);
      }
      expectOwner(converted, owner);
    }
  });

  test("omitting a detached story's base loses its captured properties", () => {
    const source = paragraphWithSource("header source", "header", "12345678");
    const proseDocument = headerFooterToProseDoc([source]);
    const rebuiltProseDocument = proseDocument.type.schema.nodeFromJSON(
      proseDocument.toJSON(),
    );

    const convertedWithoutBase = proseDocToBlocks(rebuiltProseDocument).at(0);
    if (convertedWithoutBase?.type !== "paragraph") {
      panic("The omitted-base mutation lost its paragraph.");
    }
    expect(serializeParagraph(convertedWithoutBase)).not.toContain(
      'x:owner="header"',
    );

    const convertedWithBase = proseDocToBlocks(rebuiltProseDocument, [
      source,
    ]).at(0);
    if (convertedWithBase?.type !== "paragraph") {
      panic("The base-aware mutation lost its paragraph.");
    }
    expectOwner(convertedWithBase, "header");
  });
});
