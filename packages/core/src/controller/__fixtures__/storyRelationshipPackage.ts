/**
 * A package whose body, header, footer, footnote, endnote and comment each
 * hold a paragraph with a picture and a link, every part naming them by the
 * same ids, `rIdPic` and `rIdLink`, in its own relationships, each naming
 * that part's own picture and link target. Content read in one part and
 * placed into another names the wrong picture and link there unless its ids
 * are re-homed.
 */

import { panic } from "better-result";

import {
  OP_STORIES,
  storyBody,
  storyParagraphs,
} from "@stll/docx-core/ops";

import { buildRawPackage, SECTION } from "../../__tests__/documentShapes";
import {
  visitInlineContentSlots,
  visitParagraphRuns,
} from "../../docx/paragraphTraversal";
import type { RelationshipStory } from "../../docx/storyRelationshipRehoming";
import type { Document, Image, Paragraph } from "../../types/document";
import { base64ToBytes } from "../../utils/textEncoding";

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R_NS =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
const WML_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml";
const NAMESPACES = [
  `xmlns:w="${W_NS}"`,
  `xmlns:r="${R_NS}"`,
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"',
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"',
  'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"',
].join(" ");

/** A one-pixel picture; each part's copy carries its own trailing bytes. */
const PIXEL =
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==";

export type StoryPartName =
  | "body"
  | "header"
  | "footer"
  | "footnote"
  | "endnote"
  | "comment";

/** The parts of the package, each with the story its content is read in. */
export const STORY_PARTS: readonly {
  name: StoryPartName;
  story: RelationshipStory;
}[] = [
  { name: "body", story: OP_STORIES.MAIN },
  { name: "header", story: { kind: "header", rId: "rIdHeader1" } },
  { name: "footer", story: { kind: "footer", rId: "rIdFooter1" } },
  { name: "footnote", story: { kind: "footnote", id: 1 } },
  { name: "endnote", story: { kind: "endnote", id: 1 } },
  { name: "comment", story: "comments" },
];

/** The words of the paragraph each part holds its picture and link in. */
export const sourceText = (name: StoryPartName): string => `${name} source`;

/** The link target each part's `rIdLink` names. */
export const linkTarget = (name: StoryPartName): string =>
  `https://${name}.example.test/`;

const pictureBytes = (name: StoryPartName): Uint8Array => {
  const pixel = base64ToBytes(PIXEL);
  const tail = new TextEncoder().encode(name);
  const bytes = new Uint8Array(pixel.length + tail.length);
  bytes.set(pixel);
  bytes.set(tail, pixel.length);
  return bytes;
};

const picture = (id: number) =>
  '<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">' +
  `<wp:extent cx="95250" cy="95250"/><wp:docPr id="${id}" name="Picture ${id}"/>` +
  '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
  `<pic:pic><pic:nvPicPr><pic:cNvPr id="${id}" name="picture${id}.png"/><pic:cNvPicPr/></pic:nvPicPr>` +
  '<pic:blipFill><a:blip r:embed="rIdPic"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>' +
  '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="95250" cy="95250"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>' +
  "</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>";

const sourceParagraph = (name: StoryPartName, id: number, lead = "") =>
  `<w:p>${lead}<w:r><w:t xml:space="preserve">${sourceText(name)} </w:t></w:r>${picture(id)}` +
  `<w:hyperlink r:id="rIdLink"><w:r><w:t>${name} link</w:t></w:r></w:hyperlink></w:p>`;

const textParagraph = (text: string) =>
  `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;

const ownRelationships = (name: StoryPartName) =>
  `${XML_DECL}<Relationships xmlns="${REL_NS}">` +
  `<Relationship Id="rIdPic" Type="${R_NS}/image" Target="media/${name}.png"/>` +
  `<Relationship Id="rIdLink" Type="${R_NS}/hyperlink" Target="${linkTarget(name)}" TargetMode="External"/>` +
  "</Relationships>";

const media = (name: StoryPartName) => ({
  path: `word/media/${name}.png`,
  body: pictureBytes(name),
});

/** The package's bytes, every paragraph identified. */
export const storyRelationshipPackage = async (): Promise<ArrayBuffer> => {
  const bytes = await buildRawPackage({
    styles: null,
    body:
      textParagraph("Intro") +
      sourceParagraph("body", 1) +
      '<w:p><w:commentRangeStart w:id="0"/><w:r><w:t xml:space="preserve">Noted</w:t></w:r>' +
      '<w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r>' +
      '<w:r><w:footnoteReference w:id="1"/></w:r><w:r><w:endnoteReference w:id="1"/></w:r></w:p>' +
      textParagraph("Closing") +
      SECTION(
        '<w:headerReference w:type="default" r:id="rIdHeader1"/><w:footerReference w:type="default" r:id="rIdFooter1"/>',
      ),
    relationships: [
      `<Relationship Id="rIdPic" Type="${R_NS}/image" Target="media/body.png"/>`,
      `<Relationship Id="rIdLink" Type="${R_NS}/hyperlink" Target="${linkTarget("body")}" TargetMode="External"/>`,
    ],
    parts: [
      {
        path: "word/header1.xml",
        contentType: `${WML_CT}.header+xml`,
        relationshipType: "header",
        relationshipId: "rIdHeader1",
        body: `${XML_DECL}<w:hdr ${NAMESPACES}>${textParagraph("Header text")}${sourceParagraph("header", 2)}</w:hdr>`,
      },
      { path: "word/_rels/header1.xml.rels", body: ownRelationships("header") },
      {
        path: "word/footer1.xml",
        contentType: `${WML_CT}.footer+xml`,
        relationshipType: "footer",
        relationshipId: "rIdFooter1",
        body: `${XML_DECL}<w:ftr ${NAMESPACES}>${textParagraph("Footer text")}${sourceParagraph("footer", 3)}</w:ftr>`,
      },
      { path: "word/_rels/footer1.xml.rels", body: ownRelationships("footer") },
      {
        path: "word/footnotes.xml",
        contentType: `${WML_CT}.footnotes+xml`,
        relationshipType: "footnotes",
        relationshipId: "rIdFootnotes",
        body:
          `${XML_DECL}<w:footnotes ${NAMESPACES}>` +
          '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>' +
          '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>' +
          `<w:footnote w:id="1">${sourceParagraph("footnote", 4, "<w:r><w:footnoteRef/></w:r>")}</w:footnote>` +
          "</w:footnotes>",
      },
      {
        path: "word/_rels/footnotes.xml.rels",
        body: ownRelationships("footnote"),
      },
      {
        path: "word/endnotes.xml",
        contentType: `${WML_CT}.endnotes+xml`,
        relationshipType: "endnotes",
        relationshipId: "rIdEndnotes",
        body:
          `${XML_DECL}<w:endnotes ${NAMESPACES}>` +
          '<w:endnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:endnote>' +
          '<w:endnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:endnote>' +
          `<w:endnote w:id="1">${sourceParagraph("endnote", 5, "<w:r><w:endnoteRef/></w:r>")}</w:endnote>` +
          "</w:endnotes>",
      },
      {
        path: "word/_rels/endnotes.xml.rels",
        body: ownRelationships("endnote"),
      },
      {
        path: "word/comments.xml",
        contentType: `${WML_CT}.comments+xml`,
        relationshipType: "comments",
        relationshipId: "rIdComments",
        body:
          `${XML_DECL}<w:comments ${NAMESPACES}>` +
          `<w:comment w:id="0" w:author="Reviewer" w:date="2026-01-01T00:00:00Z" w:initials="R">${sourceParagraph("comment", 6)}</w:comment>` +
          "</w:comments>",
      },
      {
        path: "word/_rels/comments.xml.rels",
        body: ownRelationships("comment"),
      },
      ...STORY_PARTS.map(({ name }) => media(name)),
    ],
  });
  return bytes.slice().buffer;
};

/** The paragraphs content placed into `story` lands among. */
export const partParagraphs = (
  document: Document,
  story: RelationshipStory,
): Paragraph[] => {
  if (story === "comments")
    return (document.package.document.comments ?? []).flatMap(
      ({ content }) => content,
    );
  return storyParagraphs(storyBody(document, story)).map(
    ({ paragraph }) => paragraph,
  );
};

/** The paragraphs whose words hold `text`. */
export const paragraphsHolding = (
  document: Document,
  story: RelationshipStory,
  text: string,
): Paragraph[] =>
  partParagraphs(document, story).filter((paragraph) =>
    paragraphText(paragraph).includes(text),
  );

/** A paragraph's words, its links' among them. */
export const paragraphText = (paragraph: Paragraph): string => {
  let text = "";
  visitParagraphRuns(paragraph, (run) => {
    for (const item of run.content) if (item.type === "text") text += item.text;
  });
  return text;
};

/** The picture and link a paragraph holds. */
export const pictureAndLink = (
  paragraph: Paragraph,
): { image: Image; href: string | undefined } => {
  let image: Image | undefined;
  visitParagraphRuns(paragraph, (run) => {
    for (const item of run.content)
      if (item.type === "drawing") image ??= item.image;
  });
  let href: string | undefined;
  visitInlineContentSlots(paragraph, ({ item }) => {
    if (item.type === "hyperlink") href ??= item.href;
  });
  return { image: image ?? panic("The paragraph holds no picture."), href };
};
