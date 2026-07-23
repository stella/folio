import { describe, expect, test } from "bun:test";
import JSZip from "jszip";

import { parseDocx } from "./parser";
import { RELATIONSHIP_TYPES } from "./relsParser";
import { COMMENTS_CONTENT_TYPE, COMMENTS_EXTENDED_CONTENT_TYPE, repackDocx } from "./rezip";

const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

type FixtureOptions = {
  comments: "empty" | "one";
  commentsRelationship: boolean;
  commentsExtendedRelationship: boolean;
};

const commentsXml = (state: FixtureOptions["comments"]): string =>
  state === "empty"
    ? `${XML_DECLARATION}<x:comments xmlns:x="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>`
    : `${XML_DECLARATION}<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:comment w:id="1" w:author="Reviewer"><w:p><w:r><w:t>Review note</w:t></w:r></w:p></w:comment></w:comments>`;

const commentsExtendedXml = `${XML_DECLARATION}<w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"/>`;

const documentXml = (withComment: boolean): string => {
  const commentMarkup = withComment
    ? '<w:commentRangeStart w:id="1"/><w:r><w:t>Text</w:t></w:r><w:commentRangeEnd w:id="1"/><w:r><w:commentReference w:id="1"/></w:r>'
    : "<w:r><w:t>Text</w:t></w:r>";
  return `${XML_DECLARATION}<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>${commentMarkup}</w:p><w:sectPr/></w:body></w:document>`;
};

const buildFixture = async ({
  comments,
  commentsRelationship,
  commentsExtendedRelationship,
}: FixtureOptions): Promise<ArrayBuffer> => {
  const relationships: string[] = [];
  if (commentsRelationship) {
    relationships.push(
      `<Relationship Id="rId1" Type="${RELATIONSHIP_TYPES.comments}" Target="comments.xml"/>`,
    );
  }
  if (commentsExtendedRelationship) {
    relationships.push(
      `<Relationship Id="rId2" Type="${RELATIONSHIP_TYPES.commentsExtended}" Target="commentsExtended.xml"/>`,
    );
  }

  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    `${XML_DECLARATION}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/comments.xml" ContentType="${COMMENTS_CONTENT_TYPE}"/><Override PartName="/word/commentsExtended.xml" ContentType="${COMMENTS_EXTENDED_CONTENT_TYPE}"/></Types>`,
  );
  zip.file(
    "_rels/.rels",
    `${XML_DECLARATION}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${RELATIONSHIP_TYPES.officeDocument}" Target="word/document.xml"/></Relationships>`,
  );
  zip.file(
    "word/_rels/document.xml.rels",
    `${XML_DECLARATION}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${relationships.join("")}</Relationships>`,
  );
  zip.file("word/document.xml", documentXml(comments === "one"));
  zip.file("word/comments.xml", commentsXml(comments));
  zip.file("word/commentsExtended.xml", commentsExtendedXml);
  return zip.generateAsync({ type: "arraybuffer" });
};

describe("empty comment package preservation", () => {
  test("an empty source model keeps its comment parts and relationships byte-stable", async () => {
    const source = await buildFixture({
      comments: "empty",
      commentsRelationship: false,
      commentsExtendedRelationship: true,
    });
    const parsed = await parseDocx(source, { preloadFonts: false });
    const sourceZip = await JSZip.loadAsync(source);
    const sourceRelationships = await sourceZip.file("word/_rels/document.xml.rels")!.async("text");

    expect(parsed.package.document.comments).toBeUndefined();

    const saved = await repackDocx(parsed, { updateModifiedDate: false });
    const savedZip = await JSZip.loadAsync(saved);
    expect(await savedZip.file("word/_rels/document.xml.rels")!.async("text")).toBe(
      sourceRelationships,
    );
    expect(await savedZip.file("word/comments.xml")!.async("text")).toBe(commentsXml("empty"));
    expect(await savedZip.file("word/commentsExtended.xml")!.async("text")).toBe(
      commentsExtendedXml,
    );

    const reopened = await parseDocx(saved, { preloadFonts: false });
    expect(reopened.package.relationships).toEqual(parsed.package.relationships);
  });

  test("removing the last modeled comment still empties the source part", async () => {
    const source = await buildFixture({
      comments: "one",
      commentsRelationship: true,
      commentsExtendedRelationship: false,
    });
    const parsed = await parseDocx(source, { preloadFonts: false });
    expect(parsed.package.document.comments).toHaveLength(1);

    parsed.package.document.comments = [];
    const saved = await repackDocx(parsed, { updateModifiedDate: false });
    const savedZip = await JSZip.loadAsync(saved);
    const savedComments = await savedZip.file("word/comments.xml")!.async("text");
    const savedRelationships = await savedZip.file("word/_rels/document.xml.rels")!.async("text");

    expect(savedComments).not.toContain("<w:comment ");
    expect(savedRelationships).toContain(`Type="${RELATIONSHIP_TYPES.comments}"`);
    const reopened = await parseDocx(saved, { preloadFonts: false });
    expect(reopened.package.document.comments).toBeUndefined();
  });
});
