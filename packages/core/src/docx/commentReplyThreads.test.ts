/**
 * Reply-thread round-trip tests.
 *
 * A threaded comment (a parent + replies, linked in commentsExtended.xml) must
 * survive a save→re-parse: the reply keeps its `parentId` / `done`, and its
 * `document.xml` anchor (commentRange markers + reference) is present after both
 * the selective and the full-repack save paths. Also covers the create-reply
 * API producing a reply that round-trips as a proper reply.
 */

import { describe, expect, test } from "bun:test";
import JSZip from "jszip";

import type { BlockContent, Comment, Document, Paragraph } from "../types/document";
import { parseDocx } from "./parser";
import { RELATIONSHIP_TYPES } from "./relsParser";
import { replyToComment } from "./replyToComment";
import { repackDocx } from "./rezip";
import { attemptSelectiveSave } from "./selectiveSave";
import { serializeCommentsExtended } from "./serializer/commentSerializer";

const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

const BODY_PARA_ID = "B0000001";
const PARENT_PARA_ID = "AAAA0001";
const REPLY1_PARA_ID = "AAAA0002";
const REPLY2_PARA_ID = "AAAA0003";

const PARENT_ID = 1;
const REPLY1_ID = 2;
const REPLY2_ID = 3;

type FixtureOptions = {
  /** Emit the replies' own commentRange markers in document.xml (Word's shape).
   *  When false, only the parent is anchored so save must synthesize them. */
  includeReplyMarkers: boolean;
  /** Mark the parent thread resolved (`w15:done="1"`). */
  parentDone: boolean;
};

const anchorParagraph = (includeReplyMarkers: boolean): string => {
  const starts = includeReplyMarkers
    ? `<w:commentRangeStart w:id="${PARENT_ID}"/><w:commentRangeStart w:id="${REPLY1_ID}"/><w:commentRangeStart w:id="${REPLY2_ID}"/>`
    : `<w:commentRangeStart w:id="${PARENT_ID}"/>`;
  const ref = (id: number): string =>
    `<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="${id}"/></w:r>`;
  const ends = includeReplyMarkers
    ? `<w:commentRangeEnd w:id="${PARENT_ID}"/>${ref(PARENT_ID)}<w:commentRangeEnd w:id="${REPLY1_ID}"/>${ref(REPLY1_ID)}<w:commentRangeEnd w:id="${REPLY2_ID}"/>${ref(REPLY2_ID)}`
    : `<w:commentRangeEnd w:id="${PARENT_ID}"/>${ref(PARENT_ID)}`;
  return `<w:p w14:paraId="${BODY_PARA_ID}">${starts}<w:r><w:t xml:space="preserve">anchored text</w:t></w:r>${ends}</w:p>`;
};

const documentXmlFor = (includeReplyMarkers: boolean): string =>
  `${XML_DECLARATION}
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    ${anchorParagraph(includeReplyMarkers)}
    <w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>
  </w:body>
</w:document>`;

const comment = (id: number, paraId: string, author: string, text: string): string =>
  `<w:comment w:id="${id}" w:author="${author}" w:initials="${author[0]}" w:date="2026-05-15T00:00:00Z"><w:p w14:paraId="${paraId}"><w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:annotationRef/></w:r><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p></w:comment>`;

const commentsXml = `${XML_DECLARATION}
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">${comment(PARENT_ID, PARENT_PARA_ID, "Alice", "parent question")}${comment(REPLY1_ID, REPLY1_PARA_ID, "Bob", "first reply")}${comment(REPLY2_ID, REPLY2_PARA_ID, "Cara", "second reply")}</w:comments>`;

const commentsExtendedXmlFor = (parentDone: boolean): string =>
  `${XML_DECLARATION}
<w15:commentsEx xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"><w15:commentEx w15:paraId="${PARENT_PARA_ID}" w15:done="${parentDone ? "1" : "0"}"/><w15:commentEx w15:paraId="${REPLY1_PARA_ID}" w15:paraIdParent="${PARENT_PARA_ID}" w15:done="0"/><w15:commentEx w15:paraId="${REPLY2_PARA_ID}" w15:paraIdParent="${PARENT_PARA_ID}" w15:done="0"/></w15:commentsEx>`;

const stylesXml = `${XML_DECLARATION}
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>`;

const corePropertiesXml = `${XML_DECLARATION}
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dcterms:modified xsi:type="dcterms:W3CDTF">2024-01-01T00:00:00.000Z</dcterms:modified></cp:coreProperties>`;

const contentTypesXml = `${XML_DECLARATION}
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>
  <Override PartName="/word/commentsExtended.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
</Types>`;

const packageRelsXml = `${XML_DECLARATION}
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="${RELATIONSHIP_TYPES.officeDocument}" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
</Relationships>`;

const documentRelsXml = `${XML_DECLARATION}
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="${RELATIONSHIP_TYPES.styles}" Target="styles.xml"/>
  <Relationship Id="rId2" Type="${RELATIONSHIP_TYPES.comments}" Target="comments.xml"/>
  <Relationship Id="rId3" Type="${RELATIONSHIP_TYPES.commentsExtended}" Target="commentsExtended.xml"/>
</Relationships>`;

const buildThreadedDocx = (options: FixtureOptions): Promise<ArrayBuffer> => {
  const zip = new JSZip();
  zip.file("[Content_Types].xml", contentTypesXml);
  zip.file("_rels/.rels", packageRelsXml);
  zip.file("word/_rels/document.xml.rels", documentRelsXml);
  zip.file("word/document.xml", documentXmlFor(options.includeReplyMarkers));
  zip.file("word/styles.xml", stylesXml);
  zip.file("word/comments.xml", commentsXml);
  zip.file("word/commentsExtended.xml", commentsExtendedXmlFor(options.parentDone));
  zip.file("docProps/core.xml", corePropertiesXml);
  return zip.generateAsync({ type: "arraybuffer" });
};

// A single-comment (non-threaded) document, for the create-reply-via-API case.
const singleCommentDocumentXml = `${XML_DECLARATION}
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p w14:paraId="${BODY_PARA_ID}"><w:commentRangeStart w:id="${PARENT_ID}"/><w:r><w:t xml:space="preserve">anchored text</w:t></w:r><w:commentRangeEnd w:id="${PARENT_ID}"/><w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="${PARENT_ID}"/></w:r></w:p>
    <w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>
  </w:body>
</w:document>`;

const singleCommentXml = `${XML_DECLARATION}
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">${comment(PARENT_ID, PARENT_PARA_ID, "Alice", "parent question")}</w:comments>`;

const singleCommentContentTypes = contentTypesXml.replace(
  /\s*<Override PartName="\/word\/commentsExtended.xml"[^>]*\/>/u,
  "",
);

const singleCommentDocumentRels = `${XML_DECLARATION}
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="${RELATIONSHIP_TYPES.styles}" Target="styles.xml"/>
  <Relationship Id="rId2" Type="${RELATIONSHIP_TYPES.comments}" Target="comments.xml"/>
</Relationships>`;

const buildSingleCommentDocx = (): Promise<ArrayBuffer> => {
  const zip = new JSZip();
  zip.file("[Content_Types].xml", singleCommentContentTypes);
  zip.file("_rels/.rels", packageRelsXml);
  zip.file("word/_rels/document.xml.rels", singleCommentDocumentRels);
  zip.file("word/document.xml", singleCommentDocumentXml);
  zip.file("word/styles.xml", stylesXml);
  zip.file("word/comments.xml", singleCommentXml);
  zip.file("docProps/core.xml", corePropertiesXml);
  return zip.generateAsync({ type: "arraybuffer" });
};

// ============================================================================
// MODEL PROJECTIONS
// ============================================================================

type MarkerIds = {
  rangeStart: Set<number>;
  rangeEnd: Set<number>;
  reference: Set<number>;
};

const eachParagraph = (blocks: readonly BlockContent[], visit: (p: Paragraph) => void): void => {
  for (const block of blocks) {
    if (block.type === "paragraph") {
      visit(block);
    } else if (block.type === "table") {
      for (const row of block.rows) {
        for (const cell of row.cells) {
          eachParagraph(cell.content, visit);
        }
      }
    } else {
      eachParagraph(block.content, visit);
    }
  }
};

const collectMarkerIds = (doc: Document): MarkerIds => {
  const ids: MarkerIds = { rangeStart: new Set(), rangeEnd: new Set(), reference: new Set() };
  eachParagraph(doc.package.document.content, (paragraph) => {
    for (const item of paragraph.content) {
      if (item.type === "commentRangeStart") {
        ids.rangeStart.add(item.id);
      } else if (item.type === "commentRangeEnd") {
        ids.rangeEnd.add(item.id);
      } else if (item.type === "commentReference") {
        ids.reference.add(item.id);
      }
    }
  });
  return ids;
};

const commentById = (doc: Document, id: number): Comment | undefined =>
  (doc.package.document.comments ?? []).find((c) => c.id === id);

const parse = (buffer: ArrayBuffer): Promise<Document> =>
  parseDocx(buffer, { preloadFonts: false });

const assertRepliesThreaded = (doc: Document): void => {
  expect(commentById(doc, REPLY1_ID)?.parentId).toBe(PARENT_ID);
  expect(commentById(doc, REPLY2_ID)?.parentId).toBe(PARENT_ID);
  // A reply is fully anchored: it has a range AND a reference of its own.
  const markers = collectMarkerIds(doc);
  for (const replyId of [REPLY1_ID, REPLY2_ID]) {
    expect(markers.rangeStart.has(replyId)).toBe(true);
    expect(markers.rangeEnd.has(replyId)).toBe(true);
    expect(markers.reference.has(replyId)).toBe(true);
  }
};

// ============================================================================
// TESTS
// ============================================================================

describe("comment reply threads — parse", () => {
  test("parses parentId and done from commentsExtended.xml", async () => {
    const doc = await parse(
      await buildThreadedDocx({ includeReplyMarkers: true, parentDone: true }),
    );

    expect(commentById(doc, PARENT_ID)?.parentId).toBeUndefined();
    expect(commentById(doc, PARENT_ID)?.done).toBe(true);
    expect(commentById(doc, REPLY1_ID)?.parentId).toBe(PARENT_ID);
    expect(commentById(doc, REPLY1_ID)?.done).toBe(false);
    expect(commentById(doc, REPLY2_ID)?.parentId).toBe(PARENT_ID);
  });
});

describe("comment reply threads — round-trip", () => {
  test("full repack keeps replies threaded and anchored", async () => {
    const buffer = await buildThreadedDocx({ includeReplyMarkers: true, parentDone: true });
    const doc = await parse(buffer);

    const saved = await repackDocx({ ...doc, originalBuffer: buffer });
    const reparsed = await parse(saved);

    assertRepliesThreaded(reparsed);
    expect(commentById(reparsed, PARENT_ID)?.done).toBe(true);
  });

  test("selective save keeps replies threaded (commentsExtended byte-exact)", async () => {
    const buffer = await buildThreadedDocx({ includeReplyMarkers: true, parentDone: false });
    const doc = await parse(buffer);

    const saved = await attemptSelectiveSave(doc, buffer, {
      changedParaIds: new Set(),
      structuralChange: false,
      hasUntrackedChanges: false,
    });
    expect(saved).not.toBeNull();
    if (!saved) {
      return;
    }

    const originalExtended = await readPart(buffer, "word/commentsExtended.xml");
    const savedExtended = await readPart(saved, "word/commentsExtended.xml");
    expect(savedExtended).toBe(originalExtended);

    assertRepliesThreaded(await parse(saved));
  });

  test("full repack synthesizes reply markers when only the parent is anchored", async () => {
    const buffer = await buildThreadedDocx({ includeReplyMarkers: false, parentDone: false });
    const doc = await parse(buffer);

    // The parsed model has the replies threaded but WITHOUT their own markers.
    const beforeMarkers = collectMarkerIds(doc);
    expect(beforeMarkers.rangeStart.has(REPLY1_ID)).toBe(false);
    expect(commentById(doc, REPLY1_ID)?.parentId).toBe(PARENT_ID);

    const saved = await repackDocx({ ...doc, originalBuffer: buffer });
    assertRepliesThreaded(await parse(saved));
  });

  test("selective save bails when a reply lacks its own anchor", async () => {
    const buffer = await buildThreadedDocx({ includeReplyMarkers: false, parentDone: false });
    const doc = await parse(buffer);
    const saved = await attemptSelectiveSave(doc, buffer, {
      changedParaIds: new Set(),
      structuralChange: false,
      hasUntrackedChanges: false,
    });
    // The reply's anchor must be synthesized on the parent's paragraph, which
    // the selective path cannot do — it hands off to the full repack.
    expect(saved).toBeNull();
  });
});

describe("comment reply threads — create-reply API", () => {
  test("replyToComment produces a reply that round-trips as a proper reply", async () => {
    const buffer = await buildSingleCommentDocx();
    const doc = await parse(buffer);

    const reply = replyToComment(doc, PARENT_ID, { author: "Reviewer", text: "a fresh reply" });
    expect(reply.parentId).toBe(PARENT_ID);

    // A created reply has no anchor yet, so save goes through the full repack.
    const saved = await repackDocx({ ...doc, originalBuffer: buffer });
    const reparsed = await parse(saved);

    const created = commentById(reparsed, reply.id);
    expect(created?.parentId).toBe(PARENT_ID);
    const markers = collectMarkerIds(reparsed);
    expect(markers.rangeStart.has(reply.id)).toBe(true);
    expect(markers.reference.has(reply.id)).toBe(true);
  });

  test("serializeCommentsExtended is null for a document with no threads or resolved state", () => {
    const plain: Comment[] = [
      {
        id: 1,
        author: "Alice",
        content: [{ type: "paragraph", formatting: {}, content: [] }],
      },
    ];
    expect(serializeCommentsExtended(plain)).toBeNull();
  });
});

async function readPart(buffer: ArrayBuffer, path: string): Promise<string> {
  const zip = await JSZip.loadAsync(buffer);
  const file = zip.file(path);
  if (!file) {
    throw new Error(`No ${path} in package`);
  }
  return file.async("text");
}
