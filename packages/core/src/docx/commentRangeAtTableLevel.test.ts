/**
 * A comment whose range closes between a table's rows, or between its last
 * row and the table's end, is anchored by markers folio does not model inside
 * a cell: the end marker is a child of `w:tbl` (or `w:tr`), the start and the
 * `w:commentReference` sit in cell paragraphs. This is how a reply is written
 * when its range runs to the end of a table. Editing a cell paragraph that
 * holds such a reference must keep it, or the reply loses its anchor.
 */

import { describe, expect, test } from "bun:test";

import { ZipArchive } from "@stll/docx-core/zip";

import { FolioDocxReviewer } from "../ai-edits/headless";
import { RELATIONSHIP_TYPES } from "./relsParser";

const XML_DECLARATION =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W14 = "http://schemas.microsoft.com/office/word/2010/wordml";
const W15 = "http://schemas.microsoft.com/office/word/2012/wordml";

const ROOT_ID = 7;
const REPLY_ID = 6;
const ROOT_PARA_ID = "2BBB0001";
const REPLY_PARA_ID = "2BBB0002";

const ref = (id: number): string =>
  `<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="${id}"/></w:r>`;

/** One row: the reply's range opens in the first cell and closes after the row. */
const tableXml = (endPlacement: "table" | "row"): string => {
  const replyEnd = `<w:commentRangeEnd w:id="${REPLY_ID}"/>`;
  return (
    '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>' +
    '<w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>' +
    "<w:tr>" +
    `<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr><w:p w14:paraId="3CCC0001"><w:commentRangeStart w:id="${REPLY_ID}"/><w:r><w:t>Item 3f</w:t></w:r></w:p></w:tc>` +
    `<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr><w:p w14:paraId="3CCC0002"><w:r><w:t>p</w:t></w:r><w:commentRangeStart w:id="${ROOT_ID}"/><w:r><w:t xml:space="preserve">er visit</w:t></w:r></w:p></w:tc>` +
    `<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr><w:p w14:paraId="3CCC0003"><w:r><w:t>rate only</w:t></w:r><w:commentRangeEnd w:id="${ROOT_ID}"/>${ref(ROOT_ID)}${ref(REPLY_ID)}</w:p></w:tc>` +
    (endPlacement === "row" ? replyEnd : "") +
    "</w:tr>" +
    (endPlacement === "table" ? replyEnd : "") +
    "</w:tbl>"
  );
};

const documentXml = (endPlacement: "table" | "row"): string =>
  `${XML_DECLARATION}<w:document xmlns:w="${W}" xmlns:w14="${W14}"><w:body>` +
  '<w:p w14:paraId="3CCC0000"><w:r><w:t>Schedule of rates</w:t></w:r></w:p>' +
  tableXml(endPlacement) +
  '<w:p w14:paraId="3CCC0004"><w:r><w:t>After the table.</w:t></w:r></w:p>' +
  "<w:sectPr/></w:body></w:document>";

const commentXml = (
  id: number,
  paraId: string,
  author: string,
  text: string,
): string =>
  `<w:comment w:id="${id}" w:author="${author}" w:initials="${author[0]}" w:date="2026-01-01T00:00:00Z"><w:p w14:paraId="${paraId}"><w:r><w:annotationRef/></w:r><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p></w:comment>`;

const buildDocx = async (
  endPlacement: "table" | "row",
): Promise<ArrayBuffer> => {
  const zip = new ZipArchive();
  zip.set(
    "[Content_Types].xml",
    `${XML_DECLARATION}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
      '<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>' +
      '<Override PartName="/word/commentsExtended.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml"/>' +
      "</Types>",
  );
  zip.set(
    "_rels/.rels",
    `${XML_DECLARATION}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`,
  );
  zip.set(
    "word/_rels/document.xml.rels",
    `${XML_DECLARATION}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rIdC" Type="${RELATIONSHIP_TYPES.comments}" Target="comments.xml"/>` +
      `<Relationship Id="rIdCx" Type="${RELATIONSHIP_TYPES.commentsExtended}" Target="commentsExtended.xml"/>` +
      "</Relationships>",
  );
  zip.set("word/document.xml", documentXml(endPlacement));
  zip.set(
    "word/comments.xml",
    `${XML_DECLARATION}<w:comments xmlns:w="${W}" xmlns:w14="${W14}">` +
      commentXml(
        ROOT_ID,
        ROOT_PARA_ID,
        "Reviewer",
        "Rate only, or a quantity too?",
      ) +
      commentXml(REPLY_ID, REPLY_PARA_ID, "Engineer", "Rate only is fine.") +
      "</w:comments>",
  );
  zip.set(
    "word/commentsExtended.xml",
    `${XML_DECLARATION}<w15:commentsEx xmlns:w15="${W15}">` +
      `<w15:commentEx w15:paraId="${ROOT_PARA_ID}" w15:done="0"/>` +
      `<w15:commentEx w15:paraId="${REPLY_PARA_ID}" w15:paraIdParent="${ROOT_PARA_ID}" w15:done="0"/>` +
      "</w15:commentsEx>",
  );
  const bytes = await zip.generate({ compression: "store" });
  return bytes.buffer as ArrayBuffer;
};

const MARKER_KEYS: Record<string, "start" | "end" | "reference"> = {
  RangeStart: "start",
  RangeEnd: "end",
  Reference: "reference",
};

const markers = (xml: string): Record<string, number[]> => {
  const out: Record<string, number[]> = { start: [], end: [], reference: [] };
  for (const m of xml.matchAll(
    /<w:comment(RangeStart|RangeEnd|Reference) w:id="(\d+)"/gu,
  )) {
    out[MARKER_KEYS[m[1]!]!]!.push(Number(m[2]));
  }
  for (const list of Object.values(out)) list.sort((a, b) => a - b);
  return out;
};

const savedDocumentXml = async (buffer: ArrayBuffer): Promise<string> =>
  (await (await ZipArchive.read(buffer)).entry("word/document.xml")?.text()) ??
  "";

const edit = async (
  endPlacement: "table" | "row",
  operation: (blockId: (text: string) => string) => Record<string, unknown>,
): Promise<{ xml: string; reviewer: FolioDocxReviewer }> => {
  const reviewer = await FolioDocxReviewer.fromBuffer(
    await buildDocx(endPlacement),
    {
      author: "Editor",
    },
  );
  const blocks = reviewer.snapshot().blocks;
  const blockId = (text: string): string => {
    const hit = blocks.find((block) => block.text === text);
    if (!hit) throw new Error(`no block reads ${JSON.stringify(text)}`);
    return hit.id;
  };
  const result = reviewer.applyDocumentOperations(
    { version: 1, operations: [{ id: "op1", ...operation(blockId) }] as never },
    {
      revisionStamp: { date: "2026-01-02T00:00:00Z", idSeed: 100 },
      undefinedReferences: "refuse",
    },
  );
  expect(result.skipped).toEqual([]);
  const saved = await reviewer.save();
  if (saved.type === "repackRefused") throw new Error(saved.reason);
  return { xml: await savedDocumentXml(saved.buffer), reviewer };
};

const BOTH = {
  start: [REPLY_ID, ROOT_ID],
  end: [REPLY_ID, ROOT_ID],
  reference: [REPLY_ID, ROOT_ID],
};

describe("a comment range that closes between table rows", () => {
  for (const endPlacement of ["table", "row"] as const) {
    test(`rewriting the paragraph that holds its reference keeps every marker (end at ${endPlacement} level)`, async () => {
      const { xml } = await edit(endPlacement, (blockId) => ({
        type: "replaceBlock",
        blockId: blockId("rate only"),
        text: "rate only, no quantity",
      }));
      expect(markers(xml)).toEqual(BOTH);
    });

    test(`deleting a commented cell paragraph keeps the other comment's markers (end at ${endPlacement} level)`, async () => {
      const { xml } = await edit(endPlacement, (blockId) => ({
        type: "deleteBlock",
        blockId: blockId("per visit"),
      }));
      expect(markers(xml)).toEqual(BOTH);
    });

    test(`accepting that deletion keeps both comments, whose ranges it only partly covered (end at ${endPlacement} level)`, async () => {
      const { reviewer } = await edit(endPlacement, (blockId) => ({
        type: "deleteBlock",
        blockId: blockId("per visit"),
      }));
      reviewer.acceptAll();
      const saved = await reviewer.save();
      if (saved.type === "repackRefused") throw new Error(saved.reason);
      expect(markers(await savedDocumentXml(saved.buffer))).toEqual(BOTH);
      expect(reviewer.getComments().map((comment) => comment.id)).toEqual([
        ROOT_ID,
      ]);
      expect(
        reviewer.getComments()[0]?.replies.map((reply) => reply.id),
      ).toEqual([REPLY_ID]);
    });
  }
});
