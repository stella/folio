/**
 * A comment range may open between blocks: before a content control, with
 * `w:displacedByCustomXml="next"`, its start is the control's sibling rather
 * than a child of the paragraph inside it. It is read as the start of the
 * text it opens on. Read as anything else, the comment became a point comment,
 * its range was rebuilt inside the paragraph and the save wrote a second start
 * under the same id.
 */

import { describe, expect, test } from "bun:test";

import { ZipArchive } from "@stll/docx-core/zip";

import { saveCanonicalHarnessDocument } from "../../../../test/canonicalEditorHarness";
import { FolioDocxReviewer } from "../ai-edits/headless";
import { parseDocx } from "./parser";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W14 = "http://schemas.microsoft.com/office/word/2010/wordml";
const HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const COMMENTS_TYPE =
  "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml";
const COMMENTS_RELATIONSHIP =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";

const BESIDE_A_CONTROL =
  '<w:commentRangeStart w:id="1" w:displacedByCustomXml="next"/>' +
  '<w:sdt><w:sdtPr><w:alias w:val="Payment"/><w:id w:val="7"/></w:sdtPr><w:sdtContent>' +
  '<w:p w14:paraId="1A000002"><w:r><w:t xml:space="preserve">Please remit funds. </w:t></w:r><w:commentRangeEnd w:id="1"/><w:r><w:commentReference w:id="1"/></w:r></w:p>' +
  "</w:sdtContent></w:sdt>";
const BETWEEN_PARAGRAPHS =
  '<w:commentRangeStart w:id="1"/>' +
  '<w:p w14:paraId="1A000002"><w:r><w:t xml:space="preserve">Please remit funds. </w:t></w:r><w:commentRangeEnd w:id="1"/><w:r><w:commentReference w:id="1"/></w:r></w:p>';

const commentedPackage = async (
  commented = BESIDE_A_CONTROL,
): Promise<ArrayBuffer> => {
  const zip = new ZipArchive();
  zip.set(
    "[Content_Types].xml",
    `${HEAD}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/comments.xml" ContentType="${COMMENTS_TYPE}"/></Types>`,
  );
  zip.set(
    "_rels/.rels",
    `${HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`,
  );
  zip.set(
    "word/_rels/document.xml.rels",
    `${HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId9" Type="${COMMENTS_RELATIONSHIP}" Target="comments.xml"/></Relationships>`,
  );
  zip.set(
    "word/document.xml",
    `${HEAD}<w:document xmlns:w="${W}" xmlns:w14="${W14}"><w:body>` +
      '<w:p w14:paraId="1A000001"><w:r><w:t>Opening words</w:t></w:r></w:p>' +
      commented +
      '<w:p w14:paraId="1A000003"><w:r><w:t>Closing words</w:t></w:r></w:p>' +
      "<w:sectPr/></w:body></w:document>",
  );
  zip.set(
    "word/comments.xml",
    `${HEAD}<w:comments xmlns:w="${W}"><w:comment w:id="1" w:author="Reviewer" w:date="2024-01-01T00:00:00Z"><w:p><w:r><w:t>Check the account</w:t></w:r></w:p></w:comment></w:comments>`,
  );
  const bytes = await zip.generate({ compression: "store" });
  return bytes.buffer as ArrayBuffer;
};

const markersOf = async (bytes: ArrayBuffer | Uint8Array) => {
  const buffer = bytes instanceof Uint8Array ? bytes.slice().buffer : bytes;
  const xml =
    (await (
      await ZipArchive.read(buffer)
    )
      .entry("word/document.xml")
      ?.text()) ?? "";
  return [
    ...xml.matchAll(/<w:comment(RangeStart|RangeEnd|Reference) w:id="(\d+)"/gu),
  ].map(([, kind, id]) => `${kind}:${id}`);
};

const AS_AUTHORED = ["RangeStart:1", "RangeEnd:1", "Reference:1"];

describe("a comment range that opens between blocks", () => {
  test("is read as the comment's range, not a second reference", async () => {
    const parsed = await parseDocx(await commentedPackage(), {
      preloadFonts: false,
    });
    const sdt = parsed.package.document.content.find(
      (block) => block.type === "blockSdt",
    );
    const paragraph = sdt?.type === "blockSdt" ? sdt.content.at(0) : undefined;
    expect(
      paragraph?.type === "paragraph"
        ? paragraph.content.map((item) => item.type)
        : [],
    ).toEqual([
      "commentRangeStart",
      "run",
      "commentRangeEnd",
      "commentReference",
    ]);
  });

  for (const [where, commented] of [
    ["beside a content control", BESIDE_A_CONTROL],
    ["between paragraphs", BETWEEN_PARAGRAPHS],
  ] as const)
    test(`keeps one start ${where} when an edit elsewhere rewrites the body`, async () => {
      const reviewer = await FolioDocxReviewer.fromBuffer(
        await commentedPackage(commented),
      );
      const blockId = reviewer.snapshot().blocks.at(-1)?.id;
      if (blockId === undefined)
        throw new Error("the fixture has no paragraph");
      const result = reviewer.applyDocumentOperations(
        {
          version: 1,
          mode: "direct",
          operations: [
            { id: "add", type: "insertAfterBlock", blockId, text: "Added." },
          ],
        },
        { undefinedReferences: "refuse" },
      );
      expect(result.applied).toHaveLength(1);
      const saved = await reviewer.toBuffer();
      expect(await markersOf(saved)).toEqual(AS_AUTHORED);
      expect(reviewer.getComments().map(({ text }) => text)).toEqual([
        "Check the account",
      ]);
    });

  // The canonical session does not open content controls yet, so it reads the
  // same boundary between two paragraphs.
  test("keeps one start through the canonical session's save", async () => {
    const parsed = await parseDocx(await commentedPackage(BETWEEN_PARAGRAPHS), {
      preloadFonts: false,
    });
    const { bytes } = await saveCanonicalHarnessDocument(parsed);
    expect(await markersOf(bytes)).toEqual(AS_AUTHORED);
  });
});
