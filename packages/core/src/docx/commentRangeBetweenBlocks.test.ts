/**
 * A comment range may open between blocks: before a content control, with
 * `w:displacedByCustomXml="next"`, its start is the control's sibling rather
 * than a child of the paragraph inside it. It is a block of its own, kept
 * between the same two blocks, and every comment reader counts it. Kept as
 * captured bytes instead, the comment became a point comment, its range was
 * rebuilt inside the paragraph and the save wrote a second start under the
 * same id.
 */

import { describe, expect, test } from "bun:test";
import { TextSelection } from "prosemirror-state";

import { ZipArchive } from "@stll/docx-core/zip";

import {
  createCanonicalEditorHarness,
  saveCanonicalHarnessDocument,
} from "../../../../test/canonicalEditorHarness";
import { FolioDocxReviewer } from "../ai-edits/headless";
import { removeCommentMarkers } from "./commentAnchorIndex";
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

/** Whether the saved start still stands between blocks rather than in a paragraph. */
const startsBetweenBlocks = async (bytes: ArrayBuffer | Uint8Array) => {
  const buffer = bytes instanceof Uint8Array ? bytes.slice().buffer : bytes;
  const xml =
    (await (
      await ZipArchive.read(buffer)
    )
      .entry("word/document.xml")
      ?.text()) ?? "";
  return /<\/w:p><w:commentRangeStart w:id="1"[^>]*\/><w:(?:sdt|p)[ >]/u.test(
    xml,
  );
};

describe("a comment range that opens between blocks", () => {
  test("is a block of its own, and the inline end stays an end", async () => {
    const parsed = await parseDocx(await commentedPackage(), {
      preloadFonts: false,
    });
    const blocks = parsed.package.document.content;
    expect(blocks.map(({ type }) => type)).toEqual([
      "paragraph",
      "commentRangeStart",
      "blockSdt",
      "paragraph",
    ]);
    const sdt = blocks[2];
    const paragraph = sdt?.type === "blockSdt" ? sdt.content.at(0) : undefined;
    expect(
      paragraph?.type === "paragraph"
        ? paragraph.content.map((item) => item.type)
        : [],
    ).toEqual(["run", "commentRangeEnd", "commentReference"]);
    expect(blocks[1]).toEqual({
      type: "commentRangeStart",
      id: 1,
      displacedByCustomXml: "next",
    });
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
      expect(await startsBetweenBlocks(saved)).toBe(true);
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
    expect(await startsBetweenBlocks(bytes)).toBe(true);
  });

  test("keeps one start after an edit through the canonical session", async () => {
    const parsed = await parseDocx(await commentedPackage(BETWEEN_PARAGRAPHS), {
      preloadFonts: false,
    });
    const driver = createCanonicalEditorHarness(parsed, "editing");
    try {
      // The caret at the end of the closing paragraph, then real typing.
      const end = driver.state.doc.content.size - 1;
      driver.dispatch(
        driver.state.tr.setSelection(
          TextSelection.near(driver.state.doc.resolve(end), -1),
        ),
      );
      driver.typeText("!");
      const { bytes } = await saveCanonicalHarnessDocument(driver.snapshot());
      expect(await markersOf(bytes)).toEqual(AS_AUTHORED);
      expect(await startsBetweenBlocks(bytes)).toBe(true);
      const xml =
        (await (
          await ZipArchive.read(bytes.slice().buffer)
        )
          .entry("word/document.xml")
          ?.text()) ?? "";
      expect(xml).toContain("Closing words!");
    } finally {
      driver.dispose();
    }
  });

  test("a retired comment takes its markers between blocks with it", async () => {
    for (const commented of [BESIDE_A_CONTROL, BETWEEN_PARAGRAPHS]) {
      const parsed = await parseDocx(await commentedPackage(commented), {
        preloadFonts: false,
      });
      const kept = removeCommentMarkers(
        parsed.package.document.content,
        new Set([1]),
      );
      const kinds: string[] = [];
      const walk = (blocks: readonly (typeof kept)[number][]) => {
        for (const block of blocks) {
          kinds.push(block.type);
          if (block.type === "paragraph")
            kinds.push(...block.content.map(({ type }) => type));
          if (block.type === "blockSdt") walk(block.content);
        }
      };
      walk(kept);
      expect(kinds.filter((kind) => kind.startsWith("comment"))).toEqual([]);
    }
  });
});
