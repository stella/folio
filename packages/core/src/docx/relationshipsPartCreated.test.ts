/**
 * A package need not carry `word/_rels/document.xml.rels` while its main
 * document points at nothing. A save that gives the document something to
 * point at, a comment say, has to create that part: other readers find the
 * comments only through it.
 */

import { describe, expect, test } from "bun:test";
import { readFileSync } from "node:fs";
import { resolve } from "node:path";

import { ZipArchive } from "@stll/docx-core/zip";

import { FolioDocxReviewer } from "../ai-edits/headless";
import { addRelationship } from "./rezip";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const COMMENTS_RELATIONSHIP =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
const COMMENTS_EXTENDED_RELATIONSHIP =
  "http://schemas.microsoft.com/office/2011/relationships/commentsExtended";
const DOCUMENT_RELS = "word/_rels/document.xml.rels";

/** A main document with no relationships part at all. */
const packageWithoutDocumentRelationships = async (): Promise<ArrayBuffer> => {
  const zip = new ZipArchive();
  zip.set(
    "[Content_Types].xml",
    `${HEAD}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`,
  );
  zip.set(
    "_rels/.rels",
    `${HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`,
  );
  zip.set(
    "word/document.xml",
    `${HEAD}<w:document xmlns:w="${W}"><w:body><w:p><w:r><w:t>Hello world</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`,
  );
  const bytes = await zip.generate({ compression: "store" });
  return bytes.buffer as ArrayBuffer;
};

const part = async (buffer: ArrayBuffer, path: string) =>
  (await ZipArchive.read(buffer)).entry(path)?.text();

const commented = async () => {
  const reviewer = await FolioDocxReviewer.fromBuffer(
    await packageWithoutDocumentRelationships(),
  );
  const blockId = reviewer.snapshot().blocks[0]?.id;
  if (blockId === undefined) throw new Error("the fixture has no paragraph");
  const result = reviewer.applyDocumentOperations(
    {
      version: 1,
      mode: "direct",
      operations: [
        {
          id: "comment",
          type: "commentOnBlock",
          blockId,
          quote: "Hello",
          comment: { text: "A note" },
        },
      ],
    },
    { undefinedReferences: "refuse" },
  );
  expect(result.applied).toHaveLength(1);
  return reviewer;
};

describe("a package whose document has no relationships part", () => {
  test("a new comment is linked from the document it annotates", async () => {
    const saved = await (await commented()).toBuffer();
    const rels = await part(saved, DOCUMENT_RELS);
    expect(rels).toContain(`Type="${COMMENTS_RELATIONSHIP}"`);
    expect(rels).toContain('Target="comments.xml"');
    expect(await part(saved, "word/comments.xml")).toContain("A note");
  });

  test("a reply is linked as well as its comment", async () => {
    const reviewer = await commented();
    const [comment] = reviewer.getComments();
    if (!comment) throw new Error("the comment was not created");
    expect(reviewer.replyTo(comment, { text: "A reply" })).not.toBeNull();
    const saved = await reviewer.toBuffer();
    const rels = await part(saved, DOCUMENT_RELS);
    expect(rels).toContain(`Type="${COMMENTS_RELATIONSHIP}"`);
    expect(rels).toContain(`Type="${COMMENTS_EXTENDED_RELATIONSHIP}"`);
    expect(await part(saved, "word/comments.xml")).toContain("A reply");
  });

  test("adding a relationship creates the part instead of failing", async () => {
    const { buffer, rId } = await addRelationship(
      await packageWithoutDocumentRelationships(),
      {
        type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        target: "https://example.org/",
        targetMode: "External",
      },
    );
    const rels = await part(buffer, DOCUMENT_RELS);
    expect(rels).toContain(`Id="${rId}"`);
    expect(rels).toContain('Target="https://example.org/"');
  });
});

test("the save paths read a relationships part to add to only through one reader", () => {
  for (const file of ["rezip.ts", "selectiveSave.ts"]) {
    const source = readFileSync(resolve(import.meta.dir, file), "utf8");
    const direct = [
      ...source.matchAll(
        /\.entry\(\s*(?:relsPath|"word\/_rels\/document\.xml\.rels")\s*\)/gu,
      ),
    ];
    // `readRelsOrStub` is the one place that reads one, in rezip.ts.
    expect(direct.length, file).toBe(file === "rezip.ts" ? 1 : 0);
  }
});
