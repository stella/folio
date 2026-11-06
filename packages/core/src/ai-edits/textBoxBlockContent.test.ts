/**
 * A text box holds what a cell can: content controls, custom XML, range markers
 * and markup Folio only carries. A rewrite of the paragraph hosting the box
 * writes it all back, and what has text shows in the box.
 */

import { describe, expect, test } from "bun:test";

import { ZipArchive } from "@stll/docx-core/zip";

import { parseDocx } from "../docx/parser";
import { createDocx } from "../docx/rezip";
import { toFlowBlocks } from "../layout-bridge/convert/toFlowBlocks";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import { createEmptyDocument } from "../utils/createDocument";
import { FolioDocxReviewer } from "./headless";

const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NS = `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:r="${R}" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"`;
const box = (inner: string) =>
  `<w:p w14:paraId="10000001"><w:r><w:t>Host</w:t></w:r><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="4000000" cy="2000000"/><wp:docPr id="1" name="Box"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="4000000" cy="2000000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr><wps:txbx><w:txbxContent>${inner}</w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`;
const held = (id: string, text: string) =>
  `<w:p w14:paraId="${id}"><w:r><w:t>${text}</w:t></w:r></w:p>`;

const CASES = [
  {
    name: "a content control",
    inner: `${held("10000003", "Plain")}<w:sdt><w:sdtPr><w:id w:val="9"/></w:sdtPr><w:sdtContent>${held("10000004", "Held")}</w:sdtContent></w:sdt>`,
    kept: ["<w:sdt>", "Held", "Plain"],
    shown: true,
  },
  {
    name: "custom XML",
    inner: `${held("10000003", "Plain")}<w:customXml w:uri="u" w:element="e">${held("10000004", "Held")}</w:customXml>`,
    kept: ["<w:customXml", "Held", "Plain"],
    shown: true,
  },
  {
    name: "a tracked block wrapper",
    inner: `${held("10000003", "Plain")}<w:ins w:id="5" w:author="a">${held("10000004", "Held")}</w:ins>`,
    kept: ["<w:ins ", "Held", "Plain"],
    shown: false,
  },
  {
    name: "a bookmark",
    inner: `<w:bookmarkStart w:id="4" w:name="bm"/>${held("10000003", "Held")}<w:bookmarkEnd w:id="4"/>`,
    kept: ['w:name="bm"', "<w:bookmarkEnd", "Held"],
    shown: true,
  },
  {
    name: "an altChunk",
    inner: `${held("10000003", "Plain")}<w:altChunk r:id="rIdC"/>`,
    kept: ["<w:altChunk", "Plain"],
    shown: false,
  },
] as const;

const packageWith = async (inner: string): Promise<ArrayBuffer> => {
  const zip = await ZipArchive.read(await createDocx(createEmptyDocument()));
  zip.set(
    "word/document.xml",
    `<?xml version="1.0" encoding="UTF-8"?><w:document ${NS}><w:body>${box(inner)}<w:p w14:paraId="10000002"><w:r><w:t>Next</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`,
  );
  const rels = (await zip.entry("word/_rels/document.xml.rels")?.text()) ?? "";
  zip.set(
    "word/_rels/document.xml.rels",
    rels.replace(
      "</Relationships>",
      `<Relationship Id="rIdC" Type="${R}/aFChunk" Target="c.htm"/></Relationships>`,
    ),
  );
  zip.set("word/c.htm", "<html><body>Chunk text</body></html>");
  const types = (await zip.entry("[Content_Types].xml")?.text()) ?? "";
  zip.set(
    "[Content_Types].xml",
    types.replace(
      "</Types>",
      '<Default Extension="htm" ContentType="text/html"/></Types>',
    ),
  );
  return (await zip.generate()).slice().buffer;
};

const editHost = async (
  bytes: ArrayBuffer,
  mode: "direct" | "tracked-changes",
): Promise<string> => {
  const reviewer = await FolioDocxReviewer.fromBuffer(bytes.slice(0));
  const { applied } = reviewer.applyDocumentOperations(
    {
      version: 1,
      mode,
      operations: [
        {
          id: "edit",
          type: "replaceInBlock",
          blockId: "10000001",
          find: "Host",
          replace: "Hostess",
        },
      ],
    } as never,
    { undefinedReferences: "refuse" },
  );
  expect(applied).toHaveLength(1);
  const saved = await ZipArchive.read(await reviewer.toBuffer());
  return (await saved.entry("word/document.xml")?.text()) ?? "";
};

const shownInBox = async (bytes: ArrayBuffer): Promise<string[]> => {
  const model = await parseDocx(bytes.slice(0), { preloadFonts: false });
  const texts: string[] = [];
  const blocks = toFlowBlocks(toProseDoc(model));
  for (const block of blocks) {
    if (block.kind !== "textBox") continue;
    for (const inner of block.content) {
      if (inner.kind !== "paragraph") continue;
      texts.push(
        inner.runs.map((run) => (run.kind === "text" ? run.text : "")).join(""),
      );
    }
  }
  return texts;
};

describe.each(CASES)("a text box holding $name", ({ inner, kept, shown }) => {
  test.each(["direct", "tracked-changes"] as const)(
    "an edit of its host paragraph keeps it, %s",
    async (mode) => {
      const xml = await editHost(await packageWith(inner), mode);
      for (const piece of kept) expect(xml).toContain(piece);
    },
  );

  test("what it holds in text shows in the box, or stays out of it", async () => {
    const shownText = await shownInBox(await packageWith(inner));
    expect(shownText.includes("Held")).toBe(shown);
  });
});

test("an altChunk in a text box shows its marker in the box", async () => {
  const inner = CASES.find(({ name }) => name === "an altChunk")?.inner ?? "";
  const shownText = await shownInBox(await packageWith(inner));
  expect(shownText.length).toBeGreaterThan(1);
});

test("a saved box reads back as the same blocks", async () => {
  const inner = CASES[0].inner;
  const reviewer = await FolioDocxReviewer.fromBuffer(
    (await packageWith(inner)).slice(0),
  );
  reviewer.applyDocumentOperations(
    {
      version: 1,
      mode: "direct",
      operations: [
        {
          id: "edit",
          type: "replaceInBlock",
          blockId: "10000001",
          find: "Host",
          replace: "Hostess",
        },
      ],
    } as never,
    { undefinedReferences: "refuse" },
  );
  const reopened = await reviewer.toBuffer();
  expect(await shownInBox(reopened)).toEqual(
    await shownInBox(await packageWith(inner)),
  );
});
