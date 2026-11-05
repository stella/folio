import { describe, expect, test } from "bun:test";

import { ZipArchive } from "@stll/docx-core/zip";

import { parseDocx } from "../../docx/parser";
import { createEmptyDocx, repackDocx } from "../../docx/rezip";
import { FOLIO_LOCALES, getFolioMessages } from "../../i18n/messages";
import type { ParagraphBlock } from "../../layout-engine/types";
import { fromProseDoc } from "../../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../../prosemirror/conversion/toProseDoc";
import { ALT_CHUNK_DEFAULT_MARKER_TEXT } from "./altChunkPlaceholder";
import { ALT_CHUNK_CONTENT } from "./flowConversionShared";
import { toFlowBlocks } from "./toFlowBlocks";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const HTML_PAYLOAD = "<html><body>Chunk markup</body></html>";
const TEXT_PAYLOAD = "First chunk line\nSecond chunk line";
const CHUNK_HTML = '<w:altChunk r:id="rIdHtml"/>';
const CHUNK_TEXT = '<w:altChunk r:id="rIdText"/>';
const paragraph = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;

const buildDocx = async (body: string): Promise<ArrayBuffer> => {
  const zip = await ZipArchive.read(await createEmptyDocx());
  zip.set(
    "word/document.xml",
    `<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body>${body}<w:sectPr/></w:body></w:document>`,
  );
  zip.set(
    "word/_rels/document.xml.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdHtml" Type="${R}/aFChunk" Target="afchunk.htm"/><Relationship Id="rIdText" Type="${R}/aFChunk" Target="afchunk.txt"/></Relationships>`,
  );
  const contentTypes =
    (await zip.entry("[Content_Types].xml")?.text())?.replace(
      "</Types>",
      '<Default Extension="htm" ContentType="text/html"/><Default Extension="txt" ContentType="text/plain"/></Types>',
    ) ?? "";
  zip.set("[Content_Types].xml", contentTypes);
  zip.set("word/afchunk.htm", HTML_PAYLOAD);
  zip.set("word/afchunk.txt", TEXT_PAYLOAD);
  return zip.generate({ compression: "store" }).then((bytes) => bytes.buffer);
};

const paragraphsOf = async (
  body: string,
  options: Parameters<typeof toFlowBlocks>[1] = {},
): Promise<ParagraphBlock[]> => {
  const parsed = await parseDocx(await buildDocx(body), {
    preloadFonts: false,
  });
  return toFlowBlocks(toProseDoc(parsed), options).filter(
    (block): block is ParagraphBlock => block.kind === "paragraph",
  );
};

const textOf = (block: ParagraphBlock): string =>
  block.runs.map((run) => (run.kind === "text" ? run.text : "")).join("");

describe("the layout of a w:altChunk", () => {
  test("marks a chunk with no text of its own, in the editor's language", async () => {
    const blocks = await paragraphsOf(CHUNK_HTML, {
      altChunkMarkerText: "Obsah se nezobrazuje",
    });
    expect(blocks.map(textOf)).toEqual(["Obsah se nezobrazuje"]);
    expect(blocks[0]?.attrs?.shading).toBeDefined();
  });

  test("falls back to the English catalogue's marker without an editor", async () => {
    const blocks = await paragraphsOf(CHUNK_HTML);
    expect(blocks.map(textOf)).toEqual([ALT_CHUNK_DEFAULT_MARKER_TEXT]);
    expect(getFolioMessages("en").folio["contentNotDisplayed"]).toBe(
      ALT_CHUNK_DEFAULT_MARKER_TEXT,
    );
  });

  test("shows a chunk's own text, one paragraph per line", async () => {
    const blocks = await paragraphsOf(CHUNK_TEXT, {
      altChunkMarkerText: "unused",
    });
    expect(blocks.map(textOf)).toEqual([
      "First chunk line",
      "Second chunk line",
    ]);
  });

  test("resolves every position in it to the boundary after the chunk", async () => {
    const blocks = await paragraphsOf(CHUNK_TEXT);
    const after = blocks[0]?.pmEnd;
    expect(after).toBe((blocks[0]?.pmStart ?? 0) + 1);
    expect(blocks[1]?.pmStart).toBeUndefined();
    for (const block of blocks) {
      expect(block.fixedPosition).toBe(after);
      for (const run of block.runs) {
        expect(run.kind).toBe("text");
        if (run.kind === "text") {
          expect([run.pmStart, run.pmEnd]).toEqual([after, after]);
        }
      }
    }
  });

  test("a marker in a right-to-left language lays out right to left", async () => {
    const [rtl] = await paragraphsOf(CHUNK_HTML, {
      altChunkMarkerText: "המכיל אינו מוצג",
    });
    const [ltr] = await paragraphsOf(CHUNK_HTML);
    expect(rtl?.attrs?.bidi).toBe(true);
    expect(ltr?.attrs?.bidi ?? false).toBe(false);
  });

  test("a story whose height moves the body shows only the marker", async () => {
    const blocks = await paragraphsOf(CHUNK_TEXT, {
      altChunkContent: ALT_CHUNK_CONTENT.markerOnly,
    });
    expect(blocks.map(textOf)).toEqual([ALT_CHUNK_DEFAULT_MARKER_TEXT]);
  });

  test("appears in a table cell too", async () => {
    const parsed = await parseDocx(
      await buildDocx(
        `<w:tbl><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc>${CHUNK_HTML}${paragraph("Cell text")}</w:tc></w:tr></w:tbl>`,
      ),
      { preloadFonts: false },
    );
    const table = toFlowBlocks(toProseDoc(parsed)).find(
      (block) => block.kind === "table",
    );
    if (table?.kind !== "table") {
      throw new TypeError("The body holds a table.");
    }
    const cellText = table.rows[0]?.cells[0]?.blocks.flatMap((block) =>
      block.kind === "paragraph" ? [textOf(block)] : [],
    );
    expect(cellText).toEqual([ALT_CHUNK_DEFAULT_MARKER_TEXT, "Cell text"]);
  });

  test("leaves other preserved blocks out of layout", async () => {
    const blocks = await paragraphsOf(
      `${paragraph("Before")}<w:customXmlDelRangeStart w:id="1"/>${paragraph("After")}`,
    );
    expect(blocks.map(textOf)).toEqual(["Before", "After"]);
  });

  test("every catalogue carries the marker", () => {
    for (const locale of FOLIO_LOCALES) {
      const marker = getFolioMessages(locale).folio["contentNotDisplayed"];
      expect(typeof marker === "string" && marker.length > 0).toBe(true);
    }
  });

  test("saving after laying it out keeps the chunk byte for byte", async () => {
    const body = `${paragraph("Visible")}${CHUNK_HTML}${CHUNK_TEXT}`;
    const input = await buildDocx(body);
    const parsed = await parseDocx(input, { preloadFonts: false });
    const proseDoc = toProseDoc(parsed);
    toFlowBlocks(proseDoc, { altChunkMarkerText: "Marker" });
    const edited = fromProseDoc(proseDoc, parsed, {
      stylesheetSource: { type: "package" },
    });

    const saved = await ZipArchive.read(
      await repackDocx(edited, { updateModifiedDate: false }),
    );
    const documentXml = await saved.entry("word/document.xml")?.text();
    expect(documentXml).toContain(CHUNK_HTML);
    expect(documentXml).toContain(CHUNK_TEXT);
    expect(documentXml).not.toContain("Marker");
    expect(documentXml).not.toContain("First chunk line");
    const encoder = new TextEncoder();
    expect(await saved.entry("word/afchunk.htm")?.bytes()).toEqual(
      encoder.encode(HTML_PAYLOAD),
    );
    expect(await saved.entry("word/afchunk.txt")?.bytes()).toEqual(
      encoder.encode(TEXT_PAYLOAD),
    );
    const rels = await saved.entry("word/_rels/document.xml.rels")?.text();
    expect(rels).toContain('Id="rIdHtml"');
    expect(rels).toContain('Id="rIdText"');
  });
});
