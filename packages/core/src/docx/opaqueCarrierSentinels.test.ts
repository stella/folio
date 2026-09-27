import { describe, expect, test } from "bun:test";
import { PARSE_WARNING_CODES } from "@stll/docx-core/model";
import JSZip from "jszip";

import { FolioDocxReviewer } from "../ai-edits/headless";
import { compareDocx } from "../compare/compare";
import { BLOCK_CONTENT_HANDLERS } from "./blockContentParser";
import { CAPTURE } from "./containerChildren";
import { parseDocx } from "./parser";
import { createEmptyDocx, repackDocx } from "./rezip";
import { docxToMarkdown } from "./server/docxToMarkdown";
import { CELL_CONTENT_HANDLERS, ROW_CONTENT_HANDLERS, TABLE_CONTENT_HANDLERS } from "./tableParser";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const SENTINEL = "Opaque carrier sentinel";

const paragraph = (text: string): string => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
const cell = (content: string): string => `<w:tc><w:tcPr/>${content}</w:tc>`;
const row = (content: string): string => `<w:tr>${content}</w:tr>`;
const table = (content: string): string =>
  `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>${content}</w:tbl>`;
const sdt = (content: string): string =>
  `<w:sdt><w:sdtPr/><w:sdtContent>${content}</w:sdtContent></w:sdt>`;
const customXml = (content: string): string =>
  `<w:customXml w:element="clause">${content}</w:customXml>`;
const revision = (kind: "ins" | "del" | "moveFrom" | "moveTo", content: string): string =>
  `<w:${kind} w:id="1" w:author="Author">${content}</w:${kind}>`;

const CASES = {
  "block/sdt": sdt(paragraph(SENTINEL)),
  "block/customXml": customXml(paragraph(SENTINEL)),
  "block/ins": revision("ins", paragraph(SENTINEL)),
  "block/del": revision("del", paragraph(SENTINEL)),
  "block/moveFrom": revision("moveFrom", paragraph(SENTINEL)),
  "block/moveTo": revision("moveTo", paragraph(SENTINEL)),
  "block/altChunk": '<w:altChunk r:id="rIdChunk"/>',
  "table/sdt": table(sdt(row(cell(paragraph(SENTINEL))))),
  "table/customXml": table(customXml(row(cell(paragraph(SENTINEL))))),
  "table/ins": table(revision("ins", row(cell(paragraph(SENTINEL))))),
  "table/del": table(revision("del", row(cell(paragraph(SENTINEL))))),
  "table/moveFrom": table(revision("moveFrom", row(cell(paragraph(SENTINEL))))),
  "table/moveTo": table(revision("moveTo", row(cell(paragraph(SENTINEL))))),
  "row/sdt": table(row(sdt(cell(paragraph(SENTINEL))))),
  "row/customXml": table(row(cell(paragraph("anchor")) + customXml(cell(paragraph(SENTINEL))))),
  "row/ins": table(row(cell(paragraph("anchor")) + revision("ins", cell(paragraph(SENTINEL))))),
  "row/del": table(row(cell(paragraph("anchor")) + revision("del", cell(paragraph(SENTINEL))))),
  "row/moveFrom": table(
    row(cell(paragraph("anchor")) + revision("moveFrom", cell(paragraph(SENTINEL)))),
  ),
  "row/moveTo": table(
    row(cell(paragraph("anchor")) + revision("moveTo", cell(paragraph(SENTINEL)))),
  ),
  "row/tr": table(row(cell(paragraph("anchor")) + row(cell(paragraph(SENTINEL))))),
  "cell/sdt": table(row(cell(sdt(paragraph(SENTINEL))))),
  "cell/customXml": table(row(cell(customXml(paragraph(SENTINEL))))),
  "cell/ins": table(row(cell(revision("ins", paragraph(SENTINEL))))),
  "cell/del": table(row(cell(revision("del", paragraph(SENTINEL))))),
  "cell/moveFrom": table(row(cell(revision("moveFrom", paragraph(SENTINEL))))),
  "cell/moveTo": table(row(cell(revision("moveTo", paragraph(SENTINEL))))),
  "cell/altChunk": table(row(cell('<w:altChunk r:id="rIdChunk"/>'))),
} as const;

const NON_CONTENT_KIND =
  /^(?:p|tbl|tc|bookmark(?:Start|End)|commentRange(?:Start|End)|customXml(?:Del|Ins|MoveFrom|MoveTo)Range(?:Start|End)|move(?:From|To)Range(?:Start|End)|perm(?:Start|End)|proofErr|customXmlPr|sectPr|tblPr|tblGrid|trPr|tblPrEx|tcPr)$/u;

const parserContentKinds = (): { all: string[]; captured: string[] } => {
  const maps = {
    block: BLOCK_CONTENT_HANDLERS,
    table: TABLE_CONTENT_HANDLERS,
    row: ROW_CONTENT_HANDLERS,
    cell: CELL_CONTENT_HANDLERS,
  };
  const all: string[] = [];
  const captured: string[] = [];
  for (const [container, handlers] of Object.entries(maps)) {
    for (const [kind, disposition] of Object.entries(handlers)) {
      if (NON_CONTENT_KIND.test(kind) || (container === "table" && kind === "tr")) {
        continue;
      }
      all.push(`${container}/${kind}`);
      if (disposition === CAPTURE) {
        captured.push(`${container}/${kind}`);
      }
    }
  }
  return { all: all.toSorted(), captured: captured.toSorted() };
};

const makeDocx = async (body: string, chunkText = SENTINEL): Promise<ArrayBuffer> => {
  const zip = await JSZip.loadAsync(await createEmptyDocx());
  zip.file(
    "word/document.xml",
    `<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body>${body}<w:sectPr/></w:body></w:document>`,
  );
  if (body.includes("<w:altChunk")) {
    zip.file("word/chunk.txt", chunkText);
    const rels = await zip.file("word/_rels/document.xml.rels")?.async("text");
    const contentTypes = await zip.file("[Content_Types].xml")?.async("text");
    if (!rels || !contentTypes) {
      throw new Error("The empty DOCX must carry relationship and content type parts");
    }
    zip.file(
      "word/_rels/document.xml.rels",
      rels.replace(
        "</Relationships>",
        `<Relationship Id="rIdChunk" Type="${R}/aFChunk" Target="chunk.txt"/></Relationships>`,
      ),
    );
    zip.file(
      "[Content_Types].xml",
      contentTypes.replace(
        "</Types>",
        '<Default Extension="txt" ContentType="text/html"/>' +
          '<Override PartName="/word/chunk.txt" ContentType="text/plain"/></Types>',
      ),
    );
  }
  return zip.generateAsync({ type: "arraybuffer" });
};

const occurrences = (text: string): number => text.split(SENTINEL).length - 1;

const warningCodeFor = (kind: string): string => {
  if (kind.endsWith("/altChunk")) return PARSE_WARNING_CODES.altChunkUnsupported;
  if (kind === "row/tr") return PARSE_WARNING_CODES.nestedRowOpaque;
  return PARSE_WARNING_CODES.revisionCarrierOpaque;
};

const opaqueFragment = (kind: string, xml: string): string | undefined => {
  if (kind.endsWith("/altChunk")) {
    return '<w:altChunk r:id="rIdChunk"/>';
  }
  const wrapper = kind.split("/").at(1);
  if (wrapper !== "ins" && wrapper !== "del" && wrapper !== "moveFrom" && wrapper !== "moveTo") {
    return undefined;
  }
  const start = xml.indexOf(`<w:${wrapper} `);
  const end = xml.indexOf(`</w:${wrapper}>`, start);
  return start === -1 || end === -1 ? undefined : xml.slice(start, end + wrapper.length + 5);
};

describe("content carrier sentinels", () => {
  test("every parser content kind, including explicit captures, has a reader case", () => {
    const cases = new Set(Object.keys(CASES));
    const { all, captured } = parserContentKinds();
    expect(all.filter((kind) => !cases.has(kind))).toEqual([]);
    expect(captured.filter((kind) => !cases.has(kind))).toEqual([]);
  });

  for (const [kind, xml] of Object.entries(CASES)) {
    test(`${kind}: readers expose the sentinel or document an opaque boundary`, async () => {
      const bytes = await makeDocx(xml);
      const parsed = await parseDocx(bytes, { preloadFonts: false });
      const warnings = parsed.parseWarnings ?? [];
      const reviewer = await FolioDocxReviewer.fromBuffer(bytes);
      const reads = {
        getContent: reviewer
          .getContent()
          .map(({ text }) => text)
          .join("\n"),
        snapshot: reviewer
          .snapshot()
          .blocks.map(({ text }) => text)
          .join("\n"),
        getContentAsText: reviewer.getContentAsText(),
        docxToMarkdown: await docxToMarkdown(bytes),
      };
      for (const [reader, read] of Object.entries(reads)) {
        const sentinelCount = occurrences(read);
        const hasDiagnostic = read.includes("[Unsupported ");
        expect(sentinelCount).toBeLessThanOrEqual(1);
        if (sentinelCount !== 1 && !hasDiagnostic) {
          throw new Error(`${reader} silently omitted ${kind}`);
        }
        if (hasDiagnostic) {
          expect(warnings.some(({ code }) => code === warningCodeFor(kind))).toBe(true);
        }
      }

      const diagnosticBlock = reviewer
        .snapshot()
        .blocks.find(({ kind: blockKind }) => blockKind === "diagnostic");
      if (Object.values(reads).some((read) => read.includes("[Unsupported "))) {
        expect(diagnosticBlock).toBeDefined();
      }
      if (diagnosticBlock) {
        const edit = reviewer.applyOperations(
          [
            {
              id: "replace-opaque-content",
              type: "replaceInBlock",
              blockId: diagnosticBlock.id,
              find: diagnosticBlock.text,
              replace: "Changed opaque content",
            },
          ],
          { mode: "direct" },
        );
        expect(edit.skipped).not.toEqual([]);
      }

      const changed = await makeDocx(
        xml.replace(SENTINEL, "Mutated carrier sentinel"),
        "Mutated carrier sentinel",
      );
      const compared = await compareDocx(bytes, changed, {
        author: "Comparison",
        timestamp: "2026-09-27T00:00:00.000Z",
      });
      expect(compared.isOk()).toBe(true);
      if (compared.isErr()) {
        return;
      }
      expect(compared.value.changes.length > 0 || compared.value.unsupported.length > 0).toBe(true);

      const fragment = opaqueFragment(kind, xml);
      if (fragment !== undefined) {
        const saved = await repackDocx(parsed, { updateModifiedDate: false });
        const savedXml = await (
          await JSZip.loadAsync(saved)
        )
          .file("word/document.xml")
          ?.async("text");
        expect(savedXml).toContain(fragment);
      }
    });
  }
});
