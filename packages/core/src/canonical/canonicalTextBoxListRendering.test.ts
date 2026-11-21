import { panic } from "better-result";
import { afterAll, beforeAll, expect, setSystemTime, test } from "bun:test";

import { OP_STORIES, paragraphsThroughTextBoxes } from "@stll/docx-core/ops";

import {
  buildRawPackage,
  SECTION,
  WORD_STYLES,
} from "../__tests__/documentShapes";
import { FolioDocxReviewer } from "../ai-edits/headless";
import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperationMode,
} from "../document-operations";
import type { Document } from "../types/document";

beforeAll(() => setSystemTime(new Date("2026-01-01T12:00:00Z")));
afterAll(() => setSystemTime());

const NAMESPACES = [
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"',
  'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"',
].join(" ");
const NUMBERED =
  '<w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>';
const box = (paragraphs: string) =>
  '<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="1828800" cy="457200"/><wp:docPr id="2" name="Box"/>' +
  '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr txBox="1"/>' +
  '<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="457200"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr>' +
  `<wps:txbx><w:txbxContent>${paragraphs}</w:txbxContent></wps:txbx>` +
  "<wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>";

/** A body paragraph drawing a text box whose paragraphs are numbered. */
const boxedList = async () =>
  Uint8Array.from(
    await buildRawPackage({
      styles: WORD_STYLES,
      body:
        `<w:p><w:r><w:t>Body</w:t></w:r>${box(
          `<w:p>${NUMBERED}<w:r><w:t>First item</w:t></w:r></w:p>` +
            `<w:p><w:r><w:t>Plain item</w:t></w:r></w:p>`,
        )}</w:p>` + SECTION(),
      parts: [
        {
          path: "word/numbering.xml",
          contentType:
            "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml",
          relationshipType: "numbering",
          relationshipId: "rIdNumbering",
          body:
            `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering ${NAMESPACES}>` +
            '<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/>' +
            '<w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum>' +
            '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>',
        },
      ],
    }),
  ).buffer;

/** What a reader sees of each text-box paragraph's numbering. */
const boxNumbering = (document: Document) =>
  paragraphsThroughTextBoxes(document, OP_STORIES.MAIN)
    .map(({ paragraph }) => paragraph)
    .filter(({ content }) => {
      const text = JSON.stringify(content);
      return text.includes("item") && !text.includes("Body");
    })
    .map(({ listRendering, formatting }) => ({
      listRendering,
      indentLeft: formatting?.indentLeft,
      indentFirstLine: formatting?.indentFirstLine,
      hangingIndent: formatting?.hangingIndent,
    }));

for (const mode of [
  "direct",
  "tracked-changes",
] as const satisfies readonly FolioDocumentOperationMode[])
  test(`canonical ${mode}: a text-box paragraph an edit numbers renders its list as a reopen reads it`, async () => {
    const live = await FolioDocxReviewer.fromBuffer(await boxedList(), {
      session: "canonical",
    });
    const target =
      live.getContent().find(({ text }) => text.includes("Plain item")) ??
      panic("The boxed paragraph is addressable.");
    const result = live.applyDocumentOperationsToStory({
      story: { type: "main" },
      batch: {
        version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
        mode,
        operations: [
          {
            id: "box-numbering",
            type: "setBlockParagraphProperties",
            blockId: target.id,
            properties: { numbering: { kind: "newList", format: "numbered" } },
          },
        ],
      },
      undefinedReferences: "refuse",
    });
    expect(result.skipped).toEqual([]);
    const reopened = await FolioDocxReviewer.fromBuffer(await live.toBuffer(), {
      session: "canonical",
    });
    const accepted = boxNumbering(live.toResolvedDocument("accept"));
    expect(accepted).toHaveLength(2);
    expect(accepted.map(({ listRendering }) => listRendering?.numFmt)).toEqual([
      "decimal",
      "decimal",
    ]);
    for (const decision of ["accept", "reject"] as const)
      expect(boxNumbering(live.toResolvedDocument(decision))).toEqual(
        boxNumbering(reopened.toResolvedDocument(decision)),
      );
  });
