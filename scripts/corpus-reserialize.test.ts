/**
 * The invariant that makes the real serializers run, over packages built here.
 *
 * Its whole value rests on one thing: removing the captures has to change what
 * folio writes. A fixture whose `w:pPr` already spells its attributes the way
 * the serializer would proves nothing, because replay and reconstruction then
 * agree byte for byte. The fixture below therefore authors `w:spacing` and
 * `w:ind` in an order no folio save produces, which is what a foreign producer
 * does and what the corpus is full of.
 */

import { describe, expect, test } from "bun:test";
import { buildBodySequenceDocx } from "@stll/folio-core/compare/__fixtures__/body-sequence";
import { parseDocx } from "@stll/folio-core/docx/parser";
import { createEmptyDocx, repackDocx } from "@stll/folio-core/docx/rezip";
import { unzipDocx } from "@stll/folio-core/docx/unzip";
import type { Document, DrawingContent } from "@stll/folio-core/types/document";
import JSZip from "jszip";

import {
  type CorpusInvariantInput,
  DEFAULT_INVARIANT_BUDGET_MS,
} from "./lib/corpus-invariants/contract";
import {
  runReserializeInvariant,
  withoutSerializerCaptures,
} from "./lib/corpus-invariants/reserialize";

const DOCUMENT_PART = "word/document.xml";
const DOCUMENT_RELS_PART = "word/_rels/document.xml.rels";

/** A 1x1 PNG, so a picture fixture has media its relationship can reach. */
const PNG_1X1_BASE64 =
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";

/** The id `buildBodySequenceDocx` writes its default header under. */
const HEADER_RELATIONSHIP_ID = "rId2";

/** `EDITED_PREVIEW_FINGERPRINT` in `packages/core/src/docx/imageRawXml.ts`. */
const EDITED_PREVIEW_FINGERPRINT = "editedPreview";

/**
 * `DRAWING_RAW_XML_MODES.PREVIEW_ONLY`. `@stll/docx-core` is not a dependency
 * of the workspace root, so the scripts see the model as types alone.
 */
const PREVIEW_ONLY_MODE = "previewOnly";

/** `REPLAY_PREFIX` in the module under test. */
const REPLAY_PREFIX = "replay hides a serializer difference at";

const WORDPROCESSING = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const OFFICE_RELATIONSHIP = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const WORDML_2010 = "http://schemas.microsoft.com/office/word/2010/wordml";
const DRAWING_WORDPROCESSING =
  "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const DRAWING_MAIN = "http://schemas.openxmlformats.org/drawingml/2006/main";
const WORD_GROUP = "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup";
const WORD_SHAPE = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";
const DRAWING_PICTURE = "http://schemas.openxmlformats.org/drawingml/2006/picture";

/**
 * A package whose only authored part is the body: `createEmptyDocx` supplies
 * the relationships, content types and a `Normal` style, so a fixture that
 * needs markup the body-sequence builder cannot author still needs no second
 * package builder.
 */
const packageWithBody = async (
  body: string,
  imageRelationshipId?: string,
): Promise<ArrayBuffer> => {
  const zip = await JSZip.loadAsync(await createEmptyDocx());
  if (imageRelationshipId !== undefined) {
    const rels = await zip.file(DOCUMENT_RELS_PART)?.async("string");
    if (rels === undefined) {
      throw new Error("the empty package has no main-part relationships");
    }
    zip.file(
      DOCUMENT_RELS_PART,
      rels.replace(
        "</Relationships>",
        `<Relationship Id="${imageRelationshipId}" Type="${OFFICE_RELATIONSHIP}/image" Target="media/art.png"/></Relationships>`,
      ),
    );
    zip.file("word/media/art.png", PNG_1X1_BASE64, { base64: true });
  }
  zip.file(
    DOCUMENT_PART,
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<w:document xmlns:w="${WORDPROCESSING}" xmlns:r="${OFFICE_RELATIONSHIP}" ` +
      `xmlns:w14="${WORDML_2010}" xmlns:wp="${DRAWING_WORDPROCESSING}" ` +
      `xmlns:a="${DRAWING_MAIN}" xmlns:wpg="${WORD_GROUP}" xmlns:wps="${WORD_SHAPE}" ` +
      `xmlns:pic="${DRAWING_PICTURE}">` +
      `<w:body>${body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body>` +
      `</w:document>`,
  );
  return zip.generateAsync({ type: "arraybuffer" });
};

/** Attribute order as the producer wrote it: replay reproduces this spelling. */
const AUTHORED_SPACING = '<w:spacing w:line="360" w:lineRule="auto" w:before="240" w:after="120"/>';
const AUTHORED_INDENT = '<w:ind w:firstLine="360" w:left="720"/>';

/** Attribute order the serializer produces from the model, once replay is gone. */
const SERIALIZED_SPACING =
  '<w:spacing w:before="240" w:after="120" w:line="360" w:lineRule="auto"/>';
const SERIALIZED_INDENT = '<w:ind w:left="720" w:firstLine="360"/>';

const DIRECTLY_FORMATTED_BODY =
  `<w:p w14:paraId="60000001"><w:pPr><w:pStyle w:val="Normal"/>` +
  `${AUTHORED_SPACING}${AUTHORED_INDENT}<w:jc w:val="both"/></w:pPr>` +
  `<w:r><w:t xml:space="preserve">First paragraph.</w:t></w:r></w:p>`;

/**
 * A bare `wpg:wgp` group: folio rasterizes it to a preview, so the drawing
 * carries all three drawing slots at once (`rawXml`, `rawXmlMode` and the
 * fingerprint the policy poisons rather than strips).
 */
const GROUPED_DRAWING_BODY =
  `<w:p><w:r><w:drawing><wp:anchor behindDoc="1">` +
  `<wp:extent cx="1000000" cy="500000"/><wp:wrapTopAndBottom/>` +
  `<a:graphic><a:graphicData uri="${WORD_GROUP}"><wpg:wgp><wps:wsp><wps:spPr>` +
  `<a:xfrm><a:off x="0" y="0"/><a:ext cx="1000000" cy="500000"/></a:xfrm>` +
  `<a:prstGeom prst="rect"/><a:solidFill><a:srgbClr val="DBEDF3"/></a:solidFill>` +
  `</wps:spPr></wps:wsp></wpg:wgp></a:graphicData></a:graphic>` +
  `</wp:anchor></w:drawing></w:r></w:p>`;

const PICTURE_RELATIONSHIP_ID = "rIdArt";

/**
 * An anchored picture, which carries `rawXml` and the fingerprint the policy
 * poisons but no `rawXmlMode`: unlike a group, the model is a projection of
 * it, so the forced path rebuilds the anchor rather than replaying it.
 */
const PICTURE_DRAWING_BODY =
  `<w:p><w:r><w:drawing><wp:anchor behindDoc="1">` +
  `<wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>` +
  `<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>` +
  `<wp:extent cx="1000000" cy="500000"/><wp:wrapTopAndBottom/>` +
  `<wp:docPr id="3" name="Art 3"/>` +
  `<a:graphic><a:graphicData uri="${DRAWING_PICTURE}"><pic:pic>` +
  `<pic:nvPicPr><pic:cNvPr id="3" name="art.png"/><pic:cNvPicPr/></pic:nvPicPr>` +
  `<pic:blipFill><a:blip r:embed="${PICTURE_RELATIONSHIP_ID}"/>` +
  `<a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
  `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000000" cy="500000"/></a:xfrm>` +
  `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>` +
  `</pic:pic></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r></w:p>`;

const documentXmlOf = async (buffer: ArrayBuffer): Promise<string> => {
  const { allXml } = await unzipDocx(buffer, { extractAllXml: true });
  const xml = allXml.get(DOCUMENT_PART);
  if (xml === undefined) {
    throw new Error("the saved package has no main part");
  }
  return xml;
};

const inputFor = async (buffer: ArrayBuffer): Promise<CorpusInvariantInput> => ({
  bytes: new Uint8Array(buffer),
  buffer,
  parsed: await parseDocx(buffer, { preloadFonts: false }),
  documentPart: DOCUMENT_PART,
  budgetMs: DEFAULT_INVARIANT_BUDGET_MS,
});

/** Every capture slot the table-and-header fixture populates, or `undefined` once stripped. */
type CaptureCensus = {
  table: string | undefined;
  grid: string | undefined;
  row: string | undefined;
  cell: string | undefined;
  header: string | undefined;
};

const captureCensus = (document: Document): CaptureCensus => {
  const table = document.package.document.content.find((block) => block.type === "table");
  if (table === undefined) {
    throw new Error("the fixture has no table");
  }
  const row = table.rows.at(0);
  const cell = row?.cells.at(0);
  return {
    table: table.formatting?.sourceXml,
    grid: table.formatting?.gridSourceXml,
    row: row?.formatting?.sourceXml,
    cell: cell?.formatting?.sourceXml,
    header: document.package.headers?.get(HEADER_RELATIONSHIP_ID)?.verbatimXml,
  };
};

/**
 * The fixture's drawing, on the preview branch that declares all three drawing
 * slots. The narrowing doubles as the assertion that `rawXmlMode` survived:
 * a clone that lost it would not reach the branch.
 */
type PreviewDrawing = Extract<DrawingContent, { rawXmlMode: typeof PREVIEW_ONLY_MODE }>;

const previewDrawing = (document: Document): PreviewDrawing => {
  const block = document.package.document.content.at(0);
  if (block?.type !== "paragraph") {
    throw new Error("the fixture has no leading paragraph");
  }
  const drawing = block.content
    .filter((content) => content.type === "run")
    .flatMap((run) => run.content)
    .find((content) => content.type === "drawing");
  if (drawing?.rawXmlMode !== PREVIEW_ONLY_MODE) {
    throw new Error("the grouped fixture did not parse as a rasterized preview");
  }
  return drawing;
};

describe("removing the captures makes the serializers run", () => {
  test("the forced path writes a main part the replay path did not", async () => {
    const buffer = await packageWithBody(DIRECTLY_FORMATTED_BODY);
    const parsed = await parseDocx(buffer, { preloadFonts: false });

    const replayed = await documentXmlOf(await repackDocx(parsed, { updateModifiedDate: false }));
    const forced = await documentXmlOf(
      await repackDocx(withoutSerializerCaptures(parsed), { updateModifiedDate: false }),
    );

    expect(forced).not.toBe(replayed);
    expect(replayed).toContain(AUTHORED_SPACING);
    expect(replayed).toContain(AUTHORED_INDENT);
    expect(forced).toContain(SERIALIZED_SPACING);
    expect(forced).toContain(SERIALIZED_INDENT);
  });

  test("the argument keeps its captures and only the clone loses them", async () => {
    const buffer = await buildBodySequenceDocx(
      [
        { kind: "paragraph", text: "Intro." },
        {
          kind: "table",
          rows: [{ cells: ["Cell A", "Cell B"], height: 400 }],
          properties: { borderSize: 4 },
        },
      ],
      { header: [{ kind: "paragraph", text: "Header paragraph." }] },
    );
    const parsed = await parseDocx(buffer, { preloadFonts: false });

    const stripped = captureCensus(withoutSerializerCaptures(parsed));
    const original = captureCensus(parsed);

    expect(original.table).toStartWith("<w:tblPr>");
    expect(original.grid).toStartWith("<w:tblGrid>");
    expect(original.row).toStartWith("<w:trPr>");
    expect(original.cell).toStartWith("<w:tcPr>");
    expect(original.header).toContain("<w:hdr");
    expect(stripped).toEqual({
      table: undefined,
      grid: undefined,
      row: undefined,
      cell: undefined,
      header: undefined,
    });
  });

  test("a drawing's fingerprint is poisoned while its raw XML and mode are kept", async () => {
    const parsed = await parseDocx(await packageWithBody(GROUPED_DRAWING_BODY), {
      preloadFonts: false,
    });
    const drawing = previewDrawing(parsed);

    const cloned = previewDrawing(withoutSerializerCaptures(parsed));

    // Poisoning rather than clearing is what marks the drawing as one an edit
    // has reached, which is what `classifyDrawingSafety` reads.
    expect(drawing.rawImageFingerprint).not.toBe(EDITED_PREVIEW_FINGERPRINT);
    expect(cloned.rawImageFingerprint).toBe(EDITED_PREVIEW_FINGERPRINT);
    expect(cloned.rawXml).toBe(drawing.rawXml);
  });

  test("a package the serializers rebuild faithfully reports nothing", async () => {
    const outcome = await runReserializeInvariant(
      await inputFor(
        await buildBodySequenceDocx([
          { kind: "paragraph", text: "First paragraph." },
          { kind: "paragraph", text: "Second paragraph.", styleId: "Heading1" },
        ]),
      ),
    );

    expect(outcome.failures).toEqual([]);
    expect(Object.keys(outcome.timings)).toEqual([
      "strip-captures",
      "forced-save",
      "forced-parse",
      "compare",
    ]);
  });

  /**
   * The rebuilt anchor carries `image.allowOverlap`, which the one the fixture
   * authored does not, so the difference is one replay was hiding rather than
   * one a plain repack shows too. The fixture is a picture rather than a
   * group: a group's capture is the content, so poisoning its fingerprint
   * reports the edit rather than licensing a rebuild, and nothing differs.
   */
  test("a difference only the forced path shows is reported as a hidden one", async () => {
    const outcome = await runReserializeInvariant(
      await inputFor(await packageWithBody(PICTURE_DRAWING_BODY, PICTURE_RELATIONSHIP_ID)),
    );

    expect(outcome.failures.length).toBeGreaterThan(0);
    for (const failure of outcome.failures) {
      expect(failure.message).toStartWith(REPLAY_PREFIX);
    }
    expect(Object.keys(outcome.timings)).toContain("control-save");
    // Every difference, not the first: the signatures are distinct, which is
    // what stops a fix to one of them from revealing the next as a new defect.
    expect(new Set(outcome.failures.map((failure) => failure.message)).size).toBe(
      outcome.failures.length,
    );
  });

  test("a poisoned group preview replays, so the forced path reports nothing", async () => {
    const outcome = await runReserializeInvariant(
      await inputFor(await packageWithBody(GROUPED_DRAWING_BODY)),
    );

    expect(outcome.failures).toEqual([]);
  });
});
