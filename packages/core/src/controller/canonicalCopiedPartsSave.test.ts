/**
 * Kept markup naming package parts (a chart with its embedded workbook and
 * styles, an embedded object, a kept alternative-format chunk) pastes into
 * another package with copies of those parts: the save writes them under
 * fresh names with their content types, every relationship resolves after a
 * reopen, and an undo leaves none of them in the save.
 */
import { describe, expect, test } from "bun:test";
import type { Node as ProseNode } from "prosemirror-model";
import { EditorState, TextSelection } from "prosemirror-state";

import { DRAWING_RAW_XML_MODES } from "@stll/docx-core/model";
import { ZipArchive } from "@stll/docx-core/zip";

import {
  type CanonicalSession,
  createCanonicalSession,
  publishCanonicalProjection,
} from "../canonical/canonicalSession";
import { serializeCanonicalSave } from "../docx/canonicalSave";
import { parseDocx } from "../docx/parser";
import { parseRelationships, resolveRelativePath } from "../docx/relsParser";
import { createDocx } from "../docx/rezip";
import { createSimpleDocument } from "../docx/serializer/documentSerializer";
import {
  getAttribute,
  getChildElements,
  getLocalName,
  parseXmlDocument,
} from "../docx/xmlParser";
import type {
  BlockContent,
  Document,
  DrawingContent,
  Paragraph,
} from "../types/document";
import { readModel } from "../types/documentModel";
import { FOLIO_DOCX_SERIALIZATION_MODE } from "../types/docxSerialization";
import { decodeText, encodeUtf8 } from "../utils/textEncoding";
import { prepareCanonicalPaste } from "./canonicalClipboard";
import {
  readClipboardPackageParts,
  sliceKeptMarkup,
} from "./canonicalClipboardParts";

const NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const CHART_XML =
  `<w:drawing ${NS} ` +
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">' +
  '<wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="914400" cy="457200"/>' +
  '<wp:docPr id="77" name="Copied chart"/><a:graphic>' +
  '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">' +
  '<c:chart r:id="rIdChart"/></a:graphicData></a:graphic></wp:inline></w:drawing>';
const OBJECT_XML =
  `<w:object ${NS} xmlns:v="urn:schemas-microsoft-com:vml" ` +
  'xmlns:o="urn:schemas-microsoft-com:office:office" w:dxaOrig="1440" w:dyaOrig="720">' +
  '<v:rect style="width:72pt;height:36pt" fillcolor="#336699" stroked="f"/>' +
  '<o:OLEObject Type="Embed" ProgID="Package" ShapeID="_x0000_i1025" ' +
  'DrawAspect="Content" ObjectID="_1700000001" r:id="rIdObject"/></w:object>';
/** An embedded object drawn by a picture of it, the form Word writes. */
const PICTURED_OBJECT_XML =
  `<w:object ${NS} xmlns:v="urn:schemas-microsoft-com:vml" ` +
  'xmlns:o="urn:schemas-microsoft-com:office:office" w:dxaOrig="1440" w:dyaOrig="720">' +
  '<v:shape id="_x0000_i1026" type="#_x0000_t75" style="width:72pt;height:36pt">' +
  '<v:imagedata r:id="rIdPreview" o:title=""/></v:shape>' +
  '<o:OLEObject Type="Embed" ProgID="Package" ShapeID="_x0000_i1026" ' +
  'DrawAspect="Content" ObjectID="_1700000002" r:id="rIdPicturedObject"/></w:object>';
const CHUNK_XML = `<w:altChunk ${NS} r:id="rIdChunk"/>`;
const PNG = Uint8Array.from(
  atob(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==",
  ),
  (character) => character.charCodeAt(0),
);

const RELATIONSHIP_TYPE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
/** The source package's parts the pasted markup reaches. */
type Part = { type: string; body: string | Uint8Array };
const SOURCE_PARTS: Record<string, Part> = {
  "word/charts/chart1.xml": {
    type: "application/vnd.openxmlformats-officedocument.drawingml.chart+xml",
    body: '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><c:chart/><c:externalData r:id="rId1"/></c:chartSpace>',
  },
  "word/charts/_rels/chart1.xml.rels": {
    type: "application/vnd.openxmlformats-package.relationships+xml",
    body:
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      `<Relationship Id="rId1" Type="${RELATIONSHIP_TYPE}/package" Target="../embeddings/Microsoft_Excel_Worksheet.xlsx"/>` +
      '<Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2011/relationships/chartColorStyle" Target="colors1.xml"/>' +
      '<Relationship Id="rId3" Type="http://schemas.microsoft.com/office/2011/relationships/chartStyle" Target="style1.xml"/>' +
      "</Relationships>",
  },
  "word/charts/colors1.xml": {
    type: "application/vnd.ms-office.chartcolorstyle+xml",
    body: '<cs:colorStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle" meth="cycle" id="10"/>',
  },
  "word/charts/style1.xml": {
    type: "application/vnd.ms-office.chartstyle+xml",
    body: '<cs:chartStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle" id="201"/>',
  },
  "word/embeddings/Microsoft_Excel_Worksheet.xlsx": {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    body: "PK workbook bytes",
  },
  "word/embeddings/oleObject1.bin": {
    type: "application/vnd.openxmlformats-officedocument.oleObject",
    body: "OLE object bytes",
  },
  "word/embeddings/oleObject2.bin": {
    type: "application/vnd.openxmlformats-officedocument.oleObject",
    body: "Pictured OLE object bytes",
  },
  "word/media/image1.png": { type: "image/png", body: PNG },
  "word/chunk1.html": {
    type: "text/html",
    body: "<html><body><p>Chunk</p></body></html>",
  },
};
const SOURCE_RELATIONSHIPS = [
  {
    id: "rIdChart",
    type: `${RELATIONSHIP_TYPE}/chart`,
    target: "charts/chart1.xml",
  },
  {
    id: "rIdObject",
    type: `${RELATIONSHIP_TYPE}/oleObject`,
    target: "embeddings/oleObject1.bin",
  },
  {
    id: "rIdPicturedObject",
    type: `${RELATIONSHIP_TYPE}/oleObject`,
    target: "embeddings/oleObject2.bin",
  },
  {
    id: "rIdPreview",
    type: `${RELATIONSHIP_TYPE}/image`,
    target: "media/image1.png",
  },
  {
    id: "rIdChunk",
    type: `${RELATIONSHIP_TYPE}/aFChunk`,
    target: "chunk1.html",
  },
];
/** A part of the destination's own that a copy must not overwrite. */
const DESTINATION_CHART = {
  path: "word/charts/chart1.xml",
  body: '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/>',
};

const paragraph = (paraId: string, text: string): Paragraph => ({
  type: "paragraph",
  paraId,
  content: [{ type: "run", content: [{ type: "text", text }] }],
});
const drawing = (rawXml: string): DrawingContent => ({
  type: "drawing",
  image: {
    type: "image",
    size: { width: 914_400, height: 457_200 },
    wrap: { type: "inline" },
  },
  rawXml,
  rawXmlMode: DRAWING_RAW_XML_MODES.PRESERVE_ONLY,
});

/** A package built from `content`, with `parts` and `relationships` added. */
const packageOf = async (
  content: BlockContent[],
  parts: Record<string, Part>,
  relationships: readonly { id: string; type: string; target: string }[],
): Promise<Document> => {
  const model = createSimpleDocument([]);
  model.package.document.content = content;
  const zip = await ZipArchive.read(await createDocx(model));
  const overrides = Object.entries(parts)
    .filter(([path]) => !path.endsWith(".rels"))
    .map(
      ([path, { type }]) =>
        `<Override PartName="/${path}" ContentType="${type}"/>`,
    )
    .join("");
  const types = (await zip.entry("[Content_Types].xml")?.text()) ?? "";
  zip.set(
    "[Content_Types].xml",
    types.replace("</Types>", `${overrides}</Types>`),
  );
  const rels = (await zip.entry("word/_rels/document.xml.rels")?.text()) ?? "";
  zip.set(
    "word/_rels/document.xml.rels",
    rels.replace(
      "</Relationships>",
      `${relationships.map(({ id, type, target }) => `<Relationship Id="${id}" Type="${type}" Target="${target}"/>`).join("")}</Relationships>`,
    ),
  );
  for (const [path, { body }] of Object.entries(parts)) zip.set(path, body);
  return parseDocx((await zip.generate({ compression: "store" })).buffer, {
    preloadFonts: false,
  });
};

const source = () =>
  packageOf(
    [
      paragraph("10000001", "Before"),
      {
        type: "paragraph",
        paraId: "10000002",
        content: [
          {
            type: "run",
            content: [drawing(CHART_XML), { type: "text", text: "Chart" }],
          },
          {
            type: "run",
            content: [{ type: "preservedXml", xml: OBJECT_XML, text: "" }],
          },
          {
            type: "run",
            content: [
              { type: "preservedXml", xml: PICTURED_OBJECT_XML, text: "" },
            ],
          },
        ],
      },
      { type: "preservedBlock", xml: CHUNK_XML },
      paragraph("10000003", "After"),
    ],
    SOURCE_PARTS,
    SOURCE_RELATIONSHIPS,
  );
const destination = () =>
  packageOf(
    [paragraph("20000001", "Target"), paragraph("20000002", "Tail")],
    {
      [DESTINATION_CHART.path]: {
        type: "application/vnd.openxmlformats-officedocument.drawingml.chart+xml",
        body: DESTINATION_CHART.body,
      },
    },
    [
      {
        id: "rIdOwnChart",
        type: `${RELATIONSHIP_TYPE}/chart`,
        target: "charts/chart1.xml",
      },
    ],
  );

/**
 * The copied range: the source's paragraph holding the chart and the object,
 * and with `block` the kept block after it. A suggested paste records no kept
 * block markup, so a suggesting copy leaves it out.
 */
const copiedSlice = (doc: ProseNode, block: boolean) => {
  const from = doc.child(0).nodeSize;
  return doc.slice(
    from,
    block
      ? from + doc.child(1).nodeSize + doc.child(2).nodeSize
      : from + doc.child(1).nodeSize,
  );
};

const pasted = async (mode: "editing" | "suggesting") => {
  const copiedFrom = readModel(
    createCanonicalSession(await source()).unwrap().document,
  );
  const copy = createCanonicalSession(copiedFrom).unwrap();
  const slice = copiedSlice(copy.projection.doc, mode === "editing");
  const sourceParts = await readClipboardPackageParts({
    document: copiedFrom,
    markup: sliceKeptMarkup(slice),
  });
  const session = createCanonicalSession(await destination()).unwrap();
  session.setMode(
    mode === "editing"
      ? { type: "editing" }
      : { type: "suggesting", author: "Paster" },
  );
  let state = EditorState.create({ doc: session.projection.doc });
  state = state.apply(
    state.tr.setSelection(
      TextSelection.create(state.doc, state.doc.child(0).nodeSize - 1),
    ),
  );
  return { session, state, slice, copiedFrom, sourceParts };
};

const textOf = async (zip: ZipArchive, path: string) =>
  decodeText((await zip.entry(path)?.bytes()) ?? new Uint8Array(), {
    ignoreBOM: true,
  });

/** Each relationship id an office attribute names in `xml`. */
const namedIds = (xml: string) =>
  [...xml.matchAll(/\sr:id="(?<id>[^"]+)"/gu)].flatMap(({ groups }) =>
    groups?.["id"] === undefined ? [] : [groups["id"]],
  );

/** The content type the package registry gives `path`. */
const contentTypeOf = (registry: string, path: string) => {
  for (const element of getChildElements(parseXmlDocument(registry))) {
    if (
      getLocalName(element.name ?? "") === "Override" &&
      getAttribute(element, null, "PartName") === `/${path}`
    )
      return getAttribute(element, null, "ContentType");
  }
  const extension = path.slice(path.lastIndexOf(".") + 1);
  for (const element of getChildElements(parseXmlDocument(registry)))
    if (
      getLocalName(element.name ?? "") === "Default" &&
      getAttribute(element, null, "Extension") === extension
    )
      return getAttribute(element, null, "ContentType");
  return null;
};

/** Assert every internal relationship of every part resolves to a part with a content type. */
const assertResolves = async (zip: ZipArchive) => {
  const registry = await textOf(zip, "[Content_Types].xml");
  const paths = new Set(zip.paths());
  const relsPaths = zip.paths().filter((path) => path.endsWith(".rels"));
  const relsParts = await Promise.all(
    relsPaths.map(async (path) => [path, await textOf(zip, path)] as const),
  );
  for (const [relsPath, xml] of relsParts)
    for (const relationship of parseRelationships(xml).values()) {
      if (relationship.targetMode === "External") continue;
      const target = resolveRelativePath(relsPath, relationship.target);
      expect(paths.has(target), `${relsPath} ${relationship.id}`).toBe(true);
      expect(contentTypeOf(registry, target)).not.toBeNull();
    }
};

/** The names the copies take in the destination. */
const COPIED_PATHS = new Set([
  "word/charts/chart2.xml",
  "word/charts/_rels/chart2.xml.rels",
  "word/charts/colors1.xml",
  "word/charts/style1.xml",
  "word/embeddings/Microsoft_Excel_Worksheet1.xlsx",
  "word/embeddings/oleObject1.bin",
  "word/embeddings/oleObject2.bin",
  "word/media/image1.png",
  "word/chunk1.html",
]);

const save = async (session: CanonicalSession, mode: string) => {
  const saved = await serializeCanonicalSave({
    snapshot: session.captureSaveSnapshot(),
    options: {
      mode:
        mode === "full"
          ? FOLIO_DOCX_SERIALIZATION_MODE.full
          : FOLIO_DOCX_SERIALIZATION_MODE.preferSelective,
    },
    featureFlags: { selectiveSave: true },
  });
  return saved.buffer;
};

describe.each(["editing", "suggesting"] as const)("in %s", (mode) => {
  test.each(["full", "selective"])(
    "pasted parts from another package save under fresh names and resolve after reopening (%s save)",
    async (saveMode) => {
      const { session, state, slice, copiedFrom, sourceParts } =
        await pasted(mode);
      const prepared = prepareCanonicalPaste({
        session,
        state,
        slice,
        sourceDocument: copiedFrom,
        ...(sourceParts === undefined ? {} : { sourceParts }),
      });
      if (prepared.isErr()) throw prepared.error;
      const next = publishCanonicalProjection({
        session,
        state,
        commit: prepared.value,
      }).unwrap().state;

      const zip = await ZipArchive.read(await save(session, saveMode));
      await assertResolves(zip);
      const documentXml = await textOf(zip, "word/document.xml");
      const relationships = parseRelationships(
        await textOf(zip, "word/_rels/document.xml.rels"),
      );
      // Each object keeps its markup and every id it names resolves.
      expect(documentXml.match(/<o:OLEObject /gu)).toHaveLength(2);
      const named = namedIds(documentXml);
      expect(named).toHaveLength(mode === "editing" ? 5 : 4);
      const targets = new Map<string, string[]>();
      for (const id of named) {
        const relationship = relationships.get(id);
        if (relationship === undefined) throw new TypeError(id);
        const type = relationship.type.slice(
          relationship.type.lastIndexOf("/") + 1,
        );
        targets.set(type, [
          ...(targets.get(type) ?? []),
          resolveRelativePath(
            "word/_rels/document.xml.rels",
            relationship.target,
          ),
        ]);
      }
      // The destination's own chart keeps its name and bytes.
      expect(targets.get("chart")).toEqual(["word/charts/chart2.xml"]);
      expect(await textOf(zip, DESTINATION_CHART.path)).toBe(
        DESTINATION_CHART.body,
      );
      expect(await textOf(zip, "word/charts/chart2.xml")).toBe(
        SOURCE_PARTS["word/charts/chart1.xml"]?.body ?? "",
      );
      expect(targets.get("oleObject")?.toSorted()).toEqual([
        "word/embeddings/oleObject1.bin",
        "word/embeddings/oleObject2.bin",
      ]);
      expect(targets.get("image")).toEqual(["word/media/image1.png"]);
      expect(await zip.entry("word/media/image1.png")?.bytes()).toEqual(PNG);
      if (mode === "editing")
        expect(targets.get("aFChunk")).toEqual(["word/chunk1.html"]);
      // The chart's own relationships reach its copied workbook and styles.
      const chartRelationships = parseRelationships(
        await textOf(zip, "word/charts/_rels/chart2.xml.rels"),
      );
      expect(
        [...chartRelationships.values()].map(({ target }) =>
          resolveRelativePath("word/charts/_rels/chart2.xml.rels", target),
        ),
      ).toEqual([
        "word/embeddings/Microsoft_Excel_Worksheet1.xlsx",
        "word/charts/colors1.xml",
        "word/charts/style1.xml",
      ]);
      expect(
        await textOf(zip, "word/embeddings/Microsoft_Excel_Worksheet1.xlsx"),
      ).toBe(
        SOURCE_PARTS["word/embeddings/Microsoft_Excel_Worksheet.xlsx"]?.body ??
          "",
      );
      const registry = await textOf(zip, "[Content_Types].xml");
      expect(contentTypeOf(registry, "word/charts/chart2.xml")).toBe(
        SOURCE_PARTS["word/charts/chart1.xml"]?.type ?? "",
      );
      expect(contentTypeOf(registry, "word/charts/colors1.xml")).toBe(
        SOURCE_PARTS["word/charts/colors1.xml"]?.type ?? "",
      );

      // Reopened, the markup still names its parts.
      const reopened = await parseDocx(
        (await zip.generate({ compression: "store" })).buffer,
        { preloadFonts: false },
      );
      const reopenedXml = JSON.stringify(reopened.package.document.content);
      for (const id of named) {
        expect(reopenedXml).toContain(id);
        expect(reopened.package.relationships?.has(id)).toBe(true);
      }

      // Undone, the save holds none of the copies.
      publishCanonicalProjection({
        session,
        state: next,
        commit: session.prepareUndo(next).unwrap(),
      }).unwrap();
      expect(session.document.package.copiedParts ?? new Map()).toEqual(
        new Map(),
      );
      const undone = await ZipArchive.read(await save(session, saveMode));
      await assertResolves(undone);
      expect(undone.paths().filter((path) => COPIED_PATHS.has(path))).toEqual(
        [],
      );
      expect(namedIds(await textOf(undone, "word/document.xml"))).toEqual([]);
    },
  );
});

test("without the source package's parts, a paste naming them into another package refuses", async () => {
  const { session, state, slice, copiedFrom } = await pasted("editing");
  const prepared = prepareCanonicalPaste({
    session,
    state,
    slice,
    sourceDocument: copiedFrom,
  });
  expect(prepared.isErr()).toBe(true);
  if (prepared.isOk()) return;
  expect(prepared.error.ground).toEqual({ gap: "story-content-projection" });
});

test("the source parts read for a copy are the ones its markup reaches", async () => {
  const { slice, copiedFrom, sourceParts } = await pasted("editing");
  expect(sliceKeptMarkup(slice).length).toBeGreaterThan(0);
  expect(sourceParts?.source).toBe(copiedFrom);
  expect([...(sourceParts?.parts.keys() ?? [])].toSorted()).toEqual(
    Object.keys(SOURCE_PARTS).toSorted(),
  );
  expect(
    new Uint8Array(
      sourceParts?.parts.get("word/charts/chart1.xml")?.data ?? [],
    ),
  ).toEqual(encodeUtf8(SOURCE_PARTS["word/charts/chart1.xml"]?.body ?? ""));
});
