/**
 * A preview the budget does not see is a preview with no bound, which is how
 * the SmartArt raster came to retain tens of megabytes from a package of
 * twenty kilobytes. The budget sees exactly what the parse's ledger holds, so
 * these tests drive real producers through real parses and check the ledger
 * against the model in both directions: every preview the model carries was
 * charged, and nothing the model does not carry spent allowance.
 *
 * The SmartArt producer no longer emits a raster, so the character budget has
 * nothing of its to charge: the drawing is drawn from its descriptor, bounded
 * by the shape cap the parse applies.
 */

import { describe, expect, test } from "bun:test";
import { DRAWING_RAW_XML_MODES } from "@stll/docx-core/model";
import JSZip from "jszip";

import type { Document, DrawingContent, MediaFile, Run } from "../types/document";
import { parseDiagramPreview } from "./diagramPreview";
import { parseDocx, parseDocxWithPreviewBudget } from "./parser";
import { createPackagePreviewBudget } from "./previewBudget";
import { RELATIONSHIP_TYPES, parseRelationships } from "./relsParser";
import { repackDocx, validateDocx } from "./rezip";
import { parseXmlDocument } from "./xmlParser";

const RELATIONSHIPS = parseRelationships(
  `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdData" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData" Target="diagrams/data1.xml"/><Relationship Id="rIdDrawing" Type="http://schemas.microsoft.com/office/2007/relationships/diagramDrawing" Target="diagrams/drawing1.xml"/></Relationships>`,
);

const MEDIA: Map<string, MediaFile> = new Map([
  [
    "word/diagrams/drawing1.xml",
    {
      path: "word/diagrams/drawing1.xml",
      filename: "drawing1.xml",
      mimeType: "application/xml",
      data: new TextEncoder().encode(
        `<dsp:drawing xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><dsp:spTree><dsp:sp><dsp:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="400" cy="200"/></a:xfrm><a:solidFill><a:srgbClr val="70AD47"/></a:solidFill></dsp:spPr></dsp:sp></dsp:spTree></dsp:drawing>`,
      ).buffer,
    },
  ],
]);

const diagramDrawing = () => {
  const drawing = parseXmlDocument(
    `<w:drawing xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><wp:inline><wp:extent cx="400" cy="200"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram"><dgm:relIds r:dm="rIdData"/></a:graphicData></a:graphic></wp:inline></w:drawing>`,
  );
  if (!drawing) {
    throw new Error("diagram fixture did not parse");
  }
  return drawing;
};

const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

const NAMESPACES = [
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"',
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"',
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"',
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"',
  'xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"',
  'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"',
  'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"',
  'xmlns:v="urn:schemas-microsoft-com:vml"',
  'xmlns:o="urn:schemas-microsoft-com:office:office"',
].join(" ");

/**
 * A VML rectangle folio renders, labelled by its fill. Every label has the
 * same length, so every preview of one kind costs the same.
 */
const vmlRun = (fill: string): string =>
  `<w:r><w:pict><v:rect style="width:20pt;height:10pt" fillcolor="#${fill}" stroked="f"/></w:pict></w:r>`;

/** A WordprocessingGroup folio renders, labelled by its fill. */
const groupDrawing = (fill: string): string =>
  `<w:drawing><wp:inline><wp:extent cx="2000000" cy="1000000"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"><wpg:wgp><wps:wsp><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2000000" cy="500000"/></a:xfrm><a:solidFill><a:srgbClr val="${fill}"/></a:solidFill></wps:spPr></wps:wsp></wpg:wgp></a:graphicData></a:graphic></wp:inline></w:drawing>`;

const groupRun = (fill: string): string => `<w:r>${groupDrawing(fill)}</w:r>`;

/** A DrawingML text box whose content holds `inner` paragraphs. */
const textBoxRun = (inner: string): string =>
  `<w:r><w:drawing><wp:anchor behindDoc="0" distT="0" distB="0" distL="0" distR="0" simplePos="0" locked="0" layoutInCell="1" allowOverlap="1" relativeHeight="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV><wp:extent cx="1828800" cy="914400"/><wp:wrapNone/><wp:docPr id="7" name="Text Box 7"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr><wps:txbx><w:txbxContent>${inner}</w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>`;

/** A VML text box whose content holds `inner` paragraphs. */
const vmlTextBoxRun = (inner: string): string =>
  `<w:r><w:pict><v:shape style="width:100pt;height:50pt"><v:textbox><w:txbxContent>${inner}</w:txbxContent></v:textbox></v:shape></w:pict></w:r>`;

type PackageParts = {
  body: string;
  header?: string;
  footer?: string;
  footnote?: string;
  /** Raw `w:footnote` siblings written before the first one. */
  footnotesBefore?: string;
  /** Raw `w:footnote` siblings written after the first one. */
  footnotesAfter?: string;
  endnote?: string;
  comment?: string;
  /** Raw `w:comment` siblings written after the first one. */
  commentsAfter?: string;
};

const PART_TYPES = {
  header: "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml",
  footer: "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml",
  footnotes: "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml",
  endnotes: "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml",
  comments: "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml",
} as const;

const packageOf = async (parts: PackageParts): Promise<ArrayBuffer> => {
  const zip = new JSZip();
  const overrides: string[] = [];
  const relationships: string[] = [];
  const addPart = (name: string, type: string, relationshipType: string, xml: string): void => {
    zip.file(`word/${name}`, `${XML}${xml}`);
    overrides.push(`<Override PartName="/word/${name}" ContentType="${type}"/>`);
    relationships.push(
      `<Relationship Id="rId${String(relationships.length + 1)}" Type="${relationshipType}" Target="${name}"/>`,
    );
  };
  if (parts.header !== undefined) {
    addPart(
      "header1.xml",
      PART_TYPES.header,
      RELATIONSHIP_TYPES.header,
      `<w:hdr ${NAMESPACES}><w:p>${parts.header}</w:p></w:hdr>`,
    );
  }
  if (parts.footer !== undefined) {
    addPart(
      "footer1.xml",
      PART_TYPES.footer,
      RELATIONSHIP_TYPES.footer,
      `<w:ftr ${NAMESPACES}><w:p>${parts.footer}</w:p></w:ftr>`,
    );
  }
  if (parts.footnote !== undefined) {
    addPart(
      "footnotes.xml",
      PART_TYPES.footnotes,
      RELATIONSHIP_TYPES.footnotes,
      `<w:footnotes ${NAMESPACES}>${parts.footnotesBefore ?? ""}<w:footnote w:id="1"><w:p>${parts.footnote}</w:p></w:footnote>${parts.footnotesAfter ?? ""}</w:footnotes>`,
    );
  }
  if (parts.endnote !== undefined) {
    addPart(
      "endnotes.xml",
      PART_TYPES.endnotes,
      RELATIONSHIP_TYPES.endnotes,
      `<w:endnotes ${NAMESPACES}><w:endnote w:id="1"><w:p>${parts.endnote}</w:p></w:endnote></w:endnotes>`,
    );
  }
  if (parts.comment !== undefined) {
    addPart(
      "comments.xml",
      PART_TYPES.comments,
      RELATIONSHIP_TYPES.comments,
      `<w:comments ${NAMESPACES}><w:comment w:id="0" w:author="A"><w:p>${parts.comment}</w:p></w:comment>${parts.commentsAfter ?? ""}</w:comments>`,
    );
  }
  const commentAnchor =
    parts.comment === undefined
      ? ""
      : `<w:p><w:commentRangeStart w:id="0"/><w:r><w:t>a</w:t></w:r><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r></w:p>`;
  zip.file(
    "[Content_Types].xml",
    `${XML}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>${overrides.join("")}</Types>`,
  );
  zip.file(
    "_rels/.rels",
    `${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${RELATIONSHIP_TYPES.officeDocument}" Target="word/document.xml"/></Relationships>`,
  );
  zip.file(
    "word/_rels/document.xml.rels",
    `${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${relationships.join("")}</Relationships>`,
  );
  zip.file(
    "word/document.xml",
    `${XML}<w:document ${NAMESPACES}><w:body>${parts.body}${commentAnchor}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`,
  );
  return zip.generateAsync({ type: "arraybuffer" });
};

/** A model node the parse built as a preview. */
const isPreviewDrawing = (value: object): value is DrawingContent =>
  "type" in value &&
  value.type === "drawing" &&
  "rawXmlMode" in value &&
  value.rawXmlMode === DRAWING_RAW_XML_MODES.PREVIEW_ONLY;

/** Every preview-only drawing in the model, in model order. */
const previewDrawings = (document: Document): DrawingContent[] => {
  const found: DrawingContent[] = [];
  const seen = new Set<unknown>();
  const visit = (value: unknown): void => {
    if (value === null || typeof value !== "object" || seen.has(value)) {
      return;
    }
    seen.add(value);
    if (value instanceof ArrayBuffer || ArrayBuffer.isView(value)) {
      return;
    }
    if (isPreviewDrawing(value)) {
      found.push(value);
    }
    for (const child of value instanceof Map ? value.values() : Object.values(value)) {
      visit(child);
    }
  };
  visit(document.package);
  return found;
};

/** The fill a preview was labelled with, read back out of its SVG. */
const labelOf = (drawing: DrawingContent): string | undefined =>
  /%23([A-D][0-9A-F]{5})/iu
    .exec(drawing.image.src ?? "")
    ?.at(1)
    ?.toUpperCase();

const svgKindOf = (drawing: DrawingContent): "vmlShape" | "wpGroup" =>
  drawing.image.filename === "wordprocessing-group.svg" ? "wpGroup" : "vmlShape";

/**
 * One preview in every container a parse reaches a producer from, each
 * labelled so a test can say which preview is which.
 */
const EVERY_CONTAINER = {
  body: [
    `<w:p>${vmlRun("A00001")}</w:p>`,
    `<w:p>${groupRun("B00001")}</w:p>`,
    `<w:p><w:hyperlink w:anchor="x">${vmlRun("A00002")}</w:hyperlink></w:p>`,
    `<w:p><w:hyperlink w:anchor="y"><w:ins w:id="11" w:author="A">${vmlRun("A00003")}</w:ins></w:hyperlink></w:p>`,
    `<w:p><w:fldSimple w:instr=" PAGE ">${vmlRun("A00004")}</w:fldSimple></w:p>`,
    `<w:p><w:sdt><w:sdtContent>${vmlRun("A00005")}</w:sdtContent></w:sdt></w:p>`,
    `<w:p><w:ins w:id="12" w:author="A">${vmlRun("A00006")}</w:ins></w:p>`,
    `<w:p><w:smartTag w:uri="u" w:element="e">${vmlRun("A00007")}</w:smartTag></w:p>`,
    `<w:p><w:r><mc:AlternateContent><mc:Choice Requires="wpg">${groupDrawing("B00002")}</mc:Choice><mc:Fallback/></mc:AlternateContent></w:r></w:p>`,
    `<w:sdt><w:sdtContent><w:p>${vmlRun("A00008")}</w:p></w:sdtContent></w:sdt>`,
    `<w:tbl><w:tr><w:tc><w:p>${vmlRun("A00009")}</w:p><w:tbl><w:tr><w:tc><w:p>${vmlRun("A0000A")}</w:p></w:tc></w:tr></w:tbl><w:p/></w:tc></w:tr></w:tbl>`,
    `<w:p>${textBoxRun(`<w:p>${vmlRun("A0000B")}</w:p>`)}</w:p>`,
    `<w:p>${vmlTextBoxRun(`<w:p>${vmlRun("A0000C")}</w:p>`)}</w:p>`,
  ].join(""),
  header: vmlRun("A0000D"),
  footer: groupRun("B00003"),
  footnote: vmlRun("A0000E"),
  endnote: vmlRun("A0000F"),
  comment: vmlRun("A00010"),
} as const satisfies PackageParts;

const EVERY_LABEL = [
  ...Array.from({ length: 16 }, (_, index) => `A${(index + 1).toString(16).padStart(5, "0")}`),
  "B00001",
  "B00002",
  "B00003",
].map((label) => label.toUpperCase());

const labels = (drawings: DrawingContent[]): string[] =>
  drawings.map((drawing) => labelOf(drawing) ?? "unlabelled").sort();

describe("package preview budget", () => {
  test("the SmartArt producer emits a description, not a raster", () => {
    const image = parseDiagramPreview(diagramDrawing(), RELATIONSHIPS, MEDIA);
    expect(image?.src).toBeUndefined();
    expect(image?.mimeType).toBeUndefined();
    expect(image?.filename).toBeUndefined();
    expect(image?.preview?.kind).toBe("diagram");
  });

  test("the ledger charges what it built, and one kind's allowance does not spend another's", () => {
    const budget = createPackagePreviewBudget();
    const frame = { size: { width: 1, height: 1 }, wrap: { type: "inline" } } as const;
    const vml = budget.ledger.svgImage("vmlShape", "<svg/>", frame);
    const group = budget.ledger.svgImage("wpGroup", "<svg/>", frame);
    const vmlSrc = vml.src;
    expect(vml.rId).toBeUndefined();
    expect(vml.filename).not.toBe(group.filename);

    budget.enforce({ wpGroup: 0 });

    expect(group.src).toBeUndefined();
    expect(group.size).toEqual(frame.size);
    expect(vml.src).toBe(vmlSrc);
  });

  test("a package's retained preview text is bounded however many previews it has", () => {
    const budget = createPackagePreviewBudget();
    const frame = { size: { width: 1, height: 1 }, wrap: { type: "inline" } } as const;
    const previews = Array.from({ length: 8 }, () =>
      budget.ledger.svgImage("vmlShape", "<svg/>", frame),
    );
    const single = previews.at(0)?.src?.length ?? 0;
    expect(single).toBeGreaterThan(0);

    // Room for three, offered eight: the first three built keep their render.
    budget.enforce({ vmlShape: single * 3 });
    expect(previews.map((image) => image.src !== undefined)).toEqual([
      true,
      true,
      true,
      false,
      false,
      false,
      false,
      false,
    ]);
  });

  test("the package reaches a producer from every container", async () => {
    const document = await parseDocx(await packageOf(EVERY_CONTAINER), { preloadFonts: false });
    expect(labels(previewDrawings(document))).toEqual([...EVERY_LABEL].sort());
  });

  /**
   * Under-registration: a preview the ledger never saw would survive an
   * allowance of nothing.
   */
  test("every preview the model carries was charged", async () => {
    const document = await parseDocxWithPreviewBudget(
      await packageOf(EVERY_CONTAINER),
      { preloadFonts: false },
      { vmlShape: 0, wpGroup: 0 },
    );
    const drawings = previewDrawings(document);
    expect(drawings).toHaveLength(EVERY_LABEL.length);
    expect(drawings.filter((drawing) => drawing.image.src !== undefined)).toEqual([]);
  });

  /**
   * Over-registration: a preview built for content the parse then discarded
   * would spend allowance, so an allowance of exactly what the model carries
   * would drop something.
   */
  test("nothing the model does not carry spends allowance", async () => {
    const source = await packageOf(EVERY_CONTAINER);
    const retained = previewDrawings(await parseDocx(source, { preloadFonts: false }));
    const spent = { vmlShape: 0, wpGroup: 0 };
    for (const drawing of retained) {
      spent[svgKindOf(drawing)] += drawing.image.src?.length ?? 0;
    }

    const exact = previewDrawings(
      await parseDocxWithPreviewBudget(source, { preloadFonts: false }, spent),
    );
    expect(exact.filter((drawing) => drawing.image.src === undefined)).toEqual([]);
  });

  /**
   * Content a parse builds and then discards: a separator note, a note or a
   * comment repeating an id, and the paragraph hosting a header watermark.
   * Each is followed by a preview the model keeps, so a discarded preview that
   * still spent allowance would cost a retained one its render.
   */
  test("previews in discarded content spend no allowance", async () => {
    const watermark = `<w:r><w:pict><v:shape id="PowerPlusWaterMarkObject1" o:spid="_x0000_s1" type="#_x0000_t136" style="position:absolute;margin-left:0;margin-top:0;width:400pt;height:100pt;z-index:-1;mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin" fillcolor="silver" stroked="f"><v:textpath style="font-family:&quot;Calibri&quot;" string="DRAFT"/></v:shape></w:pict></w:r>`;
    const source = await packageOf({
      body: `<w:p>${vmlRun("E00001")}</w:p>`,
      header: `${watermark}${vmlRun("E00002")}`,
      footer: vmlRun("E00003"),
      footnotesBefore: `<w:footnote w:type="separator" w:id="-1"><w:p>${vmlRun("E00004")}</w:p></w:footnote>`,
      footnote: vmlRun("E00005"),
      footnotesAfter: `<w:footnote w:id="1"><w:p>${vmlRun("E00006")}</w:p></w:footnote><w:footnote w:id="2"><w:p>${vmlRun("E00007")}</w:p></w:footnote>`,
      endnote: vmlRun("E00008"),
      comment: vmlRun("E00009"),
      commentsAfter: `<w:comment w:id="0" w:author="B"><w:p>${vmlRun("E0000A")}</w:p></w:comment><w:comment w:id="1" w:author="C"><w:p>${vmlRun("E0000B")}</w:p></w:comment>`,
    });
    const retained = previewDrawings(await parseDocx(source, { preloadFonts: false }));
    const spent = retained.reduce((sum, drawing) => sum + (drawing.image.src?.length ?? 0), 0);

    const exact = previewDrawings(
      await parseDocxWithPreviewBudget(source, { preloadFonts: false }, { vmlShape: spent }),
    );
    expect(labels(exact)).toEqual(labels(retained));
    expect(exact.filter((drawing) => drawing.image.src === undefined)).toEqual([]);
  });

  /**
   * Previews are charged in the order the parse builds them: the body, a text
   * box's content after the paragraph that anchors it, headers and footers in
   * relationship order, then footnotes, endnotes and comments. One character
   * short of the whole allowance drops the last one built.
   */
  test("a package over its allowance drops the previews built last", async () => {
    const source = await packageOf(EVERY_CONTAINER);
    const retained = previewDrawings(await parseDocx(source, { preloadFonts: false }));
    const vmlCharacters = retained
      .filter((drawing) => svgKindOf(drawing) === "vmlShape")
      .reduce((sum, drawing) => sum + (drawing.image.src?.length ?? 0), 0);

    const short = previewDrawings(
      await parseDocxWithPreviewBudget(
        source,
        { preloadFonts: false },
        { vmlShape: vmlCharacters - 1 },
      ),
    );
    const dropped = retained.filter((_, index) => short[index]?.image.src === undefined);
    expect(labels(dropped)).toEqual(["A00010"]);
  });

  test("a text box's previews are charged after its paragraph's own", async () => {
    const source = await packageOf({
      body: `<w:p>${textBoxRun(`<w:p>${vmlRun("C00001")}</w:p>`)}${vmlRun("C00002")}</w:p>`,
    });
    const [single] = previewDrawings(await parseDocx(source, { preloadFonts: false }));
    const document = await parseDocxWithPreviewBudget(
      source,
      { preloadFonts: false },
      { vmlShape: single?.image.src?.length ?? 0 },
    );
    const kept = previewDrawings(document).filter((drawing) => drawing.image.src !== undefined);
    expect(labels(kept)).toEqual(["C00002"]);
  });

  /**
   * What being over budget costs, end to end: the group preview goes and
   * nothing else does. The drawing keeps the space it reserved on the page,
   * the package saves the group it was authored with, and opening the saved
   * package renders the preview again, because dropping one was an economy in
   * the model and never a loss from the file.
   */
  test("drops a group preview past its allowance and still saves the group", async () => {
    const original = await packageOf({ body: `<w:p>${groupRun("DBEDF3")}</w:p>` });
    const firstDrawing = (document: Document): DrawingContent | undefined => {
      const block = document.package.document.content.at(0);
      if (block?.type !== "paragraph") {
        return undefined;
      }
      const run = block.content.find((item): item is Run => item.type === "run");
      return run?.content.find((item): item is DrawingContent => item.type === "drawing");
    };
    const preview = firstDrawing(await parseDocx(original, { preloadFonts: false }))?.image.src;
    expect(preview).toStartWith("data:image/svg+xml");

    const document = await parseDocxWithPreviewBudget(
      original,
      { preloadFonts: false },
      { wpGroup: 0 },
    );
    const drawing = firstDrawing(document);
    expect(drawing?.rawXmlMode).toBe("previewOnly");
    expect(drawing?.image.src).toBeUndefined();
    expect(drawing?.image.size).toEqual({ width: 2_000_000, height: 1_000_000 });

    const saved = await repackDocx(document, { updateModifiedDate: false });
    expect((await validateDocx(saved)).valid).toBe(true);
    const savedXml = await (await JSZip.loadAsync(saved)).file("word/document.xml")!.async("text");
    expect(savedXml).toContain("<wpg:wgp>");
    expect(savedXml).toContain('<a:srgbClr val="DBEDF3"/>');
    expect(savedXml).not.toContain("svg");

    const reopened = await parseDocx(saved, { preloadFonts: false });
    expect(firstDrawing(reopened)?.image.src).toBe(preview);
  });
});
