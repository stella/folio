/**
 * A drawing folio cannot project is written back untouched, in every part.
 *
 * A `wpg:wgp` group has no editable projection: `parseGroupDrawing` renders it
 * to an SVG preview and the model holds that render, not the group. The
 * capture is therefore the content, and regenerating DrawingML from the
 * preview emits whichever child picture the rasterizer saw first, under that
 * child's relationship, in place of the whole group — which the census saw as
 * `image.filename: "wordprocessing-group.svg" became absent` once the rebuild
 * path ran.
 *
 * The property fixes the input class the single corpus file does not cover:
 * every part a drawing can live in (body, header, footer, footnote, endnote)
 * crossed with the three states the model can be in when the save runs — as
 * parsed, with the preview's fingerprint poisoned exactly as the `reserialize`
 * invariant poisons it, and after an editor round trip. In all fifteen, the
 * part must come back carrying the source `w:drawing` byte for byte and its
 * relationship must still resolve to the same target.
 */

import { describe, expect, test } from "bun:test";
import fc from "fast-check";
import JSZip from "jszip";

import { propertyConfig, propertyTestTimeout } from "../../../../test/property-testing";

import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import type { Document } from "../types/document";
import { EDITED_PREVIEW_FINGERPRINT } from "./imageRawXml";
import { parseDocx } from "./parser";
import { RELATIONSHIP_TYPES } from "./relsParser";
import { repackDocx } from "./rezip";

const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

const NS = [
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"',
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"',
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"',
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"',
  'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"',
  'xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"',
  'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"',
].join(" ");

const IMAGE_RELATIONSHIP_ID = "rIdGroupPicture";
const IMAGE_TARGET = "media/group-child.png";

/** A group holding a picture and a shape: no single model can stand for it. */
const GROUP_DRAWING = `<w:drawing><wp:inline><wp:extent cx="2000000" cy="1000000"/><wp:docPr id="11" name="Group 11"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"><wpg:wgp><pic:pic><pic:nvPicPr><pic:cNvPr id="12" name="child.png"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${IMAGE_RELATIONSHIP_ID}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000000" cy="500000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic><wps:wsp><wps:spPr><a:xfrm><a:off x="1000000" y="0"/><a:ext cx="1000000" cy="500000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="DBEDF3"/></a:solidFill></wps:spPr></wps:wsp></wpg:wgp></a:graphicData></a:graphic></wp:inline></w:drawing>`;

const GROUP_PARAGRAPH = `<w:p><w:r>${GROUP_DRAWING}</w:r></w:p>`;

/** Every part a run-level drawing can live in. */
const PART_PLACEMENTS = {
  body: "word/document.xml",
  header: "word/header1.xml",
  footer: "word/footer1.xml",
  footnote: "word/footnotes.xml",
  endnote: "word/endnotes.xml",
} as const;

type PartPlacement = keyof typeof PART_PLACEMENTS;

/** How the model looks by the time the save runs. */
const MODEL_STATES = ["as-parsed", "fingerprint-poisoned", "editor-round-trip"] as const;
type ModelState = (typeof MODEL_STATES)[number];

const placementCase = fc.record({
  placement: fc.constantFrom(...(Object.keys(PART_PLACEMENTS) as PartPlacement[])),
  state: fc.constantFrom(...MODEL_STATES),
});

const filler = "<w:p><w:r><w:t>text</w:t></w:r></w:p>";

const contentFor = (placement: PartPlacement, part: PartPlacement): string =>
  placement === part ? GROUP_PARAGRAPH : filler;

const PNG_1X1_BASE64 =
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";

/** Only the part carrying the group needs the image relationship. */
const partRels = (placement: PartPlacement, part: PartPlacement, extra = ""): string =>
  `${XML_DECLARATION}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${extra}${
    placement === part
      ? `<Relationship Id="${IMAGE_RELATIONSHIP_ID}" Type="${RELATIONSHIP_TYPES.image}" Target="${IMAGE_TARGET}"/>`
      : ""
  }</Relationships>`;

const docxFor = (placement: PartPlacement): Promise<ArrayBuffer> => {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    `${XML_DECLARATION}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`,
  );
  zip.file(
    "_rels/.rels",
    `${XML_DECLARATION}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${RELATIONSHIP_TYPES.officeDocument}" Target="word/document.xml"/></Relationships>`,
  );
  zip.file(
    "word/_rels/document.xml.rels",
    partRels(
      placement,
      "body",
      [
        `<Relationship Id="rIdHdr" Type="${RELATIONSHIP_TYPES.header}" Target="header1.xml"/>`,
        `<Relationship Id="rIdFtr" Type="${RELATIONSHIP_TYPES.footer}" Target="footer1.xml"/>`,
        `<Relationship Id="rIdFn" Type="${RELATIONSHIP_TYPES.footnotes}" Target="footnotes.xml"/>`,
        `<Relationship Id="rIdEn" Type="${RELATIONSHIP_TYPES.endnotes}" Target="endnotes.xml"/>`,
      ].join(""),
    ),
  );
  zip.file(
    "word/document.xml",
    `${XML_DECLARATION}<w:document ${NS}><w:body>${contentFor(placement, "body")}<w:sectPr><w:headerReference w:type="default" r:id="rIdHdr"/><w:footerReference w:type="default" r:id="rIdFtr"/></w:sectPr></w:body></w:document>`,
  );
  zip.file(
    "word/header1.xml",
    `${XML_DECLARATION}<w:hdr ${NS}>${contentFor(placement, "header")}</w:hdr>`,
  );
  zip.file(
    "word/footer1.xml",
    `${XML_DECLARATION}<w:ftr ${NS}>${contentFor(placement, "footer")}</w:ftr>`,
  );
  zip.file(
    "word/footnotes.xml",
    `${XML_DECLARATION}<w:footnotes ${NS}><w:footnote w:id="1">${contentFor(placement, "footnote")}</w:footnote></w:footnotes>`,
  );
  zip.file(
    "word/endnotes.xml",
    `${XML_DECLARATION}<w:endnotes ${NS}><w:endnote w:id="1">${contentFor(placement, "endnote")}</w:endnote></w:endnotes>`,
  );
  const SIDECAR_PARTS = {
    header1: "header",
    footer1: "footer",
    footnotes: "footnote",
    endnotes: "endnote",
  } as const satisfies Record<string, PartPlacement>;
  for (const [part, placementForPart] of Object.entries(SIDECAR_PARTS)) {
    zip.file(`word/_rels/${part}.xml.rels`, partRels(placement, placementForPart));
  }
  zip.file("word/media/group-child.png", PNG_1X1_BASE64, { base64: true });
  return zip.generateAsync({ type: "arraybuffer" });
};

/**
 * `rawImageFingerprint: "editedPreview"` is what the `reserialize` invariant
 * writes to force the serializer, and what an editor edit leaves behind.
 */
const poisonPreviewFingerprints = (value: unknown): void => {
  if (Array.isArray(value)) {
    for (const item of value) {
      poisonPreviewFingerprints(item);
    }
    return;
  }
  if (value instanceof Map) {
    for (const item of value.values()) {
      poisonPreviewFingerprints(item);
    }
    return;
  }
  if (typeof value !== "object" || value === null) {
    return;
  }
  const record: Record<string, unknown> = value as Record<string, unknown>;
  if (typeof record["rawImageFingerprint"] === "string") {
    record["rawImageFingerprint"] = EDITED_PREVIEW_FINGERPRINT;
  }
  for (const item of Object.values(record)) {
    poisonPreviewFingerprints(item);
  }
};

const applyState = (document: Document, state: ModelState): Document => {
  switch (state) {
    case "as-parsed":
      return document;
    case "fingerprint-poisoned":
      poisonPreviewFingerprints(document.package);
      return document;
    case "editor-round-trip":
      return fromProseDoc(toProseDoc(document), document);
    default:
      return state satisfies never;
  }
};

const partText = async (buffer: ArrayBuffer, path: string): Promise<string> => {
  const file = (await JSZip.loadAsync(buffer)).file(path);
  if (!file) {
    throw new Error(`saved package has no ${path}`);
  }
  return file.async("string");
};

describe("a group drawing is replayed, never rebuilt", () => {
  test(
    "every part keeps the source drawing and its relationship, in every model state",
    async () => {
      await fc.assert(
        fc.asyncProperty(placementCase, async ({ placement, state }) => {
          const source = await docxFor(placement);
          const parsed = await parseDocx(source, { detectVariables: false, preloadFonts: false });
          const saved = await repackDocx(applyState(parsed, state), {
            updateModifiedDate: false,
          });

          const path = PART_PLACEMENTS[placement];
          expect(await partText(saved, path)).toContain(GROUP_DRAWING);

          const relsPath = path.replace("word/", "word/_rels/").concat(".rels");
          const rels = await partText(saved, relsPath);
          expect(rels).toContain(`Id="${IMAGE_RELATIONSHIP_ID}"`);
          expect(rels).toContain(`Target="${IMAGE_TARGET}"`);
        }),
        propertyConfig({ numRuns: 45 }),
      );
    },
    propertyTestTimeout(60_000),
  );

  test("the preview's filename survives a rebuild of the body", async () => {
    const parsed = await parseDocx(await docxFor("body"), {
      detectVariables: false,
      preloadFonts: false,
    });
    poisonPreviewFingerprints(parsed.package);
    const reopened = await parseDocx(await repackDocx(parsed, { updateModifiedDate: false }), {
      detectVariables: false,
      preloadFonts: false,
    });

    const block = reopened.package.document.content.at(0);
    const run = block?.type === "paragraph" ? block.content.at(0) : undefined;
    const drawing = run?.type === "run" ? run.content.at(0) : undefined;
    expect(drawing?.type).toBe("drawing");
    if (drawing?.type !== "drawing") {
      return;
    }
    expect(drawing.image.filename).toBe("wordprocessing-group.svg");
  });
});
