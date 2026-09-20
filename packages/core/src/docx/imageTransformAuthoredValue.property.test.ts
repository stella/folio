/**
 * An authored transform value survives a real save, including zero.
 *
 * `rot="0"` and `flipH="0"` are OOXML's defaults, so a truthiness guard on the
 * model cannot tell "the author wrote none" from "the author wrote nothing".
 * The serializer guarded on truthiness, so an authored zero rotation was
 * dropped on save and came back as an absent transform: the absence-versus-
 * default class, where the only way to be right is to write iff authored.
 *
 * The property is over the input class rather than one example, because the
 * defect lives at the boundary between the values that are falsy and the ones
 * that are not: rotation crossed with both flips, each absent, off, or on.
 * The save is forced past replay the way the corpus gate forces it, so the
 * serializer has to build `a:xfrm` from the model.
 */

import { describe, expect, test } from "bun:test";
import fc from "fast-check";
import JSZip from "jszip";

import { propertyConfig, propertyTestTimeout } from "../../../../test/property-testing";

import type { Document, Image } from "../types/document";
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
].join(" ");

const IMAGE_RELATIONSHIP_ID = "rIdPicture";
const PNG_1X1_BASE64 =
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";

/** `rot` in 60000ths of a degree: absent, zero, either sign, and a right angle. */
const ROTATIONS = [undefined, "0", "1", "-1", "5400000"] as const;
/** `flipH` / `flipV`: absent, authored off, authored on. */
const FLIPS = [undefined, "0", "1"] as const;

type AuthoredTransform = {
  rot: (typeof ROTATIONS)[number];
  flipH: (typeof FLIPS)[number];
  flipV: (typeof FLIPS)[number];
};

const authoredTransform = fc.record({
  rot: fc.constantFrom(...ROTATIONS),
  flipH: fc.constantFrom(...FLIPS),
  flipV: fc.constantFrom(...FLIPS),
});

const xfrmAttributes = ({ rot, flipH, flipV }: AuthoredTransform): string =>
  [
    rot === undefined ? "" : ` rot="${rot}"`,
    flipH === undefined ? "" : ` flipH="${flipH}"`,
    flipV === undefined ? "" : ` flipV="${flipV}"`,
  ].join("");

const docxWith = (transform: AuthoredTransform): Promise<ArrayBuffer> => {
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
    `${XML_DECLARATION}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="${IMAGE_RELATIONSHIP_ID}" Type="${RELATIONSHIP_TYPES.image}" Target="media/picture.png"/></Relationships>`,
  );
  const drawing =
    `<w:drawing><wp:inline><wp:extent cx="1000000" cy="500000"/><wp:docPr id="1" name="Picture 1"/>` +
    `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic>` +
    `<pic:nvPicPr><pic:cNvPr id="2" name="picture.png"/><pic:cNvPicPr/></pic:nvPicPr>` +
    `<pic:blipFill><a:blip r:embed="${IMAGE_RELATIONSHIP_ID}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
    `<pic:spPr><a:xfrm${xfrmAttributes(transform)}><a:off x="0" y="0"/><a:ext cx="1000000" cy="500000"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>`;
  zip.file(
    "word/document.xml",
    `${XML_DECLARATION}<w:document ${NS}><w:body><w:p><w:r>${drawing}</w:r></w:p><w:sectPr/></w:body></w:document>`,
  );
  zip.file("word/media/picture.png", PNG_1X1_BASE64, { base64: true });
  return zip.generateAsync({ type: "arraybuffer" });
};

/** The one image in the package, whichever run it landed in. */
const imageOf = (document: Document): Image => {
  const found: Image[] = [];
  const walk = (value: unknown, seen: WeakSet<object>): void => {
    if (Array.isArray(value)) {
      for (const item of value) {
        walk(item, seen);
      }
      return;
    }
    if (typeof value !== "object" || value === null || seen.has(value)) {
      return;
    }
    seen.add(value);
    const record: Record<string, unknown> = value;
    if (typeof record["rId"] === "string" && record["size"] !== undefined) {
      found.push(record as unknown as Image);
    }
    for (const item of Object.values(record)) {
      walk(item, seen);
    }
  };
  walk(document.package.document.content, new WeakSet());
  const image = found.at(0);
  if (image === undefined) {
    throw new Error("the parsed package holds no image");
  }
  return image;
};

/**
 * Force the save past replay, exactly as the `reserialize` corpus invariant
 * does: with the capture intact the picture is written back verbatim and the
 * serializer never runs.
 */
const forceRebuild = (image: Image): void => {
  const record: Record<string, unknown> = image;
  record["rawImageFingerprint"] = EDITED_PREVIEW_FINGERPRINT;
};

describe("an authored image transform survives a save", () => {
  test(
    "rotation and both flips come back exactly as authored",
    async () => {
      await fc.assert(
        fc.asyncProperty(authoredTransform, async (transform) => {
          const parsed = await parseDocx(await docxWith(transform), { preloadFonts: false });
          const authored = imageOf(parsed);
          const expected = structuredClone(authored.transform);
          forceRebuild(authored);

          const saved = await repackDocx(parsed, { updateModifiedDate: false });
          const reparsed = await parseDocx(saved, { preloadFonts: false });

          expect(imageOf(reparsed).transform).toEqual(expected);
        }),
        propertyConfig({ numRuns: 45 }),
      );
    },
    propertyTestTimeout(),
  );

  test("an authored zero rotation reaches the model", async () => {
    const parsed = await parseDocx(await docxWith({ rot: "0", flipH: "0", flipV: undefined }), {
      preloadFonts: false,
    });

    expect(imageOf(parsed).transform).toEqual({ rotation: 0, flipH: false });
  });
});
