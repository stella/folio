/**
 * An edited drawing is rebuilt from the model, and the model has to hold the
 * anchor.
 *
 * folio replays a drawing's captured bytes while a fingerprint says the model
 * still agrees with them, so an untouched document round-trips a `wp:anchor`
 * whatever the model holds. Resize the picture and the fingerprint stops
 * matching, `serializeDrawingContent` runs, and what the model does not hold is
 * gone. It held neither half of the anchor's identity:
 *
 * - `@simplePos`, `@relativeHeight`, `@locked` were written as constants and
 *   `@hidden` not at all, so two pictures a document deliberately stacked came
 *   back on one layer and a hidden one became visible;
 * - `a:hlinkClick` kept its `r:id` and lost `tooltip`, `tgtFrame`, `history`
 *   and the rest of `CT_Hyperlink`, and `a:hlinkHover` was not read anywhere in
 *   the repository.
 *
 * The container survival census cannot see any of this. Its fixtures carry no
 * `a:blip`, so every drawing it builds has no picture relationship, takes the
 * preserve-only replay path, and comes back byte for byte whatever the
 * serializer would have written. That is why this is a property here rather
 * than a line in the baseline.
 */

import { describe, expect, test } from "bun:test";
import fc from "fast-check";
import JSZip from "jszip";

import { propertyConfig, propertyTestTimeout } from "../../../../test/property-testing";

import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import type { Document } from "../types/document";
import { parseDocx } from "./parser";
import { createEmptyDocx, repackDocx } from "./rezip";

const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const W_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R_NAMESPACE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const A_NAMESPACE = "http://schemas.openxmlformats.org/drawingml/2006/main";
const WP_NAMESPACE = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const PIC_NAMESPACE = "http://schemas.openxmlformats.org/drawingml/2006/picture";
const WPS_NAMESPACE = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";

/** A 1×1 PNG, so the drawing has a real picture relationship to resolve. */
const PNG = Buffer.from(
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==",
  "base64",
);

const IMAGE_RID = "rIdPicture";
const LINK_RID = "rIdLink";

const GRAPHIC =
  `<a:graphic><a:graphicData uri="${PIC_NAMESPACE}">` +
  '<pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="p"/><pic:cNvPicPr/></pic:nvPicPr>' +
  `<pic:blipFill><a:blip r:embed="${IMAGE_RID}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
  '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm>' +
  '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>' +
  "</a:graphicData></a:graphic>";

const anchorXml = (anchorAttributes: string, docPrChildren: string): string =>
  `<w:p><w:r><w:drawing><wp:anchor ${anchorAttributes}>` +
  '<wp:simplePos x="111" y="222"/>' +
  '<wp:positionH relativeFrom="margin"><wp:posOffset>5</wp:posOffset></wp:positionH>' +
  '<wp:positionV relativeFrom="margin"><wp:posOffset>6</wp:posOffset></wp:positionV>' +
  '<wp:extent cx="914400" cy="914400"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>' +
  `<wp:docPr id="1" name="p">${docPrChildren}</wp:docPr>` +
  `${GRAPHIC}</wp:anchor></w:drawing></w:r></w:p>`;

const buildDocx = async (body: string): Promise<ArrayBuffer> => {
  const zip = await JSZip.loadAsync(await createEmptyDocx());
  zip.file(
    "word/document.xml",
    `${XML_DECLARATION}<w:document xmlns:w="${W_NAMESPACE}" xmlns:r="${R_NAMESPACE}" ` +
      `xmlns:a="${A_NAMESPACE}" xmlns:wp="${WP_NAMESPACE}" xmlns:pic="${PIC_NAMESPACE}" ` +
      `xmlns:wps="${WPS_NAMESPACE}">` +
      `<w:body>${body}<w:sectPr/></w:body></w:document>`,
  );
  zip.file("word/media/image1.png", PNG);
  const types = await zip.file("[Content_Types].xml")?.async("text");
  const rels = await zip.file("word/_rels/document.xml.rels")?.async("text");
  if (types === undefined || rels === undefined) {
    throw new Error("the empty package lost its packaging parts");
  }
  zip.file(
    "[Content_Types].xml",
    types.replace("</Types>", '<Default Extension="png" ContentType="image/png"/></Types>'),
  );
  zip.file(
    "word/_rels/document.xml.rels",
    rels.replace(
      "</Relationships>",
      `<Relationship Id="${IMAGE_RID}" Type="${R_NAMESPACE}/image" Target="media/image1.png"/>` +
        `<Relationship Id="${LINK_RID}" Type="${R_NAMESPACE}/hyperlink" Target="https://example.org/" TargetMode="External"/>` +
        "</Relationships>",
    ),
  );
  return zip.generateAsync({ type: "arraybuffer" });
};

/**
 * Resize every picture, which is what an editor does and what stops the
 * captured bytes from being replayed. Nothing else forces the serializer on a
 * document that was only opened.
 */
const resizePictures = (document: Document): void => {
  const visit = (value: unknown, seen: WeakSet<object>): void => {
    if (Array.isArray(value)) {
      for (const item of value) {
        visit(item, seen);
      }
      return;
    }
    if (typeof value !== "object" || value === null || seen.has(value)) {
      return;
    }
    seen.add(value);
    const record = value as Record<string, unknown>;
    if (record["type"] === "image" && typeof record["size"] === "object") {
      record["size"] = { width: 457_200, height: 457_200 };
    }
    for (const item of Object.values(record)) {
      visit(item, seen);
    }
  };
  visit(document.package, new WeakSet());
};

const savedAfterResize = async (body: string): Promise<string> => {
  const parsed = await parseDocx(await buildDocx(body), { preloadFonts: false });
  resizePictures(parsed);
  const saved = await repackDocx(parsed, { updateModifiedDate: false });
  return (await (await JSZip.loadAsync(saved)).file("word/document.xml")?.async("text")) ?? "";
};

/** `CT_Anchor`'s own attributes, with a value no folio constant would produce. */
const ANCHOR_ATTRIBUTES = {
  simplePos: "1",
  relativeHeight: "7",
  locked: "1",
  hidden: "1",
  behindDoc: "1",
  layoutInCell: "0",
  allowOverlap: "0",
  distT: "1",
  distB: "2",
  distL: "3",
  distR: "4",
} as const;

const ANCHOR_ATTRIBUTE_XML = Object.entries(ANCHOR_ATTRIBUTES)
  .map(([name, value]) => `${name}="${value}"`)
  .join(" ");

const ANCHOR_ATTRIBUTE_NAMES = Object.keys(ANCHOR_ATTRIBUTES);

const HLINK_CLICK =
  `<a:hlinkClick r:id="${LINK_RID}" tooltip="a tip" tgtFrame="_blank" history="0">` +
  '<a:extLst><a:ext uri="{9DB03344-7A64-4A76-B8AB-CC76FC2DDDD4}"/></a:extLst></a:hlinkClick>';
const HLINK_HOVER = `<a:hlinkHover r:id="${LINK_RID}" tooltip="on hover"/>`;

describe("a resized drawing is rebuilt with the anchor it was authored with", () => {
  test("every CT_Anchor attribute comes back as the author stated it", async () => {
    const saved = await savedAfterResize(anchorXml(ANCHOR_ATTRIBUTE_XML, HLINK_CLICK));
    await fc.assert(
      fc.asyncProperty(fc.constantFrom(...ANCHOR_ATTRIBUTE_NAMES), (name) => {
        const written = `${name}="${ANCHOR_ATTRIBUTES[name as keyof typeof ANCHOR_ATTRIBUTES]}"`;
        expect({ name, present: saved.includes(written) }).toEqual({ name, present: true });
      }),
      propertyConfig({ numRuns: ANCHOR_ATTRIBUTE_NAMES.length }),
    );
    expect(saved).toContain('<wp:simplePos x="111" y="222"/>');
  });

  test("both wp:docPr links come back whole", async () => {
    const saved = await savedAfterResize(
      anchorXml(ANCHOR_ATTRIBUTE_XML, `${HLINK_CLICK}${HLINK_HOVER}`),
    );

    expect(saved).toContain('tooltip="a tip"');
    expect(saved).toContain('tgtFrame="_blank"');
    expect(saved).toContain("{9DB03344-7A64-4A76-B8AB-CC76FC2DDDD4}");
    expect(saved).toContain(HLINK_HOVER);
    // The click element still precedes the hover one, which is the order
    // `CT_NonVisualDrawingProps` declares and the only one Word accepts.
    expect(saved.indexOf("a:hlinkClick")).toBeLessThan(saved.indexOf("a:hlinkHover"));
  });

  test("a link folio refuses is not replayed from the source bytes", async () => {
    const zip = await JSZip.loadAsync(
      await buildDocx(anchorXml(ANCHOR_ATTRIBUTE_XML, HLINK_CLICK)),
    );
    const rels = await zip.file("word/_rels/document.xml.rels")?.async("text");
    zip.file(
      "word/_rels/document.xml.rels",
      (rels ?? "").replace("https://example.org/", "javascript:alert(1)"),
    );
    const parsed = await parseDocx(await zip.generateAsync({ type: "arraybuffer" }), {
      preloadFonts: false,
    });
    resizePictures(parsed);
    const saved =
      (await (
        await JSZip.loadAsync(await repackDocx(parsed, { updateModifiedDate: false }))
      )
        .file("word/document.xml")
        ?.async("text")) ?? "";

    expect(saved).not.toContain("a:hlinkClick");
    expect(saved).not.toContain('tooltip="a tip"');
  });
});

/**
 * The same `CT_Anchor` under each graphic folio models.
 *
 * A picture, a shape and a text box hang off one element, and the rebuild had
 * two writers for it: the picture path wrote what the author stated, the shape
 * path wrote `simplePos="0" relativeHeight="251658240" locked="0"
 * layoutInCell="1" allowOverlap="1"` and `<wp:simplePos x="0" y="0"/>` for
 * every shape and every text box it rebuilt. The property is over subsets
 * because a writer that states all six from constants passes any check that
 * only ever authors all six.
 */
const SHAPE_GRAPHIC =
  `<a:graphic><a:graphicData uri="${WPS_NAMESPACE}">` +
  '<wps:wsp><wps:cNvPr id="2" name="s"/><wps:cNvSpPr/>' +
  '<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm>' +
  '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr>' +
  "<wps:bodyPr/></wps:wsp></a:graphicData></a:graphic>";

const TEXT_BOX_GRAPHIC =
  `<a:graphic><a:graphicData uri="${WPS_NAMESPACE}">` +
  '<wps:wsp><wps:cNvPr id="3" name="t"/><wps:cNvSpPr txBox="1"/>' +
  '<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm>' +
  '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr>' +
  "<wps:txbx><w:txbxContent><w:p><w:r><w:t>box</w:t></w:r></w:p></w:txbxContent></wps:txbx>" +
  "<wps:bodyPr/></wps:wsp></a:graphicData></a:graphic>";

const GRAPHIC_BY_KIND = {
  image: GRAPHIC,
  shape: SHAPE_GRAPHIC,
  textBox: TEXT_BOX_GRAPHIC,
} as const;

type DrawingKind = keyof typeof GRAPHIC_BY_KIND;

const DRAWING_KINDS = Object.keys(GRAPHIC_BY_KIND) as DrawingKind[];

const anchorOf = (kind: DrawingKind, attributes: string): string =>
  `<w:p><w:r><w:drawing><wp:anchor ${attributes}>` +
  '<wp:simplePos x="111" y="222"/>' +
  '<wp:positionH relativeFrom="margin"><wp:posOffset>5</wp:posOffset></wp:positionH>' +
  '<wp:positionV relativeFrom="margin"><wp:posOffset>6</wp:posOffset></wp:positionV>' +
  '<wp:extent cx="914400" cy="914400"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>' +
  '<wp:docPr id="1" name="p"/>' +
  `${GRAPHIC_BY_KIND[kind]}</wp:anchor></w:drawing></w:r></w:p>`;

/** Every attribute the anchor models, with a value no folio constant produces. */
const AUTHORED_ANCHOR = {
  simplePos: "1",
  relativeHeight: "7",
  locked: "1",
  hidden: "1",
  layoutInCell: "0",
  allowOverlap: "0",
} as const;

type AuthoredAnchorName = keyof typeof AUTHORED_ANCHOR;

const AUTHORED_ANCHOR_NAMES = Object.keys(AUTHORED_ANCHOR) as AuthoredAnchorName[];

/**
 * `CT_Anchor` requires all but `@hidden`, so an attribute the property leaves
 * unauthored is still written — with the spec default, which is what the
 * fixture states for it.
 */
const SPEC_DEFAULT_ANCHOR = {
  simplePos: "0",
  relativeHeight: "0",
  locked: "0",
  hidden: undefined,
  layoutInCell: "1",
  allowOverlap: "1",
} as const satisfies Record<AuthoredAnchorName, string | undefined>;

const anchorAttributeXml = (authored: ReadonlySet<AuthoredAnchorName>): string =>
  AUTHORED_ANCHOR_NAMES.flatMap((name) => {
    const value = authored.has(name) ? AUTHORED_ANCHOR[name] : SPEC_DEFAULT_ANCHOR[name];
    return value === undefined ? [] : [`${name}="${value}"`];
  }).join(" ");

const savedAfterEdit = async (body: string, viaEditor: boolean): Promise<string> => {
  const parsed = await parseDocx(await buildDocx(body), { preloadFonts: false });
  // The editor's own leg, which is the path every edited document takes.
  const projected = viaEditor ? fromProseDoc(toProseDoc(parsed), parsed) : parsed;
  resizePictures(projected);
  const saved = await repackDocx(projected, { updateModifiedDate: false });
  return (await (await JSZip.loadAsync(saved)).file("word/document.xml")?.async("text")) ?? "";
};

describe("every drawing kind is rebuilt with the anchor it was authored with", () => {
  test(
    "an authored subset survives under a picture, a shape and a text box",
    async () => {
      await fc.assert(
        fc.asyncProperty(
          fc.constantFrom(...DRAWING_KINDS),
          fc.subarray(AUTHORED_ANCHOR_NAMES),
          fc.boolean(),
          async (kind, authoredNames, viaEditor) => {
            const authored = new Set(authoredNames);
            const saved = await savedAfterEdit(
              anchorOf(kind, anchorAttributeXml(authored)),
              viaEditor,
            );
            const missing = [...authored].filter(
              (name) => !saved.includes(`${name}="${AUTHORED_ANCHOR[name]}"`),
            );
            expect({ kind, viaEditor, missing }).toEqual({ kind, viaEditor, missing: [] });
            expect(saved).toContain('<wp:simplePos x="111" y="222"/>');
          },
        ),
        propertyConfig({ numRuns: 30 }),
      );
    },
    propertyTestTimeout(30_000),
  );
});
