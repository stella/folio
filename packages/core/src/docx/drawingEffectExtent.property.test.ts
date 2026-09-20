/**
 * `wp:effectExtent` is declared twice, and a rebuild has to write both back.
 *
 * `CT_Inline` and `CT_Anchor` each declare one, and so do `CT_WrapSquare` and
 * `CT_WrapTopBottom`. They are two values: the drawing's is the object's own
 * effect reservation, the wrap child's is the reservation the text flow is
 * computed against. folio read only the drawing's, into `Image.padding`, and
 * wrote it back on the drawing, so a wrap child's own reservation survived an
 * untouched save on the strength of its captured bytes and vanished the moment
 * anything forced the serializer.
 *
 * The property is over the product of the two anchorings, the two wrap kinds
 * that declare a reservation, and which elements authored one, because a writer
 * that collects both onto the drawing passes any check that only ever authors
 * the drawing's.
 */

import { describe, expect, test } from "bun:test";
import fc from "fast-check";
import JSZip from "jszip";

import { propertyConfig } from "../../../../test/property-testing";

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

/** A 1×1 PNG, so the drawing has a real picture relationship to resolve. */
const PNG = Buffer.from(
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==",
  "base64",
);

const IMAGE_RID = "rIdPicture";

const GRAPHIC =
  `<a:graphic><a:graphicData uri="${PIC_NAMESPACE}">` +
  '<pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="p"/><pic:cNvPicPr/></pic:nvPicPr>' +
  `<pic:blipFill><a:blip r:embed="${IMAGE_RID}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
  '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm>' +
  '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>' +
  "</a:graphicData></a:graphic>";

/** Four sides no other slot in the fixture states, so a match names its carrier. */
const DRAWING_EXTENT = { l: 12_700, t: 19_050, r: 6350, b: 25_400 } as const;
const WRAP_CHILD_EXTENT = { l: 1270, t: 2540, r: 3810, b: 5080 } as const;

type Extent = { l: number; t: number; r: number; b: number };

const extentXml = ({ l, t, r, b }: Extent): string =>
  `<wp:effectExtent l="${l}" t="${t}" r="${r}" b="${b}"/>`;

/** Which elements the fixture authors a reservation on. */
const AUTHORED_ON = ["drawing", "wrapChild", "both"] as const;
type AuthoredOn = (typeof AUTHORED_ON)[number];

/** The `EG_WrapType` members whose own type declares a `wp:effectExtent`. */
const WRAP_KINDS = ["square", "topAndBottom"] as const;
type WrapKind = (typeof WRAP_KINDS)[number];

const WRAP_ELEMENT_NAMES = {
  square: "wrapSquare",
  topAndBottom: "wrapTopAndBottom",
} as const satisfies Record<WrapKind, string>;

const drawingExtentOf = (authoredOn: AuthoredOn): string =>
  authoredOn === "wrapChild" ? "" : extentXml(DRAWING_EXTENT);

const wrapChildXml = (kind: WrapKind, authoredOn: AuthoredOn): string => {
  const name = WRAP_ELEMENT_NAMES[kind];
  const attrs = kind === "square" ? ' wrapText="bothSides"' : "";
  return authoredOn === "drawing"
    ? `<wp:${name}${attrs}/>`
    : `<wp:${name}${attrs}>${extentXml(WRAP_CHILD_EXTENT)}</wp:${name}>`;
};

const inlineBody = (authoredOn: AuthoredOn): string =>
  "<w:p><w:r><w:drawing><wp:inline>" +
  '<wp:extent cx="914400" cy="914400"/>' +
  `${drawingExtentOf(authoredOn)}<wp:docPr id="1" name="p"/>${GRAPHIC}` +
  "</wp:inline></w:drawing></w:r></w:p>";

const anchorBody = (kind: WrapKind, authoredOn: AuthoredOn): string =>
  '<w:p><w:r><w:drawing><wp:anchor simplePos="0" relativeHeight="7" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">' +
  '<wp:simplePos x="0" y="0"/>' +
  '<wp:positionH relativeFrom="margin"><wp:posOffset>5</wp:posOffset></wp:positionH>' +
  '<wp:positionV relativeFrom="margin"><wp:posOffset>6</wp:posOffset></wp:positionV>' +
  '<wp:extent cx="914400" cy="914400"/>' +
  `${drawingExtentOf(authoredOn)}${wrapChildXml(kind, authoredOn)}` +
  `<wp:docPr id="1" name="p"/>${GRAPHIC}` +
  "</wp:anchor></w:drawing></w:r></w:p>";

const buildDocx = async (body: string): Promise<ArrayBuffer> => {
  const zip = await JSZip.loadAsync(await createEmptyDocx());
  zip.file(
    "word/document.xml",
    `${XML_DECLARATION}<w:document xmlns:w="${W_NAMESPACE}" xmlns:r="${R_NAMESPACE}" ` +
      `xmlns:a="${A_NAMESPACE}" xmlns:wp="${WP_NAMESPACE}" xmlns:pic="${PIC_NAMESPACE}">` +
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
        "</Relationships>",
    ),
  );
  return zip.generateAsync({ type: "arraybuffer" });
};

/**
 * Resize every picture, which is what stops the captured bytes from being
 * replayed. Nothing else forces the serializer on a document that was opened
 * and saved, and replayed bytes prove nothing about what the model holds.
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

const partOf = async (saved: ArrayBuffer): Promise<string> =>
  (await (await JSZip.loadAsync(saved)).file("word/document.xml")?.async("text")) ?? "";

/** The two legs the container survival law runs: a forced save, and the editor. */
const LEGS = ["save", "editor"] as const;
type Leg = (typeof LEGS)[number];

const savedPart = async (body: string, leg: Leg): Promise<string> => {
  const parsed = await parseDocx(await buildDocx(body), { preloadFonts: false });
  const document = leg === "editor" ? fromProseDoc(toProseDoc(parsed), parsed) : parsed;
  resizePictures(document);
  return partOf(await repackDocx(document, { updateModifiedDate: false }));
};

/** The reservation on `wp:inline`/`wp:anchor`: the one the extent precedes. */
const drawingExtentIn = (part: string): Extent | undefined =>
  extentIn(/<wp:extent [^>]*\/>(<wp:effectExtent [^>]*\/>)/u.exec(part)?.[1]);

/** The reservation inside the wrap child, wherever the save put it. */
const wrapChildExtentIn = (part: string, kind: WrapKind): Extent | undefined => {
  const name = WRAP_ELEMENT_NAMES[kind];
  const element = new RegExp(`<wp:${name}[^>]*>(<wp:effectExtent [^>]*/>)</wp:${name}>`, "u").exec(
    part,
  );
  return extentIn(element?.[1]);
};

const extentIn = (xml: string | undefined): Extent | undefined => {
  if (xml === undefined) {
    return undefined;
  }
  const side = (name: string): number =>
    Number(new RegExp(`${name}="(-?\\d+)"`, "u").exec(xml)?.[1] ?? Number.NaN);
  return { l: side("l"), t: side("t"), r: side("r"), b: side("b") };
};

describe("a rebuilt drawing states each wp:effectExtent on the element that authored it", () => {
  test("an anchor and its wrap child keep their own reservations", async () => {
    await fc.assert(
      fc.asyncProperty(
        fc.constantFrom(...WRAP_KINDS),
        fc.constantFrom(...AUTHORED_ON),
        fc.constantFrom(...LEGS),
        async (kind, authoredOn, leg) => {
          const part = await savedPart(anchorBody(kind, authoredOn), leg);
          expect({ kind, authoredOn, leg, drawing: drawingExtentIn(part) }).toEqual({
            kind,
            authoredOn,
            leg,
            // A drawing that authored none still states one: `CT_EffectExtent`
            // requires all four sides and an omitted element reads as zero, so
            // the two spellings are the same document.
            drawing: authoredOn === "wrapChild" ? { l: 0, t: 0, r: 0, b: 0 } : DRAWING_EXTENT,
          });
          expect({ kind, authoredOn, leg, wrapChild: wrapChildExtentIn(part, kind) }).toEqual({
            kind,
            authoredOn,
            leg,
            wrapChild: authoredOn === "drawing" ? undefined : WRAP_CHILD_EXTENT,
          });
        },
      ),
      propertyConfig({ numRuns: WRAP_KINDS.length * AUTHORED_ON.length * LEGS.length }),
    );
  });

  test("an inline drawing keeps the only reservation its type declares", async () => {
    await fc.assert(
      fc.asyncProperty(fc.constantFrom(...LEGS), async (leg) => {
        const part = await savedPart(inlineBody("drawing"), leg);
        expect({ leg, drawing: drawingExtentIn(part) }).toEqual({ leg, drawing: DRAWING_EXTENT });
      }),
      propertyConfig({ numRuns: LEGS.length }),
    );
  });
});
