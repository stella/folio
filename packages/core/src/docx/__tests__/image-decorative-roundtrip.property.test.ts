/**
 * `wp:docPr` accessibility facts, over the whole space they live in.
 *
 * Two facts that had been one. `@hidden` says the drawing is not displayed;
 * the `{C183D7F6-B498-43B3-948B-1728B52AA6E4}` extension says it is displayed
 * and carries nothing a reader needs. Writing either as the other inverts what
 * the document said, and writing one anchoring's attribute and not the other's
 * loses it the moment an inline image is anchored.
 *
 * The shape here is Word's, read from the corpus rather than from a spec: the
 * header of `libreoffice-core/sw/qa/extras/ooxmlimport/data/tdf120547.docx`
 * and the body of `.../ooxmlexport/data/image_through_shape.docx` both carry
 * `<a:ext uri="{C183D7F6-…}"><adec:decorative … val="1"/></a:ext>`, the second
 * beside an unrelated `{28A0092B-C50C-407E-A947-70E740481C1C}` sibling. That
 * sibling is why the property carries an unmodeled extension: an extension
 * folio has no opinion about must come back byte for byte.
 *
 * The round trip is the real serializer, not a verbatim replay: the model is
 * re-serialized from its own fields, which is the path every edited document
 * takes.
 */

import { describe, expect, test } from "bun:test";
import fc from "fast-check";

import { propertyConfig } from "../../../../../test/property-testing";

import { fromProseDoc } from "../../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../../prosemirror/conversion/toProseDoc";
import type { Document, Image, Run } from "../../types/document";
import { parseDrawing } from "../imageParser";
import { serializeRun } from "../serializer/runSerializer";
import { parseXml, type XmlElement } from "../xmlParser";

const NS = [
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"',
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"',
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"',
  'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"',
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"',
].join(" ");

const DECORATIVE_URI = "{C183D7F6-B498-43B3-948B-1728B52AA6E4}";
const DECORATIVE_NS = "http://schemas.microsoft.com/office/drawing/2017/decorative";
/** `a16:creationId`, the sibling extension Word writes beside it. */
const CREATION_ID_URI = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}";
const CREATION_ID_EXT =
  `<a:ext uri="${CREATION_ID_URI}">` +
  '<a16:creationId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" id="{1D4E2F17-72B9-4A0C-9B0F-9C3E0A47D0F2}"/>' +
  "</a:ext>";

type Anchoring = "inline" | "anchor";

type DocPrFacts = {
  anchoring: Anchoring;
  decorative: boolean | undefined;
  hidden: boolean | undefined;
  alt: string | undefined;
  otherExtension: boolean;
};

const decorativeExt = (value: boolean) =>
  `<a:ext uri="${DECORATIVE_URI}"><adec:decorative xmlns:adec="${DECORATIVE_NS}" val="${value ? "1" : "0"}"/></a:ext>`;

const drawingXml = ({ anchoring, decorative, hidden, alt, otherExtension }: DocPrFacts): string => {
  const open =
    anchoring === "inline"
      ? "<wp:inline>"
      : '<wp:anchor simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">';
  const position =
    anchoring === "anchor"
      ? '<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>'
      : "";
  const wrap = anchoring === "anchor" ? '<wp:wrapSquare wrapText="bothSides"/>' : "";
  const hiddenAttr = hidden === undefined ? "" : ` hidden="${hidden ? "1" : "0"}"`;
  const descr = alt === undefined ? "" : ` descr="${alt}"`;
  const extensions = `${decorative === undefined ? "" : decorativeExt(decorative)}${
    otherExtension ? CREATION_ID_EXT : ""
  }`;
  const docPr =
    extensions === ""
      ? `<wp:docPr id="7" name="Picture 7"${descr}${hiddenAttr}/>`
      : `<wp:docPr id="7" name="Picture 7"${descr}${hiddenAttr}><a:extLst>${extensions}</a:extLst></wp:docPr>`;

  return (
    `<w:drawing ${NS}>${open}${position}<wp:extent cx="914400" cy="457200"/>${wrap}${docPr}` +
    '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic>' +
    '<pic:nvPicPr><pic:cNvPr id="7" name="media.png"/><pic:cNvPicPr/></pic:nvPicPr>' +
    '<pic:blipFill><a:blip r:embed="rId1"/></pic:blipFill>' +
    '<pic:spPr><a:xfrm><a:ext cx="914400" cy="457200"/></a:xfrm></pic:spPr>' +
    `</pic:pic></a:graphicData></a:graphic></wp:${anchoring}></w:drawing>`
  );
};

const parseDrawingXml = (xml: string): Image | null => {
  const drawing = (parseXml(xml).elements as XmlElement[]).at(0);
  return drawing ? parseDrawing(drawing, undefined, undefined) : null;
};

/** The real serializer, then the parser: the path an edited document takes. */
const reserialize = (image: Image): Image | null => {
  const run: Run = { type: "run", content: [{ type: "drawing", image }] };
  const root = (parseXml(`<root ${NS}>${serializeRun(run)}</root>`).elements as XmlElement[]).at(0);
  const drawing = (root?.elements as XmlElement[] | undefined)
    ?.at(0)
    ?.elements?.find((child) => child.type === "element");
  return drawing ? parseDrawing(drawing, undefined, undefined) : null;
};

const documentWithImage = (image: Image): Document => ({
  package: {
    document: {
      content: [
        { type: "paragraph", content: [{ type: "run", content: [{ type: "drawing", image }] }] },
      ],
    },
  },
});

const firstImage = (document: Document): Image | null => {
  const paragraph = document.package.document.content.at(0);
  if (paragraph?.type !== "paragraph") {
    return null;
  }
  const run = paragraph.content.at(0);
  if (run?.type !== "run") {
    return null;
  }
  const drawing = run.content.at(0);
  return drawing?.type === "drawing" ? drawing.image : null;
};

const editorRoundTrip = (image: Image): Image | null =>
  firstImage(fromProseDoc(toProseDoc(documentWithImage(image)), documentWithImage(image)));

/** The facts, as the model should hold them. */
const factsOf = (image: Image | null) => ({
  decorative: image?.decorative,
  hidden: image?.hidden,
  alt: image?.alt,
  docPrExtensions: image?.docPrExtensions,
});

const factsArbitrary = fc.record({
  anchoring: fc.constantFrom<Anchoring>("inline", "anchor"),
  decorative: fc.constantFrom(undefined, false, true),
  hidden: fc.constantFrom(undefined, false, true),
  alt: fc.constantFrom(undefined, "A photograph of the signing"),
  otherExtension: fc.boolean(),
});

describe("wp:docPr decorative and hidden", () => {
  test("Word's decorative extension parses, beside an unrelated sibling", () => {
    const image = parseDrawingXml(
      drawingXml({
        anchoring: "inline",
        decorative: true,
        hidden: undefined,
        alt: undefined,
        otherExtension: true,
      }),
    );
    expect(image?.decorative).toBe(true);
    expect(image?.hidden).toBeUndefined();
    expect(image?.docPrExtensions).toEqual([CREATION_ID_EXT]);
  });

  test("a decorative image is not written as a hidden one", () => {
    const image = parseDrawingXml(
      drawingXml({
        anchoring: "inline",
        decorative: true,
        hidden: undefined,
        alt: undefined,
        otherExtension: false,
      }),
    );
    expect(image).not.toBeNull();
    if (!image) {
      return;
    }
    const xml = serializeRun({ type: "run", content: [{ type: "drawing", image }] });
    expect(xml).not.toContain("hidden=");
    expect(xml).toContain(`<a:ext uri="${DECORATIVE_URI}">`);
  });

  test("a hidden image is not read as a decorative one", () => {
    const image = parseDrawingXml(
      drawingXml({
        anchoring: "inline",
        decorative: undefined,
        hidden: true,
        alt: undefined,
        otherExtension: false,
      }),
    );
    expect(image?.hidden).toBe(true);
    expect(image?.decorative).toBeUndefined();
  });

  test("every fact survives a real save and the editor round trip", () => {
    fc.assert(
      fc.property(factsArbitrary, (facts) => {
        const parsed = parseDrawingXml(drawingXml(facts));
        expect(parsed).not.toBeNull();
        if (!parsed) {
          return;
        }

        const expected = {
          decorative: facts.decorative,
          hidden: facts.hidden,
          alt: facts.alt,
          docPrExtensions: facts.otherExtension ? [CREATION_ID_EXT] : undefined,
        };
        expect(factsOf(parsed)).toEqual(expected);
        expect(factsOf(reserialize(parsed))).toEqual(expected);
        expect(factsOf(editorRoundTrip(parsed))).toEqual(expected);
        // And through the editor and out to XML again, which is the whole path
        // an edited image takes.
        expect(factsOf(reserialize(editorRoundTrip(parsed) ?? parsed))).toEqual(expected);
      }),
      propertyConfig({ numRuns: 200 }),
    );
  });
});
