/**
 * A rebuilt drawing keeps the wrap outline and the wrap insets it was authored
 * with.
 *
 * `serializeWrap` wrote one constant for every tight and through wrap — the
 * four corners of a 21600-unit box, `edited="0"` — and nothing read the
 * authored `wp:wrapPolygon`, so any outline a document traced around a picture
 * came back as a rectangle and the text reflowed. The wrap insets had the
 * milder form of the same defect: `ImageWrap` held one set, the rebuild wrote
 * it on `wp:anchor`, and a document that stated an inset on the wrap child got
 * it back on the drawing.
 *
 * The property runs the class the corpus cannot enumerate — arbitrary point
 * lists over every wrap kind that carries a polygon, and every placement of
 * each inset the wrap child's own type declares — through both paths a saved
 * drawing takes: the serializer (an edited document is rebuilt from the model)
 * and the ProseMirror projection.
 */

import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import { propertyConfig, propertyTestTimeout } from "../../../../test/property-testing";

import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import type { Document, Image, ImageWrap, ImageWrapPolygon } from "../types/document";
import { parseDrawing } from "./imageParser";
import { serializeRun } from "./serializer/runSerializer";
import { parseXml, type XmlElement } from "./xmlParser";

setDefaultTimeout(propertyTestTimeout(30_000));

const NS = [
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"',
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"',
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"',
  'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"',
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"',
].join(" ");

/** The wrap element each kind is written as, and the insets its type declares. */
const WRAP_KINDS = {
  square: { element: "wrapSquare", insets: ["distT", "distB", "distL", "distR"] },
  tight: { element: "wrapTight", insets: ["distL", "distR"] },
  through: { element: "wrapThrough", insets: ["distL", "distR"] },
  topAndBottom: { element: "wrapTopAndBottom", insets: ["distT", "distB"] },
} as const;

type WrapKind = keyof typeof WRAP_KINDS;

const WRAP_KIND_NAMES = Object.keys(WRAP_KINDS) as WrapKind[];

type InsetKey = "distT" | "distB" | "distL" | "distR";

const INSET_KEYS: readonly InsetKey[] = ["distT", "distB", "distL", "distR"];

/** Where an inset was stated. Both is legal and common: the two differ. */
type Placement = "none" | "drawing" | "wrapChild" | "both";

type Authored = {
  kind: WrapKind;
  /** The value each inset is stated with, and the element it is stated on. */
  insets: Record<InsetKey, { drawing: number; wrapChild: number; placement: Placement }>;
  polygon: ImageWrapPolygon;
};

/** A wrap polygon's own coordinate space; an inset is EMU. */
const coordinate = fc.integer({ min: 0, max: 21_600 });
const point = fc.record({ x: coordinate, y: coordinate });

/** An inset reaches zero: an authored zero is a value, not an absence. */
const insetValue = fc.integer({ min: 0, max: 2_000_000 });

const authored = fc.record({
  kind: fc.constantFrom(...WRAP_KIND_NAMES),
  insets: fc.record({
    distT: fc.record({
      drawing: insetValue,
      wrapChild: insetValue,
      placement: fc.constantFrom<Placement>("none", "drawing", "wrapChild", "both"),
    }),
    distB: fc.record({
      drawing: insetValue,
      wrapChild: insetValue,
      placement: fc.constantFrom<Placement>("none", "drawing", "wrapChild", "both"),
    }),
    distL: fc.record({
      drawing: insetValue,
      wrapChild: insetValue,
      placement: fc.constantFrom<Placement>("none", "drawing", "wrapChild", "both"),
    }),
    distR: fc.record({
      drawing: insetValue,
      wrapChild: insetValue,
      placement: fc.constantFrom<Placement>("none", "drawing", "wrapChild", "both"),
    }),
  }),
  // 3..40 points: a `wp:start` and the two to thirty-nine `wp:lineTo` that
  // `CT_WrapPath` admits.
  polygon: fc.record({
    edited: fc.constantFrom(undefined, false, true),
    start: point,
    lineTo: fc.array(point, { minLength: 2, maxLength: 39 }),
  }),
});

/** The insets one element states: only those the wrap kind puts there. */
const statedOn = (
  { kind, insets }: Authored,
  carrier: "drawing" | "wrapChild",
): Partial<Record<InsetKey, number>> => {
  const declared: readonly InsetKey[] =
    carrier === "drawing" ? INSET_KEYS : WRAP_KINDS[kind].insets;
  const stated: Partial<Record<InsetKey, number>> = {};
  for (const key of declared) {
    const { placement } = insets[key];
    if (placement === carrier || placement === "both") {
      stated[key] = insets[key][carrier];
    }
  }
  return stated;
};

const insetAttributes = (stated: Partial<Record<InsetKey, number>>): string =>
  INSET_KEYS.flatMap((key) => {
    const value = stated[key];
    return value === undefined ? [] : [` ${key}="${value}"`];
  }).join("");

const polygonXml = ({ edited, start, lineTo }: ImageWrapPolygon): string =>
  `<wp:wrapPolygon${edited === undefined ? "" : ` edited="${edited ? "1" : "0"}"`}>` +
  `<wp:start x="${start.x}" y="${start.y}"/>` +
  lineTo.map(({ x, y }) => `<wp:lineTo x="${x}" y="${y}"/>`).join("") +
  "</wp:wrapPolygon>";

const drawingXml = (facts: Authored): string => {
  const { element } = WRAP_KINDS[facts.kind];
  const wrapText = facts.kind === "topAndBottom" ? "" : ' wrapText="bothSides"';
  const body =
    facts.kind === "tight" || facts.kind === "through" ? polygonXml(facts.polygon) : undefined;
  const wrapChild =
    body === undefined
      ? `<wp:${element}${wrapText}${insetAttributes(statedOn(facts, "wrapChild"))}/>`
      : `<wp:${element}${wrapText}${insetAttributes(statedOn(facts, "wrapChild"))}>${body}</wp:${element}>`;

  return (
    `<w:drawing ${NS}><wp:anchor simplePos="0" relativeHeight="1" behindDoc="0" locked="0" ` +
    `layoutInCell="1" allowOverlap="1"${insetAttributes(statedOn(facts, "drawing"))}>` +
    '<wp:simplePos x="0" y="0"/>' +
    '<wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>' +
    '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>' +
    `<wp:extent cx="914400" cy="457200"/>${wrapChild}` +
    '<wp:docPr id="7" name="Picture 7"/>' +
    '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic>' +
    '<pic:nvPicPr><pic:cNvPr id="7" name="media.png"/><pic:cNvPicPr/></pic:nvPicPr>' +
    '<pic:blipFill><a:blip r:embed="rId1"/></pic:blipFill>' +
    '<pic:spPr><a:xfrm><a:ext cx="914400" cy="457200"/></a:xfrm></pic:spPr>' +
    "</pic:pic></a:graphicData></a:graphic></wp:anchor></w:drawing>"
  );
};

const parseDrawingXml = (xml: string): Image | null => {
  const drawing = (parseXml(xml).elements as XmlElement[]).at(0);
  return drawing ? parseDrawing(drawing, undefined, undefined) : null;
};

const runXml = (image: Image): string =>
  serializeRun({ type: "run", content: [{ type: "drawing", image }] });

/** The real serializer, then the parser: the path an edited document takes. */
const reserialize = (image: Image): Image | null => {
  const root = (parseXml(`<root ${NS}>${runXml(image)}</root>`).elements as XmlElement[]).at(0);
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

const editorRoundTrip = (image: Image): Image | null => {
  const source = documentWithImage(image);
  return firstImage(fromProseDoc(toProseDoc(source), source));
};

const ROUND_TRIPS = {
  rebuild: reserialize,
  editor: (image: Image): Image | null => {
    const projected = editorRoundTrip(image);
    return projected === null ? null : reserialize(projected);
  },
} as const satisfies Record<string, (image: Image) => Image | null>;

const ROUND_TRIP_NAMES = Object.keys(ROUND_TRIPS) as (keyof typeof ROUND_TRIPS)[];

/** Everything the wrap says, as the model should hold it. */
const wrapFacts = (image: Image | null): Partial<ImageWrap> | null =>
  image === null
    ? null
    : {
        type: image.wrap.type,
        distT: image.wrap.distT,
        distB: image.wrap.distB,
        distL: image.wrap.distL,
        distR: image.wrap.distR,
        distanceSlots: image.wrap.distanceSlots,
        polygon: image.wrap.polygon,
      };

describe("a rebuilt drawing keeps its wrap outline and inset slots", () => {
  test("every authored polygon and inset placement survives both round trips", () => {
    fc.assert(
      fc.property(authored, fc.constantFrom(...ROUND_TRIP_NAMES), (facts, tripName): undefined => {
        const parsed = parseDrawingXml(drawingXml(facts));
        expect(parsed).not.toBeNull();
        if (parsed === null) {
          return;
        }

        if (facts.kind === "tight" || facts.kind === "through") {
          expect(parsed.wrap.polygon).toEqual(facts.polygon);
        }

        expect(wrapFacts(ROUND_TRIPS[tripName](parsed))).toEqual(wrapFacts(parsed));
      }),
      propertyConfig({ numRuns: 300 }),
    );
  });

  test("the outline reaches the markup, rather than the constant rectangle", () => {
    fc.assert(
      fc.property(authored, (facts): undefined => {
        fc.pre(facts.kind === "tight" || facts.kind === "through");
        const parsed = parseDrawingXml(drawingXml(facts));
        expect(parsed).not.toBeNull();
        if (parsed === null) {
          return;
        }

        const saved = runXml(parsed);
        expect(saved).toContain(polygonXml(facts.polygon));
      }),
      propertyConfig({ numRuns: 200 }),
    );
  });

  test("an inset the wrap child stated is written back on the wrap child", () => {
    fc.assert(
      fc.property(authored, (facts): undefined => {
        const stated = statedOn(facts, "wrapChild");
        fc.pre(Object.keys(stated).length > 0);
        const parsed = parseDrawingXml(drawingXml(facts));
        expect(parsed).not.toBeNull();
        if (parsed === null) {
          return;
        }

        const { element } = WRAP_KINDS[facts.kind];
        const wrapTag = runXml(parsed).match(new RegExp(`<wp:${element}[^>]*`, "u"))?.[0] ?? "";
        for (const [key, value] of Object.entries(stated)) {
          expect(wrapTag).toContain(`${key}="${value}"`);
        }
      }),
      propertyConfig({ numRuns: 200 }),
    );
  });
});
