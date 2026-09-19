/**
 * A drawing's authored EMUs survive the editor projection unchanged.
 *
 * `toProseDoc` measures a drawing in pixels, because that is what the editor
 * and the node views draw in, and `fromProseDoc` converted those pixels back
 * to EMUs. Neither conversion is exact — a size rounds to whole pixels, a
 * stroke or an inset to two — so opening a document and saving it again moved
 * every size, stroke width, wrap inset and text-box margin a few hundred EMU
 * off the number its author wrote (`shape.size.height: N became N`,
 * `shape.outline.width: N became N`, `shape.textBody.margins.bottom: N became
 * N` in the corpus census). Two values fared worse: a shape with an `a:ln` but
 * no `@w` acquired 9525 EMU from the node's default pixel width
 * (`shape.outline.width: absent became N`), and a text box's authored inset of
 * zero was dropped by a truthiness test (`shape.wrap.distB: N became absent`).
 *
 * The property runs the input class the corpus cannot enumerate — arbitrary
 * EMU sizes, stroke widths, wrap insets and margins, over all three drawing
 * nodes — through `toProseDoc → fromProseDoc` and demands the numbers back
 * exactly, while a genuine editor change to the pixel attributes still reaches
 * the document.
 *
 * Values under the projection's own resolution (a size below one pixel, a
 * stroke below a hundredth of one) are out of range: they project to zero and
 * no authoring tool writes them.
 */

import { describe, expect, test } from "bun:test";
import fc from "fast-check";
import { Fragment, type Node as PMNode } from "prosemirror-model";

import { propertyConfig } from "../../../../../test/property-testing";
import type { Document, Image, ImageWrap, Shape } from "../../types/document";
import { emuToPixels, emuToStrokePixels, pixelsToEmu } from "../../utils/units";
import { fromProseDoc } from "./fromProseDoc";
import { toProseDoc } from "./toProseDoc";

/** One pixel, the smallest size the projection can carry. */
const EMU_PER_PIXEL = 9_525;

const authoredSize = fc.integer({ min: EMU_PER_PIXEL, max: 20_000_000 });

/** A hundredth of a pixel, the stroke projection's resolution. */
const authoredStrokeWidth = fc.integer({ min: Math.ceil(EMU_PER_PIXEL / 100), max: 2_000_000 });

/** Insets and margins reach zero: an authored zero is a value, not an absence. */
const authoredInset = fc.integer({ min: 0, max: 2_000_000 });

/**
 * Every route a drawing takes through the projection. A plain shape becomes a
 * `shape` node, a shape carrying a text body is extracted into a `textBox`
 * node, a picture becomes an `image` node. All three are rebuilt from pixels.
 */
const DRAWING_KINDS = ["rect", "textBox", "picture"] as const;
type DrawingKind = (typeof DRAWING_KINDS)[number];

const insets = fc.record({
  distT: authoredInset,
  distB: authoredInset,
  distL: authoredInset,
  distR: authoredInset,
});

const margins = fc.record({
  top: authoredInset,
  bottom: authoredInset,
  left: authoredInset,
  right: authoredInset,
});

type Insets = { distT: number; distB: number; distL: number; distR: number };
type Margins = { top: number; bottom: number; left: number; right: number };

type AuthoredGeometry = {
  kind: DrawingKind;
  width: number;
  height: number;
  outlineWidth?: number;
  insets: Insets;
  margins: Margins;
};

const authoredGeometry = fc.record({
  kind: fc.constantFrom(...DRAWING_KINDS),
  width: authoredSize,
  height: authoredSize,
  outlineWidth: fc.option(authoredStrokeWidth, { nil: undefined }),
  insets,
  margins,
});

const wrapOf = ({ distT, distB, distL, distR }: Insets): ImageWrap => ({
  type: "square",
  wrapText: "bothSides",
  distT,
  distB,
  distL,
  distR,
});

const textBody = {
  content: [
    {
      type: "paragraph" as const,
      content: [{ type: "run" as const, content: [{ type: "text" as const, text: "t" }] }],
    },
  ],
};

/** A 1x1 PNG, so the picture case has media to reference. */
const PNG_1X1_BASE64 =
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";

const buildShape = ({
  kind,
  width,
  height,
  outlineWidth,
  insets: i,
  margins: m,
}: AuthoredGeometry): Shape => ({
  type: "shape",
  shapeType: kind === "textBox" ? "textBox" : "rect",
  size: { width, height },
  wrap: wrapOf(i),
  outline: {
    style: "solid",
    color: { rgb: "112233" },
    ...(outlineWidth === undefined ? {} : { width: outlineWidth }),
  },
  ...(kind === "textBox" ? { textBody: { ...textBody, margins: m } } : {}),
});

const buildImage = ({ width, height, outlineWidth, insets: i }: AuthoredGeometry): Image => ({
  type: "image",
  rId: "rIdGeometry",
  size: { width, height },
  wrap: wrapOf(i),
  ...(outlineWidth === undefined
    ? {}
    : { outline: { style: "solid" as const, color: { rgb: "112233" }, width: outlineWidth } }),
});

const documentWithDrawing = (geometry: AuthoredGeometry): Document => ({
  package: {
    document: {
      content: [
        {
          type: "paragraph",
          content: [
            {
              type: "run",
              content: [
                geometry.kind === "picture"
                  ? { type: "drawing", image: buildImage(geometry) }
                  : { type: "shape", shape: buildShape(geometry) },
              ],
            },
          ],
        },
      ],
    },
    ...(geometry.kind === "picture"
      ? {
          media: new Map([
            [
              "word/media/image1.png",
              {
                path: "word/media/image1.png",
                filename: "image1.png",
                mimeType: "image/png",
                data: new ArrayBuffer(0),
                base64: PNG_1X1_BASE64,
              },
            ],
          ]),
          relationships: new Map([
            [
              "rIdGeometry",
              {
                id: "rIdGeometry",
                type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
                target: "media/image1.png",
              },
            ],
          ]),
        }
      : {}),
  },
});

/** The size, stroke width, wrap insets and margins of the drawing, in EMU. */
type RestoredGeometry = {
  size: { width: number; height: number };
  outlineWidth?: number;
  wrap?: ImageWrap;
  margins?: Partial<Margins>;
};

const geometryOf = (document: Document): RestoredGeometry => {
  for (const block of document.package.document.content) {
    if (block.type !== "paragraph") {
      continue;
    }
    for (const item of block.content) {
      if (item.type !== "run") {
        continue;
      }
      for (const content of item.content) {
        if (content.type === "shape") {
          const { size, outline, wrap, textBody: body } = content.shape;
          return {
            size,
            ...(outline?.width === undefined ? {} : { outlineWidth: outline.width }),
            ...(wrap === undefined ? {} : { wrap }),
            ...(body?.margins === undefined ? {} : { margins: body.margins }),
          };
        }
        if (content.type === "drawing") {
          const { size, outline, wrap } = content.image;
          return {
            size,
            ...(outline?.width === undefined ? {} : { outlineWidth: outline.width }),
            ...(wrap === undefined ? {} : { wrap }),
          };
        }
      }
    }
  }
  throw new Error("Expected a drawing");
};

/** Re-create every drawing node in the projection with different attributes. */
const editDrawingAttrs = (node: PMNode, overrides: Record<string, unknown>): PMNode => {
  if (node.type.name === "shape" || node.type.name === "textBox" || node.type.name === "image") {
    return node.type.create({ ...node.attrs, ...overrides }, node.content, node.marks);
  }
  const children: PMNode[] = [];
  node.content.forEach((child) => {
    children.push(editDrawingAttrs(child, overrides));
  });
  return node.copy(Fragment.fromArray(children));
};

const restore = (source: Document): Document => fromProseDoc(toProseDoc(source), source);

describe("drawing geometry survives the editor projection", () => {
  test("an untouched drawing is written back with its authored EMUs", () => {
    fc.assert(
      fc.property(authoredGeometry, (geometry) => {
        const restored = geometryOf(restore(documentWithDrawing(geometry)));

        expect(restored.size).toEqual({ width: geometry.width, height: geometry.height });
        expect(restored.outlineWidth).toBe(geometry.outlineWidth);
        expect({
          distT: restored.wrap?.distT,
          distB: restored.wrap?.distB,
          distL: restored.wrap?.distL,
          distR: restored.wrap?.distR,
        }).toEqual(geometry.insets);
        if (geometry.kind === "textBox") {
          expect(restored.margins).toEqual(geometry.margins);
        }
      }),
      propertyConfig({ numRuns: 300 }),
    );
  });

  test("a second pass over the restored document changes nothing further", () => {
    fc.assert(
      fc.property(authoredGeometry, (geometry) => {
        const restored = restore(documentWithDrawing(geometry));
        const restoredAgain = restore(restored);

        expect(geometryOf(restoredAgain)).toEqual(geometryOf(restored));
      }),
      propertyConfig({ numRuns: 200 }),
    );
  });

  test("an editor change to the pixel attributes reaches the document", () => {
    fc.assert(
      fc.property(
        authoredGeometry,
        fc.integer({ min: 1, max: 2000 }),
        fc.integer({ min: 1, max: 400 }),
        (geometry, widthPx, distTopPx) => {
          // An "edit" that lands on the pixels the authored EMUs project to is
          // no edit: the authored value legitimately wins, and several EMUs
          // project to the same pixel.
          fc.pre(widthPx !== emuToPixels(geometry.width));
          fc.pre(distTopPx !== emuToPixels(geometry.insets.distT));

          const source = documentWithDrawing(geometry);
          const edited = editDrawingAttrs(toProseDoc(source), {
            width: widthPx,
            distTop: distTopPx,
          });
          const restored = geometryOf(fromProseDoc(edited, source));

          expect(restored.size.width).toBe(pixelsToEmu(widthPx));
          expect(restored.wrap?.distT).toBe(pixelsToEmu(distTopPx));
        },
      ),
      propertyConfig({ numRuns: 200 }),
    );
  });

  test("an editor change to a stroke width reaches the document", () => {
    fc.assert(
      fc.property(authoredGeometry, fc.integer({ min: 1, max: 40 }), (geometry, outlineWidthPx) => {
        fc.pre(
          geometry.outlineWidth === undefined ||
            outlineWidthPx !== emuToStrokePixels(geometry.outlineWidth),
        );

        const source = documentWithDrawing(geometry);
        const attr = geometry.kind === "picture" ? "borderWidth" : "outlineWidth";
        const edited = editDrawingAttrs(toProseDoc(source), { [attr]: outlineWidthPx });
        const restored = geometryOf(fromProseDoc(edited, source));

        expect(restored.outlineWidth).toBe(pixelsToEmu(outlineWidthPx));
      }),
      propertyConfig({ numRuns: 200 }),
    );
  });

  test("a shape whose outline states no width does not acquire one", () => {
    const source = documentWithDrawing({
      kind: "rect",
      width: 914_400,
      height: 457_200,
      insets: { distT: 0, distB: 0, distL: 0, distR: 0 },
      margins: { top: 0, bottom: 0, left: 0, right: 0 },
    });
    const restored = geometryOf(restore(source));

    expect(restored.outlineWidth).toBeUndefined();
  });
});
