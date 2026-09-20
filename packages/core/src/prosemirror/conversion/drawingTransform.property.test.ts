/**
 * A drawing's authored `a:xfrm` survives the editor projection unchanged.
 *
 * `toProseDoc` used to carry a drawing's rotation and flips as one CSS string
 * (`rotate(90deg) scaleX(-1)`), and `fromProseDoc` parsed them back out of it.
 * That string can state neither an authored `rot="0"` nor an authored
 * `flipH="0"`: both spell the identity, and the identity is exactly what an
 * absent transform already means, so the projection folded "the author wrote
 * zero" into "the author wrote nothing" and the save path dropped the
 * attribute (`image.transform.rotation: N became absent`).
 *
 * The property runs the three authored values over their whole range - absent,
 * the identity, and a value - through `toProseDoc → fromProseDoc` on all three
 * drawing nodes, and demands them back exactly, while a rotate or flip from
 * the editor still reaches the document.
 */

import { describe, expect, test } from "bun:test";
import fc from "fast-check";
import { Fragment, type Node as PMNode } from "prosemirror-model";

import { propertyConfig } from "../../../../../test/property-testing";
import type { Document, Image, ImageTransform, Shape } from "../../types/document";
import { computeImageTransform } from "../commands/image";
import { fromProseDoc } from "./fromProseDoc";
import { toProseDoc } from "./toProseDoc";

/**
 * Absent, the identity, a degree either way, and a value far outside a single
 * turn (Word writes `rot` in sixtieth-thousandths of a degree, and a document
 * that states one in degrees must come back as it went in).
 */
const authoredRotation = fc.constantFrom(undefined, 0, 1, -1, 5_400_000);

/** Absent, an explicit "not flipped", and a flip. */
const authoredFlip = fc.constantFrom(undefined, false, true);

/**
 * Every route a drawing takes through the projection: a plain shape becomes a
 * `shape` node, a shape with a text body a `textBox` node, a picture an
 * `image` node.
 */
const DRAWING_KINDS = ["rect", "textBox", "picture"] as const;
type DrawingKind = (typeof DRAWING_KINDS)[number];

type AuthoredTransform = {
  kind: DrawingKind;
  rotation: number | undefined;
  flipH: boolean | undefined;
  flipV: boolean | undefined;
};

const authoredTransform = fc.record({
  kind: fc.constantFrom(...DRAWING_KINDS),
  rotation: authoredRotation,
  flipH: authoredFlip,
  flipV: authoredFlip,
});

const transformOf = ({ rotation, flipH, flipV }: AuthoredTransform): ImageTransform | undefined => {
  if (rotation === undefined && flipH === undefined && flipV === undefined) {
    return undefined;
  }
  return {
    ...(rotation === undefined ? {} : { rotation }),
    ...(flipH === undefined ? {} : { flipH }),
    ...(flipV === undefined ? {} : { flipV }),
  };
};

/** A 1x1 PNG, so the picture case has media to reference. */
const PNG_1X1_BASE64 =
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";

const buildShape = (authored: AuthoredTransform): Shape => ({
  type: "shape",
  shapeType: authored.kind === "textBox" ? "textBox" : "rect",
  size: { width: 914_400, height: 914_400 },
  wrap: { type: "inline" },
  ...(transformOf(authored) === undefined ? {} : { transform: transformOf(authored) }),
  ...(authored.kind === "textBox"
    ? {
        textBody: {
          content: [
            {
              type: "paragraph" as const,
              content: [{ type: "run" as const, content: [{ type: "text" as const, text: "t" }] }],
            },
          ],
        },
      }
    : {}),
});

const buildImage = (authored: AuthoredTransform): Image => ({
  type: "image",
  rId: "rIdTransform",
  size: { width: 914_400, height: 914_400 },
  wrap: { type: "inline" },
  ...(transformOf(authored) === undefined ? {} : { transform: transformOf(authored) }),
});

const documentWithDrawing = (authored: AuthoredTransform): Document => ({
  package: {
    document: {
      content: [
        {
          type: "paragraph",
          content: [
            {
              type: "run",
              content: [
                authored.kind === "picture"
                  ? { type: "drawing", image: buildImage(authored) }
                  : { type: "shape", shape: buildShape(authored) },
              ],
            },
          ],
        },
      ],
    },
    ...(authored.kind === "picture"
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
              "rIdTransform",
              {
                id: "rIdTransform",
                type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
                target: "media/image1.png",
              },
            ],
          ]),
        }
      : {}),
  },
});

const restoredTransform = (document: Document): ImageTransform | undefined => {
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
          return content.shape.transform;
        }
        if (content.type === "drawing") {
          return content.image.transform;
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

describe("drawing transform survives the editor projection", () => {
  test("an untouched drawing is written back with its authored rotation and flips", () => {
    fc.assert(
      fc.property(authoredTransform, (authored) => {
        expect(restoredTransform(restore(documentWithDrawing(authored)))).toEqual(
          transformOf(authored),
        );
      }),
      propertyConfig({ numRuns: 400 }),
    );
  });

  test("a second pass over the restored document changes nothing further", () => {
    fc.assert(
      fc.property(authoredTransform, (authored) => {
        const restored = restore(documentWithDrawing(authored));

        expect(restoredTransform(restore(restored))).toEqual(restoredTransform(restored));
      }),
      propertyConfig({ numRuns: 200 }),
    );
  });

  test("a rotate or flip from the editor reaches the document", () => {
    fc.assert(
      fc.property(
        authoredTransform,
        fc.constantFrom("rotateCW" as const, "rotateCCW" as const, "flipH" as const),
        (authored, action) => {
          const source = documentWithDrawing(authored);
          const edited = computeImageTransform(transformOf(authored), action);
          const projected = editDrawingAttrs(toProseDoc(source), {
            docxRotation: edited.rotation ?? null,
            docxFlipH: edited.flipH ?? null,
            docxFlipV: edited.flipV ?? null,
          });

          expect(restoredTransform(fromProseDoc(projected, source))).toEqual(edited);
        },
      ),
      propertyConfig({ numRuns: 300 }),
    );
  });

  test("a rotation the editor computes names the turn it is at", () => {
    fc.assert(
      fc.property(
        fc.integer({ max: 10_000, min: -10_000 }),
        fc.constantFrom("rotateCW" as const, "rotateCCW" as const),
        (rotation, action) => {
          // `%` keeps the sign of its left operand, so a negative authored
          // rotation used to come back negative: a turn nothing can state.
          const { rotation: rotated } = computeImageTransform({ rotation }, action);

          expect(rotated).toBeGreaterThanOrEqual(0);
          expect(rotated).toBeLessThan(360);
        },
      ),
      propertyConfig({ numRuns: 200 }),
    );
  });
});
