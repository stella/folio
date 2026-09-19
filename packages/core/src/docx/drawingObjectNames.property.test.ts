/**
 * `CT_NonVisualDrawingProps` names survive every save path.
 *
 * `wp:docPr` carries three authored strings: `@name`, `@descr` (alt text) and
 * `@title`. folio used to mint `Shape 3` / `TextBox 3` / `Picture 3` in the
 * serializer whenever the model had no name, which turned every name the model
 * failed to carry into a plausible English one — a shape called
 * `直接箭头连接符 2` came back as `Shape 2` through the editor round trip, and
 * no census could tell an authored name from a generated one.
 *
 * The property runs the input class the corpus cannot enumerate — arbitrary
 * unicode names, alt text and titles, over shapes, text boxes and pictures,
 * inline and anchored — through both save paths:
 *
 *   save    parse → repack → parse. A shape has no capture slot, so this is
 *           already the real serializer.
 *   editor  parse → toProseDoc → fromProseDoc → repack → parse.
 *
 * and demands the three strings back verbatim.
 *
 * `@descr` and `@title` are optional, so absent stays absent. `@name` is
 * schema-required: a drawing the model has not named writes `""`, and the
 * reader maps that one value back to absent, so an unnamed drawing stays
 * unnamed however many times it is saved. What it must never acquire is a
 * plausible generated name, because a later reader cannot tell that from an
 * authored one.
 */

import { describe, expect, test } from "bun:test";
import fc from "fast-check";

import { propertyConfig, propertyTestTimeout } from "../../../../test/property-testing";

import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import type { Document, Image, RunContent, Shape } from "../types/document";
import { parseDocx } from "./parser";
import { createEmptyDocx, repackDocx } from "./rezip";

/** The three authored strings under test. */
type DrawingNames = {
  name?: string;
  alt?: string;
  title?: string;
};

const DRAWING_KINDS = ["shape", "textBox", "picture"] as const;
type DrawingKind = (typeof DRAWING_KINDS)[number];

const PLACEMENTS = ["inline", "anchored"] as const;
type Placement = (typeof PLACEMENTS)[number];

type DrawingCase = {
  kind: DrawingKind;
  placement: Placement;
  names: DrawingNames;
};

/**
 * Names Word accepts and folio must not touch: CJK, RTL, combining marks,
 * emoji, XML metacharacters, whitespace and the empty string.
 *
 * Control characters are out of scope: XML 1.0 cannot carry NUL at all, and an
 * unescaped CR in an attribute value is normalised to a space by any
 * conformant reader, so neither is a name a document can hold.
 */
const authoredString = fc.oneof(
  fc.constant(""),
  fc.constant("直接箭头连接符 2"),
  fc.constant("سهم مستقيم"),
  fc.constant("Šípka <&> \"quoted\" 'x'"),
  fc.constant("  leading and trailing  "),
  fc.string({ unit: "grapheme", minLength: 1, maxLength: 16 }),
  fc.string({ unit: "grapheme-composite", minLength: 1, maxLength: 16 }),
);

/** Each of the three is independently present or absent. */
const drawingNames = fc.record(
  {
    name: authoredString,
    alt: authoredString,
    title: authoredString,
  },
  { requiredKeys: [] },
);

const drawingCase = fc.record({
  kind: fc.constantFrom(...DRAWING_KINDS),
  placement: fc.constantFrom(...PLACEMENTS),
  names: drawingNames,
});

const presentNames = (names: DrawingNames): DrawingNames => ({
  ...(names.name !== undefined ? { name: names.name } : {}),
  ...(names.alt !== undefined ? { alt: names.alt } : {}),
  ...(names.title !== undefined ? { title: names.title } : {}),
});

/**
 * What a save path must return: the authored strings, except that `@name` is
 * schema-required, so both "no name" and `""` come back as no name.
 */
const namesAfterSave = (names: DrawingNames): DrawingNames =>
  presentNames({ ...names, ...(names.name === "" ? { name: undefined } : {}) });

const SIZE = { width: 914_400, height: 457_200 };

const anchorWrap = { type: "square", wrapText: "bothSides" } as const;
const anchorPosition = {
  horizontal: { relativeTo: "column", posOffset: 0 },
  vertical: { relativeTo: "paragraph", posOffset: 0 },
} as const;

const buildShape = ({ kind, placement, names }: DrawingCase): Shape => ({
  type: "shape",
  shapeType: kind === "textBox" ? "textBox" : "rect",
  size: SIZE,
  ...presentNames(names),
  ...(placement === "anchored"
    ? { wrap: anchorWrap, position: anchorPosition }
    : { wrap: { type: "inline" as const } }),
  ...(kind === "textBox"
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

const buildImage = ({ placement, names }: DrawingCase): Image => ({
  type: "image",
  id: "7",
  rId: "rIdPropPicture",
  size: SIZE,
  ...(names.name !== undefined ? { docPrName: names.name } : {}),
  ...(names.alt !== undefined ? { alt: names.alt } : {}),
  ...(names.title !== undefined ? { title: names.title } : {}),
  ...(placement === "anchored"
    ? { wrap: anchorWrap, position: anchorPosition }
    : { wrap: { type: "inline" as const } }),
});

/** A 1x1 PNG, so the picture case has media to reference. */
const PNG_1X1_BASE64 =
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";

const runContentFor = (drawing: DrawingCase): RunContent =>
  drawing.kind === "picture"
    ? { type: "drawing", image: buildImage(drawing) }
    : { type: "shape", shape: buildShape(drawing) };

const withDrawing = (document: Document, drawing: DrawingCase): Document => {
  const next: Document = {
    ...document,
    package: {
      ...document.package,
      document: {
        ...document.package.document,
        content: [
          {
            type: "paragraph",
            content: [{ type: "run", content: [runContentFor(drawing)] }],
          },
        ],
      },
    },
  };
  if (drawing.kind === "picture") {
    const bytes = Uint8Array.from(
      atob(PNG_1X1_BASE64),
      (character) => character.codePointAt(0) ?? 0,
    );
    next.package.media = new Map([
      [
        "word/media/image1.png",
        {
          path: "word/media/image1.png",
          filename: "image1.png",
          mimeType: "image/png",
          data: bytes.buffer,
          base64: PNG_1X1_BASE64,
        },
      ],
    ]);
    next.package.relationships = new Map([
      ...(next.package.relationships ?? new Map()),
      [
        "rIdPropPicture",
        {
          id: "rIdPropPicture",
          type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
          target: "media/image1.png",
        },
      ],
    ]);
  }
  return next;
};

/** The names of the first drawing in the body, whichever kind it is. */
const readNames = (document: Document): DrawingNames | null => {
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
          return presentNames(content.shape);
        }
        if (content.type === "drawing") {
          return presentNames({
            ...(content.image.docPrName !== undefined ? { name: content.image.docPrName } : {}),
            ...(content.image.alt !== undefined ? { alt: content.image.alt } : {}),
            ...(content.image.title !== undefined ? { title: content.image.title } : {}),
          });
        }
      }
    }
  }
  return null;
};

const parse = (buffer: ArrayBuffer): Promise<Document> =>
  parseDocx(buffer, { detectVariables: false, preloadFonts: false });

const save = (document: Document): Promise<ArrayBuffer> =>
  repackDocx(document, { updateModifiedDate: false });

describe("drawing object names survive both save paths", () => {
  test("a non-English shape name is not replaced by a generated one", async () => {
    const template = await parse(await createEmptyDocx());
    const authored = withDrawing(template, {
      kind: "shape",
      placement: "inline",
      names: { name: "直接箭头连接符 2", alt: "直接箭头", title: "标题" },
    });

    const opened = await parse(await save(authored));
    const edited = fromProseDoc(toProseDoc(opened), opened);
    const afterEditor = await parse(await save(edited));

    expect(readNames(afterEditor)).toEqual({
      name: "直接箭头连接符 2",
      alt: "直接箭头",
      title: "标题",
    });
  });

  test.each(DRAWING_KINDS.flatMap((kind) => PLACEMENTS.map((placement) => ({ kind, placement }))))(
    "an unnamed $placement $kind stays unnamed",
    async ({ kind, placement }) => {
      const template = await parse(await createEmptyDocx());
      const unnamed = withDrawing(template, { kind, placement, names: {} });

      const savedOnce = await parse(await save(unnamed));
      expect(readNames(savedOnce)).toEqual({});

      const edited = fromProseDoc(toProseDoc(savedOnce), savedOnce);
      const afterEditor = await parse(await save(edited));
      expect(readNames(afterEditor)).toEqual({});
    },
  );

  test(
    "an authored name, alt text and title round-trip verbatim",
    async () => {
      const template = await parse(await createEmptyDocx());

      await fc.assert(
        fc.asyncProperty(drawingCase, async (drawing) => {
          const authored = withDrawing(template, drawing);
          const expected = namesAfterSave(drawing.names);

          const savedOnce = await parse(await save(authored));
          expect(readNames(savedOnce)).toEqual(expected);

          const edited = fromProseDoc(toProseDoc(savedOnce), savedOnce);
          const afterEditor = await parse(await save(edited));
          expect(readNames(afterEditor)).toEqual(expected);
        }),
        propertyConfig({ numRuns: 60 }),
      );
    },
    propertyTestTimeout(60_000),
  );
});
