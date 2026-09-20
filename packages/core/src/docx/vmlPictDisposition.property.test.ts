/**
 * A `w:pict` is modelled or preserved, never neither.
 *
 * VML has no serializer, so an unresolved `w:pict` is only ever replayed from
 * the bytes the parse captured or lost outright. Which of the two happened was
 * decided by `shouldPreserveRawVmlPict` declining for anything shaped like a
 * watermark, on the premise that `watermarkParser` had claimed it. That premise
 * was a guess about another module, and it was wrong wherever the two readers
 * disagreed: `watermarkParser` claims a direct `v:shape` child of a `w:pict`,
 * carrying a non-empty `v:textpath` or a `v:imagedata`, alone in its paragraph,
 * in a header. A `v:oval`, a shape nested in a `v:group`, a shape sharing its
 * paragraph with text, and every `w:pict` in a footer or in the body sit
 * outside it, so their artwork was declined by one owner and claimed by no
 * other. Verbatim part replay hid the loss until the part was rebuilt.
 *
 * The property generates the disagreement rather than the watermark: shape
 * kind, watermark spelling, nesting and container vary independently, and the
 * save runs with every rebuildable capture removed so the serializers, not the
 * replay, decide what comes back. Both halves are asserted, because the markup
 * surviving is worth nothing if the relationship it names does not: a
 * `v:imagedata r:id` that outlives its `.rels` entry is a picture that no
 * longer resolves.
 */

import { describe, expect, test } from "bun:test";
import fc from "fast-check";
import JSZip from "jszip";

import { propertyConfig, propertyTestTimeout } from "../../../../test/property-testing";

import { parseDocx } from "./parser";
import { createEmptyDocx, repackDocx } from "./rezip";

const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const W_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R_NAMESPACE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const V_NAMESPACE = "urn:schemas-microsoft-com:vml";
const O_NAMESPACE = "urn:schemas-microsoft-com:office:office";
const PACKAGE_RELATIONSHIPS = "http://schemas.openxmlformats.org/package/2006/relationships";

const IMAGE_RID = "rIdImg";
const HEADER_RID = "rIdHdr";
const FOOTER_RID = "rIdFtr";

/** A one-pixel PNG, so the image relationship resolves to a real part. */
const PNG = Uint8Array.from(
  atob(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==",
  ),
  (character) => character.charCodeAt(0),
);

/**
 * The VML elements that paint something. `v:textbox` is excluded on purpose:
 * it has a second owner, the text-box enrichment pass, which rebuilds it as an
 * editable shape rather than replaying it.
 */
const SHAPE_KINDS = [
  "shape",
  "rect",
  "roundrect",
  "oval",
  "line",
  "polyline",
  "curve",
  "arc",
  "image",
] as const;

/** How the shape spells "watermark", including not at all. */
const WATERMARK_SPELLINGS = {
  /** Word's WordArt shapetype, what a text watermark carries. */
  wordArtType: "wordArtType",
  /** The id prefix Word's own UI writes on a picture watermark. */
  idPrefix: "idPrefix",
  /** Neither: an ordinary piece of header or body artwork. */
  none: "none",
} as const;

type WatermarkSpelling = (typeof WATERMARK_SPELLINGS)[keyof typeof WATERMARK_SPELLINGS];

/** What the shape carries, which is what `watermarkParser` reads to claim it. */
const PAYLOADS = {
  imagedata: "imagedata",
  textpath: "textpath",
  /** Present but empty, which `readVmlTextWatermark` declines. */
  emptyTextpath: "emptyTextpath",
  none: "none",
} as const;

type Payload = (typeof PAYLOADS)[keyof typeof PAYLOADS];

/** The part the `w:pict` lives in. Only a header has a watermark reader. */
const CONTAINERS = { body: "body", header: "header", footer: "footer" } as const;

type Container = (typeof CONTAINERS)[keyof typeof CONTAINERS];

type PictCase = {
  kind: (typeof SHAPE_KINDS)[number];
  spelling: WatermarkSpelling;
  payload: Payload;
  /** Nested in a `v:group`, which the watermark reader's direct-child walk misses. */
  grouped: boolean;
  /** Sharing its paragraph with a run of text, which the watermark reader refuses. */
  sharesParagraph: boolean;
  container: Container;
};

const payloadMarkup = (payload: Payload): string => {
  switch (payload) {
    case PAYLOADS.imagedata:
      return `<v:imagedata r:id="${IMAGE_RID}"/>`;
    case PAYLOADS.textpath:
      return '<v:textpath string="DRAFT"/>';
    case PAYLOADS.emptyTextpath:
      return '<v:textpath string=""/>';
    case PAYLOADS.none:
      return "";
    default:
      return payload satisfies never;
  }
};

const shapeMarkup = ({ kind, spelling, payload }: PictCase): string => {
  const id = spelling === WATERMARK_SPELLINGS.idPrefix ? "WordPictureWatermark101" : "shape101";
  const type = spelling === WATERMARK_SPELLINGS.wordArtType ? ' type="#_x0000_t136"' : "";
  return (
    `<v:${kind} id="${id}"${type} style="width:100pt;height:50pt">` +
    `${payloadMarkup(payload)}</v:${kind}>`
  );
};

const pictMarkup = (pict: PictCase): string => {
  const shape = shapeMarkup(pict);
  const inner = pict.grouped
    ? `<v:group style="width:100pt;height:50pt">${shape}</v:group>`
    : shape;
  return `<w:pict>${inner}</w:pict>`;
};

const paragraphMarkup = (pict: PictCase): string =>
  pict.sharesParagraph
    ? `<w:p><w:r><w:t>caption</w:t></w:r><w:r>${pictMarkup(pict)}</w:r></w:p>`
    : `<w:p><w:r>${pictMarkup(pict)}</w:r></w:p>`;

const NAMESPACES = `xmlns:w="${W_NAMESPACE}" xmlns:r="${R_NAMESPACE}" xmlns:v="${V_NAMESPACE}" xmlns:o="${O_NAMESPACE}"`;

const relationships = (entries: readonly string[]): string =>
  `${XML_DECLARATION}<Relationships xmlns="${PACKAGE_RELATIONSHIPS}">${entries.join("")}</Relationships>`;

const IMAGE_RELATIONSHIP = `<Relationship Id="${IMAGE_RID}" Type="${R_NAMESPACE}/image" Target="media/image1.png"/>`;

/** Where the generated `w:pict` ends up, and how the package points at it. */
const PART_PATHS = {
  [CONTAINERS.body]: "word/document.xml",
  [CONTAINERS.header]: "word/header1.xml",
  [CONTAINERS.footer]: "word/footer1.xml",
} as const satisfies Record<Container, string>;

const relsPathOf = (partPath: string): string =>
  partPath.replace(/^(?<directory>.*\/)(?<name>[^/]+)$/u, "$<directory>_rels/$<name>.rels");

const buildDocx = async (pict: PictCase): Promise<ArrayBuffer> => {
  const zip = await JSZip.loadAsync(await createEmptyDocx());
  zip.file("word/media/image1.png", PNG);

  const inBody = pict.container === CONTAINERS.body;
  const reference =
    pict.container === CONTAINERS.header
      ? `<w:headerReference w:type="default" r:id="${HEADER_RID}"/>`
      : `<w:footerReference w:type="default" r:id="${FOOTER_RID}"/>`;
  zip.file(
    "word/document.xml",
    `${XML_DECLARATION}<w:document ${NAMESPACES}><w:body>` +
      `${inBody ? paragraphMarkup(pict) : "<w:p><w:r><w:t>x</w:t></w:r></w:p>"}` +
      `<w:sectPr>${inBody ? "" : reference}</w:sectPr></w:body></w:document>`,
  );

  const documentRelationships = [IMAGE_RELATIONSHIP];
  if (!inBody) {
    const root = pict.container === CONTAINERS.header ? "hdr" : "ftr";
    const partPath = PART_PATHS[pict.container];
    zip.file(
      partPath,
      `${XML_DECLARATION}<w:${root} ${NAMESPACES}>${paragraphMarkup(pict)}</w:${root}>`,
    );
    zip.file(relsPathOf(partPath), relationships([IMAGE_RELATIONSHIP]));
    const type = pict.container === CONTAINERS.header ? "header" : "footer";
    documentRelationships.push(
      `<Relationship Id="${pict.container === CONTAINERS.header ? HEADER_RID : FOOTER_RID}" Type="${R_NAMESPACE}/${type}" Target="${partPath.slice("word/".length)}"/>`,
    );
    const types = await zip.file("[Content_Types].xml")!.async("text");
    zip.file(
      "[Content_Types].xml",
      types.replace(
        "</Types>",
        `<Override PartName="/${partPath}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${type}+xml"/></Types>`,
      ),
    );
  }
  const existingRels = await zip.file("word/_rels/document.xml.rels")!.async("text");
  zip.file(
    "word/_rels/document.xml.rels",
    existingRels.replace("</Relationships>", `${documentRelationships.join("")}</Relationships>`),
  );

  const types = await zip.file("[Content_Types].xml")!.async("text");
  if (!types.includes('Extension="png"')) {
    zip.file(
      "[Content_Types].xml",
      types.replace("</Types>", '<Default Extension="png" ContentType="image/png"/></Types>'),
    );
  }
  return zip.generateAsync({ type: "arraybuffer" });
};

/**
 * The capture slots a rebuild can reconstruct, removed so the serializers run.
 *
 * `rawXml` and `rawWatermarkXml` are left alone: they hold markup no model
 * stands behind, so removing them would test deletion rather than
 * serialization. This mirrors the corpus gate's `reserialize` policy.
 */
const REBUILDABLE_CAPTURES = new Set([
  "sourceXml",
  "gridSourceXml",
  "verbatimXml",
  "verbatimFingerprint",
  "rawPropertiesXml",
  "rawEndPropertiesXml",
]);

const stripRebuildableCaptures = (value: unknown, seen = new WeakSet<object>()): void => {
  if (typeof value !== "object" || value === null || seen.has(value)) {
    return;
  }
  seen.add(value);
  if (Array.isArray(value)) {
    for (const item of value) {
      stripRebuildableCaptures(item, seen);
    }
    return;
  }
  if (value instanceof Map) {
    for (const item of value.values()) {
      stripRebuildableCaptures(item, seen);
    }
    return;
  }
  for (const key of Object.keys(value)) {
    if (REBUILDABLE_CAPTURES.has(key)) {
      Reflect.deleteProperty(value, key);
      continue;
    }
    stripRebuildableCaptures((value as Record<string, unknown>)[key], seen);
  }
};

const textOf = async (buffer: ArrayBuffer, path: string): Promise<string> =>
  (await (await JSZip.loadAsync(buffer)).file(path)?.async("text")) ?? "";

const pictCase = fc.record({
  kind: fc.constantFrom(...SHAPE_KINDS),
  spelling: fc.constantFrom(...Object.values(WATERMARK_SPELLINGS)),
  payload: fc.constantFrom(...Object.values(PAYLOADS)),
  grouped: fc.boolean(),
  sharesParagraph: fc.boolean(),
  container: fc.constantFrom(...Object.values(CONTAINERS)),
});

describe("a w:pict is modelled or preserved, never neither", () => {
  test(
    "the artwork and the relationship it names both survive a forced save",
    async () => {
      await fc.assert(
        fc.asyncProperty(pictCase, async (pict) => {
          const source = await buildDocx(pict);
          const partPath = PART_PATHS[pict.container];

          const parsed = await parseDocx(source, { preloadFonts: false });
          stripRebuildableCaptures(parsed);
          const saved = await repackDocx(parsed, { updateModifiedDate: false });

          const savedPart = await textOf(saved, partPath);
          expect({ pict, kept: savedPart.includes(`<v:${pict.kind}`) }).toEqual({
            pict,
            kept: true,
          });

          // The markup is worth nothing without the relationship it names.
          if (pict.payload === PAYLOADS.imagedata) {
            const savedRels = await textOf(saved, relsPathOf(partPath));
            expect({ pict, resolves: savedRels.includes(`Id="${IMAGE_RID}"`) }).toEqual({
              pict,
              resolves: true,
            });
          }

          // A second save changes nothing: what was preserved once is
          // preserved the same way, and no owner claims it on the way back.
          const reparsed = await parseDocx(saved, { preloadFonts: false });
          stripRebuildableCaptures(reparsed);
          const again = await repackDocx(reparsed, { updateModifiedDate: false });
          expect({ pict, part: await textOf(again, partPath) }).toEqual({
            pict,
            part: savedPart,
          });
        }),
        propertyConfig({ numRuns: 120 }),
      );
    },
    propertyTestTimeout(120_000),
  );

  test("a watermark the header reader claims is still written once", async () => {
    const pict = {
      kind: "shape",
      spelling: WATERMARK_SPELLINGS.wordArtType,
      payload: PAYLOADS.textpath,
      grouped: false,
      sharesParagraph: false,
      container: CONTAINERS.header,
    } as const satisfies PictCase;

    const parsed = await parseDocx(await buildDocx(pict), { preloadFonts: false });
    expect(parsed.package.headers?.get(HEADER_RID)?.watermark?.kind).toBe("text");

    stripRebuildableCaptures(parsed);
    const saved = await textOf(
      await repackDocx(parsed, { updateModifiedDate: false }),
      PART_PATHS.header,
    );
    expect(saved.split("<v:shape").length - 1).toBe(1);
  });
});
