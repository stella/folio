/**
 * Page-border art keeps its relationship references.
 *
 * `w:pgBorders` sides carry relationship ids for Word's border art: `r:id` on
 * every side (`CT_PageBorder`), plus `r:topLeft` / `r:topRight` on the top
 * (`CT_TopPageBorder`) and `r:bottomLeft` / `r:bottomRight` on the bottom
 * (`CT_BottomPageBorder`). folio read them and wrote them back into the `w`
 * namespace, which is a different attribute: Word ignored it, the art was gone
 * on reload, and the container survival census recorded all eight as
 * `parsed-but-not-serialized`.
 *
 * The property sweeps every side crossed with every id it can carry, over
 * arbitrary relationship ids, through the save and the editor round trip, and
 * demands both the model value and the `r`-namespaced attribute back.
 */

import { describe, expect, test } from "bun:test";
import fc from "fast-check";
import JSZip from "jszip";

import { propertyConfig, propertyTestTimeout } from "../../../../test/property-testing";

import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import type { BorderSpec, Document } from "../types/document";
import { parseDocx } from "./parser";
import { RELATIONSHIP_TYPES } from "./relsParser";
import { repackDocx } from "./rezip";

const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const NS = [
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"',
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"',
].join(" ");

/** Each side, with the art attributes its schema type allows. */
const SIDE_ART_ATTRIBUTES = {
  top: ["id", "topLeft", "topRight"],
  bottom: ["id", "bottomLeft", "bottomRight"],
  left: ["id"],
  right: ["id"],
} as const satisfies Record<string, readonly string[]>;

type Side = keyof typeof SIDE_ART_ATTRIBUTES;

const SIDES = Object.keys(SIDE_ART_ATTRIBUTES) as Side[];

/** The model field each attribute lands on. */
const ART_MODEL_FIELDS = {
  id: "artRelationshipId",
  topLeft: "topLeftArtRelationshipId",
  topRight: "topRightArtRelationshipId",
  bottomLeft: "bottomLeftArtRelationshipId",
  bottomRight: "bottomRightArtRelationshipId",
} as const satisfies Record<string, keyof BorderSpec>;

type ArtAttribute = keyof typeof ART_MODEL_FIELDS;

/** `r:id` values are xsd:ID-shaped relationship ids; Word writes `rIdN`. */
const relationshipId = fc
  .string({ unit: "grapheme-ascii", minLength: 1, maxLength: 8 })
  .map((suffix) => `rId${suffix.replaceAll(/[^A-Za-z0-9]/gu, "")}`)
  .filter((id) => id.length > 3);

type ArtCase = Record<Side, Partial<Record<ArtAttribute, string>>>;

const sideArt = (side: Side): fc.Arbitrary<Partial<Record<ArtAttribute, string>>> =>
  fc.record(
    Object.fromEntries(
      SIDE_ART_ATTRIBUTES[side].map((attribute) => [attribute, relationshipId]),
    ) as Record<ArtAttribute, fc.Arbitrary<string>>,
    { requiredKeys: [] },
  );

const artCase: fc.Arbitrary<ArtCase> = fc.record({
  top: sideArt("top"),
  bottom: sideArt("bottom"),
  left: sideArt("left"),
  right: sideArt("right"),
});

const allIds = (artCase_: ArtCase): string[] => [
  ...new Set(SIDES.flatMap((side) => Object.values(artCase_[side]))),
];

const sideXml = (side: Side, art: Partial<Record<ArtAttribute, string>>): string => {
  const attributes = Object.entries(art)
    .map(([attribute, id]) => ` r:${attribute}="${id}"`)
    .join("");
  return `<w:${side} w:val="single" w:sz="24" w:space="24"${attributes}/>`;
};

const docxFor = (art: ArtCase): Promise<ArrayBuffer> => {
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
    `${XML_DECLARATION}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${allIds(
      art,
    )
      .map(
        (id) =>
          `<Relationship Id="${id}" Type="${RELATIONSHIP_TYPES.image}" Target="media/${id}.png"/>`,
      )
      .join("")}</Relationships>`,
  );
  zip.file(
    "word/document.xml",
    `${XML_DECLARATION}<w:document ${NS}><w:body><w:p><w:r><w:t>bordered</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgBorders w:offsetFrom="page">${SIDES.map(
      (side) => sideXml(side, art[side]),
    ).join("")}</w:pgBorders></w:sectPr></w:body></w:document>`,
  );
  // Every art id's relationship must resolve to a real part, or package
  // reconciliation (rightly) treats it as dangling and prunes both the
  // relationship and the `r:id` naming it — this fixture is exercising
  // relationship-id wiring, not image content, so the bytes are arbitrary.
  for (const id of allIds(art)) {
    zip.file(`word/media/${id}.png`, new Uint8Array([0x89, 0x50, 0x4e, 0x47]));
  }
  return zip.generateAsync({ type: "arraybuffer" });
};

/** The art ids the model holds, side by side. */
const readArt = (document: Document): ArtCase => {
  const borders = document.package.document.sections?.at(-1)?.properties.pageBorders;
  const read = (side: Side): Partial<Record<ArtAttribute, string>> => {
    const spec = borders?.[side];
    const art: Partial<Record<ArtAttribute, string>> = {};
    for (const attribute of SIDE_ART_ATTRIBUTES[side]) {
      const value = spec?.[ART_MODEL_FIELDS[attribute]];
      if (typeof value === "string") {
        art[attribute] = value;
      }
    }
    return art;
  };
  return { top: read("top"), bottom: read("bottom"), left: read("left"), right: read("right") };
};

const parse = (buffer: ArrayBuffer): Promise<Document> =>
  parseDocx(buffer, { detectVariables: false, preloadFonts: false });

const save = (document: Document): Promise<ArrayBuffer> =>
  repackDocx(document, { updateModifiedDate: false });

const documentXml = async (buffer: ArrayBuffer): Promise<string> => {
  const file = (await JSZip.loadAsync(buffer)).file("word/document.xml");
  if (!file) {
    throw new Error("saved package has no main part");
  }
  return file.async("string");
};

/** Every art attribute the saved `w:pgBorders` carries, prefix included. */
const artAttributeNames = async (buffer: ArrayBuffer): Promise<string[]> => {
  const xml = await documentXml(buffer);
  const borders = /<w:pgBorders\b[\s\S]*?<\/w:pgBorders>|<w:pgBorders\b[^>]*\/>/u.exec(xml);
  return [
    ...(borders?.[0] ?? "").matchAll(/\s(\w+:(?:id|topLeft|topRight|bottomLeft|bottomRight))=/gu),
  ]
    .map((match) => match[1] ?? "")
    .toSorted();
};

const expectedAttributeNames = (art: ArtCase): string[] =>
  SIDES.flatMap((side) => Object.keys(art[side]).map((attribute) => `r:${attribute}`)).toSorted();

describe("page-border art relationship ids survive a rebuild", () => {
  test("every art id is written back in the relationships namespace", async () => {
    const art: ArtCase = {
      top: { id: "rIdArt", topLeft: "rIdTopLeft", topRight: "rIdTopRight" },
      bottom: { id: "rIdArt", bottomLeft: "rIdBottomLeft", bottomRight: "rIdBottomRight" },
      left: { id: "rIdArt" },
      right: { id: "rIdArt" },
    };
    const saved = await save(await parse(await docxFor(art)));
    const xml = await documentXml(saved);

    for (const attribute of ["id", "topLeft", "topRight", "bottomLeft", "bottomRight"] as const) {
      expect(xml).toContain(`r:${attribute}="rId`);
    }
    expect(xml).not.toContain("w:topLeft=");
    expect(xml).not.toContain("w:bottomRight=");
  });

  test(
    "every side keeps every art id it carries, through the save and the editor",
    async () => {
      await fc.assert(
        fc.asyncProperty(artCase, async (art) => {
          const opened = await parse(await docxFor(art));
          expect(readArt(opened)).toEqual(art);

          const saved = await save(opened);
          expect(readArt(await parse(saved))).toEqual(art);
          // The model round trip alone cannot see the defect: folio's own
          // reader falls back to any prefix, so a `w:id` it wrote it reads
          // back. Only Word can tell, so the written namespace is asserted.
          expect(await artAttributeNames(saved)).toEqual(expectedAttributeNames(art));

          const edited = fromProseDoc(toProseDoc(opened), opened);
          expect(readArt(await parse(await save(edited)))).toEqual(art);
        }),
        propertyConfig({ numRuns: 40 }),
      );
    },
    propertyTestTimeout(90_000),
  );
});
