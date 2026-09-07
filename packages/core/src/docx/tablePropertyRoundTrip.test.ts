/**
 * A save that changed nothing must write a table's properties back unchanged.
 *
 * The typed model covers what the editor understands. A document carries more:
 * conditional-format flags in the attribute form Word writes, properties a
 * later revision of the format added, properties a producer wrote that nothing
 * here reads. Rebuilding `w:tblPr`, `w:trPr` and `w:tcPr` from the model alone
 * dropped every one of them, so a full repack — which is what a structural
 * edit falls back to — rewrote every table in the document, including the ones
 * nobody touched.
 *
 * These cases compare the property elements themselves, canonicalized so that
 * attribute order is not the subject: what is asserted is that the same
 * elements, with the same attributes and children, come back out.
 */

import { describe, expect, test } from "bun:test";
import { readFileSync } from "node:fs";
import path from "node:path";

import type { Document } from "../types/document";
import { parseDocx } from "./parser";
import { repackDocx } from "./rezip";
import { unzipDocx } from "./unzip";
import { getChildElements, getLocalName, parseXmlDocument, type XmlElement } from "./xmlParser";

const VISUAL_FIXTURES = path.resolve(import.meta.dir, "../../../../tests/visual/fixtures");
const CORPUS_FIXTURES = path.join(import.meta.dir, "__tests__/__fixtures__/corpus");

/** Real documents the repository already keeps, each carrying tables. */
const FIXTURES = [
  { dir: VISUAL_FIXTURES, name: "sample.docx" },
  { dir: VISUAL_FIXTURES, name: "docx-editor-demo.docx" },
  { dir: VISUAL_FIXTURES, name: "podily-bps.docx" },
  { dir: CORPUS_FIXTURES, name: "upstream-with-tables.docx" },
] as const;

const readFixture = ({ dir, name }: { dir: string; name: string }): ArrayBuffer => {
  const bytes = readFileSync(path.join(dir, name));
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
};

const documentPart = async (buffer: ArrayBuffer): Promise<XmlElement> => {
  const { documentXml: xml } = await unzipDocx(buffer);
  if (typeof xml !== "string") {
    throw new Error("The package has no main document part.");
  }
  const parsed = parseXmlDocument(xml);
  if (!parsed) {
    throw new Error("The main document part could not be parsed.");
  }
  return parsed;
};

/** Tag, attributes in a fixed order, then children — the element's own shape. */
const canonicalElement = (element: XmlElement): string => {
  const attributes = Object.entries(element.attributes ?? {})
    .toSorted(([left], [right]) => left.localeCompare(right))
    .map(([name, value]) => `${name}=${String(value)}`)
    .join(" ");
  return `<${element.name ?? ""} ${attributes}>${getChildElements(element)
    .map(canonicalElement)
    .join("")}</>`;
};

const PROPERTY_ELEMENTS = new Set(["tblPr", "tblGrid", "trPr", "tcPr"]);

/** Every table property element of a part, in document order. */
const tableProperties = (root: XmlElement): string[] => {
  const found: string[] = [];
  const walk = (element: XmlElement): void => {
    for (const child of getChildElements(element)) {
      if (PROPERTY_ELEMENTS.has(getLocalName(child.name))) {
        found.push(canonicalElement(child));
        continue;
      }
      walk(child);
    }
  };
  walk(root);
  return found;
};

const noEditRepack = async (buffer: ArrayBuffer): Promise<ArrayBuffer> => {
  const parsed: Document = await parseDocx(buffer, { preloadFonts: false });
  return await repackDocx({ ...parsed, originalBuffer: buffer });
};

describe("table properties survive a no-edit full repack", () => {
  test.each(FIXTURES.map((fixture) => [fixture.name, fixture] as const))(
    "%s",
    async (_name, fixture) => {
      const buffer = readFixture(fixture);
      const before = tableProperties(await documentPart(buffer));
      // A fixture with no table proves nothing here, and every one listed has
      // some; assert that rather than passing an empty comparison silently.
      expect(before.length).toBeGreaterThan(0);

      // Every element the source stated comes back, as it was written. A
      // document missing a `w:tblGrid` the schema requires gets one, so the
      // check is that nothing is lost or altered rather than that nothing is
      // added.
      const after = tableProperties(await documentPart(await noEditRepack(buffer)));
      const remaining = [...after];
      const lost = before.filter((element) => {
        const index = remaining.indexOf(element);
        if (index === -1) {
          return true;
        }
        remaining.splice(index, 1);
        return false;
      });
      expect(lost).toEqual([]);
    },
    30_000,
  );
});
