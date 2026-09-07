/**
 * A save never replays Strict markup under the Transitional root it rebuilds.
 *
 * folio writes one conformance class, so a rebuilt part must carry neither a
 * Strict namespace nor a value spelled the way Strict spells it — a length with
 * its unit attached, or a percentage with its sign — in a slot whose
 * Transitional type is a number. Defects already present in a fixture's own
 * part are subtracted, so a malformed source stays the source's problem while
 * anything folio adds fails here.
 *
 * The whole corpus runs through `test.each`, so a Strict fixture added later is
 * covered with no edit here. Fixture provenance and licensing: see
 * `__fixtures__/corpus/PROVENANCE.md`.
 */

import { describe, expect, test } from "bun:test";
import JSZip from "jszip";
import { readFileSync, readdirSync } from "node:fs";
import path from "node:path";

import { parseDocx } from "../parser";
import { serializeDocument } from "../serializer/documentSerializer";
import { toTransitionalNamespaceUri, transitionalSlotEncoding } from "../transitionalSpelling";
import { getLocalName, parseXmlDocument, type XmlElement } from "../xmlParser";

const CORPUS_DIR = path.join(import.meta.dir, "__fixtures__", "corpus");
const FIXTURES: string[] = readdirSync(CORPUS_DIR)
  .filter((name) => name.endsWith(".docx"))
  .sort()
  .map((name) => path.join(CORPUS_DIR, name));

const STRICT_URI_PREFIX = "http://purl.oclc.org/ooxml/";
const UNIVERSAL_MEASURE = /^-?[0-9]+(?:\.[0-9]+)?(?:mm|cm|in|pt|pc|pi)$/u;
const PERCENTAGE = /^-?[0-9]+(?:\.[0-9]+)?%$/u;

const readFixture = (file: string): ArrayBuffer => {
  const bytes = readFileSync(file);
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
};

/** Strict spellings in one part, as stable descriptions so two parts can be diffed. */
const strictSpellings = (xml: string): Set<string> => {
  const found = new Set<string>();
  const root = parseXmlDocument(xml);
  if (root === null) {
    return found;
  }

  const visit = (element: XmlElement): void => {
    const uri = element.namespaceUri;
    const transitional = uri === undefined ? "" : toTransitionalNamespaceUri(uri);
    const slotPrefix = `${transitional} ${getLocalName(element.name)}`;

    for (const [name, value] of Object.entries(element.attributes ?? {})) {
      if (typeof value !== "string") {
        continue;
      }
      if (value.startsWith(STRICT_URI_PREFIX)) {
        found.add(`Strict namespace ${value}`);
        continue;
      }
      const slot = `${slotPrefix} @${getLocalName(name)}`;
      const encoding = transitionalSlotEncoding(
        uri,
        getLocalName(element.name),
        getLocalName(name),
      );
      if (encoding?.measure !== undefined && UNIVERSAL_MEASURE.test(value)) {
        found.add(`${slot} carries the length ${value}`);
      }
      if (encoding?.percent !== undefined && PERCENTAGE.test(value)) {
        found.add(`${slot} carries the percentage ${value}`);
      }
    }
    if (uri?.startsWith(STRICT_URI_PREFIX) === true) {
      found.add(`Strict namespace ${uri}`);
    }

    for (const child of element.elements ?? []) {
      if (child.type === "element") {
        visit(child);
      }
    }
  };
  visit(root);
  return found;
};

const STRICT_NS = {
  a: "http://purl.oclc.org/ooxml/drawingml/main",
  w: "http://purl.oclc.org/ooxml/wordprocessingml/main",
  wp: "http://purl.oclc.org/ooxml/drawingml/wordprocessingDrawing",
};
const WP14_NS = "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing";
const WPS_NS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";

/**
 * A Strict package whose body only folio's verbatim replay paths can carry
 * across: a table whose widths are lengths, and a floating shape whose
 * DrawingML payload declares its own Strict namespace and sizes itself with a
 * percentage.
 */
const strictPackage = async (): Promise<ArrayBuffer> => {
  const cell = (width: string) =>
    `<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc>`;
  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${STRICT_NS.w}" xmlns:wp="${STRICT_NS.wp}" xmlns:wp14="${WP14_NS}" xmlns:wps="${WPS_NS}">
  <w:body>
    <w:tbl>
      <w:tblPr><w:tblW w:w="311.70pt" w:type="dxa"/></w:tblPr>
      <w:tblGrid><w:gridCol w:w="155.85pt"/><w:gridCol w:w="155.85pt"/></w:tblGrid>
      <w:tr>${cell("155.85pt")}${cell("155.85pt")}</w:tr>
    </w:tbl>
    <w:p><w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV><wp:extent cx="1905000" cy="635000"/><wp:wrapNone/><wp:docPr id="1" name="Text Box 1"/><a:graphic xmlns:a="${STRICT_NS.a}"><a:graphicData uri="${WPS_NS}"><wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1905000" cy="635000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="FFFFFF"><a:alpha val="60%"/></a:srgbClr></a:solidFill></wps:spPr><wps:txbx><w:txbxContent><w:p><w:r><w:t>Box</w:t></w:r></w:p></w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic><wp14:sizeRelH relativeFrom="margin"><wp14:pctWidth>40%</wp14:pctWidth></wp14:sizeRelH><wp14:sizeRelV relativeFrom="margin"><wp14:pctHeight>12.5%</wp14:pctHeight></wp14:sizeRelV></wp:anchor></w:drawing></w:r></w:p>
  </w:body>
</w:document>`;

  const zip = new JSZip();
  zip.file("word/document.xml", documentXml);
  return zip.generateAsync({ type: "arraybuffer" });
};

describe("Strict content replayed into a Transitional part", () => {
  test.each(FIXTURES.map((file) => [path.basename(file), file] as const))(
    "%s introduces no Strict spelling",
    async (_name, file) => {
      const original = readFixture(file);
      const source = await JSZip.loadAsync(original)
        .then((zip) => zip.file("word/document.xml")?.async("text"))
        .then((xml) => (xml === undefined ? new Set<string>() : strictSpellings(xml)));

      const saved = serializeDocument(await parseDocx(original));

      const introduced = [...strictSpellings(saved)].filter((defect) => !source.has(defect));
      expect(introduced).toEqual([]);
    },
    30_000,
  );

  test("a Strict table, drawing and text box are re-spelled Transitional", async () => {
    const saved = serializeDocument(await parseDocx(await strictPackage()));

    expect(saved).not.toContain(STRICT_URI_PREFIX);
    // 155.85pt is 3117 twips; 311.70pt is 6234.
    expect(saved).toContain('<w:tcW w:w="3117" w:type="dxa"/>');
    expect(saved).toContain('<w:tblW w:w="6234" w:type="dxa"/>');
    expect(saved).toContain('<w:gridCol w:w="3117"/>');
    expect(strictSpellings(saved).size).toBe(0);
  });

  test("a shape with no text body still closes its content model", async () => {
    const saved = serializeDocument(await parseDocx(await strictPackage()));

    expect(saved).toContain("<wps:bodyPr");
  });
});
