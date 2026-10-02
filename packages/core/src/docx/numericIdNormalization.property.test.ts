import { expect, test, setDefaultTimeout } from "bun:test";
import fc from "fast-check";
import { assertValidOoxmlNumericIds } from "@stll/docx-core";
import JSZip from "jszip";

import {
  assertProperty,
  propertyConfig,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import { normalizeImportedNumericIds } from "./numericIdNormalization";
import { parseDocx } from "./parser";
import { createEmptyDocument } from "../utils/createDocument";
import type { Document } from "../types/document";
import { createDocx, createEmptyDocx, repackDocx } from "./rezip";
import { attemptSelectiveSave } from "./selectiveSave";
import { parseXmlWithFastXmlParser, type XmlElement } from "./xmlParser";

setDefaultTimeout(propertyTestTimeout(30_000));

const WORD_NAMESPACES = [
  "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
  "http://purl.oclc.org/ooxml/wordprocessingml/main",
] as const;
const invalidId = fc
  .oneof(
    fc.bigInt({ min: 2_147_483_648n, max: 99_999_999_999_999_999_999n }),
    fc.bigInt({ min: -99_999_999_999_999_999_999n, max: -2_147_483_649n }),
  )
  .map(String);

test("in-range packages reuse their source map without a reservation or rewrite pass", () => {
  fc.assert(
    fc.property(
      fc.integer({ min: -2_147_483_648, max: 2_147_483_647 }),
      fc.integer({ min: 0, max: 4_294_967_295 }),
      fc.constantFrom(...WORD_NAMESPACES),
      (id, drawingId, namespace) => {
        const parts = new Map([
          [
            "word/document.xml",
            `<x:document xmlns:x="${namespace}" xmlns:d="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><x:commentReference x:id="${id}"/><x:numId x:val="${id}"/><d:docPr id="${drawingId}"/></x:document>`,
          ],
          [
            "word/comments.xml",
            `<x:comments xmlns:x="${namespace}"><x:comment x:id="${id}"/></x:comments>`,
          ],
        ]);
        let visited = false;
        expect(
          normalizeImportedNumericIds(parts, {
            onParsedDocument: () => {
              visited = true;
            },
          }),
        ).toBe(parts);
        expect(visited).toBe(false);
      },
    ),
    propertyConfig({ numRuns: 100 }),
  );
});

test("the retained repair tree equals an independent parse of the normalized XML", () => {
  fc.assert(
    fc.property(
      invalidId,
      fc.constantFrom(...WORD_NAMESPACES),
      fc.boolean(),
      (id, namespace, encoded) => {
        const spelling = encoded
          ? [...id].map((character) => `&#${character.charCodeAt(0)};`).join("")
          : id;
        const xml = `<x:document xmlns:x="${namespace}"><x:body><x:p><x:bookmarkStart x:id="1"/><x:commentRangeStart x:id='${spelling}'/><x:r><x:rPr><x:color x:val="123456"/></x:rPr><x:t>A&amp;B</x:t></x:r><x:commentRangeEnd x:id="${id}"/></x:p></x:body></x:document>`;
        let tree: XmlElement | undefined;
        const parts = new Map([["word/document.xml", xml]]);
        const normalized = normalizeImportedNumericIds(parts, {
          onParsedDocument: (parsed) => {
            tree = parsed;
          },
        });
        const rewritten = normalized.get("word/document.xml");
        expect(rewritten).toBeDefined();
        expect(tree).toBeDefined();
        if (rewritten === undefined) return;
        expect(tree).toEqual(parseXmlWithFastXmlParser(rewritten));
        expect(parts.get("word/document.xml")).toBe(xml);
      },
    ),
    propertyConfig({ numRuns: 100 }),
  );
});

test("imported numeric identities stay paired, avoid occupied ids, and reach a fixed point across spaces", () => {
  fc.assert(
    fc.property(invalidId, fc.constantFrom(...WORD_NAMESPACES), (id, namespace) => {
      const parts = new Map([
        [
          "word/document.xml",
          `<x:document xmlns:x="${namespace}"><x:body><x:bookmarkStart x:id="1" x:name="kept"/><x:bookmarkEnd x:id="1"/><x:commentRangeStart x:id="${id}"/><x:commentReference x:id="${id}"/><x:commentRangeEnd x:id="${id}"/><x:footnoteReference x:id="${id}"/><x:endnoteReference x:id="${id}"/><x:numId x:val="${id}"/><x:id x:val="${id}"/></x:body></x:document>`,
        ],
        [
          "word/comments.xml",
          `<x:comments xmlns:x="${namespace}"><x:comment x:id="${id}"><x:p/></x:comment></x:comments>`,
        ],
        [
          "word/footnotes.xml",
          `<x:footnotes xmlns:x="${namespace}"><x:footnote x:id="-1" x:type="separator"><x:p/></x:footnote><x:footnote x:id="0" x:type="continuationSeparator"><x:p/></x:footnote><x:footnote x:id="${id}"><x:p/></x:footnote></x:footnotes>`,
        ],
        [
          "word/endnotes.xml",
          `<x:endnotes xmlns:x="${namespace}"><x:endnote x:id="${id}"><x:p/></x:endnote></x:endnotes>`,
        ],
        [
          "word/numbering.xml",
          `<x:numbering xmlns:x="${namespace}"><x:abstractNum x:abstractNumId="${id}"><x:lvl x:ilvl="0"><x:lvlPicBulletId x:val="${id}"/></x:lvl></x:abstractNum><x:num x:numId="${id}"><x:abstractNumId x:val="${id}"/></x:num><x:numPicBullet x:numPicBulletId="${id}"/></x:numbering>`,
        ],
      ]);
      const normalized = normalizeImportedNumericIds(parts);
      for (const [path, xml] of normalized)
        expect(() => assertValidOoxmlNumericIds(xml, path)).not.toThrow();
      expect(normalized.get("word/document.xml")).toBe(
        parts
          .get("word/document.xml")
          ?.replaceAll(`x:id="${id}"`, 'x:id="2"')
          .replace('<x:footnoteReference x:id="2"/>', '<x:footnoteReference x:id="1"/>')
          .replace('<x:endnoteReference x:id="2"/>', '<x:endnoteReference x:id="1"/>')
          .replaceAll(`x:val="${id}"`, 'x:val="1"'),
      );
      expect(normalized.get("word/comments.xml")).toBe(
        parts.get("word/comments.xml")?.replaceAll(`x:id="${id}"`, 'x:id="2"'),
      );
      expect(normalized.get("word/footnotes.xml")).toBe(
        parts.get("word/footnotes.xml")?.replaceAll(`x:id="${id}"`, 'x:id="1"'),
      );
      expect(normalized.get("word/endnotes.xml")).toBe(
        parts.get("word/endnotes.xml")?.replaceAll(`x:id="${id}"`, 'x:id="1"'),
      );
      expect(normalized.get("word/numbering.xml")).toBe(
        parts.get("word/numbering.xml")?.replaceAll(id, "1"),
      );
      expect(normalizeImportedNumericIds(normalized)).toEqual(normalized);
      expect(normalizeImportedNumericIds(new Map([...parts].toReversed()))).toEqual(normalized);
    }),
    propertyConfig({ numRuns: 40 }),
  );
});

test("normalization follows nested namespace bindings and canonical integer aliases, preserving every other byte", () => {
  const xml = `<?xml version="1.0"?><root xmlns:x="${WORD_NAMESPACES[0]}" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><!-- retained --><x:bookmarkStart x:id="2147483648" x:name="marker"/><x:bookmarkEnd x:id="+02147483648"/><wp:docPr id="4294967296"/><x:permStart x:id="everyone"/><foreign xmlns:x="urn:other"><x:bookmarkStart x:id="2147483648"/></foreign></root>`;
  const normalized = normalizeImportedNumericIds(new Map([["word/document.xml", xml]]));
  expect(normalized.get("word/document.xml")).toBe(
    xml
      .replace('x:id="2147483648"', 'x:id="1"')
      .replace('x:id="+02147483648"', 'x:id="1"')
      .replace('id="4294967296"', 'id="1"'),
  );
});

test("normalization splices decoded integer aliases without changing surrounding entity spellings", () => {
  const characterReference = fc.constantFrom("decimal", "hexadecimal");
  const encode = (value: string, spelling: string) =>
    [...value]
      .map((character) => {
        const codePoint = character.charCodeAt(0);
        return spelling === "decimal" ? `&#${codePoint};` : `&#x${codePoint.toString(16)};`;
      })
      .join("");
  fc.assert(
    fc.property(
      invalidId,
      fc.constantFrom(...WORD_NAMESPACES),
      characterReference,
      (id, namespace, spelling) => {
        const encodedId = encode(id, spelling);
        const encodedNamespace = encode(namespace, spelling);
        const xml = `<root xmlns:x="${encodedNamespace}"><!-- keep --><x:bookmarkStart x:name="A&#38;B" x:id='1'/><x:bookmarkEnd x:id='1'/><x:commentRangeStart x:id='${encodedId}'/><x:commentReference x:id="${id}"/><x:commentRangeEnd x:id='${encodedId}'/><x:num x:numId='${encodedId}' x:abstractNumId="${id}"/></root>`;
        const parts = new Map([["word/document.xml", xml]]);
        const normalized = normalizeImportedNumericIds(parts);
        expect(normalized.get("word/document.xml")).toBe(
          xml
            .replaceAll(`x:id='${encodedId}'`, "x:id='2'")
            .replace(`x:id="${id}"`, 'x:id="2"')
            .replace(`x:numId='${encodedId}'`, "x:numId='1'")
            .replace(`x:abstractNumId="${id}"`, 'x:abstractNumId="1"'),
        );
        expect(normalizeImportedNumericIds(normalized)).toEqual(normalized);
        expect(parts.get("word/document.xml")).toBe(xml);
      },
    ),
    propertyConfig({ numRuns: 40 }),
  );
});

test("range repair leaves malformed lexical identities for the parser's existing drop-and-warn path", () => {
  const malformedId = fc.oneof(
    fc.constantFrom("", "bare", "1.5", "1e3", "+", "--1"),
    fc.integer().map((prefix) => `${prefix}invalid`),
  );
  assertProperty(
    fc.property(
      malformedId,
      invalidId,
      fc.constantFrom(...WORD_NAMESPACES),
      (malformed, id, namespace) => {
        const xml = `<x:document xmlns:x="${namespace}" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><x:comment x:id="${malformed}"/><x:bookmarkStart x:id="${malformed}"/><x:numId x:val="${malformed}"/><wp:docPr id="${malformed}"/><x:comment x:id="${id}"/></x:document>`;
        const parts = new Map([["word/document.xml", xml]]);
        const normalized = normalizeImportedNumericIds(parts);
        expect(normalized.get("word/document.xml")).toBe(xml.replace(`x:id="${id}"`, 'x:id="1"'));
        expect(normalizeImportedNumericIds(normalized)).toEqual(normalized);
      },
    ),
    { numRuns: 40 },
  );
});

test("invalid imported notes are readable and saveable through both save paths", async () => {
  const zip = await JSZip.loadAsync(await createEmptyDocx());
  zip.file(
    "word/footnotes.xml",
    `<x:footnotes xmlns:x="${WORD_NAMESPACES[0]}"><x:footnote x:type="separator" x:id="2147483648"><x:p/></x:footnote></x:footnotes>`,
  );
  const source = await zip.generateAsync({ type: "arraybuffer" });
  const document = await parseDocx(source, { preloadFonts: false });
  const raw = await JSZip.loadAsync(document.originalBuffer);
  expect(await raw.file("word/footnotes.xml")?.async("text")).toContain('x:id="1"');
  const full = await repackDocx(document, { updateModifiedDate: false });
  const selective = await attemptSelectiveSave(document, document.originalBuffer, {
    changedParaIds: new Set(),
    structuralChange: false,
    hasUntrackedChanges: false,
  });
  expect(selective).not.toBeNull();
  for (const saved of [full, selective]) {
    if (saved === null) continue;
    const reopened = await parseDocx(saved, { preloadFonts: false });
    expect(reopened.package.footnotes).toEqual(document.package.footnotes);
  }
});

const annotationCensus = (document: Document) => ({
  comments: document.package.document.comments?.map(({ id }) => id),
  markers: document.package.document.content.flatMap((block) => {
    if (block.type !== "paragraph") return [];
    return block.content.flatMap((content) => {
      switch (content.type) {
        case "bookmarkStart":
        case "bookmarkEnd":
        case "commentRangeStart":
        case "commentRangeEnd":
        case "commentReference":
          return [content.id];
        case "insertion":
          return [content.info.id];
        default:
          return [];
      }
    });
  }),
});

test("the same invalid source id in separate annotation families stays stable across both save paths", async () => {
  await fc.assert(
    fc.asyncProperty(invalidId, async (id) => {
      const document = createEmptyDocument();
      document.package.document.comments = [
        {
          id: 1,
          author: "Reviewer",
          content: [
            {
              type: "paragraph",
              content: [{ type: "run", content: [{ type: "text", text: "Remark" }] }],
            },
          ],
        },
      ];
      document.package.document.content = [
        {
          type: "paragraph",
          content: [
            { type: "bookmarkStart", id: 1, name: "marker" },
            { type: "commentRangeStart", id: 1 },
            {
              type: "insertion",
              info: { id: 1, author: "Reviewer" },
              content: [{ type: "run", content: [{ type: "text", text: "Anchor" }] }],
            },
            { type: "commentRangeEnd", id: 1 },
            { type: "bookmarkEnd", id: 1 },
          ],
        },
      ];
      const zip = await JSZip.loadAsync(await createDocx(document));
      for (const path of ["word/document.xml", "word/comments.xml"]) {
        const xml = await zip.file(path)?.async("text");
        expect(xml).toBeDefined();
        if (xml === undefined) return;
        zip.file(path, xml.replaceAll(/w:id="\d+"/gu, `w:id="${id}"`));
      }
      const imported = await parseDocx(await zip.generateAsync({ type: "arraybuffer" }), {
        preloadFonts: false,
      });
      const before = annotationCensus(imported);
      expect(new Set(before.markers).size).toBe(3);
      const selective = await attemptSelectiveSave(imported, imported.originalBuffer, {
        changedParaIds: new Set(),
        structuralChange: false,
        hasUntrackedChanges: false,
      });
      expect(selective).not.toBeNull();
      for (const saved of [await repackDocx(imported, { updateModifiedDate: false }), selective]) {
        if (saved === null) continue;
        const reopened = await parseDocx(saved, { preloadFonts: false });
        expect(annotationCensus(reopened)).toEqual(before);
      }
    }),
    propertyConfig({ numRuns: 8 }),
  );
});
