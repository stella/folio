/**
 * The full repack must hand back every part of the source package it did not
 * model, byte for byte — macro projects and ActiveX controls included, because
 * the document belongs to its author — and must never leave a relationship or
 * content-type override naming a part the output does not hold.
 *
 * The one thing a save refuses is an entry PATH that would escape the package,
 * which is a property of the archive rather than of the document.
 *
 * The fixtures are built here rather than checked in: a package with an
 * embedded workbook and a binary media part is three files and a relationship,
 * and spelling it out makes the shape under test readable.
 */

import { describe, expect, test } from "bun:test";
import fc from "fast-check";
import JSZip from "jszip";

import { propertyConfig, propertyTestTimeout } from "../../../../../test/property-testing";

import {
  isUnsafePackagePath,
  reconcilePackageReferences,
  removeUnsafeEntries,
} from "../packageParts";
import { parseDocx } from "../parser";
import { repackDocx } from "../rezip";
import { extractFile, getFileList, unzipDocx } from "../unzip";

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

const RELATIONSHIP_NAMESPACE = "http://schemas.openxmlformats.org/package/2006/relationships";
const OFFICE_RELATIONSHIP = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

const DOCUMENT_XML = `${XML_DECL}
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:r="${OFFICE_RELATIONSHIP}">
  <w:body>
    <w:p w14:paraId="60000001"><w:r><w:t>Hello world</w:t></w:r></w:p>
    <w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>
  </w:body>
</w:document>`;

const STYLES_XML = `${XML_DECL}<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>`;

const CORE_PROPS_XML = `${XML_DECL}<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>fx</dc:title><dcterms:modified xsi:type="dcterms:W3CDTF">2024-01-01T00:00:00.000Z</dcterms:modified></cp:coreProperties>`;

const PACKAGE_RELS = `${XML_DECL}<Relationships xmlns="${RELATIONSHIP_NAMESPACE}"><Relationship Id="rId1" Type="${OFFICE_RELATIONSHIP}/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/></Relationships>`;

/** One extra part to hang off `word/document.xml.rels`. */
type ExtraPart = {
  /** Path inside the package. */
  path: string;
  /** Bytes the repack must hand back unchanged. */
  bytes: Uint8Array;
  /** Relationship type URI, or null for a part nothing references. */
  relationshipType: string | null;
  /** `<Override>` content type, or null when a `<Default>` extension covers it. */
  contentType: string | null;
  /** Extension for a `<Default>` declaration, when the part needs one. */
  defaultExtension?: { extension: string; contentType: string };
};

const EMBEDDED_WORKBOOK: ExtraPart = {
  path: "word/embeddings/Microsoft_Excel_Worksheet.xlsx",
  bytes: new Uint8Array([0x50, 0x4b, 0x03, 0x04, 0x00, 0x11, 0x22, 0x33]),
  relationshipType: `${OFFICE_RELATIONSHIP}/package`,
  contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
};

const OLE_MEDIA: ExtraPart = {
  path: "word/media/image9.bin",
  bytes: new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]),
  relationshipType: `${OFFICE_RELATIONSHIP}/image`,
  contentType: null,
  defaultExtension: {
    extension: "bin",
    contentType: "application/vnd.openxmlformats-officedocument.oleObject",
  },
};

const DOCX_MAIN_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
/** A `.docm`'s main part. Word decides a document is macro-enabled by this. */
const DOCM_MAIN_CONTENT_TYPE = "application/vnd.ms-word.document.macroEnabled.main+xml";

type BuildPackageOptions = {
  extras?: readonly ExtraPart[];
  documentContentType?: string;
};

const buildPackage = async ({
  extras = [],
  documentContentType = DOCX_MAIN_CONTENT_TYPE,
}: BuildPackageOptions = {}): Promise<ArrayBuffer> => {
  const defaults = new Map<string, string>([
    ["rels", "application/vnd.openxmlformats-package.relationships+xml"],
    ["xml", "application/xml"],
  ]);
  const overrides: string[] = [
    `<Override PartName="/word/document.xml" ContentType="${documentContentType}"/>`,
    `<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>`,
    `<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>`,
  ];
  const relationships: string[] = [];

  extras.forEach((extra, index) => {
    if (extra.defaultExtension) {
      defaults.set(extra.defaultExtension.extension, extra.defaultExtension.contentType);
    }
    if (extra.contentType) {
      overrides.push(`<Override PartName="/${extra.path}" ContentType="${extra.contentType}"/>`);
    }
    if (extra.relationshipType) {
      relationships.push(
        `<Relationship Id="rIdExtra${String(index)}" Type="${extra.relationshipType}" Target="${extra.path.replace(/^word\//u, "")}"/>`,
      );
    }
  });

  const contentTypes = `${XML_DECL}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">${[
    ...defaults,
  ]
    .map(
      ([extension, contentType]) =>
        `<Default Extension="${extension}" ContentType="${contentType}"/>`,
    )
    .join("")}${overrides.join("")}</Types>`;

  const zip = new JSZip();
  zip.file("[Content_Types].xml", contentTypes);
  zip.file("_rels/.rels", PACKAGE_RELS);
  zip.file(
    "word/_rels/document.xml.rels",
    `${XML_DECL}<Relationships xmlns="${RELATIONSHIP_NAMESPACE}">${relationships.join("")}</Relationships>`,
  );
  zip.file("word/document.xml", DOCUMENT_XML);
  zip.file("word/styles.xml", STYLES_XML);
  zip.file("docProps/core.xml", CORE_PROPS_XML);
  for (const extra of extras) {
    zip.file(extra.path, extra.bytes);
  }
  return zip.generateAsync({ type: "arraybuffer" });
};

const repack = async (buffer: ArrayBuffer): Promise<JSZip> => {
  const document = await parseDocx(buffer, { detectVariables: false, preloadFonts: false });
  return JSZip.loadAsync(await repackDocx(document, { updateModifiedDate: false }));
};

const partBytes = async (zip: JSZip, path: string): Promise<Uint8Array | null> => {
  const entry = zip.file(path);
  return entry ? entry.async("uint8array") : null;
};

const MACRO_PROJECT: ExtraPart = {
  path: "word/vbaProject.bin",
  bytes: new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 0x5a, 0x5a, 0x5a, 0x5a]),
  relationshipType: "http://schemas.microsoft.com/office/2006/relationships/vbaProject",
  contentType: "application/vnd.ms-office.vbaProject",
};

const ACTIVEX_CONTROL: ExtraPart = {
  path: "word/activeX/activeX1.xml",
  bytes: new TextEncoder().encode('<ax:ocx xmlns:ax="urn:activex" ax:classid="{0002}"/>'),
  relationshipType: `${OFFICE_RELATIONSHIP}/control`,
  contentType: "application/vnd.ms-office.activeX+xml",
};

describe("the repack carries the parts the model does not represent", () => {
  test("an embedded workbook and a binary media part survive byte for byte", async () => {
    const extras = [EMBEDDED_WORKBOOK, OLE_MEDIA];
    const output = await repack(await buildPackage({ extras }));

    for (const extra of extras) {
      expect(await partBytes(output, extra.path)).toEqual(extra.bytes);
    }

    const rels = await output.file("word/_rels/document.xml.rels")?.async("text");
    expect(rels).toContain('Target="embeddings/Microsoft_Excel_Worksheet.xlsx"');
    expect(rels).toContain('Target="media/image9.bin"');

    const contentTypes = await output.file("[Content_Types].xml")?.async("text");
    expect(contentTypes).toContain('PartName="/word/embeddings/Microsoft_Excel_Worksheet.xlsx"');
    expect(contentTypes).toContain('Extension="bin"');
  });

  test("the output holds every part the source package declares", async () => {
    const output = await repack(
      await buildPackage({
        extras: [EMBEDDED_WORKBOOK, OLE_MEDIA, MACRO_PROJECT, ACTIVEX_CONTROL],
      }),
    );

    // Reconciliation reports exactly the declarations the package cannot
    // satisfy, so an empty repair is "every declared part is here".
    expect(await reconcilePackageReferences(output, 6)).toEqual({
      danglingRelationships: [],
      danglingOverrides: [],
    });
  });

  test("a macro project and an ActiveX control come back untouched", async () => {
    const extras = [MACRO_PROJECT, ACTIVEX_CONTROL];
    const output = await repack(await buildPackage({ extras }));

    for (const extra of extras) {
      expect(await partBytes(output, extra.path)).toEqual(extra.bytes);
    }

    const rels = await output.file("word/_rels/document.xml.rels")?.async("text");
    expect(rels).toContain('Target="vbaProject.bin"');
    expect(rels).toContain('Target="activeX/activeX1.xml"');

    const contentTypes = await output.file("[Content_Types].xml")?.async("text");
    expect(contentTypes).toContain('PartName="/word/vbaProject.bin"');
    expect(contentTypes).toContain('PartName="/word/activeX/activeX1.xml"');
  });

  test("a macro-enabled document stays macro-enabled", async () => {
    const output = await repack(
      await buildPackage({
        extras: [MACRO_PROJECT],
        documentContentType: DOCM_MAIN_CONTENT_TYPE,
      }),
    );

    const contentTypes = await output.file("[Content_Types].xml")?.async("text");
    expect(contentTypes).toContain(DOCM_MAIN_CONTENT_TYPE);
    expect(contentTypes).not.toContain(DOCX_MAIN_CONTENT_TYPE);
  });
});

describe("folio reads document content, it does not interpret it", () => {
  test("macro, ActiveX and OLE parts are never decoded into the model", async () => {
    const buffer = await buildPackage({
      extras: [MACRO_PROJECT, ACTIVEX_CONTROL, EMBEDDED_WORKBOOK, OLE_MEDIA],
    });
    const content = await unzipDocx(buffer);

    for (const extra of [MACRO_PROJECT, ACTIVEX_CONTROL, EMBEDDED_WORKBOOK, OLE_MEDIA]) {
      expect(content.allXml.has(extra.path)).toBe(false);
      expect(content.media.has(extra.path)).toBe(false);
      expect(await extractFile(content, extra.path)).toBeNull();
      expect(getFileList(content)).not.toContain(extra.path);
    }

    // The archive still holds them, byte for byte, for the save path to carry.
    expect(await partBytes(content.originalZip, MACRO_PROJECT.path)).toEqual(MACRO_PROJECT.bytes);
  });
});

describe("a path that would escape the package is refused", () => {
  test.each(["../outside.xml", "/absolute.xml", "word\\backslash.xml", "word/../../escape.bin"])(
    "%s is unsafe",
    (path) => {
      expect(isUnsafePackagePath(path)).toBe(true);
    },
  );

  test.each(["word/document.xml", "[Content_Types].xml", "word/embeddings/book.xlsx"])(
    "%s is safe",
    (path) => {
      expect(isUnsafePackagePath(path)).toBe(false);
    },
  );

  test("an escaping entry leaves with its relationship and its override", async () => {
    const zip = await JSZip.loadAsync(await buildPackage({ extras: [EMBEDDED_WORKBOOK] }));
    zip.file("../escape.bin", new Uint8Array([1, 2, 3]));
    zip.file(
      "word/_rels/document.xml.rels",
      `${XML_DECL}<Relationships xmlns="${RELATIONSHIP_NAMESPACE}"><Relationship Id="rIdEscape" Type="${OFFICE_RELATIONSHIP}/package" Target="../../escape.bin"/></Relationships>`,
    );

    removeUnsafeEntries(zip);
    expect(zip.file("../escape.bin")).toBeNull();
    expect(zip.file("word/embeddings/Microsoft_Excel_Worksheet.xlsx")).not.toBeNull();

    const repair = await reconcilePackageReferences(zip, 6);
    expect(repair.danglingRelationships).toEqual(["escape.bin"]);
    const rels = await zip.file("word/_rels/document.xml.rels")?.async("text");
    expect(rels).not.toContain("escape.bin");
  });
});

describe("a repacked package needs no further repair", () => {
  const extensionArb = fc.constantFrom("bin", "xlsx", "docx", "emf", "dat", "vml");
  const partArb = fc
    .tuple(
      fc.constantFrom("word/embeddings", "word/media", "customXml", "word/charts", "word/unknown"),
      fc.stringMatching(/^[a-z][a-z0-9]{0,7}$/u),
      extensionArb,
      fc.uint8Array({ minLength: 1, maxLength: 24 }),
      fc.boolean(),
      fc.boolean(),
    )
    .map(([directory, name, extension, bytes, referenced, overridden]) => ({
      path: `${directory}/${name}.${extension}`,
      bytes,
      relationshipType: referenced ? `${OFFICE_RELATIONSHIP}/customXml` : null,
      contentType: overridden ? "application/octet-stream" : null,
      defaultExtension: { extension, contentType: "application/octet-stream" },
    }));

  test(
    "every source part comes back with identical bytes and no dangling reference",
    async () => {
      await fc.assert(
        fc.asyncProperty(
          fc.uniqueArray(partArb, {
            minLength: 1,
            maxLength: 5,
            selector: (part) => part.path,
          }),
          async (extras) => {
            // A relationship target is resolved against `word/`, so only parts
            // under it can be referenced from `document.xml.rels`.
            const referenceable = extras.map((extra) => ({
              ...extra,
              relationshipType: extra.path.startsWith("word/") ? extra.relationshipType : null,
            }));
            const output = await repack(await buildPackage({ extras: referenceable }));

            for (const extra of referenceable) {
              expect(await partBytes(output, extra.path)).toEqual(extra.bytes);
            }

            // Reconciliation is a fixed point on a healthy package: a second
            // pass over the output finds nothing to remove.
            expect(await reconcilePackageReferences(output, 6)).toEqual({
              danglingRelationships: [],
              danglingOverrides: [],
            });
          },
        ),
        propertyConfig({ numRuns: 30 }),
      );
    },
    propertyTestTimeout(30_000),
  );
});
