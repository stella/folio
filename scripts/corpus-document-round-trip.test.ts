import { describe, expect, test } from "bun:test";
import { parseDocx } from "@stll/folio-core/docx/parser";
import JSZip from "jszip";

import { FAMILY_BASELINE_FAMILIES, familyBaselinePath } from "./lib/corpus-family-baseline";
import {
  CORPUS_INVARIANT_FAMILIES,
  DEFAULT_INVARIANT_BUDGET_MS,
  isGatingFamily,
} from "./lib/corpus-invariants/contract";
import {
  documentPartDifferences,
  runDocumentRoundTripInvariant,
  withoutFirstBodyParagraph,
} from "./lib/corpus-invariants/document-round-trip";

const WORD = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const STRICT_WORD = "http://purl.oclc.org/ooxml/wordprocessingml/main";
const CONTENT_TYPES =
  '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>';
const RELS =
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>';

const packageBuffer = async (): Promise<ArrayBuffer> => {
  const zip = new JSZip();
  zip.file("[Content_Types].xml", CONTENT_TYPES);
  zip.file("_rels/.rels", RELS);
  zip.file(
    "word/document.xml",
    `<w:document xmlns:w="${WORD}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body><w:p w14:paraId="60000001"><w:r><w:t>Hello</w:t></w:r></w:p><w:p w14:paraId="60000002"><w:r><w:t>Untouched</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`,
  );
  zip.file("custom/data.bin", new Uint8Array([0, 255, 1]));
  return zip.generateAsync({ type: "arraybuffer" });
};

describe("Document round-trip ownership", () => {
  test("the family gates and owns a separate baseline path", () => {
    const family = CORPUS_INVARIANT_FAMILIES.documentRoundTrip;
    expect(isGatingFamily(family)).toBe(true);
    expect(FAMILY_BASELINE_FAMILIES).toContain(family);
    expect(familyBaselinePath(family)).toEndWith("corpus/baselines/document-round-trip.json");
  });
});

describe("the original-part oracle", () => {
  test("equal-size binary changes, additions, and removals all remain visible", () => {
    const before = new Map([
      ["custom/data.bin", new Uint8Array([0, 255])],
      ["custom/removed.xml", new Uint8Array([1])],
    ]);
    const after = new Map([
      ["custom/data.bin", new Uint8Array([0, 254])],
      ["custom/added.xml", new Uint8Array([1])],
    ]);
    expect(documentPartDifferences(before, after)).toEqual([
      "a part changed bytes: custom/data.bin",
      "a part was added: custom/added.xml",
      "a part was removed: custom/removed.xml",
    ]);
    expect(documentPartDifferences(before, before)).toEqual([]);
  });

  test("only the declared touched part is excluded", () => {
    const before = new Map([
      ["word/document.xml", new Uint8Array([1])],
      ["word/header7.xml", new Uint8Array([1])],
    ]);
    const after = new Map([
      ["word/document.xml", new Uint8Array([2])],
      ["word/header7.xml", new Uint8Array([2])],
    ]);
    expect(documentPartDifferences(before, after, "word/document.xml")).toEqual([
      "a part changed bytes: word/headerN.xml",
    ]);
  });
});

describe("the declared-block lexical oracle", () => {
  for (const namespace of [WORD, STRICT_WORD]) {
    for (const prefix of ["w", "q"]) {
      test(`resolves ${prefix} in ${namespace} and preserves every surrounding byte`, () => {
        const before = `<${prefix}:document xmlns:${prefix}="${namespace}" xmlns:x="urn:other"><${prefix}:body><!-- <${prefix}:p/> --><x:p/><${prefix}:p a="&gt;">one<![CDATA[<fake/>]]></${prefix}:p><${prefix}:p>two</${prefix}:p><${prefix}:sectPr/></${prefix}:body></${prefix}:document>`;
        const edited = before.replace("one", "changed");
        expect(withoutFirstBodyParagraph(before)).toBe(withoutFirstBodyParagraph(edited));
        expect(withoutFirstBodyParagraph(before)).not.toBe(
          withoutFirstBodyParagraph(edited.replace("two", "lost")),
        );
      });
    }
  }

  test("handles an empty touched paragraph and fails when it is absent", () => {
    const xml = `<w:document xmlns:w="${WORD}"><w:body><w:p/><w:p>second</w:p></w:body></w:document>`;
    expect(withoutFirstBodyParagraph(xml)).toBe(xml.replace("<w:p/>", ""));
    expect(() => withoutFirstBodyParagraph("<document><body><p/></body></document>")).toThrow();
  });

  test("a foreign body with the same lexical name cannot become the touched block", () => {
    const xml = `<q:document xmlns:q="${WORD}"><q:body xmlns:q="urn:other"><q:p>foreign</q:p></q:body><q:body><q:p>touched</q:p><q:p>untouched</q:p></q:body></q:document>`;
    expect(withoutFirstBodyParagraph(xml)).toBe(xml.replace("<q:p>touched</q:p>", ""));
  });
});

test("the Document path runs no-edit preservation and one declared edit on a tiny package", async () => {
  const buffer = await packageBuffer();
  const parsed = await parseDocx(buffer, { preloadFonts: false });
  const originalContent = structuredClone(parsed.package.document.content);
  const result = await runDocumentRoundTripInvariant({
    bytes: new Uint8Array(buffer),
    buffer,
    parsed,
    documentPart: "word/document.xml",
    budgetMs: DEFAULT_INVARIANT_BUDGET_MS,
  });
  expect(parsed.package.document.content).toEqual(originalContent);
  expect(Object.keys(result.timings)).toEqual([
    "tracked-parse",
    "original-parts",
    "no-edit-save",
    "no-edit-read",
    "canonical-edit",
    "edit-save",
    "edit-read",
  ]);
  expect(result.failures.every((failure) => failure.frame === "-")).toBe(true);
  expect(result.failures.some((failure) => failure.message.includes("custom/data.bin"))).toBe(
    false,
  );
});
