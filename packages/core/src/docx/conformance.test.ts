import { describe, expect, test } from "bun:test";
import JSZip from "jszip";
import { DOCX_CONFORMANCE_CLASSES } from "@stll/docx-core/model";

import { createEmptyDocument } from "../utils/createDocument";
import { parseDocx } from "./parser";
import { detectDocxConformanceClass } from "./conformance";

const STRICT_MAIN_NAMESPACE = "http://purl.oclc.org/ooxml/wordprocessingml/main";
const TRANSITIONAL_MAIN_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

describe("DOCX conformance detection", () => {
  test("detects Strict and Transitional main document namespaces", () => {
    expect(
      detectDocxConformanceClass(
        `<x:document xmlns:x="${STRICT_MAIN_NAMESPACE}"><x:body/></x:document>`,
      ),
    ).toBe(DOCX_CONFORMANCE_CLASSES.STRICT);
    expect(
      detectDocxConformanceClass(
        `<document xmlns="${TRANSITIONAL_MAIN_NAMESPACE}"><body/></document>`,
      ),
    ).toBe(DOCX_CONFORMANCE_CLASSES.TRANSITIONAL);
  });

  test("does not infer a profile from an unrelated namespace declaration", () => {
    expect(
      detectDocxConformanceClass(
        `<x:document xmlns:x="urn:example" xmlns:w="${STRICT_MAIN_NAMESPACE}"/>`,
      ),
    ).toBe(DOCX_CONFORMANCE_CLASSES.UNKNOWN);
    expect(detectDocxConformanceClass("<not-document/>")).toBe(DOCX_CONFORMANCE_CLASSES.UNKNOWN);
    expect(detectDocxConformanceClass(null)).toBe(DOCX_CONFORMANCE_CLASSES.UNKNOWN);
  });

  test("reads only the actual root start tag", () => {
    expect(
      detectDocxConformanceClass(
        `<?xml version="1.0"?>
        <!-- <x:document xmlns:x="${STRICT_MAIN_NAMESPACE}"/> -->
        <x:document data-comparison="1 > 0" xmlns:x = '${TRANSITIONAL_MAIN_NAMESPACE}'>
          <x:body xmlns:x="${STRICT_MAIN_NAMESPACE}"/>
        </x:document>`,
      ),
    ).toBe(DOCX_CONFORMANCE_CLASSES.TRANSITIONAL);
  });

  test("falls back to full parsing for uncommon content before the root", () => {
    expect(
      detectDocxConformanceClass(`not XML <x:document xmlns:x="${STRICT_MAIN_NAMESPACE}"/>`),
    ).toBe(DOCX_CONFORMANCE_CLASSES.STRICT);
    expect(
      detectDocxConformanceClass(
        `<!DOCTYPE x:document [<!ELEMENT x:document EMPTY>]><x:document xmlns:x="${STRICT_MAIN_NAMESPACE}"/>`,
      ),
    ).toBe(DOCX_CONFORMANCE_CLASSES.STRICT);
  });

  test("bounds the root scan without changing full-parser fallback behavior", () => {
    expect(
      detectDocxConformanceClass(
        `<!-- ${"padding".repeat(10_000)} --><x:document xmlns:x="${STRICT_MAIN_NAMESPACE}"/>`,
      ),
    ).toBe(DOCX_CONFORMANCE_CLASSES.STRICT);
  });

  test("stores the detected class on parsed package metadata", async () => {
    const zip = new JSZip();
    zip.file(
      "word/document.xml",
      `<x:document xmlns:x="${STRICT_MAIN_NAMESPACE}"><x:body><x:p><x:r><x:t>Text</x:t></x:r></x:p></x:body></x:document>`,
    );

    const document = await parseDocx(await zip.generateAsync({ type: "arraybuffer" }), {
      preloadFonts: false,
    });

    expect(document.package.conformanceClass).toBe(DOCX_CONFORMANCE_CLASSES.STRICT);
  });

  test("marks newly created packages as Transitional", () => {
    expect(createEmptyDocument().package.conformanceClass).toBe(
      DOCX_CONFORMANCE_CLASSES.TRANSITIONAL,
    );
  });

  test("scans a large quote-heavy tag prefix with no closing '>' in linear time", () => {
    // A regex-based scanner with `(?:[^>"']|"[^"]*"|'[^']*')*` here is
    // susceptible to catastrophic backtracking once no terminating `>`
    // exists; the hand-rolled scanner must resolve this in one linear pass.
    const hostile = `<x:document ${'"'.repeat(65_000)}`;
    const start = performance.now();

    expect(detectDocxConformanceClass(hostile)).toBe(DOCX_CONFORMANCE_CLASSES.UNKNOWN);
    expect(performance.now() - start).toBeLessThan(1000);
  });
});
