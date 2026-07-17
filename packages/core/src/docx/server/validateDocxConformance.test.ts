import { describe, expect, test } from "bun:test";
import JSZip from "jszip";

import { createEmptyDocx } from "../rezip";
import type {
  FolioDocxConformanceCheckId,
  FolioDocxConformanceReport,
} from "./validateDocxConformance";
import {
  FOLIO_DOCX_CONFORMANCE_PROFILE,
  FOLIO_DOCX_CONFORMANCE_REPORT_VERSION,
  validateDocxConformance,
} from "./validateDocxConformance";

const checkStatus = (report: FolioDocxConformanceReport, id: FolioDocxConformanceCheckId) =>
  report.checks.find((check) => check.id === id)?.status;

type MutatePackage = (zip: JSZip) => void;

const mutateEmptyPackage = async (mutate: MutatePackage): Promise<Uint8Array> => {
  const zip = await JSZip.loadAsync(await createEmptyDocx());
  mutate(zip);
  return await zip.generateAsync({ type: "uint8array" });
};

describe("validateDocxConformance", () => {
  test("reports the exact profile and unverified dimensions for a supported package", async () => {
    const report = await validateDocxConformance(await createEmptyDocx());

    expect(report).toMatchObject({
      version: FOLIO_DOCX_CONFORMANCE_REPORT_VERSION,
      profile: FOLIO_DOCX_CONFORMANCE_PROFILE,
      status: "conformant",
      conformanceClass: "transitional",
      issues: [],
    });
    expect(report.checks.every(({ status }) => status === "passed")).toBe(true);
    expect(report.unverifiedStandardsDimensions).toEqual([
      "complete-schema-constraints",
      "markup-compatibility-processing",
      "consumer-specific-rendering",
    ]);
  });

  test("fails a malformed XML part and does not run semantic checks", async () => {
    const bytes = await mutateEmptyPackage((zip) => {
      zip.file("word/document.xml", "<w:document>");
    });

    const report = await validateDocxConformance(bytes);

    expect(report.status).toBe("invalid");
    expect(checkStatus(report, "xml-well-formedness")).toBe("failed");
    expect(checkStatus(report, "package-roots")).toBe("not-run");
    expect(report.issues).toContainEqual({
      check: "xml-well-formedness",
      code: "xml-not-well-formed",
      message: "An XML package part is not well formed.",
      severity: "error",
      part: "word/document.xml",
    });
  });

  test("rejects document type declarations in XML package parts", async () => {
    const bytes = await mutateEmptyPackage((zip) => {
      zip.file("custom.xml", "<!DOCTYPE root [<!ELEMENT root EMPTY>]><root/>");
    });

    const report = await validateDocxConformance(bytes);

    expect(report.status).toBe("invalid");
    expect(report.issues.at(0)).toMatchObject({
      check: "xml-well-formedness",
      code: "xml-doctype-forbidden",
      part: "custom.xml",
    });
  });

  test("fails missing required parts with their package paths", async () => {
    const bytes = await mutateEmptyPackage((zip) => {
      zip.remove("_rels/.rels");
    });

    const report = await validateDocxConformance(bytes);

    expect(report.status).toBe("invalid");
    expect(checkStatus(report, "required-parts")).toBe("failed");
    expect(report.issues).toContainEqual({
      check: "required-parts",
      code: "required-part-missing",
      message: "A required package part is missing.",
      severity: "error",
      part: "_rels/.rels",
    });
  });

  test("fails an invalid main document package binding", async () => {
    const bytes = await mutateEmptyPackage((zip) => {
      zip.file(
        "_rels/.rels",
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>',
      );
    });

    const report = await validateDocxConformance(bytes);

    expect(report.status).toBe("invalid");
    expect(checkStatus(report, "package-roots")).toBe("failed");
    expect(report.issues.some(({ part }) => part === "_rels/.rels")).toBe(true);
  });

  test("accepts prefixed package vocabulary roots", async () => {
    const bytes = await mutateEmptyPackage((zip) => {
      zip.file(
        "[Content_Types].xml",
        `<ct:Types xmlns:ct="http://schemas.openxmlformats.org/package/2006/content-types">
          <ct:Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
        </ct:Types>`,
      );
      zip.file(
        "_rels/.rels",
        `<r:Relationships xmlns:r="http://schemas.openxmlformats.org/package/2006/relationships">
          <r:Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
        </r:Relationships>`,
      );
    });

    const report = await validateDocxConformance(bytes);

    expect(report.status).toBe("conformant");
    expect(checkStatus(report, "package-roots")).toBe("passed");
  });

  test("accepts a package-root-relative main document target", async () => {
    const bytes = await mutateEmptyPackage((zip) => {
      zip.file(
        "_rels/.rels",
        `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="/word/document.xml"/>
        </Relationships>`,
      );
    });

    const report = await validateDocxConformance(bytes);

    expect(report.status).toBe("conformant");
    expect(checkStatus(report, "package-roots")).toBe("passed");
  });

  test("reports an unknown main document namespace as indeterminate", async () => {
    const bytes = await mutateEmptyPackage((zip) => {
      zip.file(
        "word/document.xml",
        '<x:document xmlns:x="urn:example"><x:body><x:p/></x:body></x:document>',
      );
    });

    const report = await validateDocxConformance(bytes);

    expect(report.status).toBe("indeterminate");
    expect(report.conformanceClass).toBe("unknown");
    expect(checkStatus(report, "conformance-class")).toBe("indeterminate");
    expect(report.issues).toContainEqual({
      check: "conformance-class",
      code: "conformance-class-unknown",
      message: "The main document namespace does not identify a supported conformance class.",
      severity: "warning",
      part: "word/document.xml",
    });
  });

  test("reports encrypted containers as indeterminate before archive loading", async () => {
    const report = await validateDocxConformance(
      new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]),
    );

    expect(report.status).toBe("indeterminate");
    expect(checkStatus(report, "archive-safety")).toBe("indeterminate");
    expect(checkStatus(report, "required-parts")).toBe("not-run");
  });

  test("fails packages that exceed configured archive limits", async () => {
    const report = await validateDocxConformance(await createEmptyDocx(), {
      archive: { maxEntries: 1 },
    });

    expect(report.status).toBe("invalid");
    expect(checkStatus(report, "archive-safety")).toBe("failed");
    expect(report.issues.at(0)?.code).toBe("archive-too-many-entries");
  });
});
