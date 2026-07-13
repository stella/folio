import path from "node:path";

import { describe, expect, test } from "bun:test";

import { buildLibreOfficeExportArgs } from "../libreOfficeTruth";

describe("LibreOffice reference renderer", () => {
  test("uses headless PDF export with an isolated user profile", () => {
    const args = buildLibreOfficeExportArgs({
      binary: "/Applications/LibreOffice.app/Contents/MacOS/soffice",
      inputPath: "/tmp/folio/input.docx",
      outputDir: "/tmp/folio/output",
      profileDir: "/tmp/folio/profile with spaces",
    });

    expect(args).toContain("--headless");
    expect(args).toContain("pdf:writer_pdf_Export");
    expect(args).toContain("/tmp/folio/input.docx");
    expect(args).toContain("/tmp/folio/output");

    const profileArg = args.find((arg) => arg.startsWith("-env:UserInstallation="));
    expect(profileArg).toBeDefined();
    expect(profileArg).toContain("file://");
    expect(profileArg).toContain("profile%20with%20spaces");
  });

  test("passes paths as separate process arguments", () => {
    const inputPath = path.join("/tmp", "folio fixtures", "contract.docx");
    const args = buildLibreOfficeExportArgs({
      binary: "soffice",
      inputPath,
      outputDir: "/tmp/output with spaces",
      profileDir: "/tmp/profile",
    });

    expect(args.at(-1)).toBe(inputPath);
    expect(args.at(-2)).toBe("/tmp/output with spaces");
  });
});
