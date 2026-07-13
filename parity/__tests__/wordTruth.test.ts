import { describe, expect, test } from "bun:test";
import path from "node:path";

import {
  buildCloseStagedDocumentScript,
  buildExportScript,
  isParityStagedDocumentPath,
} from "../wordTruth";

const wordContainerTmp = path.join(
  process.env["HOME"] ?? "",
  "Library",
  "Containers",
  "com.microsoft.Word",
  "Data",
  "tmp",
);

describe("Word automation scripts", () => {
  test("binds the opened document by its exact staged path", () => {
    const stagedPath = path.join(wordContainerTmp, "parity-test-token.docx");
    const script = buildExportScript({
      docxPath: stagedPath,
      pdfPath: path.join(wordContainerTmp, "output.pdf"),
    });

    expect(script).toContain("set openDocuments to get every document");
    expect(script).toContain("set candidatePath to POSIX path of");
    expect(script).toContain("if candidatePath is stagedDocumentPath then");
    expect(script).toContain("set theDoc to contents of candidateDocument");
    expect(script).not.toContain("active document");
  });

  test("cleanup closes only a document whose full path matches staging", () => {
    const stagedPath = path.join(wordContainerTmp, "parity-test-token.docx");
    const script = buildCloseStagedDocumentScript(stagedPath);

    expect(script).toContain("if candidatePath is stagedDocumentPath then");
    expect(script).toContain("close candidateDocument saving no");
    expect(script).not.toContain("close every document");
    expect(script).not.toContain("active document");
  });

  test("recognizes only parity-owned docx paths in Word staging", () => {
    expect(isParityStagedDocumentPath(path.join(wordContainerTmp, "parity-test.docx"))).toBe(true);
    expect(isParityStagedDocumentPath(path.join(wordContainerTmp, "ordinary.docx"))).toBe(false);
    expect(isParityStagedDocumentPath("/tmp/parity-test.docx")).toBe(false);
    expect(isParityStagedDocumentPath(path.join(wordContainerTmp, "parity-test.pdf"))).toBe(false);
  });

  test("escapes staged paths before AppleScript interpolation", () => {
    const stagedPath = path.join(wordContainerTmp, 'parity-a"b\\c.docx');
    const script = buildExportScript({
      docxPath: stagedPath,
      pdfPath: path.join(wordContainerTmp, 'a"b\\c.pdf'),
    });

    expect(script).toContain('parity-a\\"b\\\\c.docx');
    expect(script).toContain('a\\"b\\\\c.pdf');
  });
});
