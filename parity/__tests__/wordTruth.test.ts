import { describe, expect, test } from "bun:test";
import path from "node:path";

import JSZip from "jszip";

import {
  buildCloseStagedDocumentScript,
  buildExportScript,
  isParityStagedDocumentPath,
  stageableDocxBytes,
  stripEditRestrictions,
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
    expect(script).toContain("set openDocuments to {}");
    expect(script).toContain("set candidatePath to POSIX path of");
    expect(script).toContain("if candidatePath is stagedDocumentPath then");
    expect(script).toContain("set theDoc to contents of candidateDocument");
    expect(script).toContain("set print revisions of theDoc to false");
    expect(script).toContain("set documentView to view of active window of theDoc");
    expect(script).toContain("set revisions view of documentView to revisions view final");
    expect(script).toContain("set show revisions and comments of documentView to false");
    expect(script).not.toContain("set revisions mode of documentView");
    expect(script).not.toContain("active document");
  });

  test("waits up to two minutes for a large staged document to open", () => {
    const script = buildExportScript({
      docxPath: path.join(wordContainerTmp, "parity-test-token.docx"),
      pdfPath: path.join(wordContainerTmp, "output.pdf"),
    });

    const attempts = Number(/repeat (\d+) times/.exec(script)?.[1]);
    const delaySeconds = Number(/delay ([\d.]+)/.exec(script)?.[1]);
    expect(delaySeconds).toBeGreaterThanOrEqual(0.5);
    expect(delaySeconds).toBeLessThanOrEqual(1);
    expect(attempts * delaySeconds).toBeGreaterThanOrEqual(120);
    expect(script).toContain("within 120 seconds of opening it");
    const appleEventTimeout = Number(/with timeout of (\d+) seconds/.exec(script)?.[1]);
    expect(appleEventTimeout).toBeGreaterThan(attempts * delaySeconds);
  });

  test("exports All Markup with inline revisions enabled for comparison", () => {
    const stagedPath = path.join(wordContainerTmp, "parity-test-token.docx");
    const script = buildExportScript({
      docxPath: stagedPath,
      pdfPath: path.join(wordContainerTmp, "output.pdf"),
      reviewView: "all-markup",
    });

    expect(script).toContain("set print revisions of theDoc to true");
    expect(script).toContain("set revisions view of documentView to revisions view final");
    expect(script).toContain("set revisions mode of documentView to in line revisions");
    expect(script).toContain("set show insertions and deletions of documentView to true");
    expect(script).toContain("set show comments of documentView to false");
    expect(script).toContain("set show revisions and comments of documentView to true");
    expect(script.indexOf("set show revisions and comments of documentView to true")).toBeLessThan(
      script.indexOf("set show comments of documentView to false"),
    );
  });

  test("cleanup closes only a document whose full path matches staging", () => {
    const stagedPath = path.join(wordContainerTmp, "parity-test-token.docx");
    const script = buildCloseStagedDocumentScript(stagedPath);

    expect(script).toContain("if candidatePath is stagedDocumentPath then");
    expect(script).toContain("close candidateDocument saving no");
    expect(script).not.toContain("set openDocuments to {}");
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

const W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

const settingsXml = (body: string): string =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings ${W_NS}><w:zoom w:percent="100"/>${body}<w:defaultTabStop w:val="720"/></w:settings>`;

const buildDocx = async (settings: string | undefined): Promise<Uint8Array> => {
  const zip = new JSZip();
  zip.file(
    "word/document.xml",
    `<w:document ${W_NS}><w:body><w:p><w:r><w:t>Text</w:t></w:r></w:p></w:body></w:document>`,
  );
  if (settings !== undefined) zip.file("word/settings.xml", settings);
  return await zip.generateAsync({ type: "uint8array" });
};

describe("export staging", () => {
  test("drops document and write protection from the settings part", () => {
    const stripped = stripEditRestrictions(
      settingsXml(
        '<w:writeProtection w:recommended="1"/><w:documentProtection w:edit="forms" w:enforcement="1"/>',
      ),
    );

    expect(stripped).toBe(settingsXml(""));
  });

  test("leaves unrestricted settings untouched", () => {
    expect(stripEditRestrictions(settingsXml("<w:trackRevisions/>"))).toBeUndefined();
    expect(
      stripEditRestrictions(settingsXml('<w:documentProtectionNote w:val="1"/>')),
    ).toBeUndefined();
  });

  test("stages a restricted document as an unrestricted copy", async () => {
    const source = await buildDocx(
      settingsXml('<w:documentProtection w:edit="readOnly" w:enforcement="1"/>'),
    );
    const sourceCopy = source.slice();

    const staged = await stageableDocxBytes(source);
    const stagedZip = await JSZip.loadAsync(staged);

    expect(await stagedZip.file("word/settings.xml")?.async("string")).toBe(settingsXml(""));
    expect(await stagedZip.file("word/document.xml")?.async("string")).toContain("<w:t>Text</w:t>");
    expect(source).toEqual(sourceCopy);
  });

  test("stages unrestricted documents byte for byte", async () => {
    const withSettings = await buildDocx(settingsXml("<w:trackRevisions/>"));
    const withoutSettings = await buildDocx(undefined);

    expect(await stageableDocxBytes(withSettings)).toBe(withSettings);
    expect(await stageableDocxBytes(withoutSettings)).toBe(withoutSettings);
  });
});
