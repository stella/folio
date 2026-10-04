/**
 * End-to-end watermark header coverage: a title-page section without its own
 * first-page header should still show the watermark on the cover page. The
 * coverage transform creates the header in the model; the save pipeline
 * materializes it into a real part. eigenpal/docx-editor#684 (BUG2).
 */
import { describe, expect, test } from "bun:test";
import JSZip from "jszip";
import fc from "fast-check";
import { assertProperty, propertyTestTimeout } from "../../../../test/property-testing";
import { getParagraphPropertySource } from "./paragraphPropertySource";

import type { HeaderFooter } from "../types/document";
import { setDocumentWatermark, ensureWatermarkHeaderCoverage } from "../watermark/index";
import { parseDocx } from "./parser";
import { createEmptyDocx, repackDocx, validateDocx } from "./rezip";

describe("watermark header coverage on save (eigenpal #684)", () => {
  test("a titlePg section without a first header gets one carrying the watermark", async () => {
    const base = await createEmptyDocx();
    const doc = await parseDocx(base, { preloadFonts: false });

    // A default header plus a title-page final section that lacks a first header.
    const defaultHeader: HeaderFooter = {
      type: "header",
      hdrFtrType: "default",
      content: [],
    };
    doc.package.headers = new Map([["rIdHdr", defaultHeader]]);
    doc.package.document.finalSectionProperties = {
      ...doc.package.document.finalSectionProperties,
      titlePg: true,
      headerReferences: [{ type: "default", rId: "rIdHdr" }],
    };

    const withWatermark = setDocumentWatermark(doc, {
      kind: "text",
      text: "CONFIDENTIAL",
    });
    const out = await repackDocx(withWatermark, { updateModifiedDate: false });

    expect((await validateDocx(out)).valid).toBe(true);

    const zip = await JSZip.loadAsync(out);
    const headerFiles = Object.keys(zip.files).filter((p) => /^word\/header\d+\.xml$/u.test(p));
    // Default header + the coverage-created first-page header.
    expect(headerFiles).toHaveLength(2);
    for (const path of headerFiles) {
      // oxlint-disable-next-line no-await-in-loop -- sequential test assertion over each extracted header entry
      expect(await zip.file(path)!.async("text")).toContain("CONFIDENTIAL");
    }

    const docXml = await zip.file("word/document.xml")!.async("text");
    expect(docXml).toMatch(/<w:headerReference[^>]*w:type="first"/u);

    const reparsed = await parseDocx(out, { preloadFonts: false });
    expect(reparsed.package.headers?.size).toBe(2);
  });
});

// The old coverage fixture had only a final section, so it missed mid-body source transfer.
test(
  "generated legacy watermark coverage preserves private paragraph property sources",
  async () => {
    await assertProperty(
      fc.asyncProperty(fc.boolean(), async (evenPages) => {
        for (const mode of ["set", "coverage"] as const) {
          const zip = await JSZip.loadAsync(await createEmptyDocx());
          zip.file(
            "word/document.xml",
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body><w:p w14:paraId="00000001"><w:pPr><w:keepNext/><w:sectPr><w:titlePg/></w:sectPr></w:pPr><w:r><w:t>First section</w:t></w:r></w:p><w:p w14:paraId="00000002"><w:r><w:t>Last section</w:t></w:r></w:p><w:sectPr/></w:body></w:document>',
          );
          const source = await parseDocx(await zip.generateAsync({ type: "arraybuffer" }), {
            preloadFonts: false,
          });
          source.package.settings = {
            ...source.package.settings,
            defaultTabStop: 720,
            evenAndOddHeaders: evenPages,
          };
          const before = source.package.document.content.at(0);
          if (before?.type !== "paragraph") throw new Error("Expected parsed section carrier");
          const capture = getParagraphPropertySource(before);
          expect(capture).toBeDefined();
          const watermark = { kind: "text", text: "DRAFT" } as const;
          const changed =
            mode === "set"
              ? setDocumentWatermark(source, watermark)
              : ensureWatermarkHeaderCoverage(source, watermark);
          const after = changed.package.document.content.at(0);
          if (after?.type !== "paragraph") throw new Error("Expected changed section carrier");
          expect(after).not.toBe(before);
          expect(after.sectionProperties?.headerReferences).toHaveLength(evenPages ? 3 : 2);
          expect(getParagraphPropertySource(after)).toBe(capture);
        }
      }),
      { seed: 20261020, numRuns: 12 },
    );
  },
  propertyTestTimeout(30_000),
);
