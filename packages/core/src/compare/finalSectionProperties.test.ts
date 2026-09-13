import { describe, expect, test } from "bun:test";
import JSZip from "jszip";

import { FolioDocxReviewer } from "../ai-edits/headless";
import { createDocx } from "../docx/rezip";
import type { Document, SectionProperties } from "../types/document";
import { createEmptyDocument } from "../utils/createDocument";
import { compareDocx } from "./compare";

const OPTIONS = { author: "Section reviewer", timestamp: "2026-09-13T12:00:00.000Z" } as const;

const documentWithFinalSection = (finalSectionProperties: SectionProperties): Document => {
  const document = createEmptyDocument();
  return {
    ...document,
    package: { ...document.package, document: { ...document.package.document, finalSectionProperties } },
  };
};

const finalSection = (reviewer: FolioDocxReviewer): SectionProperties | undefined =>
  reviewer.toDocument().package.document.finalSectionProperties;

describe("final section property comparison", () => {
  test("tracks modeled body section properties and resolves through save and reopen", async () => {
    const base = documentWithFinalSection({
      pageWidth: 12_240,
      pageHeight: 15_840,
      marginLeft: 1_440,
      marginRight: 1_440,
      docGrid: { linePitch: 360 },
      titlePg: false,
    });
    const target = documentWithFinalSection({
      pageWidth: 11_000,
      pageHeight: 16_000,
      marginLeft: 1_800,
      marginRight: 1_800,
      columnCount: 2,
      columnSpace: 720,
      docGrid: { type: "lines", linePitch: 480, charSpace: 20 },
      titlePg: true,
      textDirection: "tbRl",
      formProtection: true,
      pageNumbering: { start: 4, format: "decimal" },
    });
    const [baseBuffer, targetBuffer] = await Promise.all([createDocx(base), createDocx(target)]);
    const result = await compareDocx(baseBuffer, targetBuffer, OPTIONS);
    if (result.isErr()) throw result.error;

    expect(result.value.verification).toEqual({ status: "verified" });
    expect(result.value.changes).toContainEqual({
      kind: "section-properties",
      location: { story: { type: "main" } },
    });
    const pendingXml = await (await JSZip.loadAsync(result.value.buffer))
      .file("word/document.xml")
      ?.async("text");
    expect(pendingXml).toContain("<w:sectPrChange");

    const accepting = await FolioDocxReviewer.fromBuffer(result.value.buffer);
    expect(accepting.acceptAll()).toBe(1);
    const acceptedReopened = await FolioDocxReviewer.fromBuffer(await accepting.toBuffer());
    expect(finalSection(acceptedReopened)).toEqual(finalSection(await FolioDocxReviewer.fromBuffer(targetBuffer)));

    const rejecting = await FolioDocxReviewer.fromBuffer(result.value.buffer);
    expect(rejecting.rejectAll()).toBe(1);
    const rejectedReopened = await FolioDocxReviewer.fromBuffer(await rejecting.toBuffer());
    expect(finalSection(rejectedReopened)).toEqual(finalSection(await FolioDocxReviewer.fromBuffer(baseBuffer)));
  });

  test("resolves individual final section revisions without hiding the remaining chain", async () => {
    const buffer = await createDocx(
      documentWithFinalSection({
        pageWidth: 12_240,
        propertyChanges: [
          {
            type: "sectionPropertyChange",
            info: { id: 71, author: OPTIONS.author, date: OPTIONS.timestamp },
            previousProperties: { pageWidth: 10_000 },
          },
          {
            type: "sectionPropertyChange",
            info: { id: 72, author: OPTIONS.author, date: OPTIONS.timestamp },
            previousProperties: { pageWidth: 11_000 },
          },
        ],
      }),
    );

    const rejectingEarlier = await FolioDocxReviewer.fromBuffer(buffer);
    expect(rejectingEarlier.getChanges().map(({ id }) => id)).toEqual([71, 72]);
    expect(rejectingEarlier.rejectChange(71)).toBe(true);
    expect(finalSection(rejectingEarlier)?.pageWidth).toBe(12_240);
    expect(rejectingEarlier.getChanges().map(({ id }) => id)).toEqual([72]);
    expect(rejectingEarlier.rejectChange(72)).toBe(true);
    expect(finalSection(rejectingEarlier)?.pageWidth).toBe(10_000);

    const acceptingEarlier = await FolioDocxReviewer.fromBuffer(buffer);
    expect(acceptingEarlier.acceptChange(71)).toBe(true);
    expect(acceptingEarlier.getChanges().map(({ id }) => id)).toEqual([72]);
    expect(acceptingEarlier.rejectChange(72)).toBe(true);
    expect(finalSection(acceptingEarlier)?.pageWidth).toBe(11_000);
  });
  test("resolves a revision shared by body content and final section properties", async () => {
    const info = { id: 73, author: OPTIONS.author, date: OPTIONS.timestamp };
    const document = documentWithFinalSection({
      pageWidth: 12_240,
      propertyChanges: [{
        type: "sectionPropertyChange",
        info,
        previousProperties: { pageWidth: 10_000 },
      }],
    });
    document.package.document.content = [{
      type: "paragraph",
      content: [{
        type: "insertion",
        info,
        content: [{ type: "run", content: [{ type: "text", text: "Inserted" }] }],
      }],
    }];
    const archive = await JSZip.loadAsync(await createDocx(document));
    const xml = await archive.file("word/document.xml")?.async("text");
    if (xml === undefined) throw new Error("Missing synthetic document part");
    archive.file("word/document.xml", xml.replace(/(<w:sectPrChange[^>]*w:id=")[^"]+/, "$173"));
    const buffer = await archive.generateAsync({ type: "arraybuffer" });
    for (const mode of ["accept", "reject"] as const) {
      const reviewer = await FolioDocxReviewer.fromBuffer(buffer);
      expect(mode === "accept" ? reviewer.acceptChange(73) : reviewer.rejectChange(73)).toBe(true);
      expect(reviewer.getChanges()).toEqual([]);
      expect(finalSection(reviewer)?.pageWidth).toBe(mode === "accept" ? 12_240 : 10_000);
      const reopened = await FolioDocxReviewer.fromBuffer(await reviewer.toBuffer());
      expect(reopened.getChanges()).toEqual([]);
      expect(finalSection(reopened)?.pageWidth).toBe(mode === "accept" ? 12_240 : 10_000);
    }
  });

});
