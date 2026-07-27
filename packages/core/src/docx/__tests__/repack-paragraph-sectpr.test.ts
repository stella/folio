import { describe, expect, test } from "bun:test";
import { readFileSync } from "node:fs";
import path from "node:path";

import type { Paragraph } from "../../types/document";
import { parseDocx } from "../parser";
import { repackDocx } from "../rezip";

const FIXTURE_PATH = path.join(
  import.meta.dir,
  "__fixtures__",
  "regressions",
  "repack-paragraph-sectpr.docx",
);

const readFixture = (): ArrayBuffer => {
  const bytes = readFileSync(FIXTURE_PATH);
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
};

const countParagraphSectPrs = (
  blocks: readonly { type: string; sectionProperties?: unknown }[],
): number =>
  blocks.filter((b) => b.type === "paragraph" && b.sectionProperties !== undefined).length;

// Regression: this real-world contract places its body-level A4 sectPr before
// all content, followed by an empty paragraph-level section break. Repacking
// must canonicalize the element locations without reversing the effective A4
// then default-page sequence.
describe("repack preserves paragraph-level section properties", () => {
  test("round-tripping the fixture preserves section count and effective order", async () => {
    const buf = readFixture();
    const doc = await parseDocx(buf);

    const originalParagraphSectPrs = countParagraphSectPrs(doc.package.document.content);
    expect(originalParagraphSectPrs).toBeGreaterThan(0);
    const originalSectionBreak = doc.package.document.content.find(
      (block): block is Paragraph =>
        block.type === "paragraph" && block.sectionProperties !== undefined,
    );
    expect(originalSectionBreak?.sectionProperties).toMatchObject({
      pageWidth: 11_906,
      pageHeight: 16_838,
    });
    expect(doc.package.document.finalSectionProperties?.pageWidth).toBeUndefined();

    const repacked = await repackDocx(doc, { updateModifiedDate: false });
    const reparsed = await parseDocx(repacked);

    expect(countParagraphSectPrs(reparsed.package.document.content)).toBe(originalParagraphSectPrs);
    const reparsedSectionBreak = reparsed.package.document.content.find(
      (block): block is Paragraph =>
        block.type === "paragraph" && block.sectionProperties !== undefined,
    );
    expect(reparsedSectionBreak?.sectionProperties).toMatchObject({
      pageWidth: 11_906,
      pageHeight: 16_838,
    });
    expect(reparsed.package.document.finalSectionProperties?.pageWidth).toBeUndefined();
  });
});
