import { describe, expect, test } from "bun:test";
import JSZip from "jszip";

import { FolioDocxReviewer } from "../ai-edits/headless";
import { createEmptyDocx } from "./rezip";

const WORDPROCESSINGML_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

const sourceWithDeletedRunPropertyChange = async (
  wrapper: "del" | "moveFrom",
): Promise<ArrayBuffer> => {
  const zip = await JSZip.loadAsync(await createEmptyDocx());
  const documentPart = zip.file("word/document.xml");
  const documentXml = await documentPart?.async("text");
  if (!documentXml) {
    throw new Error("Expected the generated package to contain word/document.xml");
  }
  const body = `<w:body>
    <x:p xmlns:x="${WORDPROCESSINGML_NAMESPACE}">
      <x:${wrapper} x:id="91" x:author="Reviewer" x:date="2026-09-09T00:00:00Z">
        <x:r>
          <x:rPr>
            <x:b/>
            <x:rPrChange x:id="92" x:author="Formatter" x:date="2026-09-09T00:01:00Z">
              <x:rPr><x:i/></x:rPr>
            </x:rPrChange>
          </x:rPr>
          <x:delText>removed</x:delText>
        </x:r>
      </x:${wrapper}>
    </x:p>
    <w:sectPr/>
  </w:body>`;
  zip.file("word/document.xml", documentXml.replace(/<w:body>[\s\S]*<\/w:body>/u, body));
  return zip.generateAsync({ type: "arraybuffer" });
};

const documentXml = async (buffer: ArrayBuffer): Promise<string> => {
  const part = (await JSZip.loadAsync(buffer)).file("word/document.xml");
  if (!part) {
    throw new Error("Expected word/document.xml");
  }
  return part.async("text");
};

describe("deleted-run property-change lifecycle", () => {
  test.each(["del", "moveFrom"] as const)(
    "preserves x:%s rPrChange metadata through reopen and both resolutions",
    async (wrapper) => {
      const source = await sourceWithDeletedRunPropertyChange(wrapper);
      const firstOpen = await FolioDocxReviewer.fromBuffer(source);
      expect(firstOpen.getChanges()).toEqual(
        expect.arrayContaining([
          expect.objectContaining({ id: 91, author: "Reviewer", type: "deletion" }),
          expect.objectContaining({ id: 92, author: "Formatter", type: "formatting" }),
        ]),
      );

      const pending = await firstOpen.toBuffer();
      const reopened = await FolioDocxReviewer.fromBuffer(pending);
      expect(reopened.getChanges()).toEqual(
        expect.arrayContaining([
          expect.objectContaining({ id: 91, author: "Reviewer", type: "deletion" }),
          expect.objectContaining({ id: 92, author: "Formatter", type: "formatting" }),
        ]),
      );
      expect(await documentXml(pending)).toMatch(
        /<(?:w|x):rPrChange (?:w|x):id="92" (?:w|x):author="Formatter" (?:w|x):date="2026-09-09T00:01:00Z">/u,
      );

      const accepting = await FolioDocxReviewer.fromBuffer(pending);
      accepting.acceptAll();
      const accepted = await accepting.toBuffer();
      expect((await FolioDocxReviewer.fromBuffer(accepted)).getChanges()).toHaveLength(0);
      expect(await documentXml(accepted)).not.toContain("removed");

      const rejecting = await FolioDocxReviewer.fromBuffer(pending);
      rejecting.rejectAll();
      const rejected = await rejecting.toBuffer();
      const rejectedXml = await documentXml(rejected);
      const rejectedReopen = await FolioDocxReviewer.fromBuffer(rejected);
      expect(rejectedReopen.getChanges()).toHaveLength(0);
      expect(rejectedReopen.snapshot().blocks.at(0)?.text).toBe("removed");
      expect(rejectedXml).toContain("<w:i/>");
      expect(rejectedXml).not.toContain("<w:b/>");
      expect(rejectedXml).not.toMatch(/<w:(?:del|moveFrom|rPrChange)\b/u);
    },
  );
});
