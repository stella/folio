import { describe, expect, test } from "bun:test";
import { panic } from "better-result";
import JSZip from "jszip";

import { FolioDocxReviewer } from "../ai-edits/headless";
import { createDocx } from "../docx/rezip";
import type { Paragraph } from "../types/document";
import { createEmptyDocument } from "../utils/createDocument";
import { compareDocx } from "./compare";

const OPTIONS = { author: "compare", timestamp: "2026-09-13T00:00:00.000Z" } as const;
const TEXT = "The agreement remains effective for the stated term.";

const indentationDocument = (indentLeft: number): Promise<ArrayBuffer> => {
  const document = createEmptyDocument();
  const paragraph: Paragraph = {
    type: "paragraph",
    paraId: "12345678",
    textId: "12345678",
    formatting: { indentLeft },
    content: [{ type: "run", content: [{ type: "text", text: TEXT }] }],
  };
  document.package.document.content = [paragraph];
  return createDocx(document);
};

const mainDocumentXml = async (buffer: ArrayBuffer): Promise<string> => {
  const part = (await JSZip.loadAsync(buffer)).file("word/document.xml");
  if (!part) {
    panic("expected word/document.xml");
  }
  return await part.async("string");
};

describe("paragraph indentation comparison", () => {
  test("tracks direct indentation and preserves accepted and rejected raw pPr", async () => {
    const result = await compareDocx(
      await indentationDocument(720),
      await indentationDocument(1440),
      OPTIONS,
    );
    if (result.isErr()) {
      throw result.error;
    }

    expect(result.value.verification).toEqual({ status: "verified" });
    expect(result.value.changes).toEqual([
      expect.objectContaining({
        kind: "paragraph-format",
        properties: { indentation: { indentLeft: 1440, hangingIndent: false } },
      }),
    ]);
    const pending = await mainDocumentXml(result.value.buffer);
    expect(pending).toContain('<w:ind w:left="1440"/>');
    expect(pending).toContain("<w:pPrChange ");
    expect(pending).toContain('<w:ind w:left="720"/>');

    const accepting = await FolioDocxReviewer.fromBuffer(result.value.buffer);
    expect(accepting.acceptAll()).toBe(1);
    const accepted = await mainDocumentXml(await accepting.toBuffer());
    expect(accepted).not.toContain("<w:pPrChange");
    expect(accepted).toContain('<w:ind w:left="1440"/>');

    const rejecting = await FolioDocxReviewer.fromBuffer(result.value.buffer);
    expect(rejecting.rejectAll()).toBe(1);
    const rejected = await mainDocumentXml(await rejecting.toBuffer());
    expect(rejected).not.toContain("<w:pPrChange");
    expect(rejected).toContain('<w:ind w:left="720"/>');
  });
});
