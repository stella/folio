import { describe, expect, test } from "bun:test";

import { ZipArchive } from "@stll/docx-core/zip";

import { FolioDocxReviewer } from "../ai-edits/headless";
import type { HeadlessReviewerSession } from "../types/canonicalCapabilities";
import { buildBodySequenceDocx } from "./__fixtures__/body-sequence";
import { compareDocx } from "./compare";

const documentWithParagraphs = async (
  paragraphs: readonly string[],
): Promise<ArrayBuffer> => {
  const source = await buildBodySequenceDocx([
    { kind: "paragraph", text: "placeholder" },
  ]);
  const zip = await ZipArchive.read(source);
  const document = zip.entry("word/document.xml");
  if (!document) {
    throw new Error("The synthetic DOCX fixture has no primary document part.");
  }
  const xml = await document.text();
  const sectionProperties = /<w:sectPr\b[\s\S]*?<\/w:sectPr>/u.exec(xml)?.[0];
  if (!sectionProperties) {
    throw new Error("The synthetic DOCX fixture has no section properties.");
  }
  zip.set(
    "word/document.xml",
    xml.replace(
      /<w:body>[\s\S]*?<\/w:body>/u,
      `<w:body>${paragraphs
        .map(
          (inlines, index) =>
            `<w:p w14:paraId="${String(index + 1).padStart(8, "0")}" w14:textId="${String(index + 1).padStart(8, "0")}">${inlines}</w:p>`,
        )
        .join("")}${sectionProperties}</w:body>`,
    ),
  );
  return (await zip.generate({ compression: "store" })).buffer;
};

const rawControlElements = async (buffer: ArrayBuffer): Promise<string[]> => {
  const zip = await ZipArchive.read(buffer);
  const document = zip.entry("word/document.xml");
  if (!document) {
    throw new Error("The compared DOCX has no primary document part.");
  }
  return [...(await document.text()).matchAll(/<w:(tab|br)\b/gu)].map(
    ([, element]) => element ?? "",
  );
};

const SESSIONS: readonly HeadlessReviewerSession[] = ["canonical", "legacy"];

const firstBlockText = async (
  buffer: ArrayBuffer,
  session: HeadlessReviewerSession,
): Promise<string> => {
  const block = (await FolioDocxReviewer.fromBuffer(buffer, { session }))
    .getContent()
    .at(0);
  if (!block) {
    throw new Error("Expected the synthetic DOCX fixture to produce a block.");
  }
  return block.text;
};

describe.each(SESSIONS)("comparison control characters (%s)", (session) => {
  test("accept and reject preserve tabs and hard breaks in a replaced paragraph", async () => {
    const base = await documentWithParagraphs([
      "<w:r><w:t>alpha</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>base</w:t></w:r><w:r><w:br/></w:r><w:r><w:t>end</w:t></w:r>",
    ]);
    const target = await documentWithParagraphs([
      "<w:r><w:t>alpha</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>target</w:t></w:r><w:r><w:br/></w:r><w:r><w:t>end</w:t></w:r>",
    ]);
    const baseText = await firstBlockText(base, session);
    const targetText = await firstBlockText(target, session);
    expect(baseText).toBe("alpha\tbase\nend");
    expect(targetText).toBe("alpha\ttarget\nend");

    const compared = await compareDocx(base, target, {
      author: "compare",
      timestamp: "2024-03-01T00:00:00.000Z",
      session,
    });
    expect(compared.isErr()).toBe(false);
    if (compared.isErr()) {
      throw compared.error;
    }

    const accepting = await FolioDocxReviewer.fromBuffer(
      compared.value.buffer,
      {
        session,
      },
    );
    accepting.acceptAll();
    const accepted = await accepting.toBuffer();
    expect(await firstBlockText(accepted, session)).toBe(targetText);
    expect(await rawControlElements(accepted)).toEqual(["tab", "br"]);

    const rejecting = await FolioDocxReviewer.fromBuffer(
      compared.value.buffer,
      {
        session,
      },
    );
    rejecting.rejectAll();
    const rejected = await rejecting.toBuffer();
    expect(await firstBlockText(rejected, session)).toBe(baseText);
    expect(await rawControlElements(rejected)).toEqual(["tab", "br"]);
  });

  test("accept retains controls in an inserted paragraph and reject removes it", async () => {
    const base = await documentWithParagraphs(["<w:r><w:t>anchor</w:t></w:r>"]);
    const target = await documentWithParagraphs([
      "<w:r><w:t>anchor</w:t></w:r>",
      "<w:r><w:t>inserted</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>value</w:t></w:r><w:r><w:br/></w:r><w:r><w:t>end</w:t></w:r>",
    ]);
    const compared = await compareDocx(base, target, {
      author: "compare",
      timestamp: "2024-03-01T00:00:00.000Z",
      session,
    });
    expect(compared.isErr()).toBe(false);
    if (compared.isErr()) {
      throw compared.error;
    }

    const accepting = await FolioDocxReviewer.fromBuffer(
      compared.value.buffer,
      {
        session,
      },
    );
    accepting.acceptAll();
    const accepted = await accepting.toBuffer();
    expect(await firstBlockText(accepted, session)).toBe("anchor");
    expect(
      (await FolioDocxReviewer.fromBuffer(accepted, { session }))
        .getContent()
        .map(({ text }) => text),
    ).toEqual(["anchor", "inserted\tvalue\nend"]);
    expect(await rawControlElements(accepted)).toEqual(["tab", "br"]);

    const rejecting = await FolioDocxReviewer.fromBuffer(
      compared.value.buffer,
      {
        session,
      },
    );
    rejecting.rejectAll();
    const rejected = await rejecting.toBuffer();
    expect(
      (await FolioDocxReviewer.fromBuffer(rejected, { session }))
        .getContent()
        .map(({ text }) => text),
    ).toEqual(["anchor"]);
    expect(await rawControlElements(rejected)).toEqual([]);
  });
});
