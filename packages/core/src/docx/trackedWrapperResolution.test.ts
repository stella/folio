/**
 * Resolving a revision resolves the text the wrapper holds.
 *
 * The demonstrated defect: `<w:ins><w:bdo>x</w:bdo></w:ins>` lifted the
 * wrapper out of the revision, so `x` was neither inserted nor deleted and
 * both answers kept it. Accepting must keep `x` and drop the `w:ins`;
 * rejecting must drop `x`, and must not leave the wrapper standing empty.
 */

import { describe, expect, test } from "bun:test";
import JSZip from "jszip";

import { FolioDocxReviewer } from "../ai-edits/headless";
import type { Document } from "../types/document";
import { createEmptyDocument } from "../utils/createDocument";
import { createDocx } from "./rezip";

const insertedInsideWrapper = (): Document => {
  const template = createEmptyDocument();
  return {
    ...template,
    package: {
      ...template.package,
      document: {
        ...template.package.document,
        content: [
          {
            type: "paragraph",
            content: [
              {
                type: "insertion",
                info: { id: 1, author: "Reviewer", date: "2026-01-01T00:00:00Z" },
                content: [
                  {
                    type: "bidiWrapper",
                    control: "override",
                    direction: "rtl",
                    content: [{ type: "run", content: [{ type: "text", text: "x" }] }],
                  },
                ],
              },
            ],
          },
        ],
      },
    },
  };
};

const documentXml = async (buffer: ArrayBuffer): Promise<string> => {
  const part = (await JSZip.loadAsync(buffer)).file("word/document.xml");
  if (!part) {
    throw new Error("missing word/document.xml");
  }
  return part.async("text");
};

describe("resolving a revision that holds a bidirectional wrapper", () => {
  test("the authored package nests the wrapper inside the revision", async () => {
    const xml = await documentXml(await createDocx(insertedInsideWrapper()));
    expect(xml).toMatch(/<w:ins\b[^>]*><w:bdo w:val="rtl"><w:r><w:t>x<\/w:t><\/w:r><\/w:bdo>/u);
  });

  test("accepting keeps the text and drops the revision", async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(await createDocx(insertedInsideWrapper()));
    expect(reviewer.getChanges()).toHaveLength(1);
    expect(reviewer.acceptAll()).toBe(1);

    const xml = await documentXml(await reviewer.toBuffer());
    expect(xml).toContain("<w:t>x</w:t>");
    expect(xml).not.toContain("<w:ins ");
  });

  test("rejecting drops the text and leaves no empty wrapper behind", async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(await createDocx(insertedInsideWrapper()));
    expect(reviewer.getChanges()).toHaveLength(1);
    expect(reviewer.rejectAll()).toBe(1);

    const xml = await documentXml(await reviewer.toBuffer());
    expect(xml).not.toContain(">x<");
    expect(xml).not.toContain("<w:bdo");
    expect(xml).not.toContain("<w:ins ");
  });
});
