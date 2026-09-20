/**
 * Every pass that walks a paragraph reads through a transparent wrapper.
 *
 * `w:bdo` and `w:dir` say how their content is laid out and nothing else, so a
 * pass that stops at one sees a paragraph that is not there. The relationship
 * collector stopped at it: a link authored inside a bidirectional wrapper got
 * no `r:id`, which is the whole of how an `href` is saved, so the saved package
 * held a hyperlink pointing nowhere.
 */

import { describe, expect, test } from "bun:test";
import JSZip from "jszip";

import type { Document, Hyperlink, Paragraph } from "../types/document";
import { createEmptyDocument } from "../utils/createDocument";
import { createDocx } from "./rezip";

// Built per test: the save assigns the relationship id by mutating the link,
// so a shared literal would arrive at the second test already carrying one.
const link = (): Hyperlink => ({
  type: "hyperlink",
  href: "https://example.invalid/clause",
  children: [{ type: "run", content: [{ type: "text", text: "clause" }] }],
});

const documentWith = (content: Paragraph["content"]): Document => {
  const template = createEmptyDocument();
  return {
    ...template,
    package: {
      ...template.package,
      document: {
        ...template.package.document,
        content: [{ type: "paragraph", content }],
      },
    },
  };
};

const partText = async (buffer: ArrayBuffer, name: string): Promise<string> => {
  const part = (await JSZip.loadAsync(buffer)).file(name);
  if (!part) {
    throw new Error(`missing ${name}`);
  }
  return part.async("text");
};

describe("a link inside a bidirectional wrapper", () => {
  test("is saved with the relationship its href needs", async () => {
    const buffer = await createDocx(
      documentWith([
        { type: "bidiWrapper", control: "override", direction: "rtl", content: [link()] },
      ]),
    );

    expect(await partText(buffer, "word/_rels/document.xml.rels")).toContain(
      "https://example.invalid/clause",
    );
    expect(await partText(buffer, "word/document.xml")).toMatch(
      /<w:bdo w:val="rtl"><w:hyperlink r:id="/u,
    );
  });

  test("is saved with its relationship from inside a revision too", async () => {
    const buffer = await createDocx(
      documentWith([
        {
          type: "insertion",
          info: { id: 1, author: "Reviewer", date: "2026-01-01T00:00:00Z" },
          content: [
            { type: "bidiWrapper", control: "override", direction: "rtl", content: [link()] },
          ],
        },
      ]),
    );

    expect(await partText(buffer, "word/_rels/document.xml.rels")).toContain(
      "https://example.invalid/clause",
    );
  });
});
