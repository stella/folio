/**
 * The input boundaries where a character XML cannot hold enters folio.
 *
 * The escaper drops such a character whatever happens, so a package always
 * opens. These are the boundaries that drop it earlier, while the value is
 * still attached to the request that carried it — which is the only place a
 * warning can name what was lost.
 */

import { describe, expect, test } from "bun:test";

import { compileMarkdownToContent } from "../markdown/content";
import { parseLegalSource } from "../legal-source/parser";
import { validateDocumentModel } from "../validate/docx";
import type { Document, Paragraph } from "../model/document";

const NUL = String.fromCodePoint(0);
const LONE_HIGH_SURROGATE = String.fromCodePoint(0xd8_00);

const firstParagraphText = (blocks: readonly unknown[]): string => {
  const paragraph = blocks.at(0) as Paragraph | undefined;
  const run = paragraph?.content.at(0);
  if (run?.type !== "run") {
    return "";
  }
  const text = run.content.at(0);
  return text?.type === "text" ? text.text : "";
};

describe("markdown", () => {
  test("drops what XML cannot hold and keeps the rest", () => {
    const { content } = compileMarkdownToContent(`Signed${NUL} here${LONE_HIGH_SURROGATE}`);
    expect(firstParagraphText(content)).toBe("Signed here\uFFFD");
  });
});

describe("legal source", () => {
  test("drops what XML cannot hold and says so", () => {
    const { draft, diagnostics } = parseLegalSource(`@title Deed${NUL} of Assignment`);
    expect(draft.meta.title).toBe("Deed of Assignment");
    expect(diagnostics.map((diagnostic) => diagnostic.code)).toContain(
      "illegal-xml-characters-dropped",
    );
  });

  test("says nothing about a source XML can hold", () => {
    const { diagnostics } = parseLegalSource("@title Deed of Assignment");
    expect(diagnostics.map((diagnostic) => diagnostic.code)).not.toContain(
      "illegal-xml-characters-dropped",
    );
  });
});

describe("a model built in memory", () => {
  const documentWithText = (text: string): Document => ({
    package: {
      document: {
        content: [
          { type: "paragraph", content: [{ type: "run", content: [{ type: "text", text }] }] },
        ],
      },
    },
  });

  test("is reported, not rejected: the save still produces a package that opens", () => {
    const { valid, issues } = validateDocumentModel(documentWithText(`a${NUL}b`));
    expect(valid).toBe(true);
    expect(issues).toContainEqual({
      path: "package.document.content[0].content[0].content[0].text",
      message: "Text holds characters XML 1.0 cannot represent; they will be dropped on save.",
      severity: "warning",
    });
  });

  test("says nothing about text XML can hold", () => {
    expect(validateDocumentModel(documentWithText("a\tb\nc")).issues).toEqual([]);
  });
});
