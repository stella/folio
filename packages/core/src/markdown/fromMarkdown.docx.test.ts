// fromMarkdown must be self-consistent: it mints numId 1, 2, … for the lists
// it emits (see fromMarkdown.ts's `blocksFromTokens(tokens, { next: 1 })`),
// so it must also synthesize the `document.package.numbering` those numIds
// point to. Without that, `createDocx` throws `DocxModelValidationError:
// Numbering definition N is missing` — this only surfaces once the document
// is actually serialized to DOCX bytes, not when just walking the in-memory
// model (see fromMarkdown.test.ts, which never calls createDocx).

import { describe, expect, test } from "bun:test";

import { docxToMarkdown } from "../docx/server/docxToMarkdown";
import { createDocx } from "../docx/rezip";
import { fromMarkdown } from "./fromMarkdown";

const CLEAN = {
  annotations: "strip",
  trackedChanges: "clean",
  comments: "strip",
  hyperlinks: "inline",
  footnotes: "strip",
} as const;

describe("fromMarkdown synthesizes self-consistent numbering", () => {
  test("document.package.numbering is populated for markdown lists", () => {
    const doc = fromMarkdown("1. a\n2. b\n- x");
    expect(doc.package.numbering).toBeDefined();
    expect(doc.package.numbering?.abstractNums.length).toBeGreaterThan(0);
    expect(doc.package.numbering?.nums.length).toBeGreaterThan(0);
  });

  test("a document with no lists carries no numbering", () => {
    const doc = fromMarkdown("Just a paragraph, no lists here.");
    expect(doc.package.numbering).toBeUndefined();
  });

  test("createDocx(fromMarkdown(...)) does not throw for ordered + bullet lists", async () => {
    const doc = fromMarkdown("1. a\n2. b\n- x");
    await expect(createDocx(doc)).resolves.toBeInstanceOf(ArrayBuffer);
  });

  test("createDocx(fromMarkdown(...)) round-trips through actual DOCX bytes", async () => {
    const bytes = await createDocx(fromMarkdown("1. a\n2. b\n- x"));
    const markdown = await docxToMarkdown(bytes, CLEAN);
    expect(markdown).toBe("1. a\n2. b\n- x");
  });

  test("nested lists round-trip through actual DOCX bytes", async () => {
    const source = "- a\n  1. nested\n- b";
    const bytes = await createDocx(fromMarkdown(source));
    const markdown = await docxToMarkdown(bytes, CLEAN);
    expect(markdown).toBe(source);
  });
});
