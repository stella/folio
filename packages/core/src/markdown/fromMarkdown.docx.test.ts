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

// A nested list's indent must be wide enough for the parent's own marker
// (CommonMark: the child needs to start at or past the column where the
// parent item's content begins), not a fixed two spaces: "1. " is 3 columns,
// "10. " is 4, and only a bullet parent's "- " happens to be 2. Getting this
// wrong doesn't fail to parse — it reimports as a *flatter* document (the
// child silently promotes to the parent's own level), so every case here
// round-trips through actual DOCX bytes twice, the way a save/reopen/save
// would, and checks the second cycle still matches the first.
describe("nested lists under an ordered parent keep CommonMark's indentation", () => {
  const roundTrip = async (markdown: string): Promise<string> => {
    const bytes = await createDocx(fromMarkdown(markdown));
    const result = await docxToMarkdown(bytes, CLEAN);
    return typeof result === "string" ? result : result.markdown;
  };

  test("a 1-digit ordered parent indents its child by 3 spaces", async () => {
    const source = "1. Parent\n   - Child\n2. Next";
    const first = await roundTrip(source);
    expect(first).toBe(source);
    expect(await roundTrip(first)).toBe(source);
  });

  test("a 2-digit ordered parent indents its child by 4 spaces", async () => {
    const source = "10. Parent\n    - Child\n11. Next";
    const first = await roundTrip(source);
    expect(first).toBe(source);
    expect(await roundTrip(first)).toBe(source);
  });

  test("an ordered parent nesting an ordered child keeps its own numbering", async () => {
    const source = "1. Parent\n   1. Child\n   2. Sibling\n2. Next";
    const first = await roundTrip(source);
    expect(first).toBe(source);
    expect(await roundTrip(first)).toBe(source);
  });

  test("three levels of mixed ordered/bullet nesting round-trip", async () => {
    const source = "1. A\n   1. B\n      - C\n2. D";
    const first = await roundTrip(source);
    expect(first).toBe(source);
    expect(await roundTrip(first)).toBe(source);
  });

  test("three levels: bullet, then ordered, then bullet, round-trip", async () => {
    const source = "- A\n  1. B\n     - C\n- D";
    const first = await roundTrip(source);
    expect(first).toBe(source);
    expect(await roundTrip(first)).toBe(source);
  });

  test("a loose ordered item's continuation is idempotent, keeps both texts", async () => {
    // A continuation paragraph folds into the item's own paragraph as a soft
    // break (see fromMarkdown.test.ts's equivalent bullet-parent case), so
    // this isn't byte-identical to the source; it must still be a stable,
    // lossless normal form — the pre-existing behaviour this fix must not
    // regress.
    const source = "1. Parent\n\n   continuation text\n2. Next";
    const first = await roundTrip(source);
    expect(first).toContain("Parent");
    expect(first).toContain("continuation text");
    expect(first).toContain("2. Next");
    expect(await roundTrip(first)).toBe(first);
  });
});

describe("a table inside a list item", () => {
  test("survives as a following block, with a warning, through actual DOCX bytes", async () => {
    const source = "- Parent\n\n  | A | B |\n  | --- | --- |\n  | X | Y |\n\n- Next";
    const doc = fromMarkdown(source);
    // Never silently dropped: the model can't nest a table inside a list
    // item's paragraph, so it says so.
    expect(doc.warnings?.some((warning) => warning.includes("table"))).toBe(true);
    const bytes = await createDocx(doc);
    const result = await docxToMarkdown(bytes, CLEAN);
    const markdown = typeof result === "string" ? result : result.markdown;
    expect(markdown).toBe("- Parent\n\n| A | B |\n| --- | --- |\n| X | Y |\n\n- Next");
  });
});
