import { describe, expect, test } from "bun:test";

import { normalizeFailureMessage } from "./lib/corpus-signature";

/**
 * A signature erases what varies per machine and keeps what names the defect.
 *
 * Part names are the second kind: an invariant that reports which part folio
 * broke is useless if every part under a directory reduces to the same token.
 */
describe("normalizeFailureMessage over paths", () => {
  test("erases an absolute path", () => {
    expect(normalizeFailureMessage("cannot read /Users/someone/cache/a.docx")).toBe(
      "cannot read <path>",
    );
    expect(normalizeFailureMessage("cannot read C:\\Users\\someone\\a.docx")).toBe(
      "cannot read <path>",
    );
  });

  test("keeps a package part name whole, however deep", () => {
    expect(normalizeFailureMessage("word/media/image12.png changed")).toBe(
      "word/media/imageN.png changed",
    );
    expect(normalizeFailureMessage("word/theme/theme1.xml changed")).toBe(
      "word/theme/themeN.xml changed",
    );
  });

  test("two parts under one directory keep two signatures", () => {
    expect(normalizeFailureMessage("word/media/image1.png changed")).not.toBe(
      normalizeFailureMessage("word/theme/theme1.xml changed"),
    );
  });
});
