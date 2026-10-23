import { expect, test } from "bun:test";

import { PARSE_WARNING_CODES } from "@stll/docx-core/model";

import { importedIdCollisionFixtures } from "../../scripts/imported-id-collision-fixtures";
import { createNumberingMap } from "./numberingParser";
import { parseDocx } from "./parser";
import { repackDocx } from "./rezip";

for (const { filename } of importedIdCollisionFixtures) {
  test(`${filename} opens with an identity diagnostic and reaches a save fixed point`, async () => {
    const kind = filename.slice("duplicate-imported-".length, -".docx".length);
    const source = await Bun.file(
      `${import.meta.dir}/__tests__/__fixtures__/corpus/${filename}`,
    ).arrayBuffer();
    const first = await parseDocx(source);
    expect(
      first.parseWarnings?.some(
        ({ code, identityKind }) =>
          code === PARSE_WARNING_CODES.duplicateImportedId &&
          identityKind === kind,
      ),
    ).toBe(true);
    if (kind === "comment")
      expect(first.package.document.comments).toHaveLength(1);
    if (kind === "footnote") expect(first.package.footnotes).toHaveLength(1);
    if (kind === "endnote") expect(first.package.endnotes).toHaveLength(1);
    if (kind === "comment" || kind === "footnote" || kind === "endnote") {
      const notes = (() => {
        if (kind === "comment") return first.package.document.comments;
        if (kind === "footnote") return first.package.footnotes;
        return first.package.endnotes;
      })();
      expect(JSON.stringify(notes)).toContain(
        kind === "comment" ? "First comment" : "First note",
      );
      expect(JSON.stringify(notes)).not.toContain(
        kind === "comment" ? "Later comment" : "Later note",
      );
    }
    if (kind === "relationship")
      expect(first.package.relationships?.get("rIdLink")?.target).toBe(
        "https://example.org/first",
      );
    if (kind === "numId" || kind === "abstractNumId") {
      expect(first.package.numbering).toBeDefined();
      if (first.package.numbering)
        expect(
          createNumberingMap(first.package.numbering).getLevel(73, 0)?.numFmt,
        ).toBe("decimal");
    }
    const reopened = await parseDocx(
      await repackDocx(first, { updateModifiedDate: false }),
    );
    const stable = await parseDocx(
      await repackDocx(reopened, { updateModifiedDate: false }),
    );
    expect(reopened.package.numbering).toEqual(first.package.numbering);
    expect(stable.package.numbering).toEqual(reopened.package.numbering);
    expect(stable.package.document).toEqual(reopened.package.document);
    expect(stable.parseWarnings).toEqual(reopened.parseWarnings);
  });
}
