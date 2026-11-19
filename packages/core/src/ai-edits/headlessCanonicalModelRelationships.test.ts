/**
 * The canonical reviewer's model states the package relationships its save
 * writes: a list started in a package with no numbering part relates the part
 * the save adds, in the live model and in each resolved view, as it does
 * after the save is reopened.
 */
import { expect, test } from "bun:test";

import { FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION } from "../document-operations";
import { createDocx } from "../docx/rezip";
import { fromMarkdown } from "../markdown/fromMarkdown";
import { FolioDocxReviewer } from "./headless";

const NUMBERING =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering";

const numberingTargets = (
  document: ReturnType<FolioDocxReviewer["toDocument"]>,
): string[] =>
  [...(document.package.relationships?.values() ?? [])]
    .filter(({ type }) => type === NUMBERING)
    .map(({ target }) => target);

test.each(["direct", "tracked-changes"] as const)(
  "a new list in a package without numbering relates the numbering part (%s)",
  async (mode) => {
    const source = fromMarkdown("First clause.\n\nSecond clause.");
    expect(source.package.numbering).toBeUndefined();
    const reviewer = await FolioDocxReviewer.fromBuffer(
      await createDocx(source),
      { author: "Editor", session: "canonical" },
    );
    const [first] = reviewer.snapshot().blocks;
    if (first === undefined) throw new Error("the fixture has no block");
    const result = reviewer.applyDocumentOperations(
      {
        version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
        mode,
        operations: [
          {
            id: "list",
            type: "setBlockParagraphProperties",
            blockId: first.id,
            properties: { numbering: { kind: "newList", format: "bullet" } },
          },
        ],
      },
      { undefinedReferences: "refuse" },
    );
    expect(result.applied.map(({ id }) => id)).toEqual(["list"]);
    const reopened = await FolioDocxReviewer.fromBuffer(
      await reviewer.toBuffer(),
      { author: "Editor", session: "canonical" },
    );
    expect(numberingTargets(reopened.toDocument())).toEqual(["numbering.xml"]);
    expect(numberingTargets(reviewer.toDocument())).toEqual(["numbering.xml"]);
    if (mode === "tracked-changes")
      expect(numberingTargets(reviewer.toResolvedDocument("accept"))).toEqual([
        "numbering.xml",
      ]);
  },
);
