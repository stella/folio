import { expect, test } from "bun:test";
import { expectParagraphBlock } from "../../../../test/paragraphBlock";

import { createDocx } from "../docx/rezip";
import { fromMarkdown } from "../markdown/fromMarkdown";
import { paragraphNumberingReference } from "../docx/numberingReference";
import { FolioDocxReviewer } from "./headless";
import { isFolioAIContentBlock } from "./snapshot";

const CASES = (["direct", "tracked-changes"] as const).flatMap((mode) =>
  (["set-properties", "insert", "remove-properties", "insert-removed"] as const).flatMap(
    (operation) =>
      (["inherited", "zero", "explicit"] as const).flatMap((source) =>
        (["preserve", "clear", "zero"] as const).flatMap((indentation) =>
          (operation === "remove-properties" || operation === "insert-removed"
            ? [false, true]
            : [true]
          ).map((withStyle) => ({ mode, operation, source, indentation, withStyle })),
        ),
      ),
  ),
);

// Separate style and numbering examples missed their combined write order:
// the later style patch replaced the numbering-derived provenance.
test.each(CASES)(
  "$operation $source/$indentation indentation survives numbering (style=$withStyle) in $mode",
  async ({ mode, operation, source, indentation, withStyle }) => {
    const document = fromMarkdown("1. Anchor");
    const paragraph = document.package.document.content.at(0);
    const definition = document.package.numbering?.abstractNums.at(0);
    const template = definition?.levels.at(0);
    const numPr = paragraph?.type === "paragraph" ? paragraph.formatting?.numPr : undefined;
    if (paragraph?.type !== "paragraph" || !definition || !template || numPr?.kind !== "reference")
      throw new TypeError("Combined list fixture lacks numbering.");
    const styles = document.package.styles;
    if (!styles) throw new TypeError("Combined list fixture lacks styles.");
    styles.styles.push({
      type: "paragraph",
      styleId: "IndentBaseline",
      name: "Indent baseline",
      pPr: { indentLeft: 100, indentFirstLine: 20, hangingIndent: false },
    });
    definition.levels = [0, 1].map((ilvl) =>
      Object.assign({}, template, {
        ilvl,
        lvlText: `%${String(ilvl + 1)}.`,
        pPr: { indentLeft: 720 * (ilvl + 1), indentFirstLine: -360, hangingIndent: true },
      }),
    );
    const sourceDirect =
      source === "inherited"
        ? undefined
        : {
            indentLeft: source === "zero" ? 0 : 1440,
            indentFirstLine: 0,
            hangingIndent: false,
          };
    paragraph.paraId = "12345678";
    paragraph.formatting = {
      numPr: paragraphNumberingReference({ numId: numPr.numId, ilvl: 0 }),
      ...sourceDirect,
    };
    const reviewer = await FolioDocxReviewer.fromBuffer(await createDocx(document));
    const anchor = reviewer.getContent().find(({ text }) => text === "Anchor");
    if (!anchor) throw new TypeError("Combined list fixture lost its anchor.");
    const explicitIndentation = (() => {
      switch (indentation) {
        case "preserve":
          return undefined;
        case "clear":
          return null;
        case "zero":
          return { indentLeft: 0, indentFirstLine: 0, hangingIndent: false };
      }
    })();
    const expectedDirect =
      indentation === "preserve" ? sourceDirect : (explicitIndentation ?? undefined);
    const removesNumbering = operation === "remove-properties" || operation === "insert-removed";
    const inserts = operation === "insert" || operation === "insert-removed";
    const properties = {
      ...(withStyle && { styleId: "IndentBaseline" }),
      numbering: removesNumbering
        ? { kind: "none" as const }
        : { kind: "reference" as const, numId: numPr.numId, ilvl: 1 },
      ...(explicitIndentation !== undefined && { indentation: explicitIndentation }),
    };
    const result = reviewer.applyDocumentOperations({
      version: 1,
      mode,
      operations: [
        !inserts
          ? {
              id: "change",
              type: "setBlockParagraphProperties",
              blockId: anchor.id,
              properties,
            }
          : {
              id: "change",
              type: "insertAfterBlock",
              blockId: anchor.id,
              text: "Inserted",
              ...properties,
            },
      ],
    });
    expect(result.status).toBe("committed");
    expect(result.issues).toEqual([]);
    const text = inserts ? "Inserted" : "Anchor";
    const assertContent = (current: FolioDocxReviewer) => {
      const observedBlock = current
        .getContent()
        .filter(isFolioAIContentBlock)
        .find((entry) => entry.text === text);
      if (!observedBlock) throw new TypeError("Combined list operation lost its target.");
      const block = expectParagraphBlock(observedBlock);
      expect(block.styleId).toBe(withStyle ? "IndentBaseline" : anchor.styleId);
      expect(block.listReference).toEqual(
        removesNumbering ? undefined : { numId: numPr.numId, level: 1 },
      );
      expect(block.directIndentation).toEqual(expectedDirect);
    };
    assertContent(reviewer);
    assertContent(await FolioDocxReviewer.fromBuffer(await reviewer.toBuffer()));
    if (mode === "tracked-changes") {
      expect(reviewer.acceptAll()).toBeGreaterThan(0);
      assertContent(reviewer);
      assertContent(await FolioDocxReviewer.fromBuffer(await reviewer.toBuffer()));
    }
  },
);
