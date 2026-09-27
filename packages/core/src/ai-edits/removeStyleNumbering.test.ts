/**
 * An operation that takes a paragraph out of a list its style numbers states
 * the cancellation (`w:numId="0"`): clearing the paragraph's own numbering
 * would uncover the style's, and the save would number it again.
 */

import { describe, expect, test } from "bun:test";

import { paragraphNumberingFromSlots } from "@stll/docx-core/model";

import type { FolioDocumentOperation } from "../document-operations";
import { ensureParaIds } from "../docx/ensureParaIds";
import { createDocx } from "../docx/rezip";
import { fromMarkdown } from "../markdown/fromMarkdown";
import { FolioDocxReviewer } from "./headless";

const CLAUSES = 5;

/** `Heading 2` numbered `1.`, `2.` through its style's `w:numPr`. */
const styleNumberedReviewer = async (): Promise<FolioDocxReviewer> => {
  const document = fromMarkdown(["## Scope", "Body.", "## Payment", "Closing."].join("\n\n"));
  const pkg = document.package;
  pkg.numbering = {
    abstractNums: [
      {
        abstractNumId: CLAUSES,
        multiLevelType: "multilevel",
        levels: [
          {
            ilvl: 0,
            start: 1,
            numFmt: "decimal",
            lvlText: "%1.",
            suffix: "space",
            pPr: { indentLeft: 0, indentFirstLine: 0 },
          },
        ],
      },
    ],
    nums: [{ numId: CLAUSES, abstractNumId: CLAUSES }],
  };
  const heading2 = pkg.styles?.styles.find(({ styleId }) => styleId === "Heading2");
  if (!heading2) {
    throw new Error("fixture style Heading2 is missing");
  }
  heading2.pPr = {
    ...heading2.pPr,
    numPr: paragraphNumberingFromSlots({ numId: CLAUSES, ilvl: 0 }),
  };
  const { docx } = await ensureParaIds(await createDocx(document));
  return FolioDocxReviewer.fromBuffer(docx, { author: "Agent" });
};

const labels = (reviewer: FolioDocxReviewer): string[] =>
  reviewer.getContent().map(({ displayLabel, styleId, text }) =>
    // An unnumbered heading's label is its style id.
    `${displayLabel === styleId ? "" : (displayLabel ?? "")} ${text}`.trim(),
  );

const blockId = (reviewer: FolioDocxReviewer, text: string): string => {
  const block = reviewer.getContent().find((candidate) => candidate.text === text);
  if (!block) {
    throw new Error(`no block reads "${text}"`);
  }
  return block.id;
};

const removals: Record<string, (reviewer: FolioDocxReviewer) => FolioDocumentOperation> = {
  "numbering: null": (reviewer) => ({
    id: "1",
    type: "setBlockParagraphProperties",
    blockId: blockId(reviewer, "Scope"),
    properties: { numbering: null },
  }),
  "listLevel: null": (reviewer) => ({
    id: "1",
    type: "setBlockParagraphProperties",
    blockId: blockId(reviewer, "Scope"),
    properties: { listLevel: null },
  }),
};

describe("removing numbering a paragraph style supplies", () => {
  for (const [name, removal] of Object.entries(removals)) {
    for (const mode of ["direct", "tracked-changes"] as const) {
      test(`${name} (${mode}) stays removed across a save`, async () => {
        const reviewer = await styleNumberedReviewer();
        expect(labels(reviewer)).toEqual(["1. Scope", "Body.", "2. Payment", "Closing."]);

        const result = reviewer.applyDocumentOperations({
          version: 1,
          mode,
          operations: [removal(reviewer)],
        });
        expect(result.issues).toEqual([]);
        expect(result.status).toBe("committed");

        const expected = ["Scope", "Body.", "1. Payment", "Closing."];
        expect(labels(reviewer)).toEqual(expected);
        const reopened = await FolioDocxReviewer.fromBuffer(await reviewer.toBuffer());
        expect(labels(reopened)).toEqual(expected);
      });
    }
  }

  test("an inserted paragraph that keeps the anchor's style but not its numbering stays unnumbered", async () => {
    const reviewer = await styleNumberedReviewer();
    const result = reviewer.applyDocumentOperations({
      version: 1,
      mode: "direct",
      operations: [
        {
          id: "1",
          type: "insertAfterBlock",
          blockId: blockId(reviewer, "Scope"),
          text: "Scope continued",
          numbering: null,
        },
      ],
    });
    expect(result.issues).toEqual([]);

    const expected = ["1. Scope", "Scope continued", "Body.", "2. Payment", "Closing."];
    expect(labels(reviewer)).toEqual(expected);
    const reopened = await FolioDocxReviewer.fromBuffer(await reviewer.toBuffer());
    expect(labels(reopened)).toEqual(expected);
    expect(reopened.getContent()[1]).toMatchObject({ kind: "heading", styleId: "Heading2" });
  });
});

describe("a paragraph inserted with a style of its own", () => {
  const insertStyled = async (styleId: string): Promise<FolioDocxReviewer> => {
    const reviewer = await styleNumberedReviewer();
    const result = reviewer.applyDocumentOperations({
      version: 1,
      mode: "direct",
      operations: [
        {
          id: "1",
          type: "insertAfterBlock",
          blockId: blockId(reviewer, "Body."),
          text: "Inserted",
          styleId,
        },
      ],
    });
    expect(result.issues).toEqual([]);
    return reviewer;
  };

  test("after a style-numbered heading, takes the numbering its new style gives", async () => {
    const reviewer = await insertStyled("Heading2");
    const expected = ["1. Scope", "Body.", "2. Inserted", "3. Payment", "Closing."];
    expect(labels(reviewer)).toEqual(expected);
    const reopened = await FolioDocxReviewer.fromBuffer(await reviewer.toBuffer());
    expect(labels(reopened)).toEqual(expected);
  });

  test("with numbering: null, stays out of the list its new style numbers", async () => {
    const reviewer = await styleNumberedReviewer();
    const result = reviewer.applyDocumentOperations({
      version: 1,
      mode: "direct",
      operations: [
        {
          id: "1",
          type: "insertAfterBlock",
          blockId: blockId(reviewer, "Body."),
          text: "Inserted",
          styleId: "Heading2",
          numbering: null,
        },
      ],
    });
    expect(result.issues).toEqual([]);
    const expected = ["1. Scope", "Body.", "Inserted", "2. Payment", "Closing."];
    expect(labels(reviewer)).toEqual(expected);
    const reopened = await FolioDocxReviewer.fromBuffer(await reviewer.toBuffer());
    expect(labels(reopened)).toEqual(expected);
  });

  test("anchored on a style-numbered heading, leaves that style's numbering behind", async () => {
    const reviewer = await styleNumberedReviewer();
    const result = reviewer.applyDocumentOperations({
      version: 1,
      mode: "direct",
      operations: [
        {
          id: "1",
          type: "insertAfterBlock",
          blockId: blockId(reviewer, "Scope"),
          text: "Inserted",
          styleId: "Normal",
        },
      ],
    });
    expect(result.issues).toEqual([]);
    const expected = ["1. Scope", "Inserted", "Body.", "2. Payment", "Closing."];
    expect(labels(reviewer)).toEqual(expected);
    const reopened = await FolioDocxReviewer.fromBuffer(await reviewer.toBuffer());
    expect(labels(reopened)).toEqual(expected);
  });
});

describe("restyling a paragraph", () => {
  const restyle = async (text: string, styleId: string | null): Promise<string[]> => {
    const reviewer = await styleNumberedReviewer();
    const result = reviewer.applyDocumentOperations({
      version: 1,
      mode: "tracked-changes",
      operations: [
        {
          id: "1",
          type: "setBlockParagraphProperties",
          blockId: blockId(reviewer, text),
          properties: { styleId },
        },
      ],
    });
    expect(result.issues).toEqual([]);
    const live = reviewer.getContent().map(({ kind, text: blockText }) => `${kind} ${blockText}`);
    const reopened = await FolioDocxReviewer.fromBuffer(await reviewer.toBuffer());
    expect(reopened.getContent().map(({ kind, text: t }) => `${kind} ${t}`)).toEqual(live);
    expect(labels(reopened)).toEqual(labels(reviewer));
    return labels(reviewer);
  };

  test("into a numbered style takes that style's numbering", async () => {
    expect(await restyle("Body.", "Heading2")).toEqual([
      "1. Scope",
      "2. Body.",
      "3. Payment",
      "Closing.",
    ]);
  });

  test("into a numbered style at a list level numbers it at that level of the style's list", async () => {
    const reviewer = await styleNumberedReviewer();
    const result = reviewer.applyDocumentOperations({
      version: 1,
      mode: "direct",
      operations: [
        {
          id: "1",
          type: "setBlockParagraphProperties",
          blockId: blockId(reviewer, "Body."),
          properties: { styleId: "Heading2", listLevel: 0 },
        },
      ],
    });
    expect(result.issues).toEqual([]);
    const expected = ["1. Scope", "2. Body.", "3. Payment", "Closing."];
    expect(labels(reviewer)).toEqual(expected);
    expect(labels(await FolioDocxReviewer.fromBuffer(await reviewer.toBuffer()))).toEqual(expected);
  });

  test("out of a numbered heading style leaves its numbering and its heading level", async () => {
    expect(await restyle("Scope", null)).toEqual(["Scope", "Body.", "1. Payment", "Closing."]);
  });
});

describe("a paragraph inserted with a style of its own after a heading", () => {
  test("takes its outline level from its new style, not the anchor's style", async () => {
    const document = fromMarkdown("Intro.\n\nStyled heading\n\nClosing.");
    const heading = document.package.document.content[1];
    if (heading?.type !== "paragraph") {
      throw new Error("fixture paragraph is missing");
    }
    // A heading only through its style: no outline level of its own.
    heading.formatting = { ...heading.formatting, styleId: "Heading4" };
    const { docx } = await ensureParaIds(await createDocx(document));
    const reviewer = await FolioDocxReviewer.fromBuffer(docx, { author: "Agent" });
    const result = reviewer.applyDocumentOperations({
      version: 1,
      mode: "direct",
      operations: [
        {
          id: "1",
          type: "insertAfterBlock",
          blockId: blockId(reviewer, "Styled heading"),
          text: "Body under it",
          styleId: "Normal",
        },
      ],
    });
    expect(result.issues).toEqual([]);

    const kindOf = (current: FolioDocxReviewer) =>
      current.getContent().find(({ text }) => text === "Body under it")?.kind;
    expect(kindOf(reviewer)).toBe("paragraph");
    expect(kindOf(await FolioDocxReviewer.fromBuffer(await reviewer.toBuffer()))).toBe("paragraph");
  });
});
