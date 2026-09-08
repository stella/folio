import { describe, expect, test } from "bun:test";
import JSZip from "jszip";

import { parseDocx } from "../docx/parser";
import { createDocx, repackDocx } from "../docx/rezip";
import type { TextFormatting } from "../types/document";
import { createEmptyDocument } from "../utils/createDocument";
import { FolioDocxReviewer } from "./headless";
import { createFolioAITextRangeHandle } from "./snapshot";
import type { FolioAIInlineFormatting } from "./types";

type CreateFormattingBaselineOptions = {
  formatting?: TextFormatting;
};

const createFormattingBaseline = async ({
  formatting,
}: CreateFormattingBaselineOptions = {}): Promise<ArrayBuffer> => {
  const document = createEmptyDocument();
  document.package.document.content = [
    {
      type: "paragraph",
      paraId: "A1000001",
      content: [
        {
          type: "run",
          ...(formatting && { formatting }),
          content: [{ type: "text", text: "Formatting target" }],
        },
      ],
    },
  ];
  return createDocx(document);
};

const createSameValuedDirectFormattingDocument = async (): Promise<ArrayBuffer> => {
  const document = createEmptyDocument();
  const formatting: TextFormatting = {
    fontFamily: { ascii: "Georgia", hAnsi: "Georgia" },
    fontSize: 21,
    color: { rgb: "C00000" },
  };
  document.package.styles = {
    ...document.package.styles,
    docDefaults: { ...document.package.styles?.docDefaults, rPr: formatting },
    styles: document.package.styles?.styles.map((style) =>
      style.styleId === "Normal" ? { ...style, rPr: formatting } : style,
    ),
  };
  document.package.document.content = [
    {
      type: "paragraph",
      paraId: "A1000003",
      content: [
        { type: "run", content: [{ type: "text", text: "Inherited" }] },
        {
          type: "run",
          formatting,
          content: [{ type: "text", text: "Direct" }],
        },
      ],
    },
  ];
  return createDocx(document);
};

const createParagraphFontWithDirectBoldDocument = async (): Promise<ArrayBuffer> => {
  const document = createEmptyDocument();
  document.package.document.content = [
    {
      type: "paragraph",
      paraId: "A1000004",
      formatting: {
        runProperties: {
          fontFamily: { ascii: "Arial", hAnsi: "Arial" },
          fontSize: 21,
          color: { rgb: "C00000" },
        },
      },
      content: [
        {
          type: "run",
          formatting: { bold: true },
          content: [{ type: "text", text: "Direct bold" }],
        },
      ],
    },
  ];
  return createDocx(document);
};

type ApplyTrackedFormattingOptions = {
  formatting: FolioAIInlineFormatting;
  baselineFormatting?: TextFormatting;
};

const applyTrackedFormatting = async ({
  formatting,
  baselineFormatting,
}: ApplyTrackedFormattingOptions): Promise<ArrayBuffer> => {
  const reviewer = await FolioDocxReviewer.fromBuffer(
    await createFormattingBaseline({ formatting: baselineFormatting }),
    { author: "Reviewer" },
  );
  const block = reviewer.snapshot().blocks.at(0);
  const range = block
    ? createFolioAITextRangeHandle({
        blockId: block.id,
        text: block.text,
        startOffset: 0,
        endOffset: "Formatting".length,
      })
    : null;
  if (!range) {
    throw new Error("expected a formatting range");
  }

  const result = reviewer.applyOperations([
    { id: "format", type: "formatRange", range, formatting },
  ]);
  expect(result.skipped).toEqual([]);
  expect(result.applied.at(0)?.revisionId).toBeNumber();
  return reviewer.toBuffer();
};

const HEADER_RELATIONSHIP_ID = "rIdFormattingHeader";

const createHeaderFormattingBaseline = async (): Promise<ArrayBuffer> => {
  const seed = await createDocx(createEmptyDocument());
  const document = await parseDocx(seed, { detectVariables: false, preloadFonts: false });
  document.package.headers = new Map([
    [
      HEADER_RELATIONSHIP_ID,
      {
        type: "header" as const,
        hdrFtrType: "default" as const,
        content: [
          {
            type: "paragraph" as const,
            paraId: "A1000002",
            content: [
              {
                type: "run" as const,
                content: [{ type: "text" as const, text: "Header target" }],
              },
            ],
          },
        ],
      },
    ],
  ]);
  document.package.document.finalSectionProperties = {
    ...document.package.document.finalSectionProperties,
    headerReferences: [{ type: "default", rId: HEADER_RELATIONSHIP_ID }],
  };
  return repackDocx(document, { updateModifiedDate: false });
};

const documentXml = async (buffer: ArrayBuffer): Promise<string> => {
  const zip = await JSZip.loadAsync(buffer);
  const part = zip.file("word/document.xml");
  if (!part) {
    throw new Error("missing word/document.xml");
  }
  return part.async("text");
};

const applyTrackedBold = async (): Promise<ArrayBuffer> => {
  return applyTrackedFormatting({ formatting: { bold: true } });
};

describe("tracked run formatting", () => {
  test("snapshot preserves same-valued direct font properties", async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(
      await createSameValuedDirectFormattingDocument(),
    );

    expect(reviewer.snapshot().blocks.at(0)?.previewRuns).toEqual([
      {
        text: "Inherited",
        fontFamily: "Georgia",
        fontSizePt: 10.5,
        color: "#C00000",
      },
      {
        text: "Direct",
        fontFamily: "Georgia",
        fontSizePt: 10.5,
        color: "#C00000",
        directFormatting: {
          fontFamily: "Georgia",
          fontSizePt: 10.5,
          color: "#C00000",
        },
      },
    ]);
  });

  test("snapshot does not promote paragraph font properties beside direct bold", async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(
      await createParagraphFontWithDirectBoldDocument(),
    );

    expect(reviewer.snapshot().blocks.at(0)?.previewRuns?.at(0)?.directFormatting).toEqual({
      bold: true,
    });
  });

  test("tracks same-valued inherited font properties as new direct formatting", async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(
      await createSameValuedDirectFormattingDocument(),
      { author: "Reviewer" },
    );
    const block = reviewer.snapshot().blocks.at(0);
    const range = block
      ? createFolioAITextRangeHandle({
          blockId: block.id,
          text: block.text,
          startOffset: 0,
          endOffset: "Inherited".length,
        })
      : null;
    if (!range) {
      throw new Error("expected an inherited formatting range");
    }

    const result = reviewer.applyOperations([
      {
        id: "format-inherited",
        type: "formatRange",
        range,
        formatting: { fontFamily: "Georgia", fontSizePt: 10.5, color: "C00000" },
      },
    ]);

    expect(result.skipped).toEqual([]);
    expect(reviewer.readReviewedStory({ view: "current-markup" })?.changes).toEqual([
      expect.objectContaining({ type: "formatting", text: "Inherited" }),
    ]);
    expect(
      reviewer.readReviewedStory({ view: "final" })?.snapshot.blocks.at(0)?.previewRuns?.at(0)
        ?.directFormatting,
    ).toEqual({ fontFamily: "Georgia", fontSizePt: 10.5, color: "#C00000" });
  });

  test("clears only direct font properties in a mixed inherited range", async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(
      await createSameValuedDirectFormattingDocument(),
      { author: "Reviewer" },
    );
    const block = reviewer.snapshot().blocks.at(0);
    const range = block
      ? createFolioAITextRangeHandle({
          blockId: block.id,
          text: block.text,
          startOffset: 0,
          endOffset: block.text.length,
        })
      : null;
    if (!range) {
      throw new Error("expected a mixed formatting range");
    }

    const result = reviewer.applyOperations([
      {
        id: "clear-direct",
        type: "formatRange",
        range,
        formatting: { fontFamily: null, fontSizePt: null, color: null },
      },
    ]);

    expect(result.skipped).toEqual([]);
    expect(reviewer.readReviewedStory({ view: "current-markup" })?.changes).toEqual([
      expect.objectContaining({ type: "formatting", text: "Direct" }),
    ]);
    expect(
      reviewer.readReviewedStory({ view: "final" })?.snapshot.blocks.at(0)?.previewRuns,
    ).toEqual([
      {
        text: "InheritedDirect",
        fontFamily: "Georgia",
        fontSizePt: 10.5,
        color: "#C00000",
      },
    ]);
  });

  test("saves, reopens, accepts, and rejects a formatting revision", async () => {
    const tracked = await applyTrackedBold();
    const trackedXml = await documentXml(tracked);
    expect(trackedXml).toContain("<w:rPrChange ");
    expect(trackedXml).toContain("<w:b/>");

    const accepting = await FolioDocxReviewer.fromBuffer(tracked);
    const acceptedChange = accepting.getChanges().find(({ type }) => type === "formatting");
    if (!acceptedChange) {
      throw new Error("expected a formatting change after reopen");
    }
    expect(accepting.acceptChange(acceptedChange)).toBe(true);
    const acceptedXml = await documentXml(await accepting.toBuffer());
    expect(acceptedXml).not.toContain("<w:rPrChange ");
    expect(acceptedXml).toContain("<w:b/>");

    const rejecting = await FolioDocxReviewer.fromBuffer(tracked);
    const rejectedChange = rejecting.getChanges().find(({ type }) => type === "formatting");
    if (!rejectedChange) {
      throw new Error("expected a formatting change after reopen");
    }
    expect(rejecting.rejectChange(rejectedChange)).toBe(true);
    const rejectedXml = await documentXml(await rejecting.toBuffer());
    expect(rejectedXml).not.toContain("<w:rPrChange ");
    expect(rejectedXml).not.toContain("<w:b/>");
  });

  test("uses the same operation in a secondary document story", async () => {
    const story = { type: "header" as const, relationshipId: HEADER_RELATIONSHIP_ID };
    const reviewer = await FolioDocxReviewer.fromBuffer(await createHeaderFormattingBaseline(), {
      author: "Reviewer",
    });
    const snapshot = reviewer.snapshotStory(story);
    const block = snapshot?.blocks.at(0);
    const range = block
      ? createFolioAITextRangeHandle({
          blockId: block.id,
          text: block.text,
          startOffset: 0,
          endOffset: "Header".length,
        })
      : null;
    if (!snapshot || !range) {
      throw new Error("expected a header formatting range");
    }

    const result = reviewer.applyDocumentOperationsToStory({
      story,
      batch: {
        version: 1,
        mode: "tracked-changes",
        operations: [
          { id: "format-header", type: "formatRange", range, formatting: { italic: true } },
        ],
      },
      snapshot,
    });
    expect(result.skipped).toEqual([]);

    const reopened = await FolioDocxReviewer.fromBuffer(await reviewer.toBuffer());
    expect(reopened.readReviewedStory({ story, view: "current-markup" })?.changes).toEqual([
      expect.objectContaining({ type: "formatting", text: "Header" }),
    ]);
  });

  test("tracks target font face, size, and color through save and reopen", async () => {
    const tracked = await applyTrackedFormatting({
      formatting: { fontFamily: "Georgia", fontSizePt: 10.5, color: "c00000" },
    });
    const trackedXml = await documentXml(tracked);
    expect(trackedXml).toContain("<w:rPrChange ");
    expect(trackedXml).toContain('<w:rFonts w:ascii="Georgia" w:hAnsi="Georgia"');
    expect(trackedXml).toContain('<w:color w:val="C00000"');
    expect(trackedXml).toContain('<w:sz w:val="21"');

    const reopened = await FolioDocxReviewer.fromBuffer(tracked);
    expect(
      reopened.readReviewedStory({ view: "final" })?.snapshot.blocks.at(0)?.previewRuns?.at(0),
    ).toMatchObject({
      fontFamily: "Georgia",
      fontSizePt: 10.5,
      color: "#C00000",
      directFormatting: {
        fontFamily: "Georgia",
        fontSizePt: 10.5,
        color: "#C00000",
      },
    });
    expect(
      reopened.readReviewedStory({ view: "original" })?.snapshot.blocks.at(0)?.previewRuns?.at(0),
    ).toMatchObject({ fontFamily: "Arial", fontSizePt: 11 });
    expect(
      reopened.readReviewedStory({ view: "original" })?.snapshot.blocks.at(0)?.previewRuns?.at(0)
        ?.directFormatting,
    ).toBeUndefined();
  });

  test("tracks clearing direct font properties", async () => {
    const tracked = await applyTrackedFormatting({
      formatting: { fontFamily: null, fontSizePt: null, color: null },
      baselineFormatting: {
        fontFamily: { ascii: "Georgia", hAnsi: "Georgia" },
        fontSize: 21,
        color: { rgb: "C00000" },
      },
    });
    const reopened = await FolioDocxReviewer.fromBuffer(tracked);
    expect(
      reopened.readReviewedStory({ view: "final" })?.snapshot.blocks.at(0)?.previewRuns?.at(0),
    ).toMatchObject({ fontFamily: "Arial", fontSizePt: 11 });
    expect(
      reopened.readReviewedStory({ view: "final" })?.snapshot.blocks.at(0)?.previewRuns?.at(0)
        ?.directFormatting,
    ).toBeUndefined();
    expect(
      reopened.readReviewedStory({ view: "original" })?.snapshot.blocks.at(0)?.previewRuns?.at(0),
    ).toMatchObject({
      fontFamily: "Georgia",
      fontSizePt: 10.5,
      color: "#C00000",
      directFormatting: {
        fontFamily: "Georgia",
        fontSizePt: 10.5,
        color: "#C00000",
      },
    });
  });
});
