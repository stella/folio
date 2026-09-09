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

const packagePart = async (buffer: ArrayBuffer, path: string): Promise<string | null> => {
  const zip = await JSZip.loadAsync(buffer);
  return zip.file(path)?.async("text") ?? null;
};

const documentXml = async (buffer: ArrayBuffer): Promise<string> => {
  const xml = await packagePart(buffer, "word/document.xml");
  if (xml === null) {
    throw new Error("missing word/document.xml");
  }
  return xml;
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

  test("never serializes two unresolved formatting revisions on one run", async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(await createFormattingBaseline(), {
      author: "Reviewer",
    });
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
      { id: "bold", type: "formatRange", range, formatting: { bold: true } },
      { id: "italic", type: "formatRange", range, formatting: { italic: true } },
    ]);
    expect(result.applied.map(({ id }) => id)).toEqual(["italic"]);
    expect(result.skipped).toEqual([{ id: "bold", reason: "pendingRunPropertyChange" }]);

    const tracked = await reviewer.toBuffer();
    const trackedXml = await documentXml(tracked);
    expect(trackedXml.match(/<w:rPrChange /gu)).toHaveLength(1);
    expect(trackedXml).toContain("<w:i/>");
    expect(trackedXml).not.toContain("<w:b/>");

    const accepting = await FolioDocxReviewer.fromBuffer(tracked);
    accepting.acceptAll();
    const acceptedXml = await documentXml(await accepting.toBuffer());
    expect(acceptedXml).not.toContain("<w:rPrChange ");
    expect(acceptedXml).toContain("<w:i/>");

    const rejecting = await FolioDocxReviewer.fromBuffer(tracked);
    rejecting.rejectAll();
    const rejectedXml = await documentXml(await rejecting.toBuffer());
    expect(rejectedXml).not.toContain("<w:rPrChange ");
    expect(rejectedXml).not.toContain("<w:i/>");
  });

  test("a reopened formatting owner refuses replacement without side effects", async () => {
    const baseline = await createFormattingBaseline({ formatting: { highlight: "yellow" } });
    const first = await FolioDocxReviewer.fromBuffer(baseline, { author: "Reviewer" });
    const firstBlock = first.snapshot().blocks.at(0);
    const firstRange = firstBlock
      ? createFolioAITextRangeHandle({
          blockId: firstBlock.id,
          text: firstBlock.text,
          startOffset: 0,
          endOffset: "Formatting".length,
        })
      : null;
    if (!firstRange) {
      throw new Error("expected an initial formatting range");
    }
    const firstResult = first.applyDocumentOperations(
      {
        version: 1,
        mode: "tracked-changes",
        operations: [
          { id: "bold", type: "formatRange", range: firstRange, formatting: { bold: true } },
        ],
      },
      { revisionStamp: { date: "2026-01-02T03:04:05.000Z", idSeed: 700 } },
    );
    expect(firstResult.nextRevisionId).toBe(701);

    const tracked = await first.toBuffer();
    const reopened = await FolioDocxReviewer.fromBuffer(tracked, { author: "Reviewer" });
    const beforeRefusal = await reopened.toBuffer();
    const snapshotBeforeRefusal = reopened.snapshot();
    const reopenedBlock = reopened.snapshot().blocks.at(0);
    if (!reopenedBlock) {
      throw new Error("expected a reopened formatting block");
    }
    const replacementResult = reopened.applyDocumentOperations(
      {
        version: 1,
        mode: "tracked-changes",
        operations: [
          {
            id: "replace",
            type: "replaceInBlock",
            blockId: reopenedBlock.id,
            find: "Formatting",
            replace: "Changed",
            comment: { text: "Explain this replacement." },
          },
        ],
      },
      { revisionStamp: { date: "2026-01-02T03:04:06.000Z", idSeed: 701 } },
    );

    expect(replacementResult.status).toBe("committed");
    expect(replacementResult.applied).toEqual([]);
    expect(replacementResult.skipped).toEqual([
      { id: "replace", reason: "pendingRunPropertyChange" },
    ]);
    expect(replacementResult.nextRevisionId).toBe(701);
    expect(reopened.getComments()).toEqual([]);
    expect(reopened.snapshot()).toEqual(snapshotBeforeRefusal);
    const afterRefusal = await reopened.toBuffer();
    expect(await documentXml(afterRefusal)).toBe(await documentXml(beforeRefusal));
    expect(await packagePart(afterRefusal, "word/comments.xml")).toBe(
      await packagePart(beforeRefusal, "word/comments.xml"),
    );

    const followupBlock = reopened.snapshot().blocks.at(0);
    const followupRange = followupBlock
      ? createFolioAITextRangeHandle({
          blockId: followupBlock.id,
          text: followupBlock.text,
          startOffset: "Formatting ".length,
          endOffset: "Formatting target".length,
        })
      : null;
    if (!followupRange) {
      throw new Error("expected a follow-up formatting range");
    }
    const followupResult = reopened.applyDocumentOperations(
      {
        version: 1,
        mode: "tracked-changes",
        operations: [
          {
            id: "italic",
            type: "formatRange",
            range: followupRange,
            formatting: { italic: true },
          },
        ],
      },
      { revisionStamp: { date: "2026-01-02T03:04:07.000Z", idSeed: 701 } },
    );
    expect(followupResult.applied.at(0)?.revisionIds).toEqual([701]);
    expect(followupResult.nextRevisionId).toBe(702);

    const accepting = await FolioDocxReviewer.fromBuffer(await reopened.toBuffer());
    accepting.acceptAll();
    const acceptedBuffer = await accepting.toBuffer();
    const accepted = await FolioDocxReviewer.fromBuffer(acceptedBuffer);
    const acceptedXml = await documentXml(acceptedBuffer);
    expect(acceptedXml).toContain("<w:b/>");
    expect(acceptedXml).toContain("<w:i/>");
    expect(acceptedXml).not.toContain("<w:rPrChange ");
    expect(accepted.getChanges()).toEqual([]);

    const rejecting = await FolioDocxReviewer.fromBuffer(await reopened.toBuffer());
    rejecting.rejectAll();
    const rejectedBuffer = await rejecting.toBuffer();
    const rejected = await FolioDocxReviewer.fromBuffer(rejectedBuffer);
    const rejectedXml = await documentXml(rejectedBuffer);
    expect(rejectedXml).not.toContain("<w:b/>");
    expect(rejectedXml).not.toContain("<w:i/>");
    expect(rejectedXml).not.toContain("<w:rPrChange ");
    expect(rejected.getChanges()).toEqual([]);
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
