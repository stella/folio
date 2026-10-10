/**
 * Redline-generator tests, built around the two invariants that define a
 * correct redline document:
 *
 * 1. **Accept-all equals revised.** The redline's as-accepted view (what a
 *    reviewer sees after accepting every tracked change) must reproduce the
 *    revised document's block texts.
 * 2. **Reject-all equals base.** Rejecting every tracked change must restore
 *    the base document's block texts.
 *
 * Buffers are built from the typed `Document` model like the comparer tests,
 * and the generated redline is inspected through a fresh `FolioDocxReviewer`
 * (whose snapshot IS the as-accepted view).
 */

import { describe, expect, test } from "bun:test";
import JSZip from "jszip";

import {
  insertedTextBoxDocument,
  INSERTED_TEXT_BOX_STYLE,
} from "./__tests__/insertedTextBoxDocument";
import { buildTextBoxTableDocument, findTextBoxShape } from "./__tests__/textBoxTableDocument";
import { FolioDocxReviewer } from "./ai-edits/headless";
import { compareDocx } from "./compare/compare";
import { parseDocx } from "./docx/parser";
import { createDocx } from "./docx/rezip";
import { repackDocx } from "./docx/rezip";
import {
  generateRedlineDocx,
  GenerateRedlineDocxResourceImportError,
  InvalidGenerateRedlineDocxOptionsError,
} from "./redline";
import { paragraphNumberingReference } from "@stll/docx-core/model";
import {
  GenerateRedlineDocxOperationLimitError,
  MAX_GENERATED_REDLINE_OPERATIONS,
} from "./redlineOperationLimit";
import type { HeaderFooter, Paragraph } from "./types/document";
import { createEmptyDocument } from "./utils/createDocument";
import { expectParagraphBlock } from "../../../test/paragraphBlock";

type InlineFormattingSpec = { bold?: boolean; italic?: boolean };

type ParagraphSpec = {
  text: string;
  paraId?: string;
  formatting?: InlineFormattingSpec;
};

const buildInsertedParagraphs = (count: number): ParagraphSpec[] =>
  Array.from({ length: count }, (_, index) => ({ text: `Inserted paragraph ${index}` }));

const expectOperationLimitError = async (promise: Promise<unknown>) => {
  const error = await promise.then(
    () => null,
    (reason) => reason,
  );

  expect(error).toBeInstanceOf(GenerateRedlineDocxOperationLimitError);
  expect(error).toMatchObject({
    _tag: "GenerateRedlineDocxOperationLimitError",
    message: "The document comparison exceeds the generated operation limit.",
  });
};

const CORE_PROPERTIES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="urn:properties" xmlns:dc="urn:descriptive" xmlns:dcterms="urn:terms"><dc:title>Private title</dc:title><dc:creator>Private creator</dc:creator><cp:lastModifiedBy>Private modifier</cp:lastModifiedBy><cp:revision>7</cp:revision><dcterms:created>2026-07-01T10:30:00Z</dcterms:created><dcterms:modified>2026-07-02T11:45:00Z</dcterms:modified></cp:coreProperties>`;

const buildDocxBuffer = (paragraphs: readonly ParagraphSpec[]): Promise<ArrayBuffer> => {
  const template = createEmptyDocument();
  return createDocx({
    ...template,
    package: {
      ...template.package,
      document: {
        ...template.package.document,
        content: paragraphs.map(({ text, paraId, formatting }) => ({
          type: "paragraph",
          content: [
            {
              type: "run",
              ...(formatting !== undefined && { formatting }),
              content: [{ type: "text", text }],
            },
          ],
          ...(paraId !== undefined && { paraId }),
        })),
      },
    },
  });
};

type FormattedRunSpec = {
  text: string;
  formatting?: InlineFormattingSpec;
};

const buildFormattedRunDocxBuffer = (
  runs: readonly FormattedRunSpec[],
  paraId: string,
): Promise<ArrayBuffer> => {
  const template = createEmptyDocument();
  return createDocx({
    ...template,
    package: {
      ...template.package,
      document: {
        ...template.package.document,
        content: [
          {
            type: "paragraph",
            paraId,
            content: runs.map(({ text, formatting }) => ({
              type: "run",
              ...(formatting !== undefined && { formatting }),
              content: [{ type: "text", text }],
            })),
          },
        ],
      },
    },
  });
};

const blockTexts = (reviewer: FolioDocxReviewer): string[] =>
  reviewer.snapshot().blocks.map((block) => block.text);

const withCoreProperties = async (buffer: ArrayBuffer): Promise<ArrayBuffer> => {
  const zip = await JSZip.loadAsync(buffer);
  zip.file("docProps/core.xml", CORE_PROPERTIES_XML);
  return zip.generateAsync({ type: "arraybuffer", compression: "DEFLATE" });
};

const storyParagraph = (text: string, paraId: string): Paragraph => ({
  type: "paragraph",
  paraId,
  content: [{ type: "run", content: [{ type: "text", text }] }],
});

const headerFooterStory = (
  type: "header" | "footer",
  text: string,
  paraId: string,
): HeaderFooter => ({
  type,
  hdrFtrType: "default",
  content: [storyParagraph(text, paraId)],
});

type StoryDocumentOptions = {
  bodyText: string;
  headerText?: string;
  footerText?: string;
  footnoteText?: string;
  endnoteText?: string;
};

const buildStoryDocument = async ({
  bodyText,
  headerText,
  footerText,
  footnoteText,
  endnoteText,
}: StoryDocumentOptions): Promise<ArrayBuffer> => {
  const source = await buildDocxBuffer([{ text: bodyText, paraId: "21000001" }]);
  const document = await parseDocx(source, { detectVariables: false, preloadFonts: false });
  if (headerText !== undefined) {
    document.package.headers = new Map([
      ["rIdHeader", headerFooterStory("header", headerText, "31000001")],
    ]);
    document.package.document.finalSectionProperties = {
      ...document.package.document.finalSectionProperties,
      headerReferences: [{ type: "default", rId: "rIdHeader" }],
    };
  }
  if (footerText !== undefined) {
    document.package.footers = new Map([
      ["rIdFooter", headerFooterStory("footer", footerText, "41000001")],
    ]);
    document.package.document.finalSectionProperties = {
      ...document.package.document.finalSectionProperties,
      footerReferences: [{ type: "default", rId: "rIdFooter" }],
    };
  }
  const materialized = await repackDocx(document, { updateModifiedDate: false });
  if (footnoteText === undefined && endnoteText === undefined) {
    return materialized;
  }

  const zip = await JSZip.loadAsync(materialized);
  const contentTypesFile = zip.file("[Content_Types].xml");
  const relationshipsFile = zip.file("word/_rels/document.xml.rels");
  if (!contentTypesFile || !relationshipsFile) {
    throw new Error("expected package metadata parts");
  }
  const contentTypes = await contentTypesFile.async("text");
  const relationships = await relationshipsFile.async("text");
  const contentTypeOverrides = [
    footnoteText === undefined
      ? ""
      : '<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>',
    endnoteText === undefined
      ? ""
      : '<Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/>',
  ].join("");
  const noteRelationships = [
    footnoteText === undefined
      ? ""
      : '<Relationship Id="rIdFootnotes" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>',
    endnoteText === undefined
      ? ""
      : '<Relationship Id="rIdEndnotes" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml"/>',
  ].join("");
  zip.file(
    "[Content_Types].xml",
    contentTypes.replace("</Types>", `${contentTypeOverrides}</Types>`),
  );
  zip.file(
    "word/_rels/document.xml.rels",
    relationships.replace("</Relationships>", `${noteRelationships}</Relationships>`),
  );
  if (footnoteText !== undefined) {
    zip.file(
      "word/footnotes.xml",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:footnote w:id="2"><w:p w14:paraId="51000001"><w:r><w:t>${footnoteText}</w:t></w:r></w:p></w:footnote></w:footnotes>`,
    );
  }
  if (endnoteText !== undefined) {
    zip.file(
      "word/endnotes.xml",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:endnote w:id="3"><w:p w14:paraId="61000001"><w:r><w:t>${endnoteText}</w:t></w:r></w:p></w:endnote></w:endnotes>`,
    );
  }
  return zip.generateAsync({ type: "arraybuffer" });
};

const storyTextByType = (
  reviewer: FolioDocxReviewer,
  view: "original" | "final",
): Record<string, string> => {
  const texts: Record<string, string> = {};
  for (const { handle } of reviewer.listStories()) {
    const story = reviewer.readReviewedStory({ story: handle, view });
    if (story) {
      texts[handle.type] = story.snapshot.blocks.map(({ text }) => text).join("\n");
    }
  }
  return texts;
};

const withPendingMainChange = async (source: ArrayBuffer, text: string): Promise<ArrayBuffer> => {
  const reviewer = await FolioDocxReviewer.fromBuffer(source, { author: "Source reviewer" });
  const target = reviewer.snapshot().blocks.at(0);
  if (!target) {
    throw new Error("expected a main-story block");
  }
  reviewer.applyOperations(
    [{ id: "source-change", type: "replaceBlock", blockId: target.id, text }],
    { mode: "tracked-changes" },
  );
  return reviewer.toBuffer();
};

describe("generateRedlineDocx", () => {
  test("package comparisons share one inline-alignment allowance across stories", async () => {
    const alternating = (first: string, second: string): string =>
      Array.from({ length: 2000 }, (_unused, index) => (index % 2 === 0 ? first : second)).join(
        " ",
      );
    const before = alternating("a", "b");
    const after = alternating("b", "a");
    const base = await buildStoryDocument({
      bodyText: before,
      headerText: before,
      footnoteText: before,
    });
    const revised = await buildStoryDocument({
      bodyText: after,
      headerText: after,
      footnoteText: after,
    });
    const buffers = [(await generateRedlineDocx(base, revised)).buffer];
    const compared = await compareDocx(base, revised, {
      author: "compare",
      timestamp: "2026-09-08T12:00:00.000Z",
      onUnverified: "emit",
    });
    if (compared.isErr()) {
      throw compared.error;
    }
    buffers.push(compared.value.buffer);

    for (const buffer of buffers) {
      const reviewer = await FolioDocxReviewer.fromBuffer(buffer);
      const hasFineGrainedChange = reviewer.listStories().map(({ handle }) => {
        const story = reviewer.readReviewedStory({ story: handle, view: "current-markup" });
        return (story?.changes ?? [])
          .filter(({ type }) => type === "insertion" || type === "deletion")
          .some(({ text }) => text.length < before.length);
      });
      expect(hasFineGrainedChange).toEqual([true, false, false]);
    }
  });

  test("redlines a nested text-box table-cell edit without changing either source", async () => {
    const base = await buildTextBoxTableDocument("Original cell value");
    const revised = await buildTextBoxTableDocument("Revised cell value");
    const baseBytes = new Uint8Array(base).slice();
    const revisedBytes = new Uint8Array(revised).slice();

    const result = await generateRedlineDocx(base, revised);

    expect(result.skipped).toEqual([]);
    expect(result.unprocessedStories).toEqual([]);
    expect(result.applied).toHaveLength(1);
    expect(new Uint8Array(base)).toEqual(baseBytes);
    expect(new Uint8Array(revised)).toEqual(revisedBytes);

    const acceptView = await FolioDocxReviewer.fromBuffer(result.buffer);
    expect(acceptView.snapshot().blocks.find(({ id }) => id === "22000003")?.text).toBe(
      "Revised cell value",
    );
    expect(acceptView.getChanges().length).toBeGreaterThan(0);

    const output = await parseDocx(result.buffer, {
      detectVariables: false,
      preloadFonts: false,
    });
    const shape = findTextBoxShape(output);
    expect(shape.size).toEqual({ width: 1_828_800, height: 914_400 });
    expect(shape.textBody?.margins).toEqual({
      top: 45_720,
      bottom: 45_720,
      left: 91_440,
      right: 91_440,
    });
    expect(shape.textBody?.content.map(({ type }) => type)).toEqual([
      "paragraph",
      "table",
      "paragraph",
    ]);

    const rejectView = await FolioDocxReviewer.fromBuffer(result.buffer);
    rejectView.rejectAll();
    expect(rejectView.snapshot().blocks.find(({ id }) => id === "22000003")?.text).toBe(
      "Original cell value",
    );
  });

  test("accept-all reproduces the revised document; reject-all restores the base", async () => {
    const base = await buildDocxBuffer([
      { text: "Alpha paragraph stays untouched.", paraId: "00000001" },
      { text: "Payment is due within thirty days.", paraId: "00000002" },
      { text: "This clause is removed entirely.", paraId: "00000003" },
      { text: "Omega paragraph closes the document.", paraId: "00000004" },
    ]);
    const revised = await buildDocxBuffer([
      { text: "Alpha paragraph stays untouched.", paraId: "00000001" },
      { text: "First inserted clause about notices.", paraId: "00000005" },
      { text: "Second inserted clause about liability.", paraId: "00000006" },
      { text: "Payment is due within sixty days.", paraId: "00000002" },
      { text: "Omega paragraph closes the document.", paraId: "00000004" },
      { text: "Trailing signature paragraph.", paraId: "00000007" },
    ]);
    const revisedTexts = [
      "Alpha paragraph stays untouched.",
      "First inserted clause about notices.",
      "Second inserted clause about liability.",
      "Payment is due within sixty days.",
      "Omega paragraph closes the document.",
      "Trailing signature paragraph.",
    ];
    const baseTexts = [
      "Alpha paragraph stays untouched.",
      "Payment is due within thirty days.",
      "This clause is removed entirely.",
      "Omega paragraph closes the document.",
    ];

    const result = await generateRedlineDocx(base, revised);

    expect(result.skipped).toEqual([]);
    expect(result.unprocessedStories).toEqual([]);
    expect(result.applied.length).toBeGreaterThan(0);

    // The redline carries real tracked changes attributed to the default author.
    const acceptView = await FolioDocxReviewer.fromBuffer(result.buffer);
    const changes = acceptView.getChanges();
    expect(changes.length).toBeGreaterThan(0);
    expect(new Set(changes.map((change) => change.author))).toEqual(new Set(["folio compare"]));

    // Invariant 1: the as-accepted view (snapshot) equals the revised
    // document, plus the paragraph the revision removed. Its runs and its
    // paragraph mark both carry `w:del`, so it stands blank where it was,
    // before the Omega paragraph, and closes away only once the deletion is
    // accepted for real.
    const removedClauseIndex = revisedTexts.indexOf("Omega paragraph closes the document.");
    expect(blockTexts(acceptView)).toEqual(revisedTexts.toSpliced(removedClauseIndex, 0, ""));

    // Invariant 2: rejecting every change restores the base document.
    const rejectView = await FolioDocxReviewer.fromBuffer(result.buffer);
    rejectView.rejectAll();
    expect(blockTexts(rejectView)).toEqual(baseTexts);
  });

  test("redlines formatting-only changes and preserves accept/reject invariants", async () => {
    const base = await buildDocxBuffer([
      { text: "Payment is due.", paraId: "00000001", formatting: { bold: true } },
    ]);
    const revised = await buildDocxBuffer([
      { text: "Payment is due.", paraId: "00000001", formatting: { italic: true } },
    ]);

    const result = await generateRedlineDocx(base, revised);
    const baseFormatting = (await FolioDocxReviewer.fromBuffer(base))
      .snapshot()
      .blocks.at(0)?.previewRuns;
    const revisedFormatting = (await FolioDocxReviewer.fromBuffer(revised))
      .snapshot()
      .blocks.at(0)?.previewRuns;

    expect(result.skipped).toEqual([]);
    expect(result.applied).toHaveLength(1);
    const output = await FolioDocxReviewer.fromBuffer(result.buffer);
    expect(output.getChanges()).toEqual([
      expect.objectContaining({
        type: "formatting",
        text: "Payment is due.",
        author: "folio compare",
      }),
    ]);
    expect(output.snapshot().blocks.at(0)?.previewRuns).toEqual(revisedFormatting);

    const rejecting = await FolioDocxReviewer.fromBuffer(result.buffer);
    rejecting.rejectAll();
    expect(rejecting.snapshot().blocks.at(0)?.previewRuns).toEqual(baseFormatting);
  });

  test("redlines only the character span whose formatting changed", async () => {
    const base = await buildFormattedRunDocxBuffer([{ text: "Payment is due." }], "00000001");
    const revised = await buildFormattedRunDocxBuffer(
      [{ text: "Payment " }, { text: "is due", formatting: { bold: true } }, { text: "." }],
      "00000001",
    );

    const result = await generateRedlineDocx(base, revised);
    const baseFormatting = (await FolioDocxReviewer.fromBuffer(base))
      .snapshot()
      .blocks.at(0)?.previewRuns;
    const revisedFormatting = (await FolioDocxReviewer.fromBuffer(revised))
      .snapshot()
      .blocks.at(0)?.previewRuns;

    expect(result.skipped).toEqual([]);
    expect(result.applied).toHaveLength(1);
    const output = await FolioDocxReviewer.fromBuffer(result.buffer);
    expect(output.getChanges()).toEqual([
      expect.objectContaining({
        type: "formatting",
        text: "is due",
      }),
    ]);
    expect(output.snapshot().blocks.at(0)?.previewRuns).toEqual(revisedFormatting);

    const rejecting = await FolioDocxReviewer.fromBuffer(result.buffer);
    rejecting.rejectAll();
    expect(rejecting.snapshot().blocks.at(0)?.previewRuns).toEqual(baseFormatting);
  });

  test("allows exactly the generated operation limit for ordinary block operations", async () => {
    const [base, revised] = await Promise.all([
      buildDocxBuffer([]),
      buildDocxBuffer(buildInsertedParagraphs(MAX_GENERATED_REDLINE_OPERATIONS)),
    ]);

    const result = await generateRedlineDocx(base, revised);

    expect(result.applied).toHaveLength(MAX_GENERATED_REDLINE_OPERATIONS);
    expect(result.skipped).toEqual([]);
  }, 60_000);

  test("rejects the first ordinary block operation beyond the limit", async () => {
    const [base, revised] = await Promise.all([
      buildDocxBuffer([]),
      buildDocxBuffer(buildInsertedParagraphs(MAX_GENERATED_REDLINE_OPERATIONS + 1)),
    ]);

    await expectOperationLimitError(generateRedlineDocx(base, revised));
  }, 60_000);

  test("bounds generated operations for highly fragmented formatting changes", async () => {
    const runCount = MAX_GENERATED_REDLINE_OPERATIONS + 1;
    const base = await buildFormattedRunDocxBuffer([{ text: "a".repeat(runCount) }], "00000001");
    const revised = await buildFormattedRunDocxBuffer(
      Array.from({ length: runCount }, (_, index) => ({
        text: "a",
        formatting: index % 2 === 0 ? { bold: true } : { italic: true },
      })),
      "00000001",
    );

    await expectOperationLimitError(generateRedlineDocx(base, revised));
  });

  test("a relocated block redlines as delete + insert and still satisfies both invariants", async () => {
    const base = await buildDocxBuffer([
      { text: "Governing law shall be Czech law.", paraId: "00000001" },
      { text: "Notices must be delivered in writing.", paraId: "00000002" },
    ]);
    const revised = await buildDocxBuffer([
      { text: "Notices must be delivered in writing.", paraId: "00000002" },
      { text: "Governing law shall be Czech law.", paraId: "00000001" },
    ]);

    const result = await generateRedlineDocx(base, revised);
    expect(result.skipped).toEqual([]);

    const acceptView = await FolioDocxReviewer.fromBuffer(result.buffer);
    // The relocated block is deleted where it stood and inserted again below.
    // The deleted paragraph closed the body, so the break that went with it is
    // the one before it: the preceding paragraph carries `w:pPr/w:rPr/w:del`.
    // Until that is accepted the emptied paragraph is still there, blank.
    expect(blockTexts(acceptView)).toEqual([
      "Notices must be delivered in writing.",
      "Governing law shall be Czech law.",
      "",
    ]);

    const rejectView = await FolioDocxReviewer.fromBuffer(result.buffer);
    rejectView.rejectAll();
    expect(blockTexts(rejectView)).toEqual([
      "Governing law shall be Czech law.",
      "Notices must be delivered in writing.",
    ]);
  });

  test("consecutive additions after the final base block keep revised-document order", async () => {
    const base = await buildDocxBuffer([
      { text: "Agreement terms.", paraId: "00000001" },
      { text: "Final existing clause.", paraId: "00000002" },
    ]);
    const revised = await buildDocxBuffer([
      { text: "Agreement terms.", paraId: "00000001" },
      { text: "Final existing clause.", paraId: "00000002" },
      { text: "First appendix.", paraId: "00000003" },
      { text: "Second appendix.", paraId: "00000004" },
      { text: "Third appendix.", paraId: "00000005" },
    ]);

    const result = await generateRedlineDocx(base, revised);

    expect(result.skipped).toEqual([]);
    const acceptView = await FolioDocxReviewer.fromBuffer(result.buffer);
    expect(blockTexts(acceptView)).toEqual([
      "Agreement terms.",
      "Final existing clause.",
      "First appendix.",
      "Second appendix.",
      "Third appendix.",
    ]);

    const rejectView = await FolioDocxReviewer.fromBuffer(result.buffer);
    rejectView.rejectAll();
    expect(blockTexts(rejectView)).toEqual(["Agreement terms.", "Final existing clause."]);
  });

  test("identical documents produce a redline with no tracked changes", async () => {
    const paragraphs: ParagraphSpec[] = [
      { text: "Alpha paragraph.", paraId: "00000001" },
      { text: "Beta paragraph.", paraId: "00000002" },
    ];
    const base = await buildDocxBuffer(paragraphs);
    const revised = await buildDocxBuffer(paragraphs);

    const result = await generateRedlineDocx(base, revised);

    expect(result.applied).toEqual([]);
    expect(result.skipped).toEqual([]);
    const view = await FolioDocxReviewer.fromBuffer(result.buffer);
    expect(view.getChanges()).toEqual([]);
    expect(blockTexts(view)).toEqual(["Alpha paragraph.", "Beta paragraph."]);
  });

  test("an empty base document redlines additions and preserves accept/reject invariants", async () => {
    const base = await buildDocxBuffer([]);
    const revised = await buildDocxBuffer([
      { text: "First paragraph.", paraId: "00000001" },
      { text: "Second paragraph.", paraId: "00000002" },
    ]);

    const result = await generateRedlineDocx(base, revised);

    expect(result.skipped).toEqual([]);
    const acceptView = await FolioDocxReviewer.fromBuffer(result.buffer);
    expect(blockTexts(acceptView)).toEqual(["First paragraph.", "Second paragraph."]);

    const rejectView = await FolioDocxReviewer.fromBuffer(result.buffer);
    rejectView.rejectAll();
    // The empty base is one blank paragraph, not nothing: a package always
    // holds at least one. Rejecting the additions restores exactly that.
    expect(blockTexts(rejectView)).toEqual([""]);
  });

  test("a custom author is recorded on the generated changes", async () => {
    const base = await buildDocxBuffer([{ text: "Payment is due.", paraId: "00000001" }]);
    const revised = await buildDocxBuffer([
      { text: "Payment is due promptly.", paraId: "00000001" },
    ]);

    const result = await generateRedlineDocx(base, revised, { author: "Jan Kubica" });

    const view = await FolioDocxReviewer.fromBuffer(result.buffer);
    const authors = new Set(view.getChanges().map((change) => change.author));
    expect(authors).toEqual(new Set(["Jan Kubica"]));
  });

  test("redlines matched body, header, footer, footnote, and endnote stories", async () => {
    const base = await buildStoryDocument({
      bodyText: "Body baseline.",
      headerText: "Header baseline.",
      footerText: "Footer baseline.",
      footnoteText: "Footnote baseline.",
      endnoteText: "Endnote baseline.",
    });
    const revised = await buildStoryDocument({
      bodyText: "Body revised.",
      headerText: "Header revised.",
      footerText: "Footer revised.",
      footnoteText: "Footnote revised.",
      endnoteText: "Endnote revised.",
    });

    const result = await generateRedlineDocx(base, revised);

    expect(result.skipped).toEqual([]);
    expect(result.unprocessedStories).toEqual([]);
    expect(result.applied).toHaveLength(5);
    const output = await FolioDocxReviewer.fromBuffer(result.buffer);
    expect(storyTextByType(output, "original")).toEqual({
      main: "Body baseline.",
      header: "Header baseline.",
      footer: "Footer baseline.",
      footnote: "Footnote baseline.",
      endnote: "Endnote baseline.",
    });
    expect(storyTextByType(output, "final")).toEqual({
      main: "Body revised.",
      header: "Header revised.",
      footer: "Footer revised.",
      footnote: "Footnote revised.",
      endnote: "Endnote revised.",
    });
  });

  test("selects original or final input views independently", async () => {
    const base = await withPendingMainChange(
      await buildDocxBuffer([{ text: "Base original.", paraId: "00000001" }]),
      "Base final.",
    );
    const revised = await withPendingMainChange(
      await buildDocxBuffer([{ text: "Revised original.", paraId: "00000001" }]),
      "Revised final.",
    );

    const result = await generateRedlineDocx(base, revised, {
      baseView: "original",
      revisedView: "original",
    });
    const output = await FolioDocxReviewer.fromBuffer(result.buffer);

    expect(output.readReviewedStory({ view: "original" })?.snapshot.blocks.at(0)?.text).toBe(
      "Base original.",
    );
    expect(output.readReviewedStory({ view: "final" })?.snapshot.blocks.at(0)?.text).toBe(
      "Revised original.",
    );
  });

  test("reports package parts that exist on only one side", async () => {
    const base = await buildStoryDocument({
      bodyText: "Body text.",
      headerText: "Removed header.",
    });
    const revised = await buildStoryDocument({
      bodyText: "Body text.",
      footerText: "Added footer.",
    });

    const result = await generateRedlineDocx(base, revised);

    expect(result.unprocessedStories).toEqual([
      {
        baseStory: { type: "header", relationshipId: "rIdHeader" },
        revisedStory: null,
        reason: "missing-revised-story",
      },
      {
        baseStory: null,
        revisedStory: { type: "footer", relationshipId: "rIdFooter" },
        reason: "missing-base-story",
      },
    ]);
  });

  test("applies package-metadata privacy transforms and returns their report", async () => {
    const base = await withCoreProperties(
      await buildDocxBuffer([{ text: "Base text.", paraId: "00000001" }]),
    );
    const revised = await buildDocxBuffer([{ text: "Revised text.", paraId: "00000001" }]);

    const result = await generateRedlineDocx(base, revised, {
      privacy: { transforms: ["remove-attribution", "remove-timestamps"] },
    });

    expect(result.privacyReport).toEqual({
      appliedTransforms: ["remove-attribution", "remove-timestamps"],
      removedMetadataProperties: ["creator", "lastModifiedBy", "created", "modified"],
    });
    const zip = await JSZip.loadAsync(result.buffer);
    const coreProperties = await zip.file("docProps/core.xml")?.async("text");
    expect(coreProperties).not.toContain("Private creator");
    expect(coreProperties).not.toContain("Private modifier");
    expect(coreProperties).not.toContain("dcterms:created");
    expect(coreProperties).not.toContain("dcterms:modified");
    expect(coreProperties).toContain("Private title");
    expect(coreProperties).toContain("<cp:revision>7</cp:revision>");
  });

  test("rejects unresolved markup as an input view", async () => {
    const document = await buildDocxBuffer([{ text: "Body text.", paraId: "00000001" }]);

    await expect(
      Reflect.apply(generateRedlineDocx, undefined, [
        document,
        document,
        { baseView: "current-markup" },
      ]),
    ).rejects.toBeInstanceOf(InvalidGenerateRedlineDocxOptionsError);
  });
});

describe("generateRedlineDocx inserted list items", () => {
  type ListParagraphSpec = { text: string; paraId: string; numId?: number; ilvl?: number };
  type NumberingSpec = { numId: number; abstractNumId: number; numFmt: string; lvlText: string };

  const buildListDocx = (
    paragraphs: readonly ListParagraphSpec[],
    numbering: readonly NumberingSpec[],
  ): Promise<ArrayBuffer> => {
    const document = createEmptyDocument();
    document.package.document.content = paragraphs.map(({ text, paraId, numId, ilvl }) => ({
      type: "paragraph",
      paraId,
      textId: paraId,
      ...(numId !== undefined && {
        formatting: { numPr: { kind: "reference", numId, ilvl: ilvl ?? 0 } },
      }),
      content: [{ type: "run", content: [{ type: "text", text }] }],
    }));
    if (numbering.length > 0) {
      document.package.numbering = {
        abstractNums: numbering.map(({ abstractNumId, numFmt, lvlText }) => ({
          abstractNumId,
          levels: [
            { ilvl: 0, numFmt, lvlText, start: 1 },
            { ilvl: 1, numFmt, lvlText, start: 1 },
          ],
        })),
        nums: numbering.map(({ numId, abstractNumId }) => ({ numId, abstractNumId })),
      };
    }
    return createDocx(document);
  };

  const BULLET = { numId: 1, abstractNumId: 1, numFmt: "bullet", lvlText: "•" };
  const DECIMAL = { numId: 2, abstractNumId: 2, numFmt: "decimal", lvlText: "%1." };

  /** Reader projection after every tracked change is accepted or rejected. */
  const resolved = async (buffer: ArrayBuffer, resolution: "accept" | "reject") => {
    const reviewer = await FolioDocxReviewer.fromBuffer(buffer);
    if (resolution === "accept") reviewer.acceptAll();
    else reviewer.rejectAll();
    const saved = await reviewer.toBuffer();
    const reopened = await FolioDocxReviewer.fromBuffer(saved);
    return {
      blocks: reopened.snapshot().blocks.map((block) => {
        const paragraph = expectParagraphBlock(block);
        return {
          text: paragraph.text,
          label: paragraph.displayLabel ?? null,
          listReference: paragraph.listReference ?? null,
          statedNumbering: paragraph.statedNumbering,
        };
      }),
      saved,
    };
  };

  /** Every `w:numId` the document part references, and every one numbering.xml defines. */
  const numIdsOf = async (buffer: ArrayBuffer) => {
    const zip = await JSZip.loadAsync(buffer);
    const documentXml = (await zip.file("word/document.xml")?.async("text")) ?? "";
    const numberingXml = (await zip.file("word/numbering.xml")?.async("text")) ?? "";
    const referenced = new Set(
      [...documentXml.matchAll(/<w:numId w:val="(\d+)"/gu)].map((match) => Number(match[1])),
    );
    const defined = new Set(
      [...numberingXml.matchAll(/<w:num w:numId="(\d+)"/gu)].map((match) => Number(match[1])),
    );
    return { referenced, defined };
  };

  const expectedRows = (
    paragraphs: readonly ListParagraphSpec[],
    numbering: readonly NumberingSpec[],
  ) => {
    const counters = new Map<string, number>();
    return paragraphs.map(({ text, numId, ilvl }) => {
      if (numId === undefined) {
        return {
          text,
          label: null,
          level: null,
          statedNumbering: { kind: "inherit" as const },
          numFmt: undefined,
        };
      }
      const definition = numbering.find((candidate) => candidate.numId === numId);
      expect(definition, `Fixture numbering ${numId} is defined`).toBeDefined();
      if (!definition) throw new Error(`Fixture numbering ${numId} is not defined`);
      const level = ilvl ?? 0;
      const label =
        definition.numFmt === "bullet"
          ? definition.lvlText
          : definition.lvlText.replace(/%(?<level>\d+)/gu, (match, value: string) => {
              const counterLevel = Number(value) - 1;
              const key = `${numId}:${counterLevel}`;
              const count = counters.get(key) ?? 0;
              if (counterLevel === level) counters.set(key, count + 1);
              return String(counterLevel === level ? count + 1 : counters.get(key) || 1);
            });
      return {
        text,
        label,
        level,
        statedNumbering: { kind: "reference" as const, ilvl: level },
        numFmt: definition.numFmt,
      };
    });
  };

  const numberingFormatsOf = async (buffer: ArrayBuffer): Promise<Map<number, string>> => {
    const document = await parseDocx(buffer, { detectVariables: false, preloadFonts: false });
    const abstracts = new Map(
      (document.package.numbering?.abstractNums ?? []).map((entry) => [entry.abstractNumId, entry]),
    );
    return new Map(
      (document.package.numbering?.nums ?? []).flatMap((entry) => {
        const abstract = abstracts.get(entry.abstractNumId);
        const level = abstract?.levels.find((candidate) => candidate.ilvl === 0);
        return level?.numFmt === undefined ? [] : [[entry.numId, level.numFmt]];
      }),
    );
  };

  const expectRowsFromFixture = async (
    actual: Awaited<ReturnType<typeof resolved>>,
    paragraphs: readonly ListParagraphSpec[],
    numbering: readonly NumberingSpec[],
  ) => {
    const expected = expectedRows(paragraphs, numbering);
    const formats = await numberingFormatsOf(actual.saved);
    expect(actual.blocks).toHaveLength(expected.length);
    for (const [index, expectedRow] of expected.entries()) {
      const row = actual.blocks[index];
      expect(row).toBeDefined();
      if (!row) throw new Error(`Resolved block ${index} is absent`);
      expect(row.text).toBe(expectedRow.text);
      expect(row.label).toBe(expectedRow.label);
      expect(row.listReference?.level ?? null).toBe(expectedRow.level);
      expect(
        row.statedNumbering.kind === "reference"
          ? { kind: row.statedNumbering.kind, ilvl: row.statedNumbering.ilvl }
          : row.statedNumbering,
      ).toEqual(expectedRow.statedNumbering);
      if (expectedRow.numFmt === undefined) {
        expect(row.listReference).toBeNull();
      } else {
        expect(row.listReference).not.toBeNull();
        expect(row.statedNumbering.kind).toBe("reference");
        if (row.statedNumbering.kind !== "reference" || row.listReference === null) {
          throw new Error(`Resolved block ${index} lost its authored list reference`);
        }
        expect(row.statedNumbering.numId).toBe(row.listReference.numId);
        expect(formats.get(row.listReference.numId)).toBe(expectedRow.numFmt);
      }
    }
  };

  /** Accepting the redline gives the revision, rejecting it the base, and no numId dangles. */
  const expectRoundTrip = async ({
    base,
    baseParagraphs,
    baseNumbering,
    revised,
    revisedParagraphs,
    revisedNumbering,
  }: {
    base: ArrayBuffer;
    baseParagraphs: readonly ListParagraphSpec[];
    baseNumbering: readonly NumberingSpec[];
    revised: ArrayBuffer;
    revisedParagraphs: readonly ListParagraphSpec[];
    revisedNumbering: readonly NumberingSpec[];
  }) => {
    const result = await generateRedlineDocx(base, revised);
    expect(result.skipped).toEqual([]);
    const accepted = await resolved(result.buffer, "accept");
    await expectRowsFromFixture(accepted, revisedParagraphs, revisedNumbering);
    const rejected = await resolved(result.buffer, "reject");
    await expectRowsFromFixture(rejected, baseParagraphs, baseNumbering);
    for (const buffer of [result.buffer, accepted.saved]) {
      const { referenced, defined } = await numIdsOf(buffer);
      for (const numId of referenced) expect(defined).toContain(numId);
    }
    return { accepted, redline: result.buffer };
  };

  test("an inserted bullet keeps its bullet and level, and rejecting removes it", async () => {
    const baseParagraphs = [
      { text: "Intro.", paraId: "10000001" },
      { text: "Existing bullet.", paraId: "10000002", numId: 1 },
      { text: "Outro.", paraId: "10000003" },
    ];
    const revisedParagraphs = [
      { text: "Intro.", paraId: "10000001" },
      { text: "Existing bullet.", paraId: "10000002", numId: 1 },
      { text: "New nested bullet.", paraId: "10000004", numId: 1, ilvl: 1 },
      { text: "Outro.", paraId: "10000003" },
      { text: "Trailing bullet.", paraId: "10000005", numId: 1 },
    ];
    const base = await buildListDocx(baseParagraphs, [BULLET]);
    const revised = await buildListDocx(revisedParagraphs, [BULLET]);
    await expectRoundTrip({
      base,
      baseParagraphs,
      baseNumbering: [BULLET],
      revised,
      revisedParagraphs,
      revisedNumbering: [BULLET],
    });
  });

  test("an inserted numbered item takes its number, and a plain insertion stays plain", async () => {
    const baseParagraphs = [
      { text: "First.", paraId: "20000001", numId: 2 },
      { text: "Third.", paraId: "20000002", numId: 2 },
    ];
    const revisedParagraphs = [
      { text: "Plain before the list.", paraId: "20000003" },
      { text: "First.", paraId: "20000001", numId: 2 },
      { text: "Second.", paraId: "20000004", numId: 2 },
      { text: "Third.", paraId: "20000002", numId: 2 },
    ];
    const base = await buildListDocx(baseParagraphs, [DECIMAL]);
    const revised = await buildListDocx(revisedParagraphs, [DECIMAL]);
    await expectRoundTrip({
      base,
      baseParagraphs,
      baseNumbering: [DECIMAL],
      revised,
      revisedParagraphs,
      revisedNumbering: [DECIMAL],
    });
  });

  test("numbering defined only in the revised package is carried into the redline", async () => {
    const baseParagraphs = [{ text: "Intro.", paraId: "30000001" }];
    const revisedParagraphs = [
      { text: "Intro.", paraId: "30000001" },
      { text: "Only numbered in the revision.", paraId: "30000002", numId: 2 },
    ];
    const base = await buildListDocx(baseParagraphs, []);
    const revised = await buildListDocx(revisedParagraphs, [DECIMAL]);
    const { accepted, redline } = await expectRoundTrip({
      base,
      baseParagraphs,
      baseNumbering: [],
      revised,
      revisedParagraphs,
      revisedNumbering: [DECIMAL],
    });
    expect((await numIdsOf(redline)).referenced.size).toBe(1);
    expect(accepted.blocks.at(-1)?.statedNumbering).toEqual({
      kind: "reference",
      numId: accepted.blocks.at(-1)?.listReference?.numId,
      ilvl: 0,
    });
  });

  test("a numId the base uses for a different list is remapped, not reused", async () => {
    const baseParagraphs = [{ text: "Numbered in the base.", paraId: "40000001", numId: 1 }];
    const baseNumbering = [{ ...DECIMAL, numId: 1, abstractNumId: 1 }];
    const revisedParagraphs = [
      { text: "Numbered in the base.", paraId: "40000001", numId: 2 },
      { text: "A bullet under the same id.", paraId: "40000002", numId: 1 },
    ];
    const revisedNumbering = [BULLET, DECIMAL];
    const base = await buildListDocx(baseParagraphs, baseNumbering);
    const revised = await buildListDocx(revisedParagraphs, revisedNumbering);
    const { accepted } = await expectRoundTrip({
      base,
      baseParagraphs,
      baseNumbering,
      revised,
      revisedParagraphs,
      revisedNumbering,
    });
    expect(accepted.blocks.map(({ label }) => label)).toEqual(["1.", "•"]);
    const [numbered, bullet] = accepted.blocks;
    expect(numbered?.statedNumbering).toEqual({
      kind: "reference",
      numId: numbered?.listReference?.numId,
      ilvl: 0,
    });
    expect(bullet?.statedNumbering).toEqual({
      kind: "reference",
      numId: bullet?.listReference?.numId,
      ilvl: 0,
    });
    expect(numbered?.listReference?.numId).not.toBe(bullet?.listReference?.numId);
  });

  test("style-only insertions keep the revised colliding style after acceptance and base on rejection", async () => {
    // Earlier fixtures generated only direct references, leaving inherited list resources unexercised.
    for (const { ilvl, stated } of [undefined, 0, 1].flatMap((styleLevel) =>
      ["inherit", "levelOnly"].map((statedKind) => ({ ilvl: styleLevel, stated: statedKind })),
    )) {
      type StyleNumberingFixtureOptions = {
        inserted: boolean;
        revised: boolean;
        revisedBullet?: boolean;
        styleClosure?: "portable" | "missingLink";
      };
      const makeDocument = ({
        inserted,
        revised,
        revisedBullet = true,
        styleClosure = "portable",
      }: StyleNumberingFixtureOptions) => {
        const document = createEmptyDocument();
        const styleNumId = revised ? 2 : 1;
        document.package.document.content = [
          {
            type: "paragraph",
            paraId: "81000001",
            textId: "81000001",
            content: [{ type: "run", content: [{ type: "text", text: "Anchor." }] }],
          },
          ...(inserted
            ? [
                {
                  type: "paragraph" as const,
                  paraId: "81000002",
                  textId: "81000002",
                  formatting: {
                    styleId: "Numbered",
                    ...(stated === "levelOnly" && {
                      numPr: { kind: "levelOnly" as const, ilvl: 0 },
                    }),
                  },
                  content: [
                    {
                      type: "run" as const,
                      content: [{ type: "text" as const, text: "Inserted by style." }],
                    },
                  ],
                },
              ]
            : []),
        ];
        document.package.styles = {
          styles: [
            { type: "paragraph", styleId: "Normal", name: "Normal", default: true },
            {
              type: "paragraph",
              styleId: "Numbered",
              name: "Numbered",
              ...(styleClosure === "missingLink" && { link: "MissingStyle" }),
              pPr: {
                numPr: paragraphNumberingReference({
                  numId: styleNumId,
                  ...(ilvl === undefined ? {} : { ilvl }),
                }),
              },
            },
          ],
        };
        document.package.numbering = {
          abstractNums: [
            {
              abstractNumId: 1,
              levels: [0, 1].map((level) => ({ ilvl: level, numFmt: "decimal", lvlText: "%1." })),
            },
            {
              abstractNumId: 2,
              levels: [0, 1].map((level) => ({
                ilvl: level,
                numFmt: revised && revisedBullet ? "bullet" : "decimal",
                lvlText: revised && revisedBullet ? "▪" : "%1.",
              })),
            },
          ],
          nums: [
            { numId: 1, abstractNumId: 1 },
            { numId: 2, abstractNumId: 2 },
          ],
        };
        return createDocx(document);
      };

      const base = await makeDocument({ inserted: false, revised: false });
      const revised = await makeDocument({ inserted: true, revised: true });
      const result = await generateRedlineDocx(base, revised);
      expect(result.skipped).toEqual([]);

      const accepting = await FolioDocxReviewer.fromBuffer(result.buffer);
      accepting.acceptAll();
      const acceptedBuffer = await accepting.toBuffer();
      const accepted = await FolioDocxReviewer.fromBuffer(acceptedBuffer);
      const acceptedBlock = expectParagraphBlock(accepted.snapshot().blocks.at(-1));
      expect(acceptedBlock.text).toBe("Inserted by style.");
      expect(acceptedBlock.statedNumbering).toEqual(
        stated === "inherit" ? { kind: "inherit" } : { kind: "levelOnly", ilvl: 0 },
      );
      expect(acceptedBlock.listReference?.level).toBe(stated === "levelOnly" ? 0 : (ilvl ?? 0));
      expect(acceptedBlock.displayLabel).toBe("▪");
      const acceptedDocument = accepted.toDocument();
      const acceptedParagraph = acceptedDocument.package.document.content.at(-1);
      if (acceptedParagraph?.type !== "paragraph") throw new Error("Expected inserted paragraph");
      expect(acceptedParagraph.formatting?.styleId).not.toBe("Numbered");
      expect(acceptedParagraph.formatting?.numPr).toEqual(
        stated === "inherit" ? undefined : { kind: "levelOnly", ilvl: 0 },
      );
      expect(acceptedParagraph.formatting?.numPrFromStyle).toEqual(
        paragraphNumberingReference({
          numId: acceptedBlock.listReference?.numId ?? 0,
          ...(ilvl === undefined ? {} : { ilvl }),
        }),
      );

      const rejecting = await FolioDocxReviewer.fromBuffer(result.buffer);
      rejecting.rejectAll();
      const rejected = await FolioDocxReviewer.fromBuffer(await rejecting.toBuffer());
      expect(rejected.snapshot().blocks.map((block) => block.text)).toEqual(["Anchor."]);
      expect(
        rejected.toDocument().package.styles?.styles.find((style) => style.styleId === "Numbered")
          ?.pPr?.numPr,
      ).toEqual(paragraphNumberingReference({ numId: 1, ...(ilvl === undefined ? {} : { ilvl }) }));

      const identicalRevision = await makeDocument({
        inserted: true,
        revised: false,
        revisedBullet: false,
      });
      const identicalResult = await generateRedlineDocx(base, identicalRevision);
      const identicalReviewer = await FolioDocxReviewer.fromBuffer(identicalResult.buffer);
      identicalReviewer.acceptAll();
      const identicalAccepted = await FolioDocxReviewer.fromBuffer(
        await identicalReviewer.toBuffer(),
      );
      const identicalParagraph = identicalAccepted.toDocument().package.document.content.at(-1);
      if (identicalParagraph?.type !== "paragraph")
        throw new Error("Expected identical-definition insertion");
      expect(identicalParagraph.formatting?.styleId).toBe("Numbered");
      expect(identicalAccepted.snapshot().blocks.at(-1)?.listReference?.numId).toBe(1);

      const unsupportedRevision = await makeDocument({
        inserted: true,
        revised: true,
        styleClosure: "missingLink",
      });
      const refusal = await generateRedlineDocx(base, unsupportedRevision).then(
        () => null,
        (error: unknown) => error,
      );
      expect(refusal).toBeInstanceOf(GenerateRedlineDocxResourceImportError);
      expect(refusal).toMatchObject({
        _tag: "GenerateRedlineDocxResourceImportError",
        detail: "referenced target style MissingStyle is missing",
      });
    }
  });
});

test("the current text-redline contract flattens new carriers and imports their paragraph styles once", async () => {
  for (const innerContent of ["paragraph", "tableCell"] as const) {
    const base = await createDocx(
      insertedTextBoxDocument({ side: "base", innerContent, suffix: "inserted" }),
    );
    const revised = await createDocx(
      insertedTextBoxDocument({ side: "revised", innerContent, suffix: "inserted" }),
    );
    const result = await generateRedlineDocx(base, revised);
    expect(result.skipped).toEqual([]);
    expect(result.unprocessedStories).toEqual([]);
    const accepting = await FolioDocxReviewer.fromBuffer(result.buffer);
    expect(accepting.acceptAll()).toBeGreaterThan(0);
    const accepted = await FolioDocxReviewer.fromBuffer(await accepting.toBuffer());
    expect(accepted.snapshot().blocks.map(({ text }) => text)).toEqual([
      "Unchanged anchor",
      "Parent text inserted",
      "Inner text inserted",
    ]);
    const inserted = accepted
      .snapshot()
      .blocks.filter(({ text }) => text !== "Unchanged anchor")
      .map((block) => expectParagraphBlock(block));
    expect(inserted).toHaveLength(2);
    const importedStyleIds = new Set(inserted.map(({ styleId }) => styleId));
    expect(importedStyleIds.size).toBe(1);
    expect(importedStyleIds.has(INSERTED_TEXT_BOX_STYLE)).toBe(false);
    for (const block of inserted) {
      expect(block.previewRuns).toContainEqual(expect.objectContaining({ bold: true }));
    }
    const styles = accepted.toDocument().package.styles?.styles ?? [];
    expect(styles.filter(({ styleId }) => importedStyleIds.has(styleId))).toHaveLength(1);
    expect(styles.find(({ styleId }) => styleId === INSERTED_TEXT_BOX_STYLE)?.rPr?.bold).toBe(
      false,
    );
    // Insertions currently carry text and paragraph properties, not container structure.
    expect(accepted.toDocument().package.document.content.map(({ type }) => type)).toEqual([
      "paragraph",
      "paragraph",
      "paragraph",
    ]);
    const rejecting = await FolioDocxReviewer.fromBuffer(result.buffer);
    expect(rejecting.rejectAll()).toBeGreaterThan(0);
    const rejected = await FolioDocxReviewer.fromBuffer(await rejecting.toBuffer());
    expect(rejected.snapshot().blocks.map(({ text }) => text)).toEqual(["Unchanged anchor"]);
  }
});
