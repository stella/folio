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

import { FolioDocxReviewer } from "./ai-edits/headless";
import { parseDocx } from "./docx/parser";
import { createDocx } from "./docx/rezip";
import { repackDocx } from "./docx/rezip";
import { generateRedlineDocx, InvalidGenerateRedlineDocxOptionsError } from "./redline";
import type { HeaderFooter, Paragraph } from "./types/document";
import { createEmptyDocument } from "./utils/createDocument";

type ParagraphSpec = { text: string; paraId?: string };

const CORE_PROPERTIES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="urn:properties" xmlns:dc="urn:descriptive" xmlns:dcterms="urn:terms"><dc:title>Private title</dc:title><dc:creator>Private creator</dc:creator><cp:lastModifiedBy>Private modifier</cp:lastModifiedBy><cp:revision>7</cp:revision><dcterms:created>2026-07-01T10:30:00Z</dcterms:created><dcterms:modified>2026-07-02T11:45:00Z</dcterms:modified></cp:coreProperties>`;

const buildDocxBuffer = (paragraphs: readonly ParagraphSpec[]): Promise<ArrayBuffer> => {
  const template = createEmptyDocument();
  return createDocx({
    ...template,
    package: {
      ...template.package,
      document: {
        ...template.package.document,
        content: paragraphs.map(({ text, paraId }) => ({
          type: "paragraph",
          content: [{ type: "run", content: [{ type: "text", text }] }],
          ...(paraId !== undefined && { paraId }),
        })),
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
  const source = await buildDocxBuffer([{ text: bodyText, paraId: "A1000001" }]);
  const document = await parseDocx(source, { detectVariables: false, preloadFonts: false });
  if (headerText !== undefined) {
    document.package.headers = new Map([
      ["rIdHeader", headerFooterStory("header", headerText, "B1000001")],
    ]);
    document.package.document.finalSectionProperties = {
      ...document.package.document.finalSectionProperties,
      headerReferences: [{ type: "default", rId: "rIdHeader" }],
    };
  }
  if (footerText !== undefined) {
    document.package.footers = new Map([
      ["rIdFooter", headerFooterStory("footer", footerText, "C1000001")],
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
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:footnote w:id="2"><w:p w14:paraId="D1000001"><w:r><w:t>${footnoteText}</w:t></w:r></w:p></w:footnote></w:footnotes>`,
    );
  }
  if (endnoteText !== undefined) {
    zip.file(
      "word/endnotes.xml",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:endnote w:id="3"><w:p w14:paraId="E1000001"><w:r><w:t>${endnoteText}</w:t></w:r></w:p></w:endnote></w:endnotes>`,
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

    // Invariant 1: the as-accepted view (snapshot) equals the revised document.
    expect(blockTexts(acceptView)).toEqual(revisedTexts);

    // Invariant 2: rejecting every change restores the base document.
    const rejectView = await FolioDocxReviewer.fromBuffer(result.buffer);
    rejectView.rejectAll();
    expect(blockTexts(rejectView)).toEqual(baseTexts);
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
    expect(blockTexts(acceptView)).toEqual([
      "Notices must be delivered in writing.",
      "Governing law shall be Czech law.",
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
    expect(blockTexts(rejectView)).toEqual([]);
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
