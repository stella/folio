/**
 * Version-diff engine tests: real-paraId alignment, the deterministic-fallback-id
 * regression guard (same text, shifted ordinal must still pair as unchanged via
 * the text-LCS pass), identical-buffer no-op, and the as-accepted semantics
 * over a revised document that carries pending tracked changes.
 *
 * Buffers are built directly from the typed `Document` model (`createEmptyDocument`
 * as a template, `createDocx` to serialize) rather than a fixture file, since each
 * case needs precise control over paragraph text/paraId and no existing corpus
 * fixture has more than two named paragraphs.
 */

import { describe, expect, test } from "bun:test";
import JSZip from "jszip";

import { buildTextBoxTableDocument } from "./__tests__/textBoxTableDocument";
import { FolioDocxReviewer, getFolioDocxComparisonAccess } from "./ai-edits/headless";
import {
  compareContent,
  createContentComparisonWorkSession,
  FOLIO_CONTENT_COMPARISON_LIMITS,
} from "./compare/content";
import {
  resolvedDocxContentBlocks,
  resolvedDocxContentSnapshot,
} from "./internal/compare/resolved-docx-story-snapshot";
import { parseDocx } from "./docx/parser";
import { createDocx } from "./docx/rezip";
import { repackDocx } from "./docx/rezip";
import type { HeaderFooter, Paragraph, ParagraphAlignment, Table } from "./types/document";
import { createEmptyDocument } from "./utils/createDocument";
import {
  applyFolioVersionDiffPrivacy,
  compareDocxVersions,
  FolioVersionComparisonLimitError,
  projectFolioContentComparisonToStory,
} from "./version-comparison";
import { InvalidFolioVersionComparisonOptionsError } from "./version-comparison";
import type { FolioBlockDiff } from "./version-comparison";

type ParagraphSpec = {
  text: string;
  paraId?: string;
  formatting?: { bold?: boolean; italic?: boolean };
  paragraphAlignment?: ParagraphAlignment;
  paragraphStyleId?: string;
};

const buildDocxBuffer = (paragraphs: readonly ParagraphSpec[]): Promise<ArrayBuffer> => {
  const template = createEmptyDocument();
  return createDocx({
    ...template,
    package: {
      ...template.package,
      document: {
        ...template.package.document,
        content: paragraphs.map(
          ({ text, paraId, formatting, paragraphAlignment, paragraphStyleId }) => ({
            type: "paragraph",
            content: [
              {
                type: "run",
                ...(formatting !== undefined && { formatting }),
                content: [{ type: "text", text }],
              },
            ],
            ...((paragraphAlignment !== undefined || paragraphStyleId !== undefined) && {
              formatting: {
                ...(paragraphAlignment !== undefined && { alignment: paragraphAlignment }),
                ...(paragraphStyleId !== undefined && { styleId: paragraphStyleId }),
              },
            }),
            ...(paraId !== undefined && { paraId }),
          }),
        ),
      },
    },
  });
};

const projectMainContent = (reviewer: FolioDocxReviewer) => {
  const story = getFolioDocxComparisonAccess(reviewer)
    .projectResolvedStories()
    .stories.find(({ handle }) => handle.type === "main")?.snapshot;
  if (!story) throw new Error("main story projection missing");
  return {
    snapshot: resolvedDocxContentSnapshot(story),
    blocks: resolvedDocxContentBlocks(story),
  };
};

type CorePropertiesFixture = {
  title?: string;
  subject?: string;
  creator?: string;
  keywords?: string;
  description?: string;
  lastModifiedBy?: string;
  revision?: number;
  created?: string;
  modified?: string;
};

const escapeXml = (value: string): string =>
  value
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&apos;");

const withCoreProperties = async (
  buffer: ArrayBuffer,
  properties: CorePropertiesFixture,
): Promise<ArrayBuffer> => {
  const zip = await JSZip.loadAsync(buffer);
  const elements = [
    properties.title === undefined ? "" : `<dc:title>${escapeXml(properties.title)}</dc:title>`,
    properties.subject === undefined
      ? ""
      : `<dc:subject>${escapeXml(properties.subject)}</dc:subject>`,
    properties.creator === undefined
      ? ""
      : `<dc:creator>${escapeXml(properties.creator)}</dc:creator>`,
    properties.keywords === undefined
      ? ""
      : `<cp:keywords>${escapeXml(properties.keywords)}</cp:keywords>`,
    properties.description === undefined
      ? ""
      : `<dc:description>${escapeXml(properties.description)}</dc:description>`,
    properties.lastModifiedBy === undefined
      ? ""
      : `<cp:lastModifiedBy>${escapeXml(properties.lastModifiedBy)}</cp:lastModifiedBy>`,
    properties.revision === undefined ? "" : `<cp:revision>${properties.revision}</cp:revision>`,
    properties.created === undefined
      ? ""
      : `<dcterms:created>${escapeXml(properties.created)}</dcterms:created>`,
    properties.modified === undefined
      ? ""
      : `<dcterms:modified>${escapeXml(properties.modified)}</dcterms:modified>`,
  ].join("");
  zip.file(
    "docProps/core.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/">${elements}</cp:coreProperties>`,
  );
  return zip.generateAsync({ type: "arraybuffer" });
};

const storyParagraph = (text: string, paraId: string): Paragraph => ({
  type: "paragraph",
  paraId,
  content: [{ type: "run", content: [{ type: "text", text }] }],
});

const buildTableCellDocxBuffer = (text: string): Promise<ArrayBuffer> => {
  const template = createEmptyDocument();
  const table: Table = {
    type: "table",
    rows: [
      {
        type: "tableRow",
        cells: [
          {
            type: "tableCell",
            content: [storyParagraph(text, "00000001")],
          },
        ],
      },
    ],
  };
  return createDocx({
    ...template,
    package: {
      ...template.package,
      document: {
        ...template.package.document,
        content: [table],
      },
    },
  });
};

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
  headerText: string;
  footnoteText: string;
  footerText?: string;
};

const buildStoryDocument = async ({
  bodyText,
  headerText,
  footnoteText,
  footerText,
}: StoryDocumentOptions): Promise<ArrayBuffer> => {
  const source = await buildDocxBuffer([{ text: bodyText, paraId: "21000001" }]);
  const document = await parseDocx(source, { detectVariables: false, preloadFonts: false });
  document.package.headers = new Map([
    ["rIdHeader", headerFooterStory("header", headerText, "31000001")],
  ]);
  document.package.document.finalSectionProperties = {
    ...document.package.document.finalSectionProperties,
    headerReferences: [{ type: "default", rId: "rIdHeader" }],
  };
  if (footerText !== undefined) {
    document.package.footers = new Map([
      ["rIdFooter", headerFooterStory("footer", footerText, "51000001")],
    ]);
    document.package.document.finalSectionProperties.footerReferences = [
      { type: "default", rId: "rIdFooter" },
    ];
  }
  const materialized = await repackDocx(document, { updateModifiedDate: false });
  const zip = await JSZip.loadAsync(materialized);
  const contentTypesFile = zip.file("[Content_Types].xml");
  const relationshipsFile = zip.file("word/_rels/document.xml.rels");
  if (!contentTypesFile || !relationshipsFile) {
    throw new Error("expected package metadata parts");
  }
  const contentTypes = await contentTypesFile.async("text");
  const relationships = await relationshipsFile.async("text");
  zip.file(
    "[Content_Types].xml",
    contentTypes.replace(
      "</Types>",
      '<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>',
    ),
  );
  zip.file(
    "word/_rels/document.xml.rels",
    relationships.replace(
      "</Relationships>",
      '<Relationship Id="rIdFootnotes" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>',
    ),
  );
  zip.file(
    "word/footnotes.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:footnote w:id="2"><w:p w14:paraId="41000001"><w:r><w:t>${footnoteText}</w:t></w:r></w:p></w:footnote></w:footnotes>`,
  );
  return zip.generateAsync({ type: "arraybuffer" });
};

const findChange = (changes: readonly FolioBlockDiff[], blockId: string): FolioBlockDiff => {
  const change = changes.find((c) => c.blockId === blockId);
  if (!change) {
    throw new Error(`no change for block ${blockId}`);
  }
  return change;
};

describe("compareDocxVersions: real w14:paraId alignment", () => {
  test("reports a nested text-box table-cell edit with stable source handles", async () => {
    const diff = await compareDocxVersions(
      await buildTextBoxTableDocument("Original cell value"),
      await buildTextBoxTableDocument("Revised cell value"),
    );

    expect(diff.summaryCounts).toMatchObject({ modified: 1 });
    expect(diff.changes).toHaveLength(1);
    const change = diff.changes.at(0);
    expect(change).toMatchObject({
      type: "modified",
      blockId: "22000003",
      baseHandle: { story: { type: "main" }, blockId: "22000003" },
      revisedHandle: { story: { type: "main" }, blockId: "22000003" },
    });
    if (change?.type !== "modified") {
      throw new Error("expected a modified cell paragraph");
    }
    expect(change.segments).toEqual([
      { type: "del", text: "Original" },
      { type: "ins", text: "Revised" },
      { type: "equal", text: " cell value" },
    ]);
  });

  test("classifies unchanged, modified, added, and deleted blocks by stable id", async () => {
    const base = await buildDocxBuffer([
      { text: "Alpha paragraph.", paraId: "00000001" },
      { text: "Beta paragraph.", paraId: "00000002" },
      { text: "Gamma paragraph.", paraId: "00000003" },
      { text: "Zeta paragraph.", paraId: "00000005" },
    ]);
    const revised = await buildDocxBuffer([
      { text: "Alpha paragraph.", paraId: "00000001" },
      { text: "Beta clause.", paraId: "00000002" },
      { text: "Zeta paragraph.", paraId: "00000005" },
      { text: "Epsilon paragraph.", paraId: "00000006" },
    ]);

    const diff = await compareDocxVersions(base, revised);

    expect(diff.summaryCounts).toEqual({
      added: 1,
      deleted: 1,
      modified: 1,
      formatChanged: 0,
      moved: 0,
      metadataChanged: 0,
      unchanged: 2,
    });
    // Unchanged blocks (Alpha, Zeta) never appear in `changes`.
    expect(diff.changes).toHaveLength(3);
    expect(diff.changes.map((c) => c.type)).toEqual(["modified", "deleted", "added"]);

    const modified = findChange(diff.changes, "00000002");
    if (modified.type !== "modified") {
      throw new Error("expected a modified change");
    }
    expect(modified.kind).toBe("paragraph");
    expect(modified.segments).toEqual([
      { type: "equal", text: "Beta" },
      { type: "del", text: " paragraph." },
      { type: "ins", text: " clause." },
    ]);

    const deleted = findChange(diff.changes, "00000003");
    expect(deleted).toEqual({
      type: "deleted",
      blockId: "00000003",
      kind: "paragraph",
      text: "Gamma paragraph.",
      baseHandle: { story: { type: "main" }, blockId: "00000003" },
    });

    const added = findChange(diff.changes, "00000006");
    expect(added).toEqual({
      type: "added",
      blockId: "00000006",
      kind: "paragraph",
      text: "Epsilon paragraph.",
      revisedHandle: { story: { type: "main" }, blockId: "00000006" },
    });
  });
});

describe("compareDocxVersions: document stories", () => {
  test("reports per-story changes with source-specific navigation handles", async () => {
    const base = await buildStoryDocument({
      bodyText: "Body text.",
      headerText: "Header baseline.",
      footnoteText: "Stable note.",
    });
    const revised = await buildStoryDocument({
      bodyText: "Body text.",
      headerText: "Header revised.",
      footnoteText: "Stable note.",
      footerText: "Added footer.",
    });

    const diff = await compareDocxVersions(base, revised);

    expect(diff.summaryCounts).toEqual({
      added: 1,
      deleted: 0,
      modified: 1,
      formatChanged: 0,
      moved: 0,
      metadataChanged: 0,
      unchanged: 2,
    });
    expect(diff.stories).toHaveLength(4);

    const header = diff.stories.find(({ revisedStory }) => revisedStory?.type === "header");
    expect(header?.baseStory).toEqual({ type: "header", relationshipId: "rIdHeader" });
    expect(header?.revisedStory).toEqual({ type: "header", relationshipId: "rIdHeader" });
    const headerChange = header?.changes.at(0);
    if (!headerChange || headerChange.type !== "modified") {
      throw new Error("expected a modified header block");
    }
    expect(headerChange.baseHandle).toEqual({
      story: { type: "header", relationshipId: "rIdHeader" },
      blockId: "31000001",
    });
    expect(headerChange.revisedHandle).toEqual({
      story: { type: "header", relationshipId: "rIdHeader" },
      blockId: "31000001",
    });

    const footer = diff.stories.find(({ revisedStory }) => revisedStory?.type === "footer");
    expect(footer?.baseStory).toBeNull();
    expect(footer?.summaryCounts.added).toBe(1);
    const footerChange = footer?.changes.at(0);
    if (!footerChange || footerChange.type !== "added") {
      throw new Error("expected an added footer block");
    }
    expect(footerChange.revisedHandle).toEqual({
      story: { type: "footer", relationshipId: "rIdFooter" },
      blockId: "51000001",
    });
  });

  test("shares one inline-alignment allowance across every story", async () => {
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

    const diff = await compareDocxVersions(base, revised);
    const modified = diff.stories.flatMap(({ changes }) =>
      changes.filter(({ type }) => type === "modified"),
    );
    expect(modified).toHaveLength(3);
    expect(
      modified.filter(({ segments }) => segments.some(({ type }) => type === "equal")),
    ).toHaveLength(1);
    expect(
      modified.filter(({ segments }) => segments.every(({ type }) => type !== "equal")),
    ).toHaveLength(2);
  });
});

describe("compareDocxVersions: deterministic fallback ids (no w14:paraId)", () => {
  test("a same-text block whose ordinal shifted still pairs as unchanged via the text-LCS pass", async () => {
    // Neither source carries a w14:paraId, so FolioDocxReviewer assigns each
    // block a deterministic id derived from text plus ordinal and marks its
    // provenance positional. The "Epsilon" insertion shifts Gamma from third
    // to fourth, changing that id even though its text is untouched. Stable-id
    // pairing must ignore these ids; only exact-text alignment recovers Gamma.
    const base = await buildDocxBuffer([
      { text: "Alpha paragraph." },
      { text: "Beta paragraph." },
      { text: "Gamma paragraph." },
    ]);
    const revised = await buildDocxBuffer([
      { text: "Alpha paragraph." },
      { text: "Beta paragraph modified." },
      { text: "Epsilon paragraph." },
      { text: "Gamma paragraph." },
    ]);

    const [baseReviewer, revisedReviewer] = await Promise.all([
      FolioDocxReviewer.fromBuffer(base),
      FolioDocxReviewer.fromBuffer(revised),
    ]);
    const baseContent = projectMainContent(baseReviewer);
    const revisedContent = projectMainContent(revisedReviewer);
    expect(
      [...baseContent.blocks, ...revisedContent.blocks].every(
        ({ identity }) => identity.type === "positional",
      ),
    ).toBe(true);

    const diff = await compareDocxVersions(base, revised);

    // Alpha (unshifted, same fallback id both sides) and Gamma (shifted,
    // recovered via text-LCS) are both unchanged and absent from `changes`.
    expect(diff.summaryCounts).toEqual({
      added: 1,
      deleted: 0,
      modified: 1,
      formatChanged: 0,
      moved: 0,
      metadataChanged: 0,
      unchanged: 2,
    });
    expect(diff.changes.map((c) => c.type)).toEqual(["modified", "added"]);

    const modified = diff.changes[0];
    if (!modified || modified.type !== "modified") {
      throw new Error("expected a modified change");
    }
    // Reconstruct the revised text from the `equal` + `ins` segments (the
    // `del` segments carry the superseded base-side text, not part of either
    // whole string on their own).
    const reconstructedRevised = modified.segments
      .filter((s) => s.type !== "del")
      .map((s) => s.text)
      .join("");
    expect(reconstructedRevised).toBe("Beta paragraph modified.");

    const added = diff.changes[1];
    if (!added || added.type !== "added") {
      throw new Error("expected an added change");
    }
    expect(added.text).toBe("Epsilon paragraph.");
  });
});

describe("compareDocxVersions: no-op", () => {
  test("identical documents produce zero changes and every block counts as unchanged", async () => {
    const paragraphs: ParagraphSpec[] = [{ text: "Alpha paragraph." }, { text: "Beta paragraph." }];
    const base = await buildDocxBuffer(paragraphs);
    const revised = await buildDocxBuffer(paragraphs);

    const diff = await compareDocxVersions(base, revised);

    expect(diff.changes).toEqual([]);
    expect(diff.summaryCounts).toEqual({
      added: 0,
      deleted: 0,
      modified: 0,
      formatChanged: 0,
      moved: 0,
      metadataChanged: 0,
      unchanged: 2,
    });
  });
});

describe("compareDocxVersions: selected scopes", () => {
  test("keeps metadata opt-in and returns typed metadata changes in stable order", async () => {
    const base = await withCoreProperties(
      await buildDocxBuffer([{ text: "Baseline text.", paraId: "00000001" }]),
      {
        title: "Initial title",
        creator: "Initial author",
        revision: 1,
        created: "2026-07-01T10:30:00Z",
      },
    );
    const revised = await withCoreProperties(
      await buildDocxBuffer([{ text: "Revised text.", paraId: "00000001" }]),
      {
        title: "Revised title",
        revision: 2,
        created: "2026-07-01T10:30:00Z",
      },
    );

    const defaultDiff = await compareDocxVersions(base, revised);
    expect(defaultDiff.metadataChanges).toEqual([]);
    expect(defaultDiff.summaryCounts.metadataChanged).toBe(0);
    expect(defaultDiff.changes.map(({ type }) => type)).toEqual(["modified"]);

    const metadataDiff = await compareDocxVersions(base, revised, { include: ["metadata"] });
    expect(metadataDiff.changes).toEqual([]);
    expect(metadataDiff.metadataChanges).toEqual([
      { property: "title", baseValue: "Initial title", revisedValue: "Revised title" },
      { property: "creator", baseValue: "Initial author", revisedValue: null },
      { property: "revision", baseValue: 1, revisedValue: 2 },
    ]);
    expect(metadataDiff.summaryCounts).toEqual({
      added: 0,
      deleted: 0,
      modified: 0,
      formatChanged: 0,
      moved: 0,
      metadataChanged: 3,
      unchanged: 1,
    });
  });

  test("selects text and formatting independently", async () => {
    const base = await buildDocxBuffer([
      { text: "Baseline text.", paraId: "00000001" },
      { text: "Stable text.", paraId: "00000002" },
    ]);
    const revised = await buildDocxBuffer([
      { text: "Revised text.", paraId: "00000001" },
      { text: "Stable text.", paraId: "00000002", formatting: { bold: true } },
    ]);

    const textDiff = await compareDocxVersions(base, revised, { include: ["text"] });
    expect(textDiff.changes.map(({ type }) => type)).toEqual(["modified"]);
    expect(textDiff.summaryCounts.modified).toBe(1);
    expect(textDiff.summaryCounts.formatChanged).toBe(0);
    expect(textDiff.summaryCounts.unchanged).toBe(1);

    const formattingDiff = await compareDocxVersions(base, revised, {
      include: ["formatting"],
    });
    expect(formattingDiff.changes.map(({ type }) => type)).toEqual(["formatChanged"]);
    expect(formattingDiff.summaryCounts.modified).toBe(0);
    expect(formattingDiff.summaryCounts.formatChanged).toBe(1);
    expect(formattingDiff.summaryCounts.unchanged).toBe(1);
  });

  test("keeps summary cardinality when every text-only event is excluded", async () => {
    const stable = { text: "Stable anchor text.", paraId: "00000001" };
    const cases = [
      {
        name: "insert",
        base: [stable],
        revised: [stable, { text: "Inserted text.", paraId: "00000002" }],
        changeTypes: ["added"],
      },
      {
        name: "delete",
        base: [stable, { text: "Deleted text.", paraId: "00000002" }],
        revised: [stable],
        changeTypes: ["deleted"],
      },
      {
        name: "move",
        base: [
          { text: "Governing law shall be Czech law.", paraId: "00000001" },
          { text: "Payment is due within thirty days.", paraId: "00000002" },
          { text: "Notices must be delivered in writing.", paraId: "00000003" },
        ],
        revised: [
          { text: "Notices must be delivered in writing.", paraId: "00000003" },
          { text: "Governing law shall be Czech law.", paraId: "00000001" },
          { text: "Payment is due within thirty days.", paraId: "00000002" },
        ],
        changeTypes: ["movedTo", "movedFrom"],
      },
      {
        name: "split",
        base: [{ text: "Alpha Beta", paraId: "00000001" }],
        revised: [
          { text: "Alpha", paraId: "00000001" },
          { text: "Beta", paraId: "00000002" },
        ],
        changeTypes: ["modified", "added"],
      },
      {
        name: "merge",
        base: [
          { text: "Alpha", paraId: "00000001" },
          { text: "Beta", paraId: "00000002" },
        ],
        revised: [{ text: "Alpha Beta", paraId: "00000001" }],
        changeTypes: ["modified", "deleted"],
      },
    ] as const;

    for (const fixture of cases) {
      const [base, revised] = await Promise.all([
        buildDocxBuffer(fixture.base),
        buildDocxBuffer(fixture.revised),
      ]);
      const textDiff = await compareDocxVersions(base, revised, { include: ["text"] });
      expect(
        textDiff.changes.map(({ type }) => type),
        fixture.name,
      ).toEqual(fixture.changeTypes);
      const expectedUnits =
        textDiff.summaryCounts.added +
        textDiff.summaryCounts.deleted +
        textDiff.summaryCounts.modified +
        textDiff.summaryCounts.moved +
        textDiff.summaryCounts.unchanged;

      const formattingDiff = await compareDocxVersions(base, revised, {
        include: ["formatting"],
      });
      expect(formattingDiff.changes, fixture.name).toEqual([]);
      expect(formattingDiff.summaryCounts, fixture.name).toEqual({
        added: 0,
        deleted: 0,
        modified: 0,
        formatChanged: 0,
        moved: 0,
        metadataChanged: 0,
        unchanged: expectedUnits,
      });
    }
  });

  test("retains formatting when a modified block's text scope is excluded", async () => {
    const base = await buildDocxBuffer([
      {
        text: "Baseline clause text.",
        paraId: "00000001",
        paragraphAlignment: "left",
      },
    ]);
    const revised = await buildDocxBuffer([
      {
        text: "Revised clause text.",
        paraId: "00000001",
        paragraphAlignment: "right",
      },
    ]);

    const combined = await compareDocxVersions(base, revised);
    expect(combined.changes).toEqual([
      expect.objectContaining({
        type: "modified",
        changedProperties: ["alignment"],
      }),
    ]);

    const formatting = await compareDocxVersions(base, revised, { include: ["formatting"] });
    expect(formatting.changes).toEqual([
      expect.objectContaining({
        type: "formatChanged",
        changedProperties: ["alignment"],
      }),
    ]);
    expect(formatting.summaryCounts).toMatchObject({
      modified: 0,
      formatChanged: 1,
      unchanged: 0,
    });
  });

  test("removes selected metadata values and reports each applied transform", async () => {
    const base = await withCoreProperties(
      await buildDocxBuffer([{ text: "Stable text.", paraId: "00000001" }]),
      {
        title: "Initial title",
        creator: "Initial author",
        lastModifiedBy: "Initial reviewer",
        description: "Initial description",
        revision: 1,
        created: "2026-07-01T10:30:00Z",
        modified: "2026-07-02T10:30:00Z",
      },
    );
    const revised = await withCoreProperties(
      await buildDocxBuffer([{ text: "Stable text.", paraId: "00000001" }]),
      {
        title: "Revised title",
        creator: "Revised author",
        lastModifiedBy: "Revised reviewer",
        description: "Revised description",
        revision: 2,
        created: "2026-07-03T10:30:00Z",
        modified: "2026-07-04T10:30:00Z",
      },
    );

    const diff = await compareDocxVersions(base, revised, {
      include: ["metadata"],
      privacy: {
        transforms: ["remove-descriptive-metadata", "remove-timestamps", "remove-attribution"],
      },
    });

    expect(diff.metadataChanges).toEqual([{ property: "revision", baseValue: 1, revisedValue: 2 }]);
    expect(diff.summaryCounts.metadataChanged).toBe(1);
    expect(diff.privacyReport).toEqual({
      appliedTransforms: ["remove-attribution", "remove-timestamps", "remove-descriptive-metadata"],
      removedMetadataProperties: [
        "title",
        "creator",
        "description",
        "lastModifiedBy",
        "created",
        "modified",
      ],
    });

    expect(
      applyFolioVersionDiffPrivacy(diff, { transforms: ["remove-attribution"] }).privacyReport,
    ).toEqual(diff.privacyReport);
  });

  test("rejects an empty scope selection", async () => {
    const buffer = await buildDocxBuffer([{ text: "Text.", paraId: "00000001" }]);

    await expect(compareDocxVersions(buffer, buffer, { include: [] })).rejects.toBeInstanceOf(
      InvalidFolioVersionComparisonOptionsError,
    );
  });

  test("surfaces content resource limits with the package story position", async () => {
    const base = await buildDocxBuffer([
      {
        text: "x".repeat(FOLIO_CONTENT_COMPARISON_LIMITS.blockCodeUnits + 1),
        paraId: "00000001",
      },
    ]);
    const revised = await buildDocxBuffer([]);

    const comparison = compareDocxVersions(base, revised);
    await expect(comparison).rejects.toMatchObject({
      input: "base",
      limit: "blockCodeUnits",
      maximum: FOLIO_CONTENT_COMPARISON_LIMITS.blockCodeUnits,
      actual: FOLIO_CONTENT_COMPARISON_LIMITS.blockCodeUnits + 1,
      storyIndex: 0,
      blockIndex: 0,
      field: "blocks[0].text",
    });
    await expect(comparison).rejects.toBeInstanceOf(FolioVersionComparisonLimitError);
  });
});

describe("compareDocxVersions: move detection", () => {
  test("a relocated block re-classifies as movedFrom/movedTo sharing a moveGroupId", async () => {
    const base = await buildDocxBuffer([
      { text: "Governing law shall be Czech law.", paraId: "00000001" },
      { text: "Payment is due within thirty days.", paraId: "00000002" },
      { text: "Notices must be delivered in writing.", paraId: "00000003" },
    ]);
    const revised = await buildDocxBuffer([
      { text: "Notices must be delivered in writing.", paraId: "00000003" },
      { text: "Governing law shall be Czech law.", paraId: "00000001" },
      { text: "Payment is due within thirty days.", paraId: "00000002" },
    ]);

    const diff = await compareDocxVersions(base, revised);

    expect(diff.summaryCounts).toEqual({
      added: 0,
      deleted: 0,
      modified: 0,
      formatChanged: 0,
      moved: 1,
      metadataChanged: 0,
      unchanged: 2,
    });
    const movedTo = diff.changes.find((c) => c.type === "movedTo");
    const movedFrom = diff.changes.find((c) => c.type === "movedFrom");
    if (!movedTo || movedTo.type !== "movedTo" || !movedFrom || movedFrom.type !== "movedFrom") {
      throw new Error("expected a movedTo + movedFrom pair");
    }
    expect(movedTo.moveGroupId).toBe(movedFrom.moveGroupId);
    expect(movedTo.text).toBe("Notices must be delivered in writing.");
    expect(movedFrom.text).toBe("Notices must be delivered in writing.");
    expect(movedTo.blockId).toBe("00000003");
    expect(movedFrom.blockId).toBe("00000003");
  });

  test("identical short boilerplate below the word floor stays added + deleted", async () => {
    // "Confidential" is one word — under the move word-count floor — so the
    // deleted instance and the (differently-identified) added instance must
    // NOT pair as a move. Two long stable anchors around it keep the
    // alignment from treating anything else as relocated.
    const base = await buildDocxBuffer([
      { text: "Confidential", paraId: "00000001" },
      { text: "Payment is due within thirty days.", paraId: "00000002" },
      { text: "Notices must be delivered in writing.", paraId: "00000003" },
    ]);
    const revised = await buildDocxBuffer([
      { text: "Payment is due within thirty days.", paraId: "00000002" },
      { text: "Notices must be delivered in writing.", paraId: "00000003" },
      { text: "Confidential", paraId: "00000009" },
    ]);

    const diff = await compareDocxVersions(base, revised);

    expect(diff.summaryCounts).toEqual({
      added: 1,
      deleted: 1,
      modified: 0,
      formatChanged: 0,
      moved: 0,
      metadataChanged: 0,
      unchanged: 2,
    });
    expect(diff.changes.map((c) => c.type).toSorted()).toEqual(["added", "deleted"]);
  });

  test("two simultaneous relocations pair independently, exercising the per-text FIFO queue for two keys", async () => {
    // The neutral move core retains a bounded FIFO queue per exact text.
    // "Notices" and "Confidentiality" relocate ahead of the two monotonic
    // anchors, exercising two live queue keys without cross-contamination.
    const base = await buildDocxBuffer([
      { text: "Governing law shall be Czech law.", paraId: "00000001" },
      { text: "Payment is due within thirty days.", paraId: "00000002" },
      { text: "Notices must be delivered in writing.", paraId: "00000003" },
      { text: "Confidentiality survives termination.", paraId: "00000004" },
    ]);
    const revised = await buildDocxBuffer([
      { text: "Confidentiality survives termination.", paraId: "00000004" },
      { text: "Notices must be delivered in writing.", paraId: "00000003" },
      { text: "Governing law shall be Czech law.", paraId: "00000001" },
      { text: "Payment is due within thirty days.", paraId: "00000002" },
    ]);

    const diff = await compareDocxVersions(base, revised);

    expect(diff.summaryCounts).toEqual({
      added: 0,
      deleted: 0,
      modified: 0,
      formatChanged: 0,
      moved: 2,
      metadataChanged: 0,
      unchanged: 2,
    });
    const movedTexts = diff.changes
      .filter((c) => c.type === "movedFrom")
      .map((c) => c.text)
      .toSorted();
    expect(movedTexts).toEqual(
      ["Confidentiality survives termination.", "Notices must be delivered in writing."].toSorted(),
    );
    // Each movedFrom must share its moveGroupId with the matching movedTo of
    // the SAME text, not the other relocated block's.
    for (const from of diff.changes.filter((c) => c.type === "movedFrom")) {
      const to = diff.changes.find(
        (c) => c.type === "movedTo" && c.moveGroupId === from.moveGroupId,
      );
      if (!to || to.type !== "movedTo" || from.type !== "movedFrom") {
        throw new Error("expected a matching movedTo for every movedFrom");
      }
      expect(to.text).toBe(from.text);
    }
  });

  test("uses the neutral core's edited-move classification", async () => {
    const base = await buildDocxBuffer([
      {
        text: "alpha beta gamma delta epsilon",
        paraId: "00000001",
        paragraphAlignment: "left",
        paragraphStyleId: "Normal",
      },
      { text: "First durable anchor text", paraId: "00000002" },
      { text: "Second durable anchor text", paraId: "00000003" },
      { text: "Third durable anchor text", paraId: "00000004" },
    ]);
    const revised = await buildDocxBuffer([
      { text: "First durable anchor text", paraId: "00000002" },
      { text: "Second durable anchor text", paraId: "00000003" },
      { text: "Third durable anchor text", paraId: "00000004" },
      {
        text: "alpha beta gamma delta zeta",
        paraId: "00000001",
        paragraphAlignment: "right",
        paragraphStyleId: "Heading2",
      },
    ]);
    const [baseReviewer, revisedReviewer] = await Promise.all([
      FolioDocxReviewer.fromBuffer(base),
      FolioDocxReviewer.fromBuffer(revised),
    ]);
    const neutral = compareContent({
      base: { blocks: projectMainContent(baseReviewer).blocks },
      revised: { blocks: projectMainContent(revisedReviewer).blocks },
    });
    if (neutral.isErr()) {
      throw neutral.error;
    }
    const neutralChanges = neutral.value.events
      .filter(({ type }) => type !== "unchanged")
      .map(({ type }) => type);

    expect(neutralChanges).toEqual(["movedFrom", "movedTo"]);
    const versionDiff = await compareDocxVersions(base, revised);
    expect(versionDiff.changes.map(({ type }) => type)).toEqual(neutralChanges);
    const neutralMovedTo = neutral.value.events.find(({ type }) => type === "movedTo");
    const versionMovedTo = versionDiff.changes.find(({ type }) => type === "movedTo");
    expect(neutralMovedTo?.type).toBe("movedTo");
    if (neutralMovedTo?.type !== "movedTo") {
      throw new Error("expected a canonical moved-to event");
    }
    expect(neutralMovedTo.move.relation.segments).toMatchObject([
      { type: "equal", text: "alpha beta gamma delta" },
      { type: "del", text: " epsilon" },
      { type: "ins", text: " zeta" },
    ]);
    expect(neutralMovedTo.move.relation.blockChanges).toEqual([
      {
        field: "blockProperties",
        changes: [
          {
            key: "headingLevel",
            base: { type: "absent" },
            revised: { type: "present", value: 2 },
          },
        ],
      },
    ]);
    expect(
      neutralMovedTo.move.relation.formatting?.paragraph.authored.map(({ key }) => key),
    ).toEqual(["alignment", "styleId"]);
    expect(
      neutralMovedTo.move.relation.formatting?.ranges.flatMap(({ formatting }) =>
        formatting.effective.map(({ key }) => key),
      ),
    ).toEqual(["bold", "fontSize"]);
    expect(versionMovedTo).toMatchObject({
      type: "movedTo",
      segments: [
        { type: "equal", text: "alpha beta gamma delta" },
        { type: "del", text: " epsilon" },
        { type: "ins", text: " zeta" },
      ],
      changedProperties: [
        "headingLevel",
        "alignment",
        "bold",
        "fontSize",
        "lineSpacing",
        "spaceAfter",
        "spaceBefore",
        "styleId",
      ],
    });
    expect(versionDiff.summaryCounts).toMatchObject({
      added: 0,
      deleted: 0,
      moved: 1,
    });

    const formattingDiff = await compareDocxVersions(base, revised, {
      include: ["formatting"],
    });
    expect(formattingDiff.changes).toEqual([
      expect.objectContaining({
        type: "formatChanged",
        blockId: "00000001",
        changedProperties: [
          "alignment",
          "bold",
          "fontSize",
          "lineSpacing",
          "spaceAfter",
          "spaceBefore",
          "styleId",
        ],
      }),
    ]);
    expect(formattingDiff.summaryCounts).toMatchObject({
      modified: 0,
      formatChanged: 1,
      moved: 0,
      unchanged: 3,
    });
  });

  test("preserves positional fallback ids when delegating move classification", async () => {
    const [base, revised] = await Promise.all([
      buildDocxBuffer([
        { text: "alpha beta gamma" },
        { text: "bravo charlie delta" },
        { text: "echo foxtrot golf" },
        { text: "hotel india juliet" },
      ]),
      buildDocxBuffer([
        { text: "bravo charlie delta" },
        { text: "hotel india juliet" },
        { text: "alpha beta gamma" },
        { text: "echo foxtrot golf" },
      ]),
    ]);
    const [baseReviewer, revisedReviewer] = await Promise.all([
      FolioDocxReviewer.fromBuffer(base),
      FolioDocxReviewer.fromBuffer(revised),
    ]);
    const baseBlocks = projectMainContent(baseReviewer).blocks;
    const revisedBlocks = projectMainContent(revisedReviewer).blocks;
    expect(
      [...baseBlocks, ...revisedBlocks].every(({ identity }) => identity.type === "positional"),
    ).toBe(true);
    const neutral = compareContent({
      base: { blocks: baseBlocks },
      revised: { blocks: revisedBlocks },
    });
    if (neutral.isErr()) {
      throw neutral.error;
    }
    const neutralChanges = neutral.value.events
      .filter(({ type }) => type !== "unchanged")
      .map(({ type }) => type);

    expect(neutralChanges).toEqual(["movedFrom", "movedFrom", "movedTo", "movedTo"]);
    const versionDiff = await compareDocxVersions(base, revised);
    expect(versionDiff.changes.map(({ type }) => type)).toEqual(neutralChanges);
    expect(versionDiff.summaryCounts).toMatchObject({
      added: 0,
      deleted: 0,
      moved: 2,
    });
  });
});

describe("compareDocxVersions: neutral split and merge projection", () => {
  test("projects paragraph formatting removals and both split siblings by scope", async () => {
    const [splitBase, splitRevised] = await Promise.all([
      buildDocxBuffer([
        {
          text: "Alpha Beta",
          paraId: "00000001",
          paragraphStyleId: "Heading1",
          paragraphAlignment: "left",
        },
      ]),
      buildDocxBuffer([
        { text: "Alpha", paraId: "00000001" },
        {
          text: "Beta",
          paraId: "00000002",
          paragraphStyleId: "Heading1",
          paragraphAlignment: "right",
        },
      ]),
    ]);

    const splitCombined = await compareDocxVersions(splitBase, splitRevised);
    expect(splitCombined.changes).toEqual([
      expect.objectContaining({
        type: "modified",
        blockId: "00000001",
        changedProperties: [
          "headingLevel",
          "alignment",
          "bold",
          "fontSize",
          "lineSpacing",
          "spaceAfter",
          "spaceBefore",
          "styleId",
        ],
      }),
      expect.objectContaining({ type: "added", blockId: "00000002" }),
    ]);
    const splitFormatting = await compareDocxVersions(splitBase, splitRevised, {
      include: ["formatting"],
    });
    expect(splitFormatting.changes).toEqual([
      expect.objectContaining({
        type: "formatChanged",
        blockId: "00000001",
        changedProperties: [
          "alignment",
          "bold",
          "fontSize",
          "lineSpacing",
          "spaceAfter",
          "spaceBefore",
          "styleId",
        ],
      }),
      expect.objectContaining({
        type: "formatChanged",
        blockId: "00000002",
        changedProperties: ["alignment"],
      }),
    ]);
    expect(splitFormatting.summaryCounts).toMatchObject({
      modified: 0,
      formatChanged: 2,
      unchanged: 0,
    });

    const [mergeBase, mergeRevised] = await Promise.all([
      buildDocxBuffer([
        {
          text: "Alpha",
          paraId: "00000001",
          paragraphStyleId: "Heading1",
          paragraphAlignment: "left",
        },
        { text: "Beta", paraId: "00000002" },
      ]),
      buildDocxBuffer([{ text: "Alpha Beta", paraId: "00000001" }]),
    ]);
    const mergeCombined = await compareDocxVersions(mergeBase, mergeRevised);
    expect(mergeCombined.changes).toEqual([
      expect.objectContaining({
        type: "modified",
        blockId: "00000001",
        changedProperties: [
          "headingLevel",
          "alignment",
          "bold",
          "fontSize",
          "lineSpacing",
          "spaceAfter",
          "spaceBefore",
          "styleId",
        ],
      }),
      expect.objectContaining({ type: "deleted", blockId: "00000002" }),
    ]);
    const mergeFormatting = await compareDocxVersions(mergeBase, mergeRevised, {
      include: ["formatting"],
    });
    expect(mergeFormatting.changes).toEqual([
      expect.objectContaining({
        type: "formatChanged",
        blockId: "00000001",
        changedProperties: [
          "alignment",
          "bold",
          "fontSize",
          "lineSpacing",
          "spaceAfter",
          "spaceBefore",
          "styleId",
        ],
      }),
    ]);
    expect(mergeFormatting.summaryCounts).toMatchObject({
      modified: 0,
      formatChanged: 1,
      unchanged: 1,
    });
  });

  test.each([
    {
      name: "split",
      base: [{ text: "A😀 B", paraId: "00000001" }],
      revised: [
        { text: "A😀", paraId: "00000001" },
        { text: "B", paraId: "00000002" },
      ],
      eventType: "split",
      changeTypes: ["modified", "added"],
    },
    {
      name: "merge",
      base: [
        { text: "A😀", paraId: "00000001" },
        { text: "B", paraId: "00000002" },
      ],
      revised: [{ text: "A😀 B", paraId: "00000001" }],
      eventType: "merge",
      changeTypes: ["modified", "deleted"],
    },
  ] as const)(
    "reuses the neutral $name segments without changing structured output order",
    async (fixture) => {
      const [base, revised] = await Promise.all([
        buildDocxBuffer(fixture.base),
        buildDocxBuffer(fixture.revised),
      ]);
      const [baseReviewer, revisedReviewer] = await Promise.all([
        FolioDocxReviewer.fromBuffer(base),
        FolioDocxReviewer.fromBuffer(revised),
      ]);
      const baseContent = projectMainContent(baseReviewer);
      const revisedContent = projectMainContent(revisedReviewer);
      const neutral = compareContent({
        base: { blocks: baseContent.blocks },
        revised: { blocks: revisedContent.blocks },
      });
      if (neutral.isErr()) {
        throw neutral.error;
      }
      const event = neutral.value.events.find(({ type }) => type === fixture.eventType);
      if (event?.type !== "split" && event?.type !== "merge") {
        throw new Error(`expected a ${fixture.eventType} event`);
      }

      const firstRelation = event.relations[0];
      const baseText = firstRelation.base.block.text.slice(
        firstRelation.base.startOffset,
        firstRelation.base.endOffset,
      );
      const revisedText = firstRelation.revised.block.text.slice(
        firstRelation.revised.startOffset,
        firstRelation.revised.endOffset,
      );
      const diffCalls: string[] = [];
      const controlledSession = createContentComparisonWorkSession({
        diffText: (baseTextInput, revisedTextInput) => {
          diffCalls.push(`${baseTextInput}\u0000${revisedTextInput}`);
          if (baseTextInput === revisedTextInput) {
            return [{ type: "equal", text: baseTextInput }];
          }
          return [
            { type: "del", text: baseTextInput },
            { type: "ins", text: revisedTextInput },
          ];
        },
      });
      const captured = controlledSession.captureComparison({
        base: baseContent.snapshot,
        revised: revisedContent.snapshot,
      });
      if (captured.isErr()) throw captured.error;
      const controlled = captured.value.compare();
      if (controlled.isErr()) throw controlled.error;
      const baseStory = baseReviewer.listStories().at(0)?.handle;
      const revisedStory = revisedReviewer.listStories().at(0)?.handle;
      if (!baseStory || !revisedStory) {
        throw new Error("expected body story handles");
      }
      const controlledEvent = controlled.value.events.find(
        ({ type }) => type === fixture.eventType,
      );
      if (controlledEvent?.type !== "split" && controlledEvent?.type !== "merge") {
        throw new Error(`expected a controlled ${fixture.eventType} event`);
      }
      const controlledFirstRelation = controlledEvent.relations[0];
      const diffCallsAfterComparison = [...diffCalls];
      const controlledProjection = projectFolioContentComparisonToStory({
        baseStory,
        revisedStory,
        comparison: controlled.value,
        firstMoveGroupId: 1,
        includeText: true,
        includeFormatting: true,
      });
      const controlledModified = controlledProjection.changes.at(0);
      expect(controlledModified?.type).toBe("modified");
      if (controlledModified?.type !== "modified") {
        throw new Error("expected controlled canonical segments to project as modified");
      }
      expect(controlledModified.segments.at(0)).toEqual(
        expect.objectContaining({
          type: controlledFirstRelation.segments.at(0)?.type,
          text: controlledFirstRelation.segments.at(0)?.text,
        }),
      );
      expect(diffCalls.filter((call) => call === `${baseText}\u0000${revisedText}`)).toHaveLength(
        1,
      );
      expect(diffCalls).toEqual(diffCallsAfterComparison);

      const version = await compareDocxVersions(base, revised);
      expect(version.changes.map(({ type }) => type)).toEqual(fixture.changeTypes);
      const modified = version.changes.at(0);
      expect(modified?.type).toBe("modified");
      if (modified?.type !== "modified") {
        throw new Error("expected the paired split or merge block to remain modified");
      }
      expect(modified.segments.at(0)).toEqual(
        expect.objectContaining({ type: event.relations[0].segments[0]?.type }),
      );
    },
  );
});

describe("compareDocxVersions: neutral structural classification", () => {
  test("does not pair identical stable-id text across body and table containers", async () => {
    const text = "Payment is due within thirty days.";
    const [base, revised] = await Promise.all([
      buildDocxBuffer([{ text, paraId: "00000001" }]),
      buildTableCellDocxBuffer(text),
    ]);
    const [baseReviewer, revisedReviewer] = await Promise.all([
      FolioDocxReviewer.fromBuffer(base),
      FolioDocxReviewer.fromBuffer(revised),
    ]);
    const neutral = compareContent({
      base: { blocks: projectMainContent(baseReviewer).blocks },
      revised: { blocks: projectMainContent(revisedReviewer).blocks },
    });
    if (neutral.isErr()) {
      throw neutral.error;
    }
    const neutralChanges = neutral.value.events
      .filter(({ type }) => type !== "unchanged")
      .map((event) => {
        if (event.type === "inserted") return "added";
        if (event.type !== "structural") return event.type;
        return event.change.type === "table-insert" ||
          event.change.type === "table-row-insert" ||
          event.change.type === "table-column-insert"
          ? "added"
          : "deleted";
      });

    expect(neutralChanges).toEqual(["added", "deleted"]);
    const versionDiff = await compareDocxVersions(base, revised);
    expect(versionDiff.changes.map(({ type }) => type)).toEqual(neutralChanges);
    expect(versionDiff.summaryCounts).toMatchObject({
      added: 1,
      deleted: 1,
      unchanged: 0,
    });
  });
});

describe("compareDocxVersions: format-only changes", () => {
  test("equal text with different run formatting reports formatChanged with the changed properties", async () => {
    const base = await buildDocxBuffer([{ text: "Payment is due.", paraId: "00000001" }]);
    const revised = await buildDocxBuffer([
      { text: "Payment is due.", paraId: "00000001", formatting: { bold: true, italic: true } },
    ]);

    const diff = await compareDocxVersions(base, revised);

    expect(diff.summaryCounts).toEqual({
      added: 0,
      deleted: 0,
      modified: 0,
      formatChanged: 1,
      moved: 0,
      metadataChanged: 0,
      unchanged: 0,
    });
    const [change] = diff.changes;
    if (!change || change.type !== "formatChanged") {
      throw new Error("expected a formatChanged change");
    }
    expect(change.blockId).toBe("00000001");
    expect(change.changedProperties).toEqual(["bold", "italic"]);
    expect(change.text).toBe("Payment is due.");
  });

  test("identical formatting on both sides stays unchanged", async () => {
    const paragraphs: ParagraphSpec[] = [
      { text: "Payment is due.", paraId: "00000001", formatting: { bold: true } },
    ];
    const base = await buildDocxBuffer(paragraphs);
    const revised = await buildDocxBuffer(paragraphs);

    const diff = await compareDocxVersions(base, revised);

    expect(diff.changes).toEqual([]);
    expect(diff.summaryCounts.unchanged).toBe(1);
  });
});

describe("compareDocxVersions: as-accepted semantics", () => {
  test("a pending tracked insertion in the revised document counts as already applied", async () => {
    const base = await buildDocxBuffer([{ text: "Payment is due." }]);

    // Build the revised buffer by applying a tracked-changes (default mode)
    // replace against a fresh reviewer over `base` — the revised buffer still
    // carries real w:ins/w:del marks on disk.
    const revisedReviewer = await FolioDocxReviewer.fromBuffer(base, { author: "AI" });
    const target = revisedReviewer.snapshot().blocks.at(0);
    if (!target) {
      throw new Error("expected a block in the base document");
    }
    revisedReviewer.applyOperations([
      {
        id: "t1",
        type: "replaceInBlock",
        blockId: target.id,
        find: "due.",
        replace: "due promptly.",
      },
    ]);
    const revised = await revisedReviewer.toBuffer();

    // Sanity: the revised snapshot's clean text is already the accepted view.
    expect(revisedReviewer.snapshot().blocks.at(0)?.text).toBe("Payment is due promptly.");

    const diff = await compareDocxVersions(base, revised);

    expect(diff.summaryCounts).toEqual({
      added: 0,
      deleted: 0,
      modified: 1,
      formatChanged: 0,
      moved: 0,
      metadataChanged: 0,
      unchanged: 0,
    });
    const [change] = diff.changes;
    if (!change || change.type !== "modified") {
      throw new Error("expected a single modified change");
    }
    // The diff runs over the as-accepted text: reconstructing the revised
    // side from `equal` + `ins` segments reproduces the clean accepted
    // string, not raw tracked-change markup.
    const reconstructedRevised = change.segments
      .filter((s) => s.type !== "del")
      .map((s) => s.text)
      .join("");
    expect(reconstructedRevised).toBe("Payment is due promptly.");
    expect(change.segments.some((s) => s.type === "ins" && s.text.includes("promptly"))).toBe(true);
  });
});
