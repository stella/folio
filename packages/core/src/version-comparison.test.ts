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

import { FolioDocxReviewer } from "./ai-edits/headless";
import { parseDocx } from "./docx/parser";
import { createDocx } from "./docx/rezip";
import { repackDocx } from "./docx/rezip";
import type { HeaderFooter, Paragraph } from "./types/document";
import { createEmptyDocument } from "./utils/createDocument";
import { compareDocxVersions, exceedsLcsBudget } from "./version-comparison";
import type { FolioBlockDiff } from "./version-comparison";

type ParagraphSpec = {
  text: string;
  paraId?: string;
  formatting?: { bold?: boolean; italic?: boolean };
};

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
  const source = await buildDocxBuffer([{ text: bodyText, paraId: "A1000001" }]);
  const document = await parseDocx(source, { detectVariables: false, preloadFonts: false });
  document.package.headers = new Map([
    ["rIdHeader", headerFooterStory("header", headerText, "B1000001")],
  ]);
  document.package.document.finalSectionProperties = {
    ...document.package.document.finalSectionProperties,
    headerReferences: [{ type: "default", rId: "rIdHeader" }],
  };
  if (footerText !== undefined) {
    document.package.footers = new Map([
      ["rIdFooter", headerFooterStory("footer", footerText, "D1000001")],
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
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:footnote w:id="2"><w:p w14:paraId="C1000001"><w:r><w:t>${footnoteText}</w:t></w:r></w:p></w:footnote></w:footnotes>`,
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
      { type: "equal", text: "Beta " },
      { type: "del", text: "paragraph." },
      { type: "ins", text: "clause." },
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
      blockId: "B1000001",
    });
    expect(headerChange.revisedHandle).toEqual({
      story: { type: "header", relationshipId: "rIdHeader" },
      blockId: "B1000001",
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
      blockId: "D1000001",
    });
  });
});

describe("compareDocxVersions: deterministic fallback ids (no w14:paraId)", () => {
  test("a same-text block whose ordinal shifted still pairs as unchanged via the text-LCS pass", async () => {
    // Neither buffer carries a w14:paraId, so FolioDocxReviewer assigns each
    // block a deterministic fallback id derived from hash(text + ordinal). The
    // "Epsilon" paragraph inserted before "Gamma" shifts Gamma's ordinal (3rd
    // -> 4th), which changes its fallback id even though its text is
    // untouched. Pass 1 (stable-id pairing) therefore CANNOT pair Gamma across
    // versions; only the text-LCS pass (pass 2) recovers it as unchanged. This
    // is the regression guard for the deterministic-id pitfall.
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

    const diff = await compareDocxVersions(base, revised);

    // Alpha (unshifted, same fallback id both sides) and Gamma (shifted,
    // recovered via text-LCS) are both unchanged and absent from `changes`.
    expect(diff.summaryCounts).toEqual({
      added: 1,
      deleted: 0,
      modified: 1,
      formatChanged: 0,
      moved: 0,
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
      unchanged: 2,
    });
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
      unchanged: 2,
    });
    expect(diff.changes.map((c) => c.type).toSorted()).toEqual(["added", "deleted"]);
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

describe("exceedsLcsBudget: pass 2's LCS cell-budget guard", () => {
  test("flags unpaired-block counts whose product would exceed the LCS cell budget (4,000,000)", () => {
    // A degenerate/adversarial document with no w14:paraIds can leave
    // thousands of blocks unpaired on both sides; this guard is what stops
    // pairByExactText from allocating an O(m*n) table for it. Exercise the
    // predicate directly rather than constructing a multi-million-block
    // fixture, which would make this test slow for no extra coverage.
    expect(exceedsLcsBudget(2000, 2000)).toBe(false); // exactly at budget: 4,000,000 cells
    expect(exceedsLcsBudget(2001, 2001)).toBe(true); // just over budget: 4,004,001 cells
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
