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

import { FolioDocxReviewer } from "./ai-edits/headless";
import { createDocx } from "./docx/rezip";
import { generateRedlineDocx } from "./redline";
import { createEmptyDocument } from "./utils/createDocument";

type ParagraphSpec = { text: string; paraId?: string };

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

  test("an empty base document reports unanchorable additions as skipped", async () => {
    const base = await buildDocxBuffer([]);
    const revised = await buildDocxBuffer([{ text: "Only paragraph.", paraId: "00000001" }]);

    const result = await generateRedlineDocx(base, revised);

    expect(result.applied).toEqual([]);
    expect(result.skipped.map((op) => op.reason)).toEqual(["missingBlock"]);
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
});
