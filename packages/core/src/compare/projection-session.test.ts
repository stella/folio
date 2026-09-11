import { expect, test } from "bun:test";
import type { Node as PMNode } from "prosemirror-model";

import { FolioDocxReviewer, getFolioDocxComparisonAccess } from "../ai-edits/headless";
import { sourceDocumentOf } from "../ai-edits/snapshot";
import { createDocx } from "../docx/rezip";
import type { HeaderFooter, Paragraph } from "../types/document";
import { createEmptyDocument } from "../utils/createDocument";
import { applyComparison, parseComparison, planComparison } from "./compare";

const OPTIONS = { author: "compare", timestamp: "2026-09-11T00:00:00.000Z" } as const;
const HEADER_RELATIONSHIP_ID = "rId_projection_header";
const FOOTER_RELATIONSHIP_ID = "rId_projection_footer";

const paragraph = (text: string, paraId: string): Paragraph => ({
  type: "paragraph",
  paraId,
  content: [{ type: "run", content: [{ type: "text", text }] }],
});

const headerFooter = (type: "header" | "footer", text: string, paraId: string): HeaderFooter => ({
  type,
  hdrFtrType: "default",
  content: [paragraph(text, paraId)],
});

const storyMatrixDocx = (label: string): Promise<ArrayBuffer> => {
  const document = createEmptyDocument();
  document.package.document.content = [paragraph(`main ${label}`, "A1000001")];
  document.package.headers = new Map([
    [HEADER_RELATIONSHIP_ID, headerFooter("header", `header ${label}`, "A1000002")],
  ]);
  document.package.footers = new Map([
    [FOOTER_RELATIONSHIP_ID, headerFooter("footer", `footer ${label}`, "A1000003")],
  ]);
  document.package.document.finalSectionProperties = {
    ...document.package.document.finalSectionProperties,
    headerReferences: [{ type: "default", rId: HEADER_RELATIONSHIP_ID }],
    footerReferences: [{ type: "default", rId: FOOTER_RELATIONSHIP_ID }],
  };
  document.package.footnotes = [
    {
      type: "footnote",
      id: 2,
      noteType: "normal",
      content: [paragraph(`footnote ${label}`, "A1000004")],
    },
  ];
  document.package.endnotes = [
    {
      type: "endnote",
      id: 3,
      noteType: "normal",
      content: [paragraph(`endnote ${label}`, "A1000005")],
    },
  ];
  return createDocx(document);
};

const instrumentStoryTraversals = (reviewer: FolioDocxReviewer): (() => number)[] => {
  const counters: (() => number)[] = [];
  for (const { handle } of reviewer.listStories()) {
    const reviewed = reviewer.readReviewedStory({ story: handle, view: "current-markup" });
    if (!reviewed) {
      throw new Error(`missing ${handle.type} story`);
    }
    const doc = sourceDocumentOf(reviewed.snapshot);
    const descendants = doc.descendants.bind(doc);
    let traversals = 0;
    Object.defineProperty(doc, "descendants", {
      configurable: true,
      value: (callback: Parameters<PMNode["descendants"]>[0]) => {
        traversals += 1;
        return descendants(callback);
      },
    });
    counters.push(() => traversals);
  }
  return counters;
};

test("comparison projection has one snapshot walk plus only the requested revision census", async () => {
  const baseReviewer = await FolioDocxReviewer.fromBuffer(await storyMatrixDocx("before"));
  const baseTraversals = instrumentStoryTraversals(baseReviewer);
  const baseProjection =
    getFolioDocxComparisonAccess(baseReviewer).projectStories("with-revision-census");

  expect(baseProjection.stories).toHaveLength(5);
  expect(baseProjection.stories.every(({ snapshot }) => snapshot !== null)).toBe(true);
  expect(baseProjection.revisions).toEqual({ highestId: 0, present: false });
  expect(baseTraversals.map((read) => read())).toEqual([2, 2, 2, 2, 2]);

  const targetReviewer = await FolioDocxReviewer.fromBuffer(await storyMatrixDocx("after"));
  const targetTraversals = instrumentStoryTraversals(targetReviewer);
  const targetProjection =
    getFolioDocxComparisonAccess(targetReviewer).projectStories("without-revision-census");

  expect(targetProjection.stories).toHaveLength(5);
  expect(targetProjection.stories.every(({ snapshot }) => snapshot !== null)).toBe(true);
  expect(targetProjection.revisions).toEqual({ highestId: 0, present: false });
  expect(targetTraversals.map((read) => read())).toEqual([1, 1, 1, 1, 1]);
});

test("comparison consumes the retained story projections through apply verification", async () => {
  const parsed = await parseComparison(
    await storyMatrixDocx("before"),
    await storyMatrixDocx("after"),
    OPTIONS,
  );
  if (parsed.isErr()) {
    throw parsed.error;
  }
  expect(parsed.value.pairs).toHaveLength(5);

  const planned = planComparison(parsed.value);
  if (planned.isErr()) {
    throw planned.error;
  }
  expect(planned.value).toHaveLength(5);
  const applied = applyComparison(parsed.value, planned.value);
  if (applied.isErr()) {
    throw applied.error;
  }

  expect(applied.value.verification).toEqual({ status: "verified" });
});
