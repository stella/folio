import { expect, test } from "bun:test";
import type { Node as PMNode } from "prosemirror-model";

import { FolioDocxReviewer, getFolioDocxComparisonAccess } from "../ai-edits/headless";
import { sourceDocumentOf } from "../ai-edits/snapshot";
import { createDocx } from "../docx/rezip";
import type { HeaderFooter, Paragraph } from "../types/document";
import { createEmptyDocument } from "../utils/createDocument";
import { compareDocxVersions } from "../version-comparison";
import { applyComparison, parseComparison, planComparison } from "./compare";

const OPTIONS = { author: "compare", timestamp: "2026-09-11T00:00:00.000Z" } as const;
const HEADER_RELATIONSHIP_ID = "rId_projection_header";
const FOOTER_RELATIONSHIP_ID = "rId_projection_footer";

const paragraph = (text: string, paraId?: string): Paragraph => ({
  type: "paragraph",
  ...(paraId !== undefined && { paraId }),
  content: [{ type: "run", content: [{ type: "text", text }] }],
});

const headerFooter = (type: "header" | "footer", text: string, paraId?: string): HeaderFooter => ({
  type,
  hdrFtrType: "default",
  content: [paragraph(text, paraId)],
});

type StoryIdentityFixture = "authored" | "synthesized";

const storyMatrixDocx = (
  label: string,
  identity: StoryIdentityFixture = "authored",
): Promise<ArrayBuffer> => {
  const paraId = (authored: string): string | undefined =>
    identity === "authored" ? authored : undefined;
  const document = createEmptyDocument();
  document.package.document.content = [paragraph(`main ${label}`, paraId("A1000001"))];
  document.package.headers = new Map([
    [HEADER_RELATIONSHIP_ID, headerFooter("header", `header ${label}`, paraId("A1000002"))],
  ]);
  document.package.footers = new Map([
    [FOOTER_RELATIONSHIP_ID, headerFooter("footer", `footer ${label}`, paraId("A1000003"))],
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
      content: [paragraph(`footnote ${label}`, paraId("A1000004"))],
    },
  ];
  document.package.endnotes = [
    {
      type: "endnote",
      id: 3,
      noteType: "normal",
      content: [paragraph(`endnote ${label}`, paraId("A1000005"))],
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

test("comparison projection performs one live package projection plus the requested census", async () => {
  const baseReviewer = await FolioDocxReviewer.fromBuffer(
    await storyMatrixDocx("before", "synthesized"),
  );
  const baseTraversals = instrumentStoryTraversals(baseReviewer);
  const baseProjection = getFolioDocxComparisonAccess(baseReviewer).normalizeSourceStories();

  expect(baseProjection.stories).toHaveLength(5);
  expect(baseProjection.stories.every(({ snapshot }) => snapshot !== null)).toBe(true);
  expect(baseProjection.revisions).toEqual({ highestId: 0, present: false });
  expect(baseTraversals.map((read) => read())).toEqual([5, 2, 2, 2, 2]);

  const targetReviewer = await FolioDocxReviewer.fromBuffer(
    await storyMatrixDocx("after", "synthesized"),
  );
  const targetTraversals = instrumentStoryTraversals(targetReviewer);
  const targetProjection = getFolioDocxComparisonAccess(targetReviewer).projectResolvedStories();

  expect(targetProjection.stories).toHaveLength(5);
  expect(targetProjection.stories.every(({ snapshot }) => snapshot !== null)).toBe(true);
  expect(targetTraversals.map((read) => read())).toEqual([4, 1, 1, 1, 1]);
});

test("no-op version comparison retains synthesized identity in every story", async () => {
  const source = await storyMatrixDocx("same", "synthesized");
  const diff = await compareDocxVersions(source, source.slice(0));

  expect(diff.changes).toEqual([]);
  expect(diff.stories).toHaveLength(5);
  expect(diff.stories.every(({ changes }) => changes.length === 0)).toBe(true);
  expect(diff.summaryCounts).toEqual({
    added: 0,
    deleted: 0,
    modified: 0,
    formatChanged: 0,
    moved: 0,
    metadataChanged: 0,
    unchanged: 5,
  });
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
  const applied = applyComparison(parsed.value, planned.value, {
    mode: "strict",
  });
  if (applied.isErr()) {
    throw applied.error;
  }

  expect(applied.value.verification).toEqual({ status: "verified" });
});
