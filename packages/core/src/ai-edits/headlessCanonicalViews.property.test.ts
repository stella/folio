/**
 * A canonical headless reviewer's resolved views read as the legacy
 * reviewer's.
 *
 * For every story of a document, the original and final views (text,
 * blocks and the changes left) and the story's tables in each view are
 * read on both engines. A canonical reviewer reads them from a second
 * session resolved to the view, and leaves its own document as it was.
 * A story resolved to its final view keeps that view and holds no change.
 */

import { describe, expect, test } from "bun:test";
import fc from "fast-check";
import { readdirSync, readFileSync } from "node:fs";
import path from "node:path";

import { reviewerDocumentArbitrary } from "../../../../test/generators/reviewerDocumentArbitraries";
import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import { CanonicalSessionError } from "../canonical/canonicalSession";
import { prepareCanonicalDocxInput } from "../docx/canonicalSessionInput";
import { createDocx } from "../docx/rezip";
import type { HeadlessReviewerSession } from "../types/canonicalCapabilities";
import type { Document } from "../types/document";
import { FOLIO_RESOLVED_REVIEWED_VIEWS, FolioDocxReviewer } from "./headless";

const FIXTURES_DIR = path.join(
  import.meta.dir,
  "../docx/__tests__/__fixtures__/corpus",
);

const FIXTURE_FILES: string[] = readdirSync(FIXTURES_DIR)
  .filter((name) => name.endsWith(".docx"))
  .sort();

const readFixture = (filename: string): ArrayBuffer => {
  const bytes = readFileSync(path.join(FIXTURES_DIR, filename));
  return bytes.buffer.slice(
    bytes.byteOffset,
    bytes.byteOffset + bytes.byteLength,
  );
};

const ingested = async (buffer: ArrayBuffer): Promise<ArrayBuffer> => {
  const prepared = await prepareCanonicalDocxInput(buffer);
  if (prepared.isErr()) throw prepared.error;
  return prepared.value.slice().buffer;
};

const open = (buffer: ArrayBuffer, session: HeadlessReviewerSession) =>
  FolioDocxReviewer.fromBuffer(buffer.slice(0), { session });

/** Every story in both resolved views. */
const views = (reviewer: FolioDocxReviewer) =>
  reviewer.listStories().flatMap(({ handle }) =>
    FOLIO_RESOLVED_REVIEWED_VIEWS.map((view) => {
      const read = reviewer.readReviewedStory({ story: handle, view });
      return {
        handle,
        view,
        text: read?.text ?? null,
        changes: read?.changes ?? null,
        blocks: read?.snapshot.blocks ?? null,
        tables: reviewer
          .storyTables({ story: handle, view })
          .map((table) => table.node.textContent),
      };
    }),
  );

type ViewOutcome = "read" | "sessionRefused" | "resolutionRefused";

const viewLaws = async (buffer: ArrayBuffer): Promise<ViewOutcome> => {
  let canonical: FolioDocxReviewer;
  try {
    canonical = await open(buffer, "canonical");
  } catch (error) {
    if (CanonicalSessionError.is(error)) return "sessionRefused";
    throw error;
  }
  const before = canonical.toDocument();
  let read: ReturnType<typeof views>;
  try {
    read = views(canonical);
  } catch (error) {
    if (CanonicalSessionError.is(error)) return "resolutionRefused";
    throw error;
  }
  expect(read).toEqual(views(await open(buffer, "legacy")));
  // Reading a view resolves nothing in the reviewer itself.
  expect(canonical.toDocument()).toEqual(before);

  // Resolving one story reads, afterwards, as that story's view did.
  for (const { handle } of canonical.listStories()) {
    const reviewer = await open(buffer, "canonical");
    const final = reviewer.readReviewedStory({ story: handle, view: "final" });
    expect(
      reviewer.resolveReviewedStory({ story: handle, view: "final" }),
    ).toBe(true);
    // Its final view is unchanged, and no change of it is left pending.
    const resolved = reviewer.readReviewedStory({
      story: handle,
      view: "final",
    });
    expect(resolved?.text).toBe(final?.text);
    expect(
      reviewer.readReviewedStory({ story: handle, view: "current-markup" })
        ?.changes,
    ).toEqual([]);
  }
  return "read";
};

const packaged = async (document: Document): Promise<ArrayBuffer | null> => {
  try {
    return await ingested(await createDocx(structuredClone(document)));
  } catch {
    return null;
  }
};

describe("a canonical reviewer's resolved views", () => {
  test.each(FIXTURE_FILES)(
    "%s reads its resolved views as the legacy reviewer does",
    async (filename) => {
      await viewLaws(await ingested(readFixture(filename)));
    },
    propertyTestTimeout(30_000),
  );

  test(
    "generated documents read their resolved views as the legacy reviewer does",
    async () => {
      await assertProperty(
        fc.asyncProperty(reviewerDocumentArbitrary, async (document) => {
          const buffer = await packaged(document);
          if (buffer !== null) await viewLaws(buffer);
        }),
        {
          id: "generated documents read their resolved views as the legacy reviewer does",
          numRuns: 25,
        },
      );
    },
    propertyTestTimeout(120_000),
  );
});
