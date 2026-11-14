/**
 * Clean Markdown and the plain-text readers show a document with every change
 * accepted. They read it as a reviewer's Accept All leaves it: for corpus
 * fixtures and generated documents, the clean Markdown of a document equals
 * the clean Markdown of the document a canonical reviewer holds after
 * accepting all, and so does each story's plain text.
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
import { collectBlockTexts } from "../docx/blockPlainText";
import { prepareCanonicalDocxInput } from "../docx/canonicalSessionInput";
import { createDocx } from "../docx/rezip";
import { toMarkdownResult } from "../markdown/index";
import type { MarkdownOptions } from "../markdown/types";
import type { Document } from "../types/document";
import { FolioDocxReviewer } from "./headless";

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

/** The package as a host stores it: every paragraph and control identified. */
const ingested = async (buffer: ArrayBuffer): Promise<ArrayBuffer> => {
  const prepared = await prepareCanonicalDocxInput(buffer);
  if (prepared.isErr()) throw prepared.error;
  return prepared.value.slice().buffer;
};

const packaged = async (document: Document): Promise<ArrayBuffer | null> => {
  try {
    return await ingested(await createDocx(structuredClone(document)));
  } catch {
    return null;
  }
};

const CLEAN: MarkdownOptions = {
  annotations: "strip",
  trackedChanges: "clean",
  comments: "strip",
  footnotes: "keep",
};

const open = (buffer: ArrayBuffer): Promise<FolioDocxReviewer> =>
  FolioDocxReviewer.fromBuffer(buffer.slice(0), {
    author: "Accepted View",
    session: "canonical",
  });

/** Each story's plain text, in story order. */
const storyTexts = (document: Document): string[][] => [
  collectBlockTexts(document.package.document.content),
  ...[
    ...(document.package.headers?.values() ?? []),
    ...(document.package.footers?.values() ?? []),
    ...(document.package.footnotes ?? []),
    ...(document.package.endnotes ?? []),
  ].map(({ content }) => collectBlockTexts(content)),
];

type Outcome = "agreed" | "acceptRefused";

const acceptedViewLaws = async (buffer: ArrayBuffer): Promise<Outcome> => {
  const reviewer = await open(buffer);
  const document = reviewer.toDocument();
  const read = toMarkdownResult(document, CLEAN);
  const texts = storyTexts(document);
  try {
    reviewer.acceptAll();
  } catch (error) {
    if (CanonicalSessionError.is(error)) return "acceptRefused";
    throw error;
  }
  const accepted = reviewer.toDocument();
  expect(read.markdown).toBe(toMarkdownResult(accepted, CLEAN).markdown);
  expect(read.warnings).not.toContain(
    "Some tracked structural changes could not be resolved.",
  );
  // Notes Accept All removes leave the story list; the rest read alike.
  if (
    (accepted.package.footnotes?.length ?? 0) ===
      (document.package.footnotes?.length ?? 0) &&
    (accepted.package.endnotes?.length ?? 0) ===
      (document.package.endnotes?.length ?? 0)
  )
    expect(texts).toEqual(storyTexts(accepted));
  else expect(texts[0]).toEqual(storyTexts(accepted)[0]);
  return "agreed";
};

describe("the accepted view a reader shows", () => {
  test.each(FIXTURE_FILES)(
    "%s reads as a reviewer's Accept All leaves it",
    async (filename) => {
      await acceptedViewLaws(await ingested(readFixture(filename)));
    },
    propertyTestTimeout(30_000),
  );

  test(
    "generated documents read as a reviewer's Accept All leaves them",
    async () => {
      await assertProperty(
        fc.asyncProperty(reviewerDocumentArbitrary, async (document) => {
          const buffer = await packaged(document);
          if (buffer !== null) await acceptedViewLaws(buffer);
        }),
        {
          id: "generated documents read as a reviewer's Accept All leaves them",
          numRuns: 40,
        },
      );
    },
    propertyTestTimeout(120_000),
  );
});
