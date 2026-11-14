/**
 * A canonical headless reviewer's save gives back what it reads.
 *
 * Each case draws a document a reviewer lists every story of (or a body with
 * review marks of every shape) and one public operation built from the
 * blocks a reader sees, applied directly or tracked. The reviewer saves
 * through the editor's canonical save, and a canonical reviewer reopened on
 * the saved package reads exactly what the saving reviewer read. A second
 * save of the reopened package is a fixed point of those reads, and
 * resolving every change after the save gives what resolving it before the
 * save gives.
 */

import { describe, expect, test } from "bun:test";
import fc from "fast-check";
import { readdirSync, readFileSync } from "node:fs";
import path from "node:path";

import { reviewerDocumentArbitrary } from "../../../../test/generators/reviewerDocumentArbitraries";
import {
  type OperationDraw,
  operationDrawArbitrary,
  operationFor,
} from "../../../../test/generators/reviewerOperationDraws";
import { checkCoverageSample } from "../../../../test/property-coverage";
import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import { CanonicalSessionError } from "../canonical/canonicalSession";
import { FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION } from "../document-operations";
import { prepareCanonicalDocxInput } from "../docx/canonicalSessionInput";
import { createDocx } from "../docx/rezip";
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

type Reviewer = FolioDocxReviewer;
type Mode = "direct" | "tracked-changes";

const open = (buffer: ArrayBuffer): Promise<Reviewer> =>
  FolioDocxReviewer.fromBuffer(buffer.slice(0), {
    author: "Canonical Reviewer",
    session: "canonical",
  });

/** Every read a save must give back. */
const reads = (reviewer: Reviewer) => ({
  content: reviewer.getContent(),
  text: reviewer.getContentAsText({ annotated: true }),
  notes: reviewer.getNotesAsText(),
  stories: reviewer.listStories(),
  readStory: reviewer
    .listStories()
    .map(({ handle }) => reviewer.readStory(handle)),
  changes: reviewer.getChanges(),
  comments: reviewer.getComments(),
  numbering: reviewer.readNumberingDefinitions(),
  properties: reviewer.getDocumentProperties(),
});

/** The body once every change is resolved, or `null` when that refuses. */
const resolvedBody = (
  reviewer: Reviewer,
  resolution: "accept" | "reject",
): string | null => {
  try {
    const { skipped } =
      resolution === "accept" ? reviewer.acceptAll() : reviewer.rejectAll();
    return skipped.length > 0 ? null : reviewer.getContentAsText();
  } catch (error) {
    if (CanonicalSessionError.is(error)) return null;
    throw error;
  }
};

type SaveOutcome =
  | "saved"
  | "notAPackage"
  | "sessionRefused"
  | "noOperation"
  | "refusedOperation";

const saveLaws = async (
  buffer: ArrayBuffer,
  draw: OperationDraw,
  mode: Mode,
): Promise<SaveOutcome> => {
  let reviewer: Reviewer;
  try {
    reviewer = await open(buffer);
  } catch (error) {
    if (CanonicalSessionError.is(error)) return "sessionRefused";
    throw error;
  }
  const operation = operationFor(reviewer, draw);
  if (operation === null) return "noOperation";
  const result = reviewer.applyDocumentOperations(
    {
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      operations: [operation],
      mode,
    },
    { undefinedReferences: "refuse" },
  );
  if (result.applied.length === 0) return "refusedOperation";

  const before = reads(reviewer);
  const saved = await reviewer.toBuffer();
  const reopened = await open(saved);
  expect(reads(reopened)).toEqual(before);
  // Saving the reopened package writes the same reads again.
  expect(reads(await open(await reopened.toBuffer()))).toEqual(before);

  // Resolving after the save gives what resolving before it gives.
  for (const resolution of ["accept", "reject"] as const) {
    const live = await open(buffer);
    live.applyDocumentOperations(
      {
        version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
        operations: [operation],
        mode,
      },
      { undefinedReferences: "refuse" },
    );
    const liveBody = resolvedBody(live, resolution);
    const savedBody = resolvedBody(await open(saved), resolution);
    if (liveBody !== null && savedBody !== null)
      expect(savedBody).toBe(liveBody);
  }
  return "saved";
};

/** The generated model as a package, or nothing when the writer refuses it. */
const packaged = async (document: Document): Promise<ArrayBuffer | null> => {
  try {
    return await ingested(await createDocx(structuredClone(document)));
  } catch {
    return null;
  }
};

const generatedSaves = (tally: Map<SaveOutcome, number>) =>
  fc.asyncProperty(
    reviewerDocumentArbitrary,
    operationDrawArbitrary,
    fc.constantFrom<Mode>("direct", "tracked-changes"),
    async (document, draw, mode) => {
      const buffer = await packaged(document);
      const outcome =
        buffer === null ? "notAPackage" : await saveLaws(buffer, draw, mode);
      tally.set(outcome, (tally.get(outcome) ?? 0) + 1);
    },
  );

const GENERATED_RUNS = 30;
const COVERAGE_RUNS = 30;

describe("a canonical reviewer's save gives back what it reads", () => {
  test("corpus fixtures are present", () => {
    expect(FIXTURE_FILES.length).toBeGreaterThan(0);
  });

  test.each(FIXTURE_FILES)(
    "%s: a tracked replacement saves and reopens to the same reads",
    async (filename) => {
      const outcome = await saveLaws(
        await ingested(readFixture(filename)),
        { kind: "replaceInBlock", block: 0, word: 0, text: "revised" },
        "tracked-changes",
      );
      expect(outcome).not.toBe("notAPackage");
    },
    propertyTestTimeout(30_000),
  );

  test(
    "generated documents save and reopen to the same reads",
    async () => {
      await assertProperty(generatedSaves(new Map()), {
        id: "generated documents save and reopen to the same reads",
        numRuns: GENERATED_RUNS,
      });
    },
    propertyTestTimeout(120_000),
  );

  test(
    "most generated cases reach the save",
    async () => {
      const tally = new Map<SaveOutcome, number>();
      await checkCoverageSample(generatedSaves(tally), COVERAGE_RUNS);
      expect(tally.get("saved") ?? 0).toBeGreaterThanOrEqual(COVERAGE_RUNS / 3);
    },
    propertyTestTimeout(120_000),
  );
});
