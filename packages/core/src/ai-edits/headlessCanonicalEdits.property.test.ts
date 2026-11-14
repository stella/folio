/**
 * A canonical headless reviewer edits and resolves through its session.
 *
 * Each case draws a body with review marks of every shape (or, less often, a
 * document a reviewer lists every story of) and one operation to stage,
 * built from the blocks a reader sees, then holds the reviewer's edits to
 * three laws:
 *
 * - rejecting every change after the operation in tracked mode gives what
 *   rejecting every change gives without it;
 * - accepting every change after the operation in tracked mode gives the
 *   words of the operation applied directly, with every change accepted;
 * - undoing the operation gives back every read and the document exactly.
 *
 * A package the session cannot hold, or an operation either mode refuses,
 * is tallied, and most cases must reach the laws.
 */

import { describe, expect, test } from "bun:test";
import fc from "fast-check";
import { readdirSync, readFileSync } from "node:fs";
import path from "node:path";

import { bodyReviewerDocumentArbitrary } from "../../../../test/generators/reviewerDocumentArbitraries";
import {
  type OperationDraw,
  operationFor,
  stagingDrawArbitrary,
} from "../../../../test/generators/reviewerOperationDraws";
import { checkCoverageSample } from "../../../../test/property-coverage";
import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import { CanonicalSessionError } from "../canonical/canonicalSession";
import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperation,
  type FolioDocumentOperationBatch,
} from "../document-operations";
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

const AUTHOR = "Canonical Reviewer";

type Reviewer = FolioDocxReviewer;

const open = (buffer: ArrayBuffer): Promise<Reviewer> =>
  FolioDocxReviewer.fromBuffer(buffer.slice(0), {
    author: AUTHOR,
    session: "canonical",
  });

const batchOf = (
  operation: FolioDocumentOperation,
  mode: "direct" | "tracked-changes",
): FolioDocumentOperationBatch => ({
  version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  operations: [operation],
  mode,
});

const apply = (
  reviewer: Reviewer,
  operation: FolioDocumentOperation,
  mode: "direct" | "tracked-changes",
) =>
  reviewer.applyDocumentOperations(batchOf(operation, mode), {
    undefinedReferences: "refuse",
  });

/** The body as a reader sees it, block by block. */
const bodyText = (reviewer: Reviewer): string[] =>
  reviewer
    .getContent()
    .filter((block) => block.kind !== "diagnostic" && block.text !== "")
    .map((block) => `${block.kind}: ${block.text}`);

/**
 * The words of a body: a pending join or split shows as the blocks it has
 * now, and a merge's separator belongs to neither block until it is
 * accepted, so blocks and whitespace are left out.
 */
const words = (text: readonly string[]): string =>
  text
    .map((block) => block.replace(/^\w+: /u, ""))
    .join("")
    .replace(/\s+/gu, "");

/** Every read an undo must give back, and the document itself. */
const reads = (reviewer: Reviewer) => ({
  content: reviewer.getContent(),
  text: reviewer.getContentAsText({ annotated: true }),
  notes: reviewer.getNotesAsText(),
  stories: reviewer.listStories(),
  changes: reviewer.getChanges(),
  comments: reviewer.getComments(),
  document: reviewer.toDocument(),
});

type EditOutcome =
  | "edited"
  | "notAPackage"
  | "sessionRefused"
  | "noOperation"
  | "refusedOperation"
  | "resolutionRefused";

/** Why a case stopped short of the laws, counted beside the outcomes. */
type Refusals = Map<string, number>;

const noteRefusal = (refusals: Refusals | undefined, detail: string) =>
  refusals?.set(detail, (refusals.get(detail) ?? 0) + 1);

/** The resolution of every change, or the session's refusal of it. */
const resolvedText = (
  reviewer: Reviewer,
  resolution: "accept" | "reject",
  refusals?: Refusals,
): string[] | null => {
  try {
    const { skipped } =
      resolution === "accept" ? reviewer.acceptAll() : reviewer.rejectAll();
    // A change a lock holds stays pending; the laws are about the rest.
    if (skipped.length > 0) noteRefusal(refusals, `${resolution}: held`);
    return skipped.length > 0 ? null : bodyText(reviewer);
  } catch (error) {
    if (CanonicalSessionError.is(error)) {
      noteRefusal(
        refusals,
        `${resolution}: ${JSON.stringify(error.ground)}: ${error.message}`,
      );
      return null;
    }
    throw error;
  }
};

const editLaws = async (
  buffer: ArrayBuffer,
  draw: OperationDraw,
  refusals?: Refusals,
): Promise<EditOutcome> => {
  let base: Reviewer;
  try {
    base = await open(buffer);
  } catch (error) {
    if (CanonicalSessionError.is(error)) {
      noteRefusal(
        refusals,
        `open: ${JSON.stringify(error.ground)}: ${error.message}`,
      );
      return "sessionRefused";
    }
    throw error;
  }
  const operation = operationFor(base, draw, {
    replaceIn: "blocksWithWords",
  });
  if (operation === null) return "noOperation";

  // Undo gives back every read and the document.
  const before = reads(base);
  const edited = apply(base, operation, "direct");
  if (edited.applied.length === 0 || edited.undoHandle === null) {
    const [skip] = edited.skipped;
    noteRefusal(
      refusals,
      `${operation.type} direct: ${skip?.reason ?? "no undo"} ${skip?.canonicalRefusal?.gap ?? ""}`,
    );
    return "refusedOperation";
  }
  expect(base.undoDocumentOperations(edited.undoHandle)).toMatchObject({
    status: "undone",
  });
  expect(reads(base)).toEqual(before);

  const tracked = await open(buffer);
  const trackedResult = apply(tracked, operation, "tracked-changes");
  if (trackedResult.applied.length === 0) {
    const [skip] = trackedResult.skipped;
    noteRefusal(
      refusals,
      `${operation.type} tracked: ${skip?.reason ?? ""} ${skip?.canonicalRefusal?.gap ?? ""}`,
    );
    return "refusedOperation";
  }
  const untouched = await open(buffer);
  const rejectedWithout = resolvedText(untouched, "reject", refusals);
  const rejected = resolvedText(tracked, "reject", refusals);
  if (rejected === null || rejectedWithout === null) return "resolutionRefused";
  expect(rejected).toEqual(rejectedWithout);

  const accepted = await open(buffer);
  apply(accepted, operation, "tracked-changes");
  const direct = await open(buffer);
  apply(direct, operation, "direct");
  const acceptedText = resolvedText(accepted, "accept", refusals);
  const directText = resolvedText(direct, "accept", refusals);
  if (acceptedText === null || directText === null) return "resolutionRefused";
  expect(words(acceptedText)).toBe(words(directText));
  return "edited";
};

/** The generated model as a package, or nothing when the writer refuses it. */
const packaged = async (
  document: Document,
  refusals?: Refusals,
): Promise<ArrayBuffer | null> => {
  try {
    return await ingested(await createDocx(structuredClone(document)));
  } catch (error) {
    noteRefusal(
      refusals,
      `package: ${error instanceof Error ? error.message.slice(0, 120) : String(error)}`,
    );
    return null;
  }
};

const generatedEdits = (tally: Map<EditOutcome, number>, refusals?: Refusals) =>
  fc.asyncProperty(
    bodyReviewerDocumentArbitrary,
    stagingDrawArbitrary,
    async (document, draw) => {
      const buffer = await packaged(document, refusals);
      const outcome =
        buffer === null
          ? "notAPackage"
          : await editLaws(buffer, draw, refusals);
      tally.set(outcome, (tally.get(outcome) ?? 0) + 1);
    },
  );

const GENERATED_RUNS = 30;
const COVERAGE_RUNS = 30;

describe("a canonical reviewer edits and resolves through its session", () => {
  test("corpus fixtures are present", () => {
    expect(FIXTURE_FILES.length).toBeGreaterThan(0);
  });

  test.each(FIXTURE_FILES)(
    "%s: a replacement undoes, rejects and accepts as the laws say",
    async (filename) => {
      const buffer = await ingested(readFixture(filename));
      const outcome = await editLaws(buffer, {
        kind: "replaceInBlock",
        block: 0,
        word: 0,
        text: "revised",
      });
      expect(outcome).not.toBe("notAPackage");
    },
    propertyTestTimeout(30_000),
  );

  test(
    "generated documents undo, reject and accept as the laws say",
    async () => {
      await assertProperty(generatedEdits(new Map()), {
        id: "generated documents undo, reject and accept as the laws say",
        numRuns: GENERATED_RUNS,
      });
    },
    propertyTestTimeout(120_000),
  );

  test(
    "most generated cases reach the laws",
    async () => {
      const tally = new Map<EditOutcome, number>();
      const refusals: Refusals = new Map();
      await checkCoverageSample(generatedEdits(tally, refusals), COVERAGE_RUNS);
      // The tally and the refusals behind it are in the failure, so a
      // shortfall names its causes.
      expect({
        reached: (tally.get("edited") ?? 0) >= COVERAGE_RUNS / 3,
        tally: Object.fromEntries(tally),
        refusals: Object.fromEntries(refusals),
      }).toMatchObject({ reached: true });
    },
    propertyTestTimeout(120_000),
  );
});
