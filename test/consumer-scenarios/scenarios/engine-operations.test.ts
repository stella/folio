/**
 * Every operation type over every synthetic fixture, on a reviewer opened on
 * the canonical session, through the packed public entry points:
 *
 * - each batch applies or refuses with an issue;
 * - an applied batch undoes to every read it changed;
 * - rejecting every change after a tracked batch gives what rejecting every
 *   change gives without it, and accepting every change gives the words of
 *   the same batch applied directly.
 *
 * Each operation is also applied on the legacy reviewer, and every
 * operation the legacy reviewer applies but the canonical one refuses is
 * counted by its typed refusal and printed, so what is left of the cutover
 * is measured on the same fixtures.
 */

import assert from "node:assert/strict";
import { after, describe, test } from "node:test";

import {
  FolioDocxReviewer,
  type FolioDocxReviewerOptions,
} from "@stll/folio-core/server";

import {
  AUTHOR,
  COLLISION_FIXTURES,
  FIXTURES,
  STORY_FIXTURES,
  toArrayBuffer,
} from "../support/documents.ts";
import {
  type Block,
  coreBatch,
  GENERATORS,
  type Operation,
  supports,
} from "../support/operations.ts";
import { createRandom } from "../support/random.ts";
import { settledText } from "../support/review.ts";

const ALL_FIXTURES = {
  ...FIXTURES,
  ...COLLISION_FIXTURES,
  ...STORY_FIXTURES,
} as const satisfies Record<string, () => Promise<Uint8Array>>;

const ENGINE_MODES = ["direct", "tracked-changes"] as const;
type EngineMode = (typeof ENGINE_MODES)[number];

type Reviewer = Awaited<ReturnType<typeof FolioDocxReviewer.fromBuffer>>;
type Engine = NonNullable<FolioDocxReviewerOptions["session"]>;

const openOn = (bytes: Uint8Array, session: Engine): Promise<Reviewer> =>
  FolioDocxReviewer.fromBuffer(toArrayBuffer(bytes), {
    author: AUTHOR,
    session,
  });

const apply = (reviewer: Reviewer, operation: Operation, mode: EngineMode) =>
  reviewer.applyDocumentOperations(
    coreBatch([{ ...operation, id: "op" }], mode) as never,
    { undefinedReferences: "refuse" },
  );

/** What the canonical reviewer refused that the legacy one applied, by cause. */
const refusals = new Map<string, number>();
const countRefusal = (cause: string) =>
  refusals.set(cause, (refusals.get(cause) ?? 0) + 1);

/** Every read an undo must give back. */
const reads = (reviewer: Reviewer) => ({
  content: reviewer.getContent(),
  text: reviewer.getContentAsText({ annotated: true }),
  notes: reviewer.getNotesAsText(),
  stories: reviewer.listStories(),
  changes: reviewer.getChanges(),
  comments: reviewer.getComments(),
});

/** Whitespace and block boundaries left out, as a reader shows a pending join. */
const words = (blocks: readonly string[]): string =>
  blocks
    .map((block) => block.replace(/^\w+: /u, ""))
    .join("")
    .replace(/\s+/gu, "");

/** Whether `error` is the canonical session's typed refusal. */
const isSessionRefusal = (error: unknown): boolean =>
  typeof error === "object" &&
  error !== null &&
  Reflect.get(error, "_tag") === "CanonicalSessionError";

/**
 * The body once every change is resolved, or `null` when a lock held one or
 * the session refused a resolution it cannot run yet (counted).
 */
const resolved = (
  reviewer: Reviewer,
  resolution: "accept" | "reject",
): string[] | null => {
  try {
    const { skipped } =
      resolution === "accept" ? reviewer.acceptAll() : reviewer.rejectAll();
    return skipped.length > 0 ? null : settledText(reviewer);
  } catch (error) {
    if (!isSessionRefusal(error)) throw error;
    countRefusal(
      `${resolution}All: ${error instanceof Error ? error.message : String(error)}`,
    );
    return null;
  }
};

after(() => {
  const summary = Object.fromEntries(
    [...refusals].sort(([left], [right]) => left.localeCompare(right)),
  );
  console.log(`canonical reviewer refusals: ${JSON.stringify(summary)}`);
});

describe("operations on the canonical reviewer", () => {
  for (const [name, build] of Object.entries(ALL_FIXTURES)) {
    for (const mode of ENGINE_MODES) {
      test(`${name} / ${mode}: every operation type applies or refuses, undoes and resolves`, async () => {
        const bytes = await build();
        const random = createRandom(name.length * 31 + mode.length);
        const log: string[] = [];
        for (const type of Object.keys(GENERATORS)) {
          if (!supports(type, mode)) continue;
          const canonical = await openOn(bytes, "canonical");
          const blocks: Block[] = canonical.getContent();
          const operation = GENERATORS[type]?.(blocks, random);
          if (!operation) continue;
          const context = `${name} / ${mode}: ${JSON.stringify(operation)}`;

          const before = reads(canonical);
          const result = apply(canonical, operation, mode);
          assert.equal(
            result.applied.length + result.skipped.length,
            1,
            `${context}: neither applied nor refused`,
          );
          const legacy = apply(await openOn(bytes, "legacy"), operation, mode);
          if (result.applied.length === 0) {
            assert.ok(result.issues.length > 0, `${context}: no issue`);
            const skip = result.skipped[0];
            if (legacy.applied.length > 0)
              countRefusal(
                `${type} / ${mode}: ${skip?.canonicalRefusal?.gap ?? skip?.reason ?? "unknown"}`,
              );
            log.push(`${type}: refused`);
            continue;
          }
          log.push(`${type}: applied`);

          // The session's journal gives every read back.
          assert.ok(result.undoHandle, `${context}: no undo handle`);
          assert.equal(
            canonical.undoDocumentOperations(result.undoHandle).status,
            "undone",
            `${context}: undo refused`,
          );
          assert.deepEqual(reads(canonical), before, `${context}: undo`);
          if (mode !== "tracked-changes") continue;

          const rejected = await openOn(bytes, "canonical");
          apply(rejected, operation, mode);
          const rejectedText = resolved(rejected, "reject");
          const untouched = resolved(
            await openOn(bytes, "canonical"),
            "reject",
          );
          if (rejectedText !== null && untouched !== null)
            assert.deepEqual(
              rejectedText,
              untouched,
              `${context}: rejecting every change does not give the document back`,
            );

          const accepted = await openOn(bytes, "canonical");
          apply(accepted, operation, mode);
          const direct = await openOn(bytes, "canonical");
          if (apply(direct, operation, "direct").applied.length === 0) continue;
          const acceptedText = resolved(accepted, "accept");
          const directText = resolved(direct, "accept");
          if (acceptedText !== null && directText !== null)
            assert.equal(
              words(acceptedText),
              words(directText),
              `${context}: accepting every change does not give the direct edit`,
            );
        }
        assert.ok(
          log.some((entry) => entry.endsWith("applied")),
          `${name} / ${mode}: nothing applied: ${log.join("; ")}`,
        );
      });
    }
  }
});
