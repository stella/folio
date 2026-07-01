/**
 * Property tests for the headless `.docx` reviewer over the full fixture corpus.
 *
 * The example-based `headless.test.ts` pins specific edits on one fixture. These
 * properties fuzz the reviewer's core invariants across every corpus `.docx`
 * with a randomly generated replacement string, so the nightly budget
 * (PROPERTY_TEST_NUM_RUNS_FACTOR) explores many fixture × block × replacement
 * combinations:
 *
 *   - Deterministic block ids: `fromBuffer(d)` twice on the same bytes yields
 *     identical block-id lists (the "snapshot on one parse, apply on another"
 *     flow depends on this).
 *   - No-op fixed point: `fromBuffer(d).toBuffer()` with no ops re-parses to the
 *     same visible block text.
 *   - Accept ≡ direct: a tracked-changes edit followed by `acceptAll()` yields
 *     the same visible text as the same edit applied in `direct` mode.
 *   - Reject restores: a tracked-changes edit followed by `rejectAll()` restores
 *     the original visible text.
 *   - Selective ≡ full repack: the selective-save buffer and a forced full
 *     repack of the same edited model re-parse to the same visible text.
 *
 * Fixture provenance and licensing: see
 * `../docx/__tests__/__fixtures__/corpus/PROVENANCE.md`.
 */

import { describe, expect, test } from "bun:test";
import fc from "fast-check";
import { readFileSync, readdirSync } from "node:fs";
import path from "node:path";

import { propertyConfig, propertyTestTimeout } from "../../../../test/property-testing";

import { repackDocx } from "../docx/rezip";
import { FolioDocxReviewer } from "./headless";
import type { FolioAIBlock, FolioAIEditOperation } from "./types";

const FIXTURES_DIR = path.join(import.meta.dir, "../docx/__tests__/__fixtures__/corpus");

const FIXTURE_FILES: string[] = readdirSync(FIXTURES_DIR)
  .filter((name) => name.endsWith(".docx"))
  .sort();

const readFixture = (filename: string): ArrayBuffer => {
  const bytes = readFileSync(path.join(FIXTURES_DIR, filename));
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
};

// Plain replacement text: letters and digits only, so the redline diff and the
// find/replace stay unambiguous and no XML-special character enters the run.
const replaceArb = fc.stringMatching(/^[A-Za-z0-9]{1,12}$/u);

const blockTexts = (reviewer: FolioDocxReviewer): string[] =>
  reviewer.getContent().map((block) => block.text);

type EditTarget = { blockId: string; find: string };

/**
 * Every body block that carries a 3+ letter word, paired with that word as a
 * `find` anchor. Block ids are deterministic across parses, so a target picked
 * from one `fromBuffer` resolves against a fresh `fromBuffer` of the same bytes.
 */
const editTargets = (blocks: FolioAIBlock[]): EditTarget[] => {
  const targets: EditTarget[] = [];
  for (const block of blocks) {
    const word = /[A-Za-z]{3,}/u.exec(block.text);
    if (word) {
      targets.push({ blockId: block.id, find: word[0] });
    }
  }
  return targets;
};

const replaceOp = (target: EditTarget, replace: string): FolioAIEditOperation => ({
  id: "1",
  type: "replaceInBlock",
  blockId: target.blockId,
  find: target.find,
  replace,
});

describe("headless reviewer property invariants (full corpus)", () => {
  test(
    "fromBuffer is deterministic: same bytes yield the same block-id list",
    async () => {
      await fc.assert(
        fc.asyncProperty(fc.constantFrom(...FIXTURE_FILES), async (filename) => {
          const buffer = readFixture(filename);
          const first = await FolioDocxReviewer.fromBuffer(buffer, { author: "AI" });
          const second = await FolioDocxReviewer.fromBuffer(buffer, { author: "AI" });
          const idsOf = (r: FolioDocxReviewer) => r.snapshot().blocks.map((b) => b.id);
          expect(idsOf(second)).toEqual(idsOf(first));
        }),
        propertyConfig({ numRuns: 60 }),
      );
    },
    propertyTestTimeout(15_000),
  );

  test(
    "toBuffer with no operations is a visible-text fixed point",
    async () => {
      await fc.assert(
        fc.asyncProperty(fc.constantFrom(...FIXTURE_FILES), async (filename) => {
          const reviewer = await FolioDocxReviewer.fromBuffer(readFixture(filename), {
            author: "AI",
          });
          const before = blockTexts(reviewer);
          const reparsed = await FolioDocxReviewer.fromBuffer(await reviewer.toBuffer(), {
            author: "AI",
          });
          expect(blockTexts(reparsed)).toEqual(before);
        }),
        propertyConfig({ numRuns: 40 }),
      );
    },
    propertyTestTimeout(15_000),
  );

  test(
    "tracked-changes edit then acceptAll equals the same edit applied directly",
    async () => {
      await fc.assert(
        fc.asyncProperty(
          fc.constantFrom(...FIXTURE_FILES),
          fc.nat(),
          replaceArb,
          async (filename, selector, replace) => {
            const buffer = readFixture(filename);
            const probe = await FolioDocxReviewer.fromBuffer(buffer, { author: "AI" });
            const targets = editTargets(probe.snapshot().blocks);
            if (targets.length === 0) {
              return;
            }
            const op = replaceOp(targets[selector % targets.length]!, replace);

            const tracked = await FolioDocxReviewer.fromBuffer(buffer, { author: "AI" });
            const applied = tracked.applyOperations([op]);
            if (applied.applied.length === 0) {
              return;
            }
            tracked.acceptAll();

            const direct = await FolioDocxReviewer.fromBuffer(buffer, { author: "AI" });
            direct.applyOperations([op], { mode: "direct" });

            expect(blockTexts(tracked)).toEqual(blockTexts(direct));
          },
        ),
        propertyConfig({ numRuns: 40 }),
      );
    },
    propertyTestTimeout(20_000),
  );

  test(
    "tracked-changes edit then rejectAll restores the original visible text",
    async () => {
      await fc.assert(
        fc.asyncProperty(
          fc.constantFrom(...FIXTURE_FILES),
          fc.nat(),
          replaceArb,
          async (filename, selector, replace) => {
            const buffer = readFixture(filename);
            const probe = await FolioDocxReviewer.fromBuffer(buffer, { author: "AI" });
            const targets = editTargets(probe.snapshot().blocks);
            if (targets.length === 0) {
              return;
            }
            const original = blockTexts(probe);
            const op = replaceOp(targets[selector % targets.length]!, replace);

            const tracked = await FolioDocxReviewer.fromBuffer(buffer, { author: "AI" });
            const applied = tracked.applyOperations([op]);
            if (applied.applied.length === 0) {
              return;
            }
            tracked.rejectAll();

            expect(blockTexts(tracked)).toEqual(original);
          },
        ),
        propertyConfig({ numRuns: 40 }),
      );
    },
    propertyTestTimeout(20_000),
  );

  test(
    "selective save and a forced full repack re-parse to the same visible text",
    async () => {
      await fc.assert(
        fc.asyncProperty(
          fc.constantFrom(...FIXTURE_FILES),
          fc.nat(),
          replaceArb,
          async (filename, selector, replace) => {
            const buffer = readFixture(filename);
            const reviewer = await FolioDocxReviewer.fromBuffer(buffer, { author: "AI" });
            const targets = editTargets(reviewer.snapshot().blocks);
            if (targets.length === 0) {
              return;
            }
            const applied = reviewer.applyOperations(
              [replaceOp(targets[selector % targets.length]!, replace)],
              { mode: "direct" },
            );
            if (applied.applied.length === 0) {
              return;
            }

            const selectiveBuffer = await reviewer.toBuffer();
            const fullRepackBuffer = await repackDocx({
              ...reviewer.toDocument(),
              originalBuffer: buffer,
            });

            const fromSelective = await FolioDocxReviewer.fromBuffer(selectiveBuffer, {
              author: "AI",
            });
            const fromFull = await FolioDocxReviewer.fromBuffer(fullRepackBuffer, { author: "AI" });
            expect(blockTexts(fromSelective)).toEqual(blockTexts(fromFull));
          },
        ),
        propertyConfig({ numRuns: 40 }),
      );
    },
    propertyTestTimeout(25_000),
  );
});
