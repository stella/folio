/**
 * Property + example tests for the headless `.docx` reviewer over the full
 * fixture corpus. The example-based `headless.test.ts` pins specific edits on
 * one fixture; this file exercises the reviewer's core invariants across every
 * corpus `.docx`.
 *
 * Split by input shape:
 *
 *   - Pure per-fixture fixed points iterate the corpus EXHAUSTIVELY (`test.each`)
 *     so every fixture is checked on every run rather than sampled:
 *       - Deterministic block ids: `fromBuffer(d)` twice on the same bytes yields
 *         identical block-id lists (the "snapshot on one parse, apply on another"
 *         flow depends on this).
 *       - No-op fixed point: `fromBuffer(d).toBuffer()` with no ops re-parses to
 *         the same visible block text.
 *   - Edits with a randomly generated replacement string stay fast-check
 *     properties, so the nightly budget (PROPERTY_TEST_NUM_RUNS_FACTOR) explores
 *     many fixture × block × replacement combinations:
 *       - Accept ≡ direct: a tracked-changes edit then `acceptAll()` yields the
 *         same visible text as the same edit applied in `direct` mode.
 *       - Reject restores: a tracked-changes edit then `rejectAll()` restores the
 *         original visible text.
 *   - Selective ≡ full repack calls `attemptSelectiveSave` directly and asserts a
 *     non-null return (null = it bailed to a full repack) before comparing it
 *     against a forced full repack, so the equivalence can never be vacuously
 *     satisfied by selective never actually running.
 *
 * Fixture provenance and licensing: see
 * `../docx/__tests__/__fixtures__/corpus/PROVENANCE.md`.
 */

import { describe, expect, test } from "bun:test";
import fc from "fast-check";
import { EditorState } from "prosemirror-state";
import { readFileSync, readdirSync } from "node:fs";
import path from "node:path";

import { propertyConfig, propertyTestTimeout } from "../../../../test/property-testing";

import { parseDocx } from "../docx/parser";
import { repackDocx } from "../docx/rezip";
import { attemptSelectiveSave } from "../docx/selectiveSave";
import { updateDocumentContent } from "../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import { ensureParaIdsInState } from "../prosemirror/extensions/features/ParaIdAllocatorExtension";
import { schema, singletonManager } from "../prosemirror/schema";
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

/**
 * Corpus fixtures ship without `w14:paraId`s; the editor allocates them on load
 * and its first (full-repack) save writes them out, so only later saves take the
 * selective path. Reproduce that first save so the reviewer under test sees a
 * paraId-bearing baseline the selective patcher can key against. Mirrors
 * `headless.test.ts`.
 */
const makeParaIdBaseline = async (source: ArrayBuffer): Promise<ArrayBuffer> => {
  const document = await parseDocx(source, { detectVariables: false, preloadFonts: false });
  const state = ensureParaIdsInState(
    EditorState.create({
      schema,
      doc: toProseDoc(document),
      plugins: singletonManager.getPlugins(),
    }),
  );
  return repackDocx({ ...updateDocumentContent(document, state.doc), originalBuffer: source });
};

describe("headless reviewer invariants (full corpus)", () => {
  // An empty / uncloned corpus must be a loud failure, not a silent pass: the
  // `test.each` blocks below generate zero cases when the list is empty.
  test("corpus fixtures are present", () => {
    expect(FIXTURE_FILES.length).toBeGreaterThan(0);
  });

  test.each(FIXTURE_FILES)(
    "fromBuffer is deterministic: same bytes yield the same block-id list (%s)",
    async (filename) => {
      const buffer = readFixture(filename);
      const idsOf = async (): Promise<string[]> =>
        (await FolioDocxReviewer.fromBuffer(buffer, { author: "AI" }))
          .snapshot()
          .blocks.map((b) => b.id);
      expect(await idsOf()).toEqual(await idsOf());
    },
  );

  test.each(FIXTURE_FILES)(
    "toBuffer with no operations is a visible-text fixed point (%s)",
    async (filename) => {
      const reviewer = await FolioDocxReviewer.fromBuffer(readFixture(filename), { author: "AI" });
      const before = blockTexts(reviewer);
      const reparsed = await FolioDocxReviewer.fromBuffer(await reviewer.toBuffer(), {
        author: "AI",
      });
      expect(blockTexts(reparsed)).toEqual(before);
    },
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
    "selective save engages (attemptSelectiveSave returns non-null) and matches a full repack",
    async () => {
      let engaged = 0;
      for (const filename of FIXTURE_FILES) {
        const baseline = await makeParaIdBaseline(readFixture(filename));
        const reviewer = await FolioDocxReviewer.fromBuffer(baseline, { author: "AI" });
        const targets = editTargets(reviewer.snapshot().blocks);
        if (targets.length === 0) {
          continue;
        }
        const target = targets[0]!;
        const applied = reviewer.applyOperations([replaceOp(target, "ZZWORD")], { mode: "direct" });
        if (applied.applied.length === 0) {
          continue;
        }
        const editedDocument = reviewer.toDocument();

        // Directly observe engagement instead of inferring it from bytes: call
        // the selective-save entry point with the same inputs the reviewer's own
        // toBuffer() derives for this edit (one non-structural in-block change to
        // a paraId-bearing block). A non-null return IS selective save engaging;
        // null means it bailed to a full repack. Inferring engagement from
        // byte-identity would false-positive because the baseline is itself a
        // repack, so a full-repack fallback can reproduce the same bytes.
        const selective = await attemptSelectiveSave(editedDocument, baseline, {
          changedParaIds: new Set([target.blockId]),
          structuralChange: false,
          hasUntrackedChanges: false,
        });
        expect(selective).not.toBeNull();
        if (selective === null) {
          continue;
        }
        engaged += 1;

        // Selective and a forced full repack of the same edited model must
        // re-parse to the same visible text.
        const fullRepack = await repackDocx({ ...editedDocument, originalBuffer: baseline });
        const fromSelective = await FolioDocxReviewer.fromBuffer(selective, { author: "AI" });
        const fromFull = await FolioDocxReviewer.fromBuffer(fullRepack, { author: "AI" });
        expect(blockTexts(fromSelective)).toEqual(blockTexts(fromFull));
      }
      // Loud failure if the selective path never engaged across the corpus: the
      // equivalence above would then be vacuous.
      expect(engaged).toBeGreaterThan(0);
    },
    propertyTestTimeout(25_000),
  );

  test(
    "insertAfterBlock never leaves a block whose text contains a line break",
    async () => {
      // A model routinely sends a whole clause — heading plus body — as one
      // `text` with embedded newlines. Word has no paragraph-with-a-line-break
      // primitive, so the applier must split it into several paragraphs
      // instead; this checks that invariant across the full fixture corpus
      // and random line shapes rather than pinning one example.
      await fc.assert(
        fc.asyncProperty(
          fc.constantFrom(...FIXTURE_FILES),
          fc.nat(),
          fc.array(fc.stringMatching(/^[A-Za-z0-9 ]{0,20}$/u), { minLength: 1, maxLength: 4 }),
          fc.constantFrom("\n", "\r\n"),
          async (filename, selector, lines, lineBreak) => {
            const text = lines.join(lineBreak);
            if (text.trim().length === 0) {
              return;
            }
            const buffer = readFixture(filename);
            const probe = await FolioDocxReviewer.fromBuffer(buffer, { author: "AI" });
            const targets = editTargets(probe.snapshot().blocks);
            if (targets.length === 0) {
              return;
            }
            const target = targets[selector % targets.length]!;
            const op: FolioAIEditOperation = {
              id: "1",
              type: "insertAfterBlock",
              blockId: target.blockId,
              text,
            };

            const reviewer = await FolioDocxReviewer.fromBuffer(buffer, { author: "AI" });
            const applied = reviewer.applyOperations([op], { mode: "direct" });
            if (applied.applied.length === 0) {
              return;
            }

            for (const blockText of blockTexts(reviewer)) {
              expect(blockText).not.toContain("\n");
              expect(blockText).not.toContain("\r");
            }
          },
        ),
        propertyConfig({ numRuns: 30 }),
      );
    },
    propertyTestTimeout(20_000),
  );

  test(
    "revision ids stay unique across a multi-paragraph tracked insert followed by another apply",
    async () => {
      // The shared revision-id range is reserved once per apply call, sized
      // from the operations in that call (see `estimateRevisionIdReservation`
      // in apply.ts). A tracked-changes insert whose `text` splits into more
      // paragraphs than the reservation anticipates must not spill into the
      // range the NEXT apply call reserves — that would stamp two different
      // marks with the same revision id.
      await fc.assert(
        fc.asyncProperty(
          fc.constantFrom(...FIXTURE_FILES),
          fc.nat(),
          fc.nat(),
          // Non-blank after trim (starts with an alnum) so every requested
          // line survives the applier's blank-line collapse, guaranteeing a
          // paragraph count that exceeds the fixed 4-id-per-operation
          // cushion `estimateRevisionIdReservation` used to fall back to.
          fc.array(fc.stringMatching(/^[A-Za-z0-9][A-Za-z0-9 ]{0,19}$/u), {
            minLength: 5,
            maxLength: 9,
          }),
          async (filename, selectorA, selectorB, lines) => {
            const buffer = readFixture(filename);
            const probe = await FolioDocxReviewer.fromBuffer(buffer, { author: "AI" });
            const targets = editTargets(probe.snapshot().blocks);
            if (targets.length < 2) {
              return;
            }
            const targetA = targets[selectorA % targets.length]!;
            const targetB = targets[selectorB % targets.length]!;

            const reviewer = await FolioDocxReviewer.fromBuffer(buffer, { author: "AI" });
            const firstApply = reviewer.applyOperations([
              {
                id: "insert-1",
                type: "insertAfterBlock",
                blockId: targetA.blockId,
                text: lines.join("\n"),
              },
            ]);
            const secondApply = reviewer.applyOperations([replaceOp(targetB, "ZZWORD")]);

            const revisionIds = [
              ...firstApply.applied.flatMap((op) => op.revisionIds ?? []),
              ...secondApply.applied.flatMap((op) => op.revisionIds ?? []),
            ];
            if (revisionIds.length === 0) {
              return;
            }
            expect(new Set(revisionIds).size).toBe(revisionIds.length);
          },
        ),
        propertyConfig({ numRuns: 30 }),
      );
    },
    propertyTestTimeout(20_000),
  );
});
