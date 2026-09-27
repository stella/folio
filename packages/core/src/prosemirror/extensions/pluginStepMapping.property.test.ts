/**
 * The plugins that carry positions through a transaction answer what they
 * answered when they walked the mapping once per paragraph and once per step.
 *
 * Each reference below is the per-step implementation the plugin used before
 * it swept its positions through the transaction in one pass (see
 * `positionSweep.ts`), kept verbatim apart from naming. Random documents and
 * random multi-step transactions — typing, deleting, splitting, joining,
 * pasting, wrapping, lifting, marking and re-marking paragraphs, one to three
 * transactions per batch — must produce identical results, down to the order
 * of the sets the selective save reads.
 */

import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import type { Node as PMNode } from "prosemirror-model";
import { EditorState, type Plugin, type Transaction } from "prosemirror-state";
import {
  AddMarkStep,
  AddNodeMarkStep,
  AttrStep,
  RemoveMarkStep,
  RemoveNodeMarkStep,
  type Mapping,
} from "prosemirror-transform";

import { propertyConfig, propertyTestTimeout } from "../../../../../test/property-testing";

import { getTransactionDirtyRange } from "../../paged-layout/transactionDirtyRange";
import {
  buildRandomBatch,
  buildRandomTransaction,
  randomDocument,
  randomOperations,
  randomSchema,
} from "../__tests__/randomTransactions";
import { AutoBidiDetectionExtension } from "./features/AutoBidiDetectionExtension";
import {
  ParagraphChangeTrackerExtension,
  paragraphChangeTrackerKey,
} from "./features/ParagraphChangeTrackerExtension";
import { mapParagraphKeepers, ParaIdAllocatorExtension } from "./features/ParaIdAllocatorExtension";
import { insertedRanges, RunIdentityExtension } from "./marks/RunIdentityExtension";

setDefaultTimeout(propertyTestTimeout(60_000));

const pluginsOf = (runtime: { plugins?: Plugin[] }): Plugin[] => runtime.plugins ?? [];
const trackerPlugins = pluginsOf(
  ParagraphChangeTrackerExtension().onSchemaReady({ schema: randomSchema }),
);
const editorPlugins = [
  ...trackerPlugins,
  ...pluginsOf(ParaIdAllocatorExtension().onSchemaReady({ schema: randomSchema })),
  ...pluginsOf(AutoBidiDetectionExtension().onSchemaReady({ schema: randomSchema })),
  ...pluginsOf(RunIdentityExtension().onSchemaReady({ schema: randomSchema })),
];

const randomBatches = fc.array(randomOperations(20), { minLength: 1, maxLength: 3 });

const isUsableParaId = (value: unknown): value is string =>
  typeof value === "string" && value.length > 0 && value !== "00000000";

// ---------------------------------------------------------------------------
// Paragraph change tracker: the per-step reference.
// ---------------------------------------------------------------------------

type TrackerFields = {
  positions: number[];
  ids: string[];
  hasUntracked: boolean;
  hasUntrackedSource: boolean;
};

type Affected = { ids: Set<string>; positions: Set<number>; hasUntracked: boolean };

const referenceCollect = (doc: PMNode, from: number, to: number): Affected => {
  const ids = new Set<string>();
  const positions = new Set<number>();
  let hasUntracked = false;
  doc.nodesBetween(from, to, (node, pos) => {
    if (node.type.name === "paragraph") {
      positions.add(pos);
      const paraId = node.attrs["paraId"];
      if (isUsableParaId(paraId)) {
        ids.add(paraId);
      } else {
        hasUntracked = true;
      }
    }
  });
  return { ids, positions, hasUntracked };
};

const referenceCollectMarkLike = (doc: PMNode, from: number, to: number): Affected => {
  const lo = Math.min(from, to);
  const hi = Math.max(from, to);
  const end = hi > lo ? hi : lo + 1;
  const primary = referenceCollect(doc, lo, end);
  if (primary.ids.size > 0 || primary.hasUntracked) {
    return primary;
  }
  const $p = doc.resolve(lo);
  for (let depth = $p.depth; depth > 0; depth--) {
    const node = $p.node(depth);
    if (node.type.name === "paragraph") {
      const paraId = node.attrs["paraId"];
      return {
        ids: new Set(isUsableParaId(paraId) ? [paraId] : []),
        positions: new Set([$p.before(depth)]),
        hasUntracked: !isUsableParaId(paraId),
      };
    }
  }
  return { ids: new Set(), positions: new Set(), hasUntracked: false };
};

const referenceMapParagraphPosition = (tr: Transaction, position: number): number | null => {
  const mapped = tr.mapping.mapResult(position + 1, 1);
  if (mapped.deletedAcross) {
    return null;
  }
  const $position = tr.doc.resolve(mapped.pos);
  for (let depth = $position.depth; depth > 0; depth--) {
    if ($position.node(depth).type.name === "paragraph") {
      return $position.before(depth);
    }
  }
  return null;
};

const referenceMapAffected = (tr: Transaction, positions: readonly number[]): Set<number> => {
  const mapped = new Set<number>();
  for (const position of positions) {
    const next = referenceMapParagraphPosition(tr, position);
    if (next !== null) {
      mapped.add(next);
    }
  }
  return mapped;
};

const mapThrough = (remap: Mapping, pos: number, assoc: 1 | -1): number =>
  // oxlint-disable-next-line unicorn/no-array-method-this-argument -- ProseMirror Mapping.map uses assoc as its second parameter.
  remap.map(pos, assoc);

const referenceTrackerApply = (
  previous: TrackerFields,
  tr: Transaction,
  structureChanged: boolean,
): TrackerFields => {
  if (!tr.docChanged) {
    return previous;
  }
  if (tr.getMeta(paragraphChangeTrackerKey) === "ignore") {
    const positions = referenceMapAffected(tr, previous.positions);
    const ids = new Set<string>();
    let unresolved = previous.hasUntrackedSource;
    for (const position of positions) {
      const paraId = tr.doc.nodeAt(position)?.attrs["paraId"];
      if (isUsableParaId(paraId)) {
        ids.add(paraId);
      } else {
        unresolved = true;
      }
    }
    return { ...previous, positions: [...positions], ids: [...ids], hasUntracked: unresolved };
  }

  const positions = referenceMapAffected(tr, previous.positions);
  const ids = new Set(previous.ids);
  let hasUntracked = previous.hasUntracked;
  let hasUntrackedSource = previous.hasUntrackedSource;
  const record = (affected: Affected): void => {
    for (const position of affected.positions) {
      positions.add(position);
    }
    for (const id of affected.ids) {
      ids.add(id);
    }
    hasUntracked ||= affected.hasUntracked;
  };

  const rangesMeta = tr.getMeta("folioChangedParagraphRanges") as
    | { batches: { ranges: { from: number; to: number }[]; mappingFrom: number }[] }
    | undefined;
  for (const batch of rangesMeta?.batches ?? []) {
    const remap = tr.mapping.slice(batch.mappingFrom);
    for (const range of batch.ranges) {
      const from = mapThrough(remap, range.from, 1);
      const to = mapThrough(remap, range.to, -1);
      if (to > from) {
        record(referenceCollectMarkLike(tr.doc, from, to));
      }
    }
  }

  for (let stepIndex = 0; stepIndex < tr.steps.length; stepIndex++) {
    // SAFETY: loop condition keeps stepIndex within tr.steps bounds.
    const step = tr.steps[stepIndex]!;
    const remap = tr.mapping.slice(stepIndex + 1);
    if (step instanceof AddMarkStep || step instanceof RemoveMarkStep) {
      const from = mapThrough(remap, step.from, 1);
      const to = mapThrough(remap, step.to, -1);
      if (to > from) {
        record(referenceCollectMarkLike(tr.doc, from, to));
      }
      continue;
    }
    if (
      step instanceof AddNodeMarkStep ||
      step instanceof RemoveNodeMarkStep ||
      step instanceof AttrStep
    ) {
      const mapped = remap.mapResult(step.pos, 1);
      if (mapped.deletedAcross) {
        // A later step joined the node away; its position lands mid-text.
        continue;
      }
      const pos = mapped.pos;
      const node = tr.doc.nodeAt(pos);
      if (node) {
        record(referenceCollect(tr.doc, pos, pos + node.nodeSize));
      }
      continue;
    }
    // oxlint-disable-next-line unicorn/no-array-for-each -- ProseMirror StepMap.forEach
    step.getMap().forEach((_oldStart, _oldEnd, newStart, newEnd) => {
      const from = mapThrough(remap, newStart, 1);
      const to = mapThrough(remap, newEnd, -1);
      if (to >= from) {
        record(referenceCollect(tr.doc, from, to));
      }
    });
  }

  if (hasUntracked || structureChanged) {
    tr.before.descendants((node, position) => {
      if (node.type.name !== "paragraph") {
        return true;
      }
      if (!isUsableParaId(node.attrs["paraId"])) {
        const mapped = referenceMapParagraphPosition(tr, position);
        if (mapped === null || positions.has(mapped)) {
          hasUntrackedSource = true;
          hasUntracked = true;
        }
      }
      return false;
    });
  }

  return { positions: [...positions], ids: [...ids], hasUntracked, hasUntrackedSource };
};

const trackerFields = (state: EditorState): TrackerFields => {
  const tracker = paragraphChangeTrackerKey.getState(state);
  if (!tracker) {
    throw new Error("change tracker missing from state");
  }
  return {
    positions: [...tracker.affectedParagraphPositions],
    ids: [...tracker.changedParaIds],
    hasUntracked: tracker.hasUntrackedChanges,
    hasUntrackedSource: tracker.hasUntrackedSourceChanges,
  };
};

const blockStructureFingerprint = (doc: PMNode): string => {
  const tracker = paragraphChangeTrackerKey.getState(
    EditorState.create({ doc, plugins: trackerPlugins }),
  );
  if (!tracker) {
    throw new Error("change tracker missing from state");
  }
  return tracker.blockStructureFingerprint;
};

// ---------------------------------------------------------------------------
// Paragraph-id allocator, dirty range and run identity: the per-step
// references.
// ---------------------------------------------------------------------------

const referenceKeepers = (
  oldDoc: PMNode,
  transactions: readonly Transaction[],
  tokens: ReadonlySet<string>,
) => {
  const paraIds = new Map<string, number>();
  const sourceTokens = new Map<string, unknown>();
  oldDoc.descendants((node, pos) => {
    if (node.type.name !== "paragraph") {
      return true;
    }
    const id = node.attrs["paraId"];
    const token = node.attrs["_docxParagraphSourceToken"];
    const mapsParaId = isUsableParaId(id) && !paraIds.has(id);
    const mapsSourceToken = typeof token === "string" && tokens.has(token);
    if (!mapsParaId && !mapsSourceToken) {
      return false;
    }
    let mapped = pos;
    let mappedInterior = pos + 1;
    let deleted = false;
    for (const transaction of transactions) {
      const ownerResult = transaction.mapping.mapResult(mapped);
      const interiorResult = transaction.mapping.mapResult(mappedInterior, -1);
      deleted ||= interiorResult.deletedAcross;
      mapped = ownerResult.pos;
      mappedInterior = interiorResult.pos;
    }
    if (mapsParaId && !deleted) {
      paraIds.set(id, mapped);
    }
    if (mapsSourceToken) {
      sourceTokens.set(
        token,
        sourceTokens.has(token)
          ? { status: "ambiguous" }
          : { pos: mapped, status: deleted ? "deleted" : "mapped" },
      );
    }
    return false;
  });
  return { paraIds: [...paraIds], sourceTokens: [...sourceTokens] };
};

const referenceDirtyRange = (transaction: Transaction) => {
  const range = { from: Number.POSITIVE_INFINITY, to: Number.NEGATIVE_INFINITY };
  transaction.mapping.maps.forEach((map, stepIndex) => {
    const followingMaps = transaction.mapping.slice(stepIndex + 1);
    const extend = (newStart: number, newEnd: number): void => {
      const finalStart = mapThrough(followingMaps, newStart, -1);
      const finalEnd = mapThrough(followingMaps, newEnd, 1);
      range.from = Math.min(range.from, finalStart, finalEnd);
      range.to = Math.max(range.to, finalStart, finalEnd);
    };
    // oxlint-disable-next-line unicorn/no-array-for-each -- ProseMirror StepMap API
    map.forEach((_oldStart, _oldEnd, newStart, newEnd) => extend(newStart, newEnd));
    const step = transaction.steps[stepIndex];
    if (step instanceof AddMarkStep || step instanceof RemoveMarkStep) {
      extend(step.from, step.to);
    }
  });
  return Number.isFinite(range.from) && Number.isFinite(range.to) ? range : null;
};

const referenceInsertedRanges = (transactions: readonly Transaction[]) => {
  const ranges: { from: number; to: number }[] = [];
  for (const [transactionIndex, transaction] of transactions.entries()) {
    const later = transactions.slice(transactionIndex + 1);
    const carryForward = (pos: number, bias: -1 | 1): number => {
      let mapped = pos;
      for (const next of later) {
        mapped = mapThrough(next.mapping, mapped, bias);
      }
      return mapped;
    };
    transaction.mapping.maps.forEach((stepMap, stepIndex) => {
      const rest = transaction.mapping.slice(stepIndex + 1);
      // oxlint-disable-next-line unicorn/no-array-for-each -- ProseMirror StepMap API
      stepMap.forEach((_oldStart, _oldEnd, newStart, newEnd) => {
        if (newEnd > newStart) {
          ranges.push({
            from: carryForward(mapThrough(rest, newStart, -1), -1),
            to: carryForward(mapThrough(rest, newEnd, 1), 1),
          });
        }
      });
    });
  }
  return ranges;
};

describe("multi-step transactions map as the per-step walk did", () => {
  for (const [name, plugins] of [
    ["the change tracker alone", trackerPlugins],
    ["the change tracker beside the allocator, bidi detection and run identity", editorPlugins],
  ] as const) {
    test(`${name}: identical tracker state and dirty-set order`, () => {
      fc.assert(
        fc.property(randomDocument, randomBatches, (doc, batches) => {
          let state = EditorState.create({ doc, plugins: [...plugins] });
          for (const operations of batches) {
            const before = state;
            const tr = buildRandomTransaction(before, operations);
            let applied: ReturnType<EditorState["applyTransaction"]>;
            try {
              applied = before.applyTransaction(tr);
            } catch {
              // Whatever throws here must throw in the per-step walk too. (An
              // attribute step whose node a later join removed used to: its
              // position lands mid-text, and both now skip it.)
              const previous = trackerFields(before);
              const structureChanged =
                blockStructureFingerprint(tr.before) !== blockStructureFingerprint(tr.doc);
              expect(() => referenceTrackerApply(previous, tr, structureChanged)).toThrow();
              return;
            }
            const { state: next, transactions } = applied;
            let expected = trackerFields(before);
            for (const transaction of transactions) {
              expected = referenceTrackerApply(
                expected,
                transaction,
                transaction.docChanged &&
                  blockStructureFingerprint(transaction.before) !==
                    blockStructureFingerprint(transaction.doc),
              );
            }
            expect(trackerFields(next)).toEqual(expected);
            state = next;
          }
        }),
        propertyConfig({ numRuns: 150 }),
      );
    });
  }

  test("the paragraph-id allocator keeps the same owners", () => {
    fc.assert(
      fc.property(
        randomDocument,
        randomBatches,
        fc.subarray(["token-a", "token-b", "token-c"]),
        (doc, batches, seeded) => {
          const transactions = buildRandomBatch(doc, batches);
          const tokens = new Set(seeded);
          const actual = mapParagraphKeepers(doc, transactions, { contract: null, tokens });
          expect({
            paraIds: [...actual.paraIds],
            sourceTokens: [...actual.sourceTokens],
          }).toEqual(referenceKeepers(doc, transactions, tokens));
        },
      ),
      propertyConfig({ numRuns: 300 }),
    );
  });

  test("bidi detection scans the same dirty range", () => {
    fc.assert(
      fc.property(randomDocument, randomOperations(30), (doc, operations) => {
        const tr = buildRandomTransaction(EditorState.create({ doc }), operations);
        expect(getTransactionDirtyRange(tr)).toEqual(referenceDirtyRange(tr));
      }),
      propertyConfig({ numRuns: 300 }),
    );
  });

  test("run identity strips the same inserted ranges", () => {
    fc.assert(
      fc.property(randomDocument, randomBatches, (doc, batches) => {
        const transactions = buildRandomBatch(doc, batches);
        expect(insertedRanges(transactions)).toEqual(referenceInsertedRanges(transactions));
      }),
      propertyConfig({ numRuns: 300 }),
    );
  });
});
