import { panic } from "better-result";
import type { Node as PMNode } from "prosemirror-model";
import { Plugin, Transaction } from "prosemirror-state";

import {
  getExplicitParagraphPropertySourceTransfers,
  getProseDocumentParagraphPropertySourceContract,
  getProseParagraphPropertySourceToken,
  paragraphPropertySourceTokenMatchesContract,
  PROSE_PARAGRAPH_SOURCE_TOKEN_ATTR,
  setProseParagraphMarkupWithPropertySource,
} from "../../docx/paragraphPropertySource";
import { createExtension } from "../../prosemirror/extensions/create";
import type {
  Extension,
  ExtensionRuntime,
} from "../../prosemirror/extensions/types";
/**
 * ParaIdAllocator — assigns a stable `w14:paraId` to every paragraph.
 *
 * Lifted from
 * https://github.com/eigenpal/docx-editor/blob/main/packages/core/src/prosemirror/extensions/features/ParaIdAllocatorExtension.ts
 * (Apache-2.0). Adapted to folio's `createExtension` + import style;
 * keep behaviour in sync upstream. folio divergences: duplicate
 * resolution maps the original paragraph's position through the
 * transaction (see below), and the appended allocation transaction is
 * excluded from paragraph change tracking like `ensureParaIdsInState`.
 *
 * Why: AI tooling, chat citation chips, and the change tracker all
 * anchor on `paraId`. A paragraph with `paraId: null` is invisible
 * to those surfaces; a duplicated paraId (the second half of an
 * Enter-split, or content pasted from another doc) silently desyncs
 * their anchors. This plugin closes both gaps by allocating fresh
 * ids in an `appendTransaction` hook after every doc-changed step.
 *
 * Id lifecycle rules (locked by the co-located tests):
 * - Load: a missing/empty/reserved-zero or duplicate id gets one derived
 *   from the paragraph's text and ordinal (`ensureParaIdsInDoc`), the same
 *   id the headless reviewer derives for the same document.
 * - Missing/empty/reserved-zero id after an edit: a fresh random id is minted.
 * - Split: the half holding the paragraph's original start position
 *   keeps the id; the other half gets a fresh one. Anchors inside the
 *   content stay resolvable on the half that carries them.
 * - Merge/join: the surviving paragraph keeps its id; the absorbed id
 *   dangles by design (consumers degrade to snapshot fallback).
 * - Paste/duplicate carrying an id already in the doc: the paragraph
 *   whose mapped pre-transaction position matches keeps it — pasting a
 *   copy above its source no longer steals the source's id. When no
 *   occurrence maps back (e.g. both are new), the first in document
 *   order keeps it.
 * - Paste of an id unknown to this doc (cross-doc paste, cut-then-
 *   paste move): the id is kept, so a moved paragraph stays anchored.
 * - Undo/redo: allocation applies with `addToHistory: false`, and
 *   ProseMirror remaps it into the originating history event. Redoing a
 *   split therefore restores the same allocated id instead of minting
 *   another one.
 */
import { ignoreTrackedChanges } from "../../prosemirror/paragraphChangeTrackerMeta";
import {
  isUsableParaId,
  type ParagraphPropertySourceSeed,
  paraIdAllocatorKey,
  visitAuthoredParagraphs,
} from "../../prosemirror/paraIds";
import {
  type PositionQuery,
  sweepPositions,
} from "../../prosemirror/positionSweep";
import { deterministicHexId, generateHexId } from "../../utils/hexId";
import { deterministicParaIdSeedKey } from "./paragraphRevision";

const deterministicSeedOf = (transaction: Transaction): string | undefined => {
  const seed: unknown = transaction.getMeta(deterministicParaIdSeedKey);
  if (typeof seed === "string") return seed;
  // The allocator may already have seen the root when a later plugin repairs
  // a table. ProseMirror gives every appended transaction its root explicitly.
  const root: unknown = transaction.getMeta("appendedTransaction");
  if (!(root instanceof Transaction)) return undefined;
  const rootSeed: unknown = root.getMeta(deterministicParaIdSeedKey);
  return typeof rootSeed === "string" ? rootSeed : undefined;
};

/**
 * A fresh id not already in `taken`: derived from `seed` and the paragraph's
 * position when the caller asked for reproducibility, random otherwise. Both
 * loops terminate — the id space is 2^31 wide and `taken` holds at most one
 * entry per paragraph.
 */
const mintParaId = (
  taken: ReadonlySet<string>,
  seed: string | null,
  pos: number,
): string => {
  if (seed === null) {
    let id = generateHexId();
    while (taken.has(id)) {
      id = generateHexId();
    }
    return id;
  }
  let id = deterministicHexId(`${seed}:${String(pos)}`);
  for (let salt = 1; taken.has(id); salt++) {
    id = deterministicHexId(`${seed}:${String(pos)}:${String(salt)}`);
  }
  return id;
};

type ParaIdUpdate = {
  pos: number;
  attrs: Record<string, unknown>;
};

type ParagraphOccurrence = {
  pos: number;
  attrs: Record<string, unknown>;
};

const collectParagraphPropertySourceSeed = (
  doc: PMNode,
): ParagraphPropertySourceSeed => {
  const contract = getProseDocumentParagraphPropertySourceContract(doc);
  if (!contract) {
    return { contract: null, tokens: new Set() };
  }
  const counts = new Map<string, number>();
  visitAuthoredParagraphs(doc, (node) => {
    const token = getProseParagraphPropertySourceToken(node);
    if (paragraphPropertySourceTokenMatchesContract(token, contract)) {
      counts.set(token, (counts.get(token) ?? 0) + 1);
    }
  });
  return {
    contract,
    tokens: new Set(
      [...counts].filter(([, count]) => count === 1).map(([token]) => token),
    ),
  };
};

type ParagraphPropertySourceKeeper =
  { status: "ambiguous" } | { pos: number; status: "deleted" | "mapped" };

type MappedParagraphKeepers = {
  paraIds: Map<string, number>;
  sourceTokens: Map<string, ParagraphPropertySourceKeeper>;
};

/**
 * Where each id's and each seeded source token's owner went: the first
 * occurrence that survives the batch keeps an id.
 *
 * @internal Exported for the per-step equivalence test.
 */
export const mapParagraphKeepers = (
  oldDoc: PMNode,
  transactions: readonly Transaction[],
  sourceSeed: ParagraphPropertySourceSeed,
): MappedParagraphKeepers => {
  const paragraphs: { pos: number; id: unknown; token: unknown }[] = [];
  visitAuthoredParagraphs(oldDoc, (node, pos) => {
    const id = node.attrs["paraId"];
    const token = getProseParagraphPropertySourceToken(node);
    if (
      isUsableParaId(id) ||
      (typeof token === "string" && sourceSeed.tokens.has(token))
    ) {
      paragraphs.push({ pos, id, token });
    }
  });

  // Every candidate goes through the batch at once: its owner position, and
  // its interior, whose deletion is what says the paragraph itself is gone.
  const swept = sweepPositions(
    transactions.map((transaction) => transaction.mapping),
    paragraphs.flatMap(({ pos }): PositionQuery[] => [
      { pos, assoc: 1, from: 0 },
      { pos: pos + 1, assoc: -1, from: 0 },
    ]),
  );

  const paraIds = new Map<string, number>();
  const sourceTokens = new Map<string, ParagraphPropertySourceKeeper>();
  for (const [index, { id, token }] of paragraphs.entries()) {
    // SAFETY: the sweep returns an owner and an interior per paragraph.
    const mapped = swept[2 * index]!.pos;
    const deleted = swept[2 * index + 1]!.deletedAcross;
    // The first occurrence that survives keeps the id.
    if (isUsableParaId(id) && !paraIds.has(id) && !deleted) {
      paraIds.set(id, mapped);
    }
    if (typeof token === "string" && sourceSeed.tokens.has(token)) {
      sourceTokens.set(
        token,
        sourceTokens.has(token)
          ? { status: "ambiguous" }
          : { pos: mapped, status: deleted ? "deleted" : "mapped" },
      );
    }
  }
  return { paraIds, sourceTokens };
};

type ParagraphCensus = {
  invalidSourceTokens: ParagraphOccurrence[];
  missingParaIds: ParagraphOccurrence[];
  paragraphs: Map<number, ParagraphOccurrence>;
  paraIds: Map<string, ParagraphOccurrence[]>;
  sourceTokens: Map<string, ParagraphOccurrence[]>;
};

const collectParagraphCensus = (
  doc: PMNode,
  sourceSeed?: ParagraphPropertySourceSeed,
): ParagraphCensus => {
  const census: ParagraphCensus = {
    invalidSourceTokens: [],
    missingParaIds: [],
    paragraphs: new Map(),
    paraIds: new Map(),
    sourceTokens: new Map(),
  };
  const contractMatchesSeed =
    sourceSeed?.contract !== null &&
    sourceSeed?.contract !== undefined &&
    getProseDocumentParagraphPropertySourceContract(doc) ===
      sourceSeed.contract;

  visitAuthoredParagraphs(doc, (node, pos) => {
    const occurrence = { pos, attrs: node.attrs };
    census.paragraphs.set(pos, occurrence);
    const id = node.attrs["paraId"];
    if (isUsableParaId(id)) {
      const occurrences = census.paraIds.get(id) ?? [];
      occurrences.push(occurrence);
      census.paraIds.set(id, occurrences);
    } else {
      census.missingParaIds.push(occurrence);
    }

    if (sourceSeed) {
      const token = getProseParagraphPropertySourceToken(node);
      if (token !== null && token !== undefined) {
        if (
          typeof token !== "string" ||
          !contractMatchesSeed ||
          !sourceSeed.contract ||
          !sourceSeed.tokens.has(token) ||
          !paragraphPropertySourceTokenMatchesContract(
            token,
            sourceSeed.contract,
          )
        ) {
          census.invalidSourceTokens.push(occurrence);
        } else {
          const occurrences = census.sourceTokens.get(token) ?? [];
          occurrences.push(occurrence);
          census.sourceTokens.set(token, occurrences);
        }
      }
    }
  });
  return census;
};

const collectParaIdUpdatesFromCensus = (
  census: ParagraphCensus,
  keeperPositions?: ReadonlyMap<string, number>,
  deterministicSeed: string | null = null,
): ParaIdUpdate[] => {
  const needFreshId = [...census.missingParaIds];
  for (const [id, occurrences] of census.paraIds) {
    if (occurrences.length === 1) {
      continue;
    }
    const keeperPos = keeperPositions?.get(id);
    const keeper =
      occurrences.find(({ pos }) => pos === keeperPos) ?? occurrences[0];
    for (const occurrence of occurrences) {
      if (occurrence !== keeper) {
        needFreshId.push(occurrence);
      }
    }
  }

  const taken = new Set(census.paraIds.keys());
  const updates: ParaIdUpdate[] = [];
  for (const { pos, attrs } of needFreshId) {
    const newId = mintParaId(taken, deterministicSeed, pos);
    taken.add(newId);
    updates.push({ pos, attrs: { ...attrs, paraId: newId } });
  }
  updates.sort((left, right) => left.pos - right.pos);
  return updates;
};

type ParagraphPropertySourceUpdate = {
  pos: number;
  token: string | null;
};

const collectParagraphPropertySourceUpdates = (
  census: ParagraphCensus,
  keeperPositions: ReadonlyMap<string, ParagraphPropertySourceKeeper>,
  transactions: readonly Transaction[],
): ParagraphPropertySourceUpdate[] => {
  const updates = new Map<number, string | null>();
  for (const { pos } of census.invalidSourceTokens) {
    updates.set(pos, null);
  }
  const explicitTransfers = transactions.flatMap((transaction) => [
    ...getExplicitParagraphPropertySourceTransfers(transaction),
  ]);
  const explicitlySelected = new Set(
    explicitTransfers
      .map(({ selectedToken }) => selectedToken)
      .filter((token): token is string => token !== null),
  );
  const explicitlyDisplaced = new Set(
    explicitTransfers
      .map(({ displacedToken }) => displacedToken)
      .filter((token): token is string => token !== null),
  );
  const keptPositions = new Set<number>();
  const keptTokens = new Set<string>();
  for (const [token, occurrences] of census.sourceTokens) {
    const mappedOwner = keeperPositions.get(token);
    const mappedKeeper =
      mappedOwner?.status === "mapped"
        ? occurrences.find(({ pos }) => pos === mappedOwner.pos)
        : undefined;
    const transferredKeeper =
      !mappedKeeper && explicitlySelected.has(token) && occurrences.length === 1
        ? occurrences.at(0)
        : undefined;
    const detachedKeeper =
      !mappedKeeper &&
      !transferredKeeper &&
      occurrences.length === 1 &&
      (mappedOwner === undefined ||
        (mappedOwner.status === "deleted" &&
          mappedOwner.pos === occurrences.at(0)?.pos))
        ? occurrences.at(0)
        : undefined;
    const keeper = mappedKeeper ?? transferredKeeper ?? detachedKeeper;
    if (keeper) {
      keptPositions.add(keeper.pos);
      keptTokens.add(token);
    }
    for (const occurrence of occurrences) {
      if (occurrence !== keeper) {
        updates.set(occurrence.pos, null);
      }
    }
  }

  // Attribute-only node replacement preserves the paragraph's content gap but
  // callers may accidentally omit its private token. Restore only the mapped
  // surviving owner. Whole-paragraph deletion marks the interior boundary as
  // deleted-across, so a replacement paragraph cannot borrow the old source.
  for (const [token, mappedOwner] of keeperPositions) {
    if (
      mappedOwner.status !== "mapped" ||
      keptTokens.has(token) ||
      explicitlyDisplaced.has(token) ||
      keptPositions.has(mappedOwner.pos) ||
      !census.paragraphs.has(mappedOwner.pos)
    ) {
      continue;
    }
    updates.set(mappedOwner.pos, token);
  }

  return [...updates]
    .map(([pos, token]) => ({ pos, token }))
    .sort((left, right) => left.pos - right.pos);
};

const createParaIdAllocatorPlugin = (): Plugin<ParagraphPropertySourceSeed> =>
  new Plugin({
    key: paraIdAllocatorKey.createKey(),
    state: {
      init: (_config, state) => collectParagraphPropertySourceSeed(state.doc),
      apply: (_transaction, seed) => seed,
    },

    appendTransaction(transactions, oldState, newState) {
      // Skip selection-only / mark-only transactions — they can't have
      // created or duplicated a paragraph.
      if (!transactions.some((t) => t.docChanged)) {
        return null;
      }

      const paragraphSourceSeed = paraIdAllocatorKey.getState(oldState);
      if (!paragraphSourceSeed) {
        panic("ParaId allocator lost its paragraph-property source seed");
      }
      const keeperPositions = mapParagraphKeepers(
        oldState.doc,
        transactions,
        paragraphSourceSeed,
      );
      const deterministicSeed =
        transactions
          .map(deterministicSeedOf)
          .find((value) => typeof value === "string") ?? null;
      const census = collectParagraphCensus(newState.doc, paragraphSourceSeed);
      const updates = collectParaIdUpdatesFromCensus(
        census,
        keeperPositions.paraIds,
        deterministicSeed,
      );
      const paragraphSourceUpdates = collectParagraphPropertySourceUpdates(
        census,
        keeperPositions.sourceTokens,
        transactions,
      );
      if (updates.length === 0 && paragraphSourceUpdates.length === 0) {
        return null;
      }

      const tr = newState.tr;
      for (const u of updates) {
        setProseParagraphMarkupWithPropertySource({
          attrs: u.attrs,
          ownership: "transfer-allocated-id",
          pos: u.pos,
          transaction: tr,
        });
      }
      for (const { pos, token } of paragraphSourceUpdates) {
        const paragraph = tr.doc.nodeAt(pos);
        if (!paragraph || paragraph.type.name !== "paragraph") {
          panic("Paragraph-property token update lost its paragraph");
        }
        setProseParagraphMarkupWithPropertySource({
          attrs: {
            ...paragraph.attrs,
            [PROSE_PARAGRAPH_SOURCE_TOKEN_ATTR]: token,
          },
          ownership: "preserve",
          pos,
          transaction: tr,
        });
      }
      // Allocation is bookkeeping, not a user edit: it must not mark
      // untouched paragraphs as changed (the user's own transaction
      // already recorded any real edits).
      ignoreTrackedChanges(tr);
      tr.setMeta(paraIdAllocatorKey, "allocated");
      tr.setMeta("addToHistory", false);
      if (newState.storedMarks !== null) {
        tr.setStoredMarks(newState.storedMarks);
      }
      return tr;
    },
  });

export const ParaIdAllocatorExtension: (
  options?: Partial<Record<never, never>> | undefined,
) => Extension = createExtension({
  name: "paraIdAllocator",
  defaultOptions: {},
  onSchemaReady(): ExtensionRuntime {
    return {
      plugins: [createParaIdAllocatorPlugin()],
    };
  },
});
