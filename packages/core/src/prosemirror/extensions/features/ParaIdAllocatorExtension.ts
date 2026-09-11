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
 * - Missing/empty/reserved-zero id: a fresh random id is minted.
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
import { panic } from "better-result";
import { Fragment, type Node as PMNode } from "prosemirror-model";
import { Plugin, PluginKey } from "prosemirror-state";
import type { EditorState, Transaction } from "prosemirror-state";

import { deterministicHexId, generateHexId } from "../../../utils/hexId";
import {
  type ParagraphPropertySourceContract,
  ParagraphPropertySourceValidationError,
  getExplicitParagraphPropertySourceTransfers,
  readProseDocumentParagraphPropertySourceContract,
  setProseParagraphMarkupWithPropertySource,
} from "../../../docx/paragraphPropertySource";
import { recreateProseNode } from "../../recreateNode";
import {
  readParagraphPropertyState,
  paragraphPropertyStateAttribute,
  transitionParagraphPropertyState,
  type ParagraphPropertyState,
} from "../../paragraphPropertyState";
import { createExtension } from "../create";
import type { ExtensionRuntime } from "../types";
import { ignoreTrackedChanges } from "./ParagraphChangeTrackerExtension";

type ParagraphPropertySourceSeed =
  | { type: "unbound" }
  | {
      contract: ParagraphPropertySourceContract;
      tokens: ReadonlySet<string>;
      type: "bound";
    };

export const paraIdAllocatorKey = new PluginKey<ParagraphPropertySourceSeed>("paraIdAllocator");

/**
 * Transaction meta asking this plugin to derive the ids it mints from a
 * caller-supplied seed instead of `Math.random()`.
 *
 * A live editor wants random ids: two people typing must not collide. A
 * generator whose whole output has to be reproducible — the document compare —
 * cannot afford them, and it is the only party that knows its run is meant to
 * be deterministic, so it says so on its own transaction rather than leaving
 * the plugin to guess.
 */
export const deterministicParaIdSeedKey = new PluginKey<string>("deterministicParaIdSeed");

/** Ask the allocator to mint reproducible ids for the paragraphs `tr` creates. */
export const requestDeterministicParaIds = (tr: Transaction, seed: string): Transaction =>
  tr.setMeta(deterministicParaIdSeedKey, seed);

/**
 * A fresh id not already in `taken`: derived from `seed` and the paragraph's
 * position when the caller asked for reproducibility, random otherwise. Both
 * loops terminate — the id space is 2^31 wide and `taken` holds at most one
 * entry per paragraph.
 */
const mintParaId = (taken: ReadonlySet<string>, seed: string | null, pos: number): string => {
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

const isUsableParaId = (value: unknown): value is string =>
  typeof value === "string" && value.length > 0 && value !== "00000000";

const paragraphPropertyStateOf = (node: PMNode): ParagraphPropertyState => {
  const state = readParagraphPropertyState(node.attrs["_paragraphPropertyState"]);
  if (state.status !== "valid") {
    throw new ParagraphPropertySourceValidationError({
      code: "invalid_state",
      message: "A paragraph contains invalid paragraph-property state.",
      ...(state.status === "invalid" ? { token: state.raw } : {}),
    });
  }
  return state.value;
};

const importedTokenOf = (state: ParagraphPropertyState): string | null =>
  state.type === "imported" ? state.token : null;

const contractAcceptsWireToken = (
  contract: ParagraphPropertySourceContract,
  token: string,
): boolean => contract.readToken(token).status === "valid";

const collectParagraphPropertySourceSeed = (doc: PMNode): ParagraphPropertySourceSeed => {
  const contract = readProseDocumentParagraphPropertySourceContract(doc);
  if (contract.status === "invalid") {
    throw new ParagraphPropertySourceValidationError({
      code: "contract_mismatch",
      message: "The ProseMirror document contains an invalid paragraph-property source contract.",
    });
  }
  if (contract.status === "absent") {
    doc.descendants((node) => {
      if (node.type.name !== "paragraph") {
        return true;
      }
      const state = paragraphPropertyStateOf(node);
      if (state.type === "imported") {
        throw new ParagraphPropertySourceValidationError({
          code: "contract_mismatch",
          message: "An imported paragraph-property state requires its source contract.",
          token: state.token,
        });
      }
      return false;
    });
    return { type: "unbound" };
  }
  const counts = new Map<string, number>();
  doc.descendants((node) => {
    if (node.type.name !== "paragraph") {
      return true;
    }
    const state = paragraphPropertyStateOf(node);
    if (state.type === "imported" && !contractAcceptsWireToken(contract.value, state.token)) {
      throw new ParagraphPropertySourceValidationError({
        code: "unknown_token",
        message: "A paragraph-property token belongs to a different source document.",
        token: state.token,
      });
    }
    if (state.type === "imported") {
      const serialized = state.token;
      counts.set(serialized, (counts.get(serialized) ?? 0) + 1);
    }
    return false;
  });
  return {
    contract: contract.value,
    tokens: new Set([...counts].filter(([, count]) => count === 1).map(([token]) => token)),
    type: "bound",
  };
};

type ParagraphPropertySourceKeeper =
  | { status: "ambiguous" }
  | { pos: number; status: "deleted" | "mapped" };

type MappedParagraphKeepers = {
  paraIds: Map<string, number>;
  sourceTokens: Map<string, ParagraphPropertySourceKeeper>;
};

const mapParagraphKeepers = (
  oldDoc: PMNode,
  transactions: readonly Transaction[],
  sourceSeed: ParagraphPropertySourceSeed,
): MappedParagraphKeepers => {
  const paraIds = new Map<string, number>();
  const sourceTokens = new Map<string, ParagraphPropertySourceKeeper>();
  oldDoc.descendants((node, pos) => {
    if (node.type.name !== "paragraph") {
      return true;
    }
    const id = node.attrs["paraId"];
    const token = importedTokenOf(paragraphPropertyStateOf(node));
    const mapsParaId = isUsableParaId(id) && !paraIds.has(id);
    const sourceToken =
      sourceSeed.type === "bound" &&
      token !== null &&
      sourceSeed.tokens.has(token)
        ? token
        : null;
    const mapsSourceToken = sourceToken !== null;
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
        sourceToken,
        sourceTokens.has(sourceToken)
          ? { status: "ambiguous" }
          : { pos: mapped, status: deleted ? "deleted" : "mapped" },
      );
    }
    return false;
  });
  return { paraIds, sourceTokens };
};

type ParagraphCensus = {
  detachedSourceStates: ParagraphOccurrence[];
  missingParaIds: ParagraphOccurrence[];
  paraIds: Map<string, ParagraphOccurrence[]>;
  sourceTokens: Map<string, ParagraphOccurrence[]>;
};

const collectParagraphCensus = (
  doc: PMNode,
  sourceSeed?: ParagraphPropertySourceSeed,
): ParagraphCensus => {
  const census: ParagraphCensus = {
    detachedSourceStates: [],
    missingParaIds: [],
    paraIds: new Map(),
    sourceTokens: new Map(),
  };
  const contract = readProseDocumentParagraphPropertySourceContract(doc);
  const contractMatchesSeed =
    sourceSeed?.type === "bound" &&
    contract.status === "valid" &&
    contract.value.serialized === sourceSeed.contract.serialized;

  doc.descendants((node, pos) => {
    if (node.type.name !== "paragraph") {
      return true;
    }
    const occurrence = { pos, attrs: node.attrs };
    const id = node.attrs["paraId"];
    if (isUsableParaId(id)) {
      const occurrences = census.paraIds.get(id) ?? [];
      occurrences.push(occurrence);
      census.paraIds.set(id, occurrences);
    } else {
      census.missingParaIds.push(occurrence);
    }

    if (sourceSeed) {
      const state = paragraphPropertyStateOf(node);
      if (state.type === "imported") {
        if (
          !contractMatchesSeed ||
          sourceSeed.type !== "bound" ||
          !sourceSeed.tokens.has(state.token) ||
          !contractAcceptsWireToken(sourceSeed.contract, state.token)
        ) {
          census.detachedSourceStates.push(occurrence);
        } else {
          const serialized = state.token;
          const occurrences = census.sourceTokens.get(serialized) ?? [];
          occurrences.push(occurrence);
          census.sourceTokens.set(serialized, occurrences);
        }
      }
    }
    return false;
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
    const keeper = occurrences.find(({ pos }) => pos === keeperPos) ?? occurrences[0];
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

const collectParaIdUpdates = (doc: PMNode): ParaIdUpdate[] =>
  collectParaIdUpdatesFromCensus(collectParagraphCensus(doc));

type ParagraphPropertySourceUpdate = {
  pos: number;
  state: ParagraphPropertyState;
};

const collectParagraphPropertySourceUpdates = (
  census: ParagraphCensus,
  keeperPositions: ReadonlyMap<string, ParagraphPropertySourceKeeper>,
  transactions: readonly Transaction[],
): ParagraphPropertySourceUpdate[] => {
  const updates = new Map<number, ParagraphPropertyState>();
  for (const { attrs, pos } of census.detachedSourceStates) {
    const state = readParagraphPropertyState(attrs["_paragraphPropertyState"]);
    if (state.status !== "valid") {
      panic("Validated paragraph-property state became malformed during allocation");
    }
    updates.set(pos, transitionParagraphPropertyState(state.value, { type: "editor-copy" }));
  }
  const explicitTransfers = transactions.flatMap((transaction) => [
    ...getExplicitParagraphPropertySourceTransfers(transaction),
  ]);
  const explicitlySelected = new Set(
    explicitTransfers
      .map(({ selectedToken }) => selectedToken)
      .filter((token): token is string => token !== null),
  );
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
        (mappedOwner.status === "deleted" && mappedOwner.pos === occurrences.at(0)?.pos))
        ? occurrences.at(0)
        : undefined;
    const keeper = mappedKeeper ?? transferredKeeper ?? detachedKeeper;
    for (const occurrence of occurrences) {
      if (occurrence !== keeper) {
        const state = readParagraphPropertyState(occurrence.attrs["_paragraphPropertyState"]);
        if (state.status !== "valid") {
          panic("Validated paragraph-property state became malformed during allocation");
        }
        updates.set(
          occurrence.pos,
          transitionParagraphPropertyState(state.value, { type: "editor-copy" }),
        );
      }
    }
  }

  return [...updates]
    .map(([pos, state]) => ({ pos, state }))
    .sort((left, right) => left.pos - right.pos);
};

type InitialParaIds = {
  needsRewrite: boolean;
  taken: Set<string>;
};

const collectInitialParaIds = (doc: PMNode): InitialParaIds => {
  let needsRewrite = false;
  const taken = new Set<string>();
  doc.descendants((node) => {
    if (node.type.name !== "paragraph") {
      return true;
    }

    const id = node.attrs["paraId"];
    if (!isUsableParaId(id) || taken.has(id)) {
      needsRewrite = true;
    } else {
      taken.add(id);
    }
    return false;
  });
  return { needsRewrite, taken };
};

const rewriteInitialParaIds = (parent: PMNode, taken: Set<string>, seen: Set<string>): Fragment => {
  let changed = false;
  const children: PMNode[] = [];
  parent.forEach((child) => {
    let next = child;
    if (child.type.name === "paragraph") {
      const id = child.attrs["paraId"];
      if (!isUsableParaId(id) || seen.has(id)) {
        let newId = generateHexId();
        while (taken.has(newId)) {
          newId = generateHexId();
        }
        taken.add(newId);
        seen.add(newId);
        next = recreateProseNode(child, {
          attrs: { ...child.attrs, paraId: newId },
        });
      } else {
        seen.add(id);
      }
    } else if (child.childCount > 0) {
      const content = rewriteInitialParaIds(child, taken, seen);
      if (content !== child.content) {
        next = recreateProseNode(child, { content });
      }
    }
    if (next !== child) {
      changed = true;
    }
    children.push(next);
  });
  return changed ? Fragment.fromArray(children) : parent.content;
};

/**
 * Allocate initial paragraph IDs before constructing editor state.
 *
 * Rebuilding only changed branches avoids applying one ProseMirror transaction
 * step per missing paragraph. Live edits still use the transaction-based plugin
 * below so history, mappings, and duplicate ownership retain their semantics.
 */
export const ensureParaIdsInDoc = (doc: PMNode): PMNode => {
  const { needsRewrite, taken } = collectInitialParaIds(doc);
  if (!needsRewrite) {
    return doc;
  }

  return recreateProseNode(doc, {
    content: rewriteInitialParaIds(doc, taken, new Set()),
  });
};

export const ensureParaIdsInState = (state: EditorState): EditorState => {
  const updates = collectParaIdUpdates(state.doc);
  if (updates.length === 0) {
    return state;
  }

  const tr = state.tr;
  for (const update of updates) {
    setProseParagraphMarkupWithPropertySource({
      attrs: update.attrs,
      ownership: "transfer-allocated-id",
      pos: update.pos,
      transaction: tr,
    });
  }
  ignoreTrackedChanges(tr);
  tr.setMeta(paraIdAllocatorKey, "allocated");
  tr.setMeta("addToHistory", false);
  return state.apply(tr);
};

const createParaIdAllocatorPlugin = (): Plugin<ParagraphPropertySourceSeed> =>
  new Plugin({
    key: paraIdAllocatorKey,
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
      const keeperPositions = mapParagraphKeepers(oldState.doc, transactions, paragraphSourceSeed);
      const deterministicSeed =
        transactions
          .map((transaction) => transaction.getMeta(deterministicParaIdSeedKey))
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
      for (const { pos, state } of paragraphSourceUpdates) {
        const paragraph = tr.doc.nodeAt(pos);
        if (!paragraph || paragraph.type.name !== "paragraph") {
          panic("Paragraph-property token update lost its paragraph");
        }
        setProseParagraphMarkupWithPropertySource({
          attrs: {
            ...paragraph.attrs,
            _paragraphPropertyState: paragraphPropertyStateAttribute(state),
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
      return tr;
    },
  });

export const ParaIdAllocatorExtension = createExtension({
  name: "paraIdAllocator",
  defaultOptions: {},
  onSchemaReady(): ExtensionRuntime {
    return {
      plugins: [createParaIdAllocatorPlugin()],
    };
  },
});
