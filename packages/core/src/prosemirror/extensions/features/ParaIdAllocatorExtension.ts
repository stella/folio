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
 * - Missing/empty id: a fresh random id is minted.
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
 * - Undo/redo: allocation applies with `addToHistory: false`, so
 *   redoing a split mints a different fresh id than before the undo.
 *   An id is stable across saves, not across redo-recreation.
 */
import type { Node as PMNode } from "prosemirror-model";
import { Plugin, PluginKey } from "prosemirror-state";
import type { EditorState, Transaction } from "prosemirror-state";

import { generateHexId } from "../../../utils/hexId";
import { createExtension } from "../create";
import type { ExtensionRuntime } from "../types";
import { ignoreTrackedChanges } from "./ParagraphChangeTrackerExtension";

export const paraIdAllocatorKey = new PluginKey("paraIdAllocator");

type ParaIdUpdate = {
  pos: number;
  attrs: Record<string, unknown>;
};

type ParagraphOccurrence = {
  pos: number;
  attrs: Record<string, unknown>;
};

/**
 * For every valid paraId in the pre-transaction doc, the position its
 * paragraph ends up at after `transactions` — the occurrence that
 * rightfully keeps the id when the new doc holds duplicates. Ids whose
 * paragraph start was deleted by the steps are left out.
 */
const mapKeeperPositions = (
  oldDoc: PMNode,
  transactions: readonly Transaction[],
): Map<string, number> => {
  const keepers = new Map<string, number>();
  oldDoc.descendants((node, pos) => {
    if (node.type.name !== "paragraph") {
      return true;
    }
    const id = node.attrs["paraId"];
    if (typeof id === "string" && id.length > 0 && !keepers.has(id)) {
      let mapped = pos;
      let deleted = false;
      for (const tr of transactions) {
        const result = tr.mapping.mapResult(mapped);
        if (result.deleted) {
          deleted = true;
          break;
        }
        mapped = result.pos;
      }
      if (!deleted) {
        keepers.set(id, mapped);
      }
    }
    return false;
  });
  return keepers;
};

const collectParaIdUpdates = (
  doc: PMNode,
  keeperPositions?: Map<string, number>,
): ParaIdUpdate[] => {
  // Pass 1: paragraphs without a usable id, and occurrences per id.
  const missing: ParagraphOccurrence[] = [];
  const occurrencesById = new Map<string, ParagraphOccurrence[]>();

  doc.descendants((node, pos) => {
    // Non-paragraph: recurse — paragraphs nested in tables / cells
    // are still in scope.
    if (node.type.name !== "paragraph") {
      return true;
    }
    const id = node.attrs["paraId"];
    if (typeof id !== "string" || id.length === 0) {
      missing.push({ pos, attrs: node.attrs });
    } else {
      const occurrences = occurrencesById.get(id) ?? [];
      occurrences.push({ pos, attrs: node.attrs });
      occurrencesById.set(id, occurrences);
    }
    // Paragraphs only contain inline content (text / runs) — nothing
    // we'd ever paraId. Skip the subtree.
    return false;
  });

  // Pass 2: keeper resolution. A unique id always stays put; among
  // duplicates, the occurrence at the id's mapped pre-transaction
  // position keeps it, falling back to document order.
  const needFreshId: ParagraphOccurrence[] = [...missing];
  for (const [id, occurrences] of occurrencesById) {
    if (occurrences.length === 1) {
      continue;
    }
    const keeperPos = keeperPositions?.get(id);
    const keeper = occurrences.find((occurrence) => occurrence.pos === keeperPos) ?? occurrences[0];
    for (const occurrence of occurrences) {
      if (occurrence !== keeper) {
        needFreshId.push(occurrence);
      }
    }
  }

  const taken = new Set(occurrencesById.keys());
  const updates: ParaIdUpdate[] = [];
  for (const { pos, attrs } of needFreshId) {
    let newId = generateHexId();
    while (taken.has(newId)) {
      newId = generateHexId();
    }
    taken.add(newId);
    updates.push({ pos, attrs: { ...attrs, paraId: newId } });
  }

  // Document order keeps transaction step order deterministic.
  updates.sort((a, b) => a.pos - b.pos);
  return updates;
};

export const ensureParaIdsInState = (state: EditorState): EditorState => {
  const updates = collectParaIdUpdates(state.doc);
  if (updates.length === 0) {
    return state;
  }

  const tr = state.tr;
  for (const update of updates) {
    tr.setNodeMarkup(update.pos, undefined, update.attrs);
  }
  ignoreTrackedChanges(tr);
  tr.setMeta(paraIdAllocatorKey, "allocated");
  tr.setMeta("addToHistory", false);
  return state.apply(tr);
};

const createParaIdAllocatorPlugin = (): Plugin =>
  new Plugin({
    key: paraIdAllocatorKey,
    appendTransaction(transactions, oldState, newState) {
      // Skip selection-only / mark-only transactions — they can't have
      // created or duplicated a paragraph.
      if (!transactions.some((t) => t.docChanged)) {
        return null;
      }

      const keeperPositions = mapKeeperPositions(oldState.doc, transactions);
      const updates = collectParaIdUpdates(newState.doc, keeperPositions);
      if (updates.length === 0) {
        return null;
      }

      const tr = newState.tr;
      for (const u of updates) {
        tr.setNodeMarkup(u.pos, undefined, u.attrs);
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
