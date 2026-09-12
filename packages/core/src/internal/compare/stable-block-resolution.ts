import type { Node as PMNode } from "prosemirror-model";

import { buildCleanBlockText } from "../../ai-edits/clean-text";
import {
  hashFolioAIBlockStructuralBoundaries,
  hashFolioAIBlockText,
  isHiddenTableRow,
  normalizeFolioAIBlockText,
} from "../../ai-edits/snapshot";
import type { FolioAIEditSnapshot } from "../../ai-edits/types";
import { getFolioParaIdFromBlockId } from "../../types/block-id";

type LiveBlock = { readonly from: number; readonly to: number; readonly node: PMNode };

export type FolioStableBlockResolution =
  | {
      readonly type: "resolved";
      readonly blockNode: PMNode;
      readonly blockFrom: number;
      readonly blockTo: number;
      readonly cleanBlock: ReturnType<typeof buildCleanBlockText>;
      readonly currentText: string;
      readonly currentTextHash: string;
    }
  | { readonly type: "unsupported"; readonly reason: "missing-block" | "changed-block" };

/** One immutable story index shared by every comparison-instruction preflight. */
export class FolioStableBlockResolver {
  readonly #anchors: FolioAIEditSnapshot["anchors"];
  readonly #byHash = new Map<string, readonly LiveBlock[]>();
  readonly #byParaId = new Map<string, LiveBlock>();
  readonly #ambiguousParaIds = new Set<string>();
  readonly #ordinalByBlockId = new Map<string, number>();

  private constructor(doc: PMNode, snapshot: FolioAIEditSnapshot) {
    this.#anchors = snapshot.anchors;
    const snapshotHashCounts = new Map<string, number>();
    for (const block of snapshot.blocks) {
      const anchor = snapshot.anchors[block.id];
      if (!anchor) continue;
      const ordinal = snapshotHashCounts.get(anchor.textHash) ?? 0;
      this.#ordinalByBlockId.set(block.id, ordinal);
      snapshotHashCounts.set(anchor.textHash, ordinal + 1);
    }

    const mutableByHash = new Map<string, LiveBlock[]>();
    doc.descendants((node, pos) => {
      if (isHiddenTableRow(node)) return false;
      if (!node.isTextblock) return true;
      const block = Object.freeze({ from: pos, to: pos + node.nodeSize, node });
      const cleanText = buildCleanBlockText(node, pos).text;
      const hash = hashFolioAIBlockText(normalizeFolioAIBlockText(cleanText));
      const bucket = mutableByHash.get(hash) ?? [];
      bucket.push(block);
      mutableByHash.set(hash, bucket);
      const paraId: unknown = node.attrs["paraId"];
      if (typeof paraId !== "string" || paraId.length === 0) return false;
      if (this.#byParaId.has(paraId)) {
        this.#byParaId.delete(paraId);
        this.#ambiguousParaIds.add(paraId);
      } else if (!this.#ambiguousParaIds.has(paraId)) {
        this.#byParaId.set(paraId, block);
      }
      return false;
    });
    for (const [hash, blocks] of mutableByHash) {
      this.#byHash.set(hash, Object.freeze(blocks));
    }
  }

  static create(doc: PMNode, snapshot: FolioAIEditSnapshot): FolioStableBlockResolver {
    return new FolioStableBlockResolver(doc, snapshot);
  }

  resolve(blockId: string): FolioStableBlockResolution {
    const anchor = this.#anchors[blockId];
    if (!anchor) return { type: "unsupported", reason: "missing-block" };
    const encodedParaId = getFolioParaIdFromBlockId(blockId);
    if (encodedParaId !== null && this.#ambiguousParaIds.has(encodedParaId)) {
      return { type: "unsupported", reason: "changed-block" };
    }
    const ordinal = this.#ordinalByBlockId.get(blockId);
    let live = encodedParaId === null ? undefined : this.#byParaId.get(encodedParaId);
    if (encodedParaId === null && ordinal !== undefined) {
      live = this.#byHash.get(anchor.textHash)?.[ordinal];
    }
    if (!live?.node.isTextblock) {
      return {
        type: "unsupported",
        reason: encodedParaId === null ? "changed-block" : "missing-block",
      };
    }
    const cleanBlock = buildCleanBlockText(live.node, live.from);
    const currentText = cleanBlock.text;
    const currentTextHash = hashFolioAIBlockText(normalizeFolioAIBlockText(currentText));
    if (
      currentTextHash !== anchor.textHash ||
      hashFolioAIBlockStructuralBoundaries(cleanBlock) !== anchor.structuralBoundaryHash
    ) {
      return { type: "unsupported", reason: "changed-block" };
    }
    return Object.freeze({
      type: "resolved",
      blockNode: live.node,
      blockFrom: live.from,
      blockTo: live.to,
      cleanBlock,
      currentText,
      currentTextHash,
    });
  }
}
