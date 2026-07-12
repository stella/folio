/**
 * Redline generator: compare two `.docx` buffers and produce a THIRD `.docx`
 * whose differences are recorded as real Word tracked changes (`w:ins` /
 * `w:del`), ready for review in Word or the folio editor — the
 * document-producing counterpart to {@link compareDocxVersions}'s structured
 * diff.
 *
 * The pipeline is a composition of existing machinery, not a new engine:
 * the base buffer is opened in a headless {@link FolioDocxReviewer}, the two
 * snapshots are aligned with {@link alignFolioBlocks} (the same three-pass
 * alignment the comparer uses), each alignment event maps onto a
 * tracked-changes {@link FolioAIEditOperation}, and the reviewer's shared
 * apply path records them as redlines. Word-level minimality comes free:
 * the apply engine word-diffs a `replaceBlock` and only marks the divergent
 * tokens.
 *
 * ## Semantics and limitations
 *
 * - **As-accepted inputs.** Pending tracked changes in either input count as
 *   applied before comparison (mirroring {@link compareDocxVersions}): the
 *   base's own pending redlines are accepted in the output, so the tracked
 *   changes it carries are exactly base → revised. This is intentionally
 *   lossy about the inputs' own revision history and authorship.
 * - **Text redlines only.** Format-only changes and relocated blocks are
 *   redlined as plain edits (a move becomes a delete + insert, like Word
 *   compare with move detection off); block-level formatting is carried via
 *   the revised block's `styleId` on inserted blocks.
 * - **Anchoring.** Insertions anchor `before` the next surviving base block
 *   so consecutive additions keep their order; additions past the last base
 *   block anchor `after` it. A base document with no content blocks offers
 *   no anchor at all — such additions surface in `skipped` rather than
 *   silently vanishing.
 */

import { FolioDocxReviewer } from "./ai-edits/headless";
import type { FolioAIEditApplyResult, FolioAIEditOperation } from "./ai-edits/types";
import { alignFolioBlocks, type FolioAlignedBlockEvent } from "./version-comparison";

/** Options for {@link generateRedlineDocx}. */
export type GenerateRedlineDocxOptions = {
  /** Author recorded on the generated tracked changes. (default: `"folio compare"`) */
  author?: string;
};

/** Result of {@link generateRedlineDocx}. */
export type GenerateRedlineDocxResult = FolioAIEditApplyResult & {
  /** The redline `.docx`: the base document with base → revised tracked changes. */
  buffer: ArrayBuffer;
};

/**
 * The base block id the next `revisedOnly` event at `startIndex` should
 * anchor before: the base side of the nearest later `pair` or `baseOnly`
 * event, or `null` when only additions remain until the end of the document.
 */
const nextBaseBlockId = (
  events: readonly FolioAlignedBlockEvent[],
  startIndex: number,
): string | null => {
  for (let i = startIndex; i < events.length; i++) {
    const event = events[i];
    if (!event) {
      continue;
    }
    if (event.type === "pair") {
      return event.baseBlock.id;
    }
    if (event.type === "baseOnly") {
      return event.block.id;
    }
  }
  return null;
};

/**
 * Compare `base` and `revised` and return the base document with every
 * difference recorded as a tracked change. See the module doc comment for
 * semantics and limitations; `skipped` reports any block the generator could
 * not redline (e.g. additions with no anchor in an empty base document).
 */
export const generateRedlineDocx = async (
  base: ArrayBuffer,
  revised: ArrayBuffer,
  options: GenerateRedlineDocxOptions = {},
): Promise<GenerateRedlineDocxResult> => {
  const [baseReviewer, revisedReviewer] = await Promise.all([
    FolioDocxReviewer.fromBuffer(base, { author: options.author ?? "folio compare" }),
    FolioDocxReviewer.fromBuffer(revised),
  ]);
  // As-accepted semantics: resolve the base's own pending redlines first so
  // the output's tracked changes describe exactly base -> revised.
  baseReviewer.acceptAll();
  const baseSnapshot = baseReviewer.snapshot();
  const revisedBlocks = revisedReviewer.snapshot().blocks;

  const events = alignFolioBlocks(baseSnapshot.blocks, revisedBlocks);

  const operations: FolioAIEditOperation[] = [];
  let operationSeq = 0;
  const nextOperationId = () => `redline-${++operationSeq}`;
  /** Additions past the last base block, inserted after it in reverse so the final order matches the revised document. */
  const trailingAdditions: { text: string; styleId?: string }[] = [];
  const lastBaseBlockId = baseSnapshot.blocks.at(-1)?.id ?? null;

  events.forEach((event, eventIndex) => {
    if (event.type === "pair") {
      if (event.baseBlock.text !== event.revisedBlock.text) {
        operations.push({
          id: nextOperationId(),
          type: "replaceBlock",
          blockId: event.baseBlock.id,
          text: event.revisedBlock.text,
        });
      }
      return;
    }
    if (event.type === "baseOnly") {
      operations.push({
        id: nextOperationId(),
        type: "deleteBlock",
        blockId: event.block.id,
      });
      return;
    }
    const anchorId = nextBaseBlockId(events, eventIndex + 1);
    if (anchorId === null) {
      trailingAdditions.push({
        text: event.block.text,
        ...(event.block.styleId !== undefined && { styleId: event.block.styleId }),
      });
      return;
    }
    operations.push({
      id: nextOperationId(),
      type: "insertBeforeBlock",
      blockId: anchorId,
      text: event.block.text,
      ...(event.block.styleId !== undefined && { styleId: event.block.styleId }),
    });
  });

  for (const addition of trailingAdditions.toReversed()) {
    operations.push({
      id: nextOperationId(),
      type: "insertAfterBlock",
      // An empty base document has no anchor block at all; an unresolvable
      // id makes the shared applier report the addition in `skipped`
      // (missingBlock) instead of dropping it silently.
      blockId: lastBaseBlockId ?? "redline-unanchored",
      text: addition.text,
      ...(addition.styleId !== undefined && { styleId: addition.styleId }),
    });
  }

  const { applied, skipped } = baseReviewer.applyOperations(operations, {
    mode: "tracked-changes",
    snapshot: baseSnapshot,
  });
  const buffer = await baseReviewer.toBuffer();
  return { buffer, applied, skipped };
};
