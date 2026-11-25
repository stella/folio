/**
 * The public option and result types of the AI edit applier, and the
 * operation types it applies in suggested mode.
 */
import type { EditorState, Transaction } from "prosemirror-state";

import type {
  FolioAIEditApplyResult,
  FolioAIResolvedEditOperation,
} from "./types";
import type { WordDiffGranularity } from "./word-diff";

/**
 * The only editor surface the apply logic touches: a current `state`
 * and a `dispatch` that swaps in the next one. A live `EditorView`
 * satisfies this structurally, and so does a headless seam
 * (`{ state, dispatch: (tr) => { state = state.apply(tr); } }`) — the
 * apply path never reaches for anything DOM-bound on the view, which is
 * what lets the same operation applier drive both the React editor and
 * the server-side reviewer in `./headless`.
 */
export type FolioAIEditView = {
  state: EditorState;
  dispatch: (transaction: Transaction) => void;
};

/**
 * Fixed revision provenance, for callers whose output must be reproducible.
 * Both halves of the stamp travel together because either one left to the
 * ambient clock (`new Date()` for the date, the shared bounded allocator for
 * the ids) makes the produced package differ between two runs over the same
 * inputs. `idSeed` starts a contiguous range the batch allocates from, so it
 * must sit above every revision id the target document already carries.
 */
export type FolioRevisionStamp = {
  /** ISO-8601 date stamped on every `w:ins` / `w:del` this batch produces. */
  date: string;
  /** First revision id the batch may allocate. */
  idSeed: number;
};

/**
 * How an apply batch cuts the redline for a replacement. Word-level by
 * default; character granularity marks the changed letters inside a token,
 * which reads well for a reference or a date. Normalization is deliberately
 * absent: a batch that leaves a difference unmarked does not accept back to
 * the text the caller asked for.
 */
export type FolioWordDiffOptions = { granularity?: WordDiffGranularity };

/**
 * What a replacement does with a highlight or `w:shd` the text it replaces
 * carries.
 *
 * `clear` is the authoring default: text typed over a highlighted placeholder
 * is new text, and keeping the marker that said "fill this in" would carry it
 * into the finished document. `keep` is for a caller that states the
 * replacement's own run properties rather than letting it inherit them — a
 * comparison reproducing a revised document holds the properties that document
 * has, and clearing them first records a run-property change it takes straight
 * back, which reaches the reader as a revision whose before and after are the
 * same.
 */
export type FolioReplacementBackground = "clear" | "keep";

/**
 * An apply result plus where the batch left the revision-id counter.
 *
 * A caller writing several batches into one package — a comparison walking
 * the main story, then each header, footer and note — has to give every batch
 * a seed above the last id the previous one used, or two stories claim the
 * same `w:id` and a consumer resolving one revision resolves the other with
 * it. The batch is the only thing that knows how many ids it took, so it says
 * so rather than making the caller guess a stride.
 */
export type FolioAIEditApplyOutcome = FolioAIEditApplyResult & {
  /** The held transaction was refused when its final geometry exposed a protected break. */
  heldRevisionConflict?: true;
  /**
   * First revision id a following batch may allocate: one past the last id
   * this batch used, or the seed it started from when it allocated none.
   */
  nextRevisionId: number;
};

/**
 * Operation types applied in `"suggested"` mode. Every produced revision — an
 * inline mark, a whole-node `_suggestedInsert` marker, or a suggested
 * `trIns`/`trDel`/`cellMarker` — is stripped from serialized DOCX until
 * accepted. Cell merge/split (not row/column ops) and comment ops stay
 * `unsupportedMode`.
 */
export const SUGGESTED_SUPPORTED_OPERATION_TYPES: ReadonlySet<
  FolioAIResolvedEditOperation["type"]
> = new Set([
  "replaceInBlock",
  "replaceRange",
  "formatRange",
  "setBlockParagraphProperties",
  "replaceBlock",
  "deleteBlock",
  "insertAfterBlock",
  "insertBeforeBlock",
  "insertSignatureTable",
  "insertTableRow",
  "deleteTableRow",
  "insertTableColumn",
  "deleteTableColumn",
]);
