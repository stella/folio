/**
 * Headless `.docx` review path: buffer -> apply AI edits -> buffer, with
 * no `EditorView` and no DOM. A queue worker or agent can read a document,
 * apply `FolioAIEditOperation`s (tracked-changes or direct), and write a
 * reviewed `.docx` back out — the server-side counterpart to the React
 * editor's live apply flow.
 *
 * The operation applier is shared, not forked: {@link applyFolioAIEditOperations}
 * only needs a {@link FolioAIEditView} seam (`{ state, dispatch }`), which a
 * headless `EditorState` satisfies via `state.apply(tr)`. The React editor and
 * this reviewer therefore run byte-for-byte the same anchor resolution,
 * word-diff redlines, and tracked-change bookkeeping.
 *
 * Save mirrors the editor's own path: a selective patch of only the changed
 * paragraphs in `document.xml` (leaving every untouched part byte-exact), with
 * a full repack as the fallback for structural edits.
 *
 * Scope: BODY blocks. Footnote / endnote note-body apply is intentionally out
 * of scope here and can build on the note serialization work separately.
 */

import { EditorState } from "prosemirror-state";
import type { Command } from "prosemirror-state";
import type { Plugin } from "prosemirror-state";
import type { Transaction } from "prosemirror-state";

import { attemptSelectiveSave } from "../docx/selectiveSave";
import { parseDocx } from "../docx/parser";
import { repackDocx } from "../docx/rezip";
import {
  acceptAIEditRevision,
  acceptAllChanges,
  rejectAIEditRevision,
  rejectAllChanges,
} from "../prosemirror/commands/comments";
import { updateDocumentContent } from "../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import {
  getChangedParagraphIds,
  hasStructuralChanges,
  hasUntrackedChanges,
  ignoreTrackedChanges,
} from "../prosemirror/extensions/features/ParagraphChangeTrackerExtension";
import { schema, singletonManager } from "../prosemirror/schema";
import type { Comment } from "../types/content";
import type { Document } from "../types/document";
import { MAX_HEX_ID_EXCLUSIVE } from "../utils/hexId";
import { applyFolioAIEditOperations } from "./apply";
import { createFolioAIEditSnapshot } from "./snapshot";
import type {
  FolioAIBlock,
  FolioAIEditApplyMode,
  FolioAIEditApplyResult,
  FolioAIEditOperation,
  FolioAIEditSnapshot,
} from "./types";

/**
 * Standalone comment-id allocator for the reviewer. The React editor uses
 * `Date.now()`-seeded ids from `commentsHelpers`; that module is DOM-bound
 * (`EditorView`, `findBodyPmAnchors`) and unimportable here, so the reviewer
 * mints its own. Seeded once per realm and incremented so ids stay unique
 * across reviewers created within the same millisecond.
 */
let commentIdCursor = Date.now();

/**
 * Build the note-free comment thread the apply layer references by id.
 * Mirrors `commentsHelpers.createComment` (a pure object literal there) so a
 * headless `commentOnBlock` / `comment` op serialises the same `comments.xml`
 * shape the editor produces.
 */
const createReviewerComment = (text: string, author: string): Comment => ({
  id: commentIdCursor++,
  author,
  date: new Date().toISOString(),
  content: [
    {
      type: "paragraph",
      formatting: {},
      content: [{ type: "run", formatting: {}, content: [{ type: "text", text }] }],
    },
  ],
});

/** Deterministic 8-char uppercase hex id (FNV-1a over `seed`), `< 0x7FFFFFFF`. */
const deterministicHexId = (seed: string): string => {
  let hash = 2_166_136_261;
  for (const character of seed) {
    hash = Math.imul(hash ^ (character.codePointAt(0) ?? 0), 16_777_619) >>> 0;
  }
  return (hash % MAX_HEX_ID_EXCLUSIVE).toString(16).toUpperCase().padStart(8, "0");
};

/**
 * Assign a stable `w14:paraId` to every body paragraph that lacks one (or
 * duplicates an earlier id), derived deterministically from the paragraph's
 * text plus its document ordinal. Re-parsing the SAME bytes mints the SAME ids,
 * so the block ids a snapshot exposes are reproducible across
 * {@link FolioDocxReviewer.fromBuffer} calls — the documented "snapshot on one
 * parse, apply on another" flow resolves instead of skipping every op as a
 * stale anchor. Word-authored paraIds are preserved; only gaps and collisions
 * get a fresh id.
 *
 * The shared `ParaIdAllocatorExtension` mints RANDOM ids (correct for freshly
 * typed paragraphs in the live editor); this load-time pass is deterministic so
 * a paraId-less corpus document anchors reproducibly. The transaction is marked
 * ignore-tracked so the change baseline stays clean.
 */
const ensureDeterministicParaIdsInState = (state: EditorState): EditorState => {
  const seen = new Set<string>();
  const updates: { pos: number; attrs: Record<string, unknown> }[] = [];
  let ordinal = 0;

  state.doc.descendants((node, pos) => {
    if (node.type.name !== "paragraph") {
      return undefined;
    }
    ordinal += 1;
    const existing = node.attrs["paraId"];
    if (typeof existing === "string" && existing.length > 0 && !seen.has(existing)) {
      seen.add(existing);
      return false;
    }
    let paraId = deterministicHexId(`${node.textContent}:${ordinal}`);
    for (let salt = 1; seen.has(paraId); salt++) {
      paraId = deterministicHexId(`${node.textContent}:${ordinal}:${salt}`);
    }
    seen.add(paraId);
    updates.push({ pos, attrs: { ...node.attrs, paraId } });
    return false;
  });

  if (updates.length === 0) {
    return state;
  }

  const tr = state.tr;
  for (const update of updates) {
    tr.setNodeMarkup(update.pos, undefined, update.attrs);
  }
  ignoreTrackedChanges(tr);
  tr.setMeta("addToHistory", false);
  return state.apply(tr);
};

/** Options for {@link FolioDocxReviewer.fromBuffer}. */
export type FolioDocxReviewerOptions = {
  /** Default author for tracked changes and comments. (default: `"AI"`) */
  author?: string;
};

/** Options for {@link FolioDocxReviewer.applyOperations}. */
export type FolioApplyOperationsOptions = {
  /** `"tracked-changes"` (default) produces ins/del redlines; `"direct"` edits in place. */
  mode?: FolioAIEditApplyMode;
  /**
   * The snapshot the `operations`' block ids were built against. Omit to
   * snapshot the reviewer's current state — correct for the common flow of
   * `snapshot()` -> build ops -> `applyOperations()` on the same reviewer.
   */
  snapshot?: FolioAIEditSnapshot;
};

export type FolioReviewChangeKind = "insertion" | "deletion";

/** A tracked change (insertion or deletion) discovered in the document body. */
export type FolioReviewChange = {
  /**
   * The tracked-change revision id — the OOXML `w:id` carried on the
   * insertion / deletion mark. Pass it (or the whole change) to
   * {@link FolioDocxReviewer.acceptChange} / `rejectChange`. A replace produces
   * two changes (a deletion side and an insertion side) with distinct ids.
   */
  id: number;
  type: FolioReviewChangeKind;
  author: string;
  /** ISO date the change was authored, or `null` when the source omitted it. */
  date: string | null;
  /** Inserted text for insertions, removed text for deletions. */
  text: string;
  /**
   * Stable id of the containing body block (Word `w14:paraId` or `seq-NNNN`),
   * matching {@link FolioDocxReviewer.getContent}. `null` when the change sits
   * in a block with no surviving visible text, which has no snapshot block.
   */
  blockId: string | null;
};

/** Filter for {@link FolioDocxReviewer.getChanges}. */
export type FolioReviewChangeFilter = {
  author?: string;
  type?: FolioReviewChangeKind;
};

export type FolioReviewCommentReply = {
  id: number;
  author: string;
  date: string | null;
  text: string;
};

/** A comment thread discovered in the document. */
export type FolioReviewComment = {
  id: number;
  author: string;
  date: string | null;
  /** The comment body text, its paragraphs joined by newlines. */
  text: string;
  /** The document text the comment is anchored to, or `""` when unanchored. */
  anchoredText: string;
  /** Stable id of the anchored body block, or `null` when the anchor is absent. */
  blockId: string | null;
  replies: FolioReviewCommentReply[];
  /** Whether the comment is marked resolved / done. */
  done: boolean;
};

/** Filter for {@link FolioDocxReviewer.getComments}. */
export type FolioReviewCommentFilter = {
  author?: string;
  done?: boolean;
};

/**
 * One LLM-ready line for a block: `[<blockId>] text`, with an `(h<level>)` tag
 * for headings and the list marker for list items, so a model can copy the
 * block id straight back into an operation.
 */
const formatBlockForLLM = (block: FolioAIBlock): string => {
  const label = `[${block.id}]`;
  if (block.kind === "heading") {
    return `${label} (h${headingLevel(block)}) ${block.text}`;
  }
  if (block.kind === "listItem") {
    return `${label} ${block.displayLabel ?? "•"} ${block.text}`;
  }
  return `${label} ${block.text}`;
};

const headingLevel = (block: FolioAIBlock): number => {
  const digits = /(\d+)/u.exec(block.styleId ?? block.displayLabel ?? "")?.[1];
  const level = digits ? Number.parseInt(digits, 10) : 1;
  return level >= 1 && level <= 9 ? level : 1;
};

const revisionIdOf = (target: FolioReviewChange | number): number =>
  typeof target === "number" ? target : target.id;

const commentPlainText = (comment: Comment): string =>
  // A comment parsed from a malformed package can omit `content`; the model
  // types it as required, so guard at this untrusted boundary.
  (comment.content ?? []).map(paragraphPlainText).join("\n");

const paragraphPlainText = (paragraph: Comment["content"][number]): string => {
  const parts: string[] = [];
  for (const item of paragraph.content ?? []) {
    if (item.type !== "run") {
      continue;
    }
    for (const runItem of item.content ?? []) {
      if (runItem.type === "text") {
        parts.push(runItem.text);
      }
    }
  }
  return parts.join("");
};

/**
 * Headless `.docx` reviewer. Parse a buffer, read blocks, apply
 * `FolioAIEditOperation`s against the document model, and write the reviewed
 * `.docx` back out — no editor instance, no DOM.
 *
 * @example
 * ```ts
 * const reviewer = await FolioDocxReviewer.fromBuffer(buffer, { author: "AI" });
 * const { blocks } = reviewer.snapshot();
 * reviewer.applyOperations([
 *   { id: "1", type: "replaceInBlock", blockId: blocks[0].id, find: "$50k", replace: "$500k" },
 * ]);
 * const reviewed = await reviewer.toBuffer();
 * ```
 */
export class FolioDocxReviewer {
  /** Default author for tracked changes and comments. */
  readonly author: string;
  private readonly baseDocument: Document;
  private readonly originalBuffer: ArrayBuffer;
  private state: EditorState;
  private readonly createdComments: Comment[] = [];

  private constructor(args: {
    baseDocument: Document;
    originalBuffer: ArrayBuffer;
    state: EditorState;
    author: string;
  }) {
    this.baseDocument = args.baseDocument;
    this.originalBuffer = args.originalBuffer;
    this.state = args.state;
    this.author = args.author;
  }

  /** Parse a `.docx` buffer into a reviewer. */
  static async fromBuffer(
    buffer: ArrayBuffer,
    options: FolioDocxReviewerOptions = {},
  ): Promise<FolioDocxReviewer> {
    const baseDocument = await parseDocx(buffer, {
      detectVariables: false,
      preloadFonts: false,
    });
    // Same plugin set the live editor mounts: the change tracker feeds the
    // selective-save key set, and the paraId allocator hands freshly inserted
    // paragraphs a stable `w14:paraId`. No plugin `view()` runs headlessly, so
    // the DOM-facing halves stay dormant.
    const plugins: Plugin[] = singletonManager.getPlugins();
    // Allocate paraIds up front (the editor does this on load) so every block
    // anchors on a stable id and the selective-save path can key changed
    // paragraphs by paraId. Deterministic (not random) allocation so a
    // paraId-less document yields the SAME block ids on every parse: ops built
    // from one parse's snapshot then resolve against another parse of the same
    // bytes instead of skipping as stale anchors.
    const state = ensureDeterministicParaIdsInState(
      EditorState.create({ schema, doc: toProseDoc(baseDocument), plugins }),
    );
    return new FolioDocxReviewer({
      baseDocument,
      originalBuffer: buffer,
      state,
      author: options.author ?? "AI",
    });
  }

  /**
   * Snapshot the current document into AI-facing blocks (id, kind, text) plus
   * the anchor map the apply layer resolves operations against.
   */
  snapshot(): FolioAIEditSnapshot {
    return createFolioAIEditSnapshot(this.state.doc);
  }

  /**
   * Apply operations against the current state. Reuses the live-editor applier
   * verbatim via a headless `{ state, dispatch }` seam; the resulting state
   * (including any comments) is retained for {@link toBuffer}.
   */
  applyOperations(
    operations: FolioAIEditOperation[],
    options: FolioApplyOperationsOptions = {},
  ): FolioAIEditApplyResult {
    const snapshot = options.snapshot ?? this.snapshot();
    const view = {
      state: this.state,
      dispatch: (transaction: Transaction) => {
        view.state = view.state.apply(transaction);
      },
    };

    const result = applyFolioAIEditOperations({
      view,
      snapshot,
      operations,
      mode: options.mode ?? "tracked-changes",
      author: this.author,
      createCommentId: (text) => {
        const comment = createReviewerComment(text, this.author);
        this.createdComments.push(comment);
        return comment.id;
      },
    });

    this.state = view.state;
    return result;
  }

  /**
   * The body blocks with their stable ids, in document order — the same
   * `FolioAIBlock` shape {@link snapshot} builds, reused verbatim so the ids
   * feed straight back into {@link applyOperations} and accept / reject.
   */
  getContent(): FolioAIBlock[] {
    return this.snapshot().blocks;
  }

  /**
   * The body as LLM-ready plain text: one line per block, each prefixed with
   * its stable block id (`[<blockId>] text`). Copyable verbatim into a prompt
   * without JSON quote-escaping.
   */
  getContentAsText(): string {
    return this.getContent().map(formatBlockForLLM).join("\n");
  }

  /**
   * The tracked changes (insertions and deletions) present in the body, read
   * from the same `insertion` / `deletion` marks the editor renders. Each
   * carries the revision id {@link acceptChange} / {@link rejectChange} resolve
   * against. Runs of one revision within a block fold into a single entry.
   */
  getChanges(filter?: FolioReviewChangeFilter): FolioReviewChange[] {
    const insertionType = this.state.schema.marks["insertion"];
    const deletionType = this.state.schema.marks["deletion"];
    const blockStarts = this.blockStartIds();
    const grouped = new Map<string, FolioReviewChange>();
    let currentBlockId: string | null = null;

    this.state.doc.descendants((node, pos) => {
      if (node.isTextblock) {
        currentBlockId = blockStarts.get(pos) ?? null;
        return true;
      }
      if (!node.isInline || node.text === undefined) {
        return undefined;
      }
      const text = node.text;
      for (const mark of node.marks) {
        if (typeof mark.attrs["revisionId"] !== "number") {
          continue;
        }
        let kind: FolioReviewChangeKind;
        if (mark.type === insertionType) {
          kind = "insertion";
        } else if (mark.type === deletionType) {
          kind = "deletion";
        } else {
          continue;
        }
        const revisionId = mark.attrs["revisionId"];
        const key = `${currentBlockId ?? ""}:${kind}:${revisionId}`;
        const existing = grouped.get(key);
        if (existing) {
          existing.text += text;
          continue;
        }
        const author = mark.attrs["author"];
        const date = mark.attrs["date"];
        grouped.set(key, {
          id: revisionId,
          type: kind,
          author: typeof author === "string" ? author : "",
          date: typeof date === "string" ? date : null,
          text,
          blockId: currentBlockId,
        });
      }
      return undefined;
    });

    const changes = [...grouped.values()];
    if (!filter) {
      return changes;
    }
    return changes.filter(
      (change) =>
        (filter.author === undefined || change.author === filter.author) &&
        (filter.type === undefined || change.type === filter.type),
    );
  }

  /**
   * The comment threads in the document — comments parsed from the package plus
   * any the reviewer authored — with their anchored text and containing block
   * id. Only run text is read from a comment body; richer comment content
   * (nested tracked changes, hyperlinks) is out of scope.
   */
  getComments(filter?: FolioReviewCommentFilter): FolioReviewComment[] {
    const definitions = [
      ...(this.baseDocument.package.document.comments ?? []),
      ...this.createdComments,
    ];
    if (definitions.length === 0) {
      return [];
    }

    const anchors = this.commentAnchors();
    const repliesByParent = new Map<number, Comment[]>();
    const topLevel: Comment[] = [];
    for (const comment of definitions) {
      if (comment.parentId === undefined) {
        topLevel.push(comment);
        continue;
      }
      const siblings = repliesByParent.get(comment.parentId) ?? [];
      siblings.push(comment);
      repliesByParent.set(comment.parentId, siblings);
    }

    const threads = topLevel.map<FolioReviewComment>((comment) => {
      const anchor = anchors.get(comment.id);
      return {
        id: comment.id,
        author: comment.author,
        date: comment.date ?? null,
        text: commentPlainText(comment),
        anchoredText: anchor?.text ?? "",
        blockId: anchor?.blockId ?? null,
        replies: (repliesByParent.get(comment.id) ?? []).map((reply) => ({
          id: reply.id,
          author: reply.author,
          date: reply.date ?? null,
          text: commentPlainText(reply),
        })),
        done: comment.done ?? false,
      };
    });

    if (!filter) {
      return threads;
    }
    return threads.filter(
      (thread) =>
        (filter.author === undefined || thread.author === filter.author) &&
        (filter.done === undefined || thread.done === filter.done),
    );
  }

  /**
   * Accept an existing tracked change, keeping its text and dropping the
   * redline. Pass a {@link FolioReviewChange} from {@link getChanges} or its
   * revision id. Reuses the editor's own accept command headlessly, so the
   * resolved state persists on {@link toBuffer}. Returns `false` when the
   * revision is no longer present (already resolved, or never existed).
   */
  acceptChange(target: FolioReviewChange | number): boolean {
    return this.runCommand(acceptAIEditRevision(revisionIdOf(target)));
  }

  /**
   * Reject an existing tracked change: an insertion's text is removed, a
   * deletion's text is restored. See {@link acceptChange} for targeting.
   */
  rejectChange(target: FolioReviewChange | number): boolean {
    return this.runCommand(rejectAIEditRevision(revisionIdOf(target)));
  }

  /**
   * Accept every tracked change in the body. Returns the number of changes
   * present before the sweep. Runs unconditionally: the underlying command also
   * resolves paragraph-level changes (`w:pPrChange`, paragraph-boundary
   * ins/del) that {@link getChanges} does not enumerate. Note-body changes are
   * out of scope.
   */
  acceptAll(): number {
    const count = this.countTrackedChanges();
    this.runCommand(acceptAllChanges());
    return count;
  }

  /** Reject every tracked change in the body. See {@link acceptAll}. */
  rejectAll(): number {
    const count = this.countTrackedChanges();
    this.runCommand(rejectAllChanges());
    return count;
  }

  /** The current document model with edits merged back in. */
  toDocument(): Document {
    const document = updateDocumentContent(this.baseDocument, this.state.doc);
    if (this.createdComments.length > 0) {
      document.package.document.comments = [
        ...(document.package.document.comments ?? []),
        ...this.createdComments,
      ];
    }
    return document;
  }

  /**
   * Serialise the reviewed document to a new `.docx` buffer. Tries a selective
   * patch first (only changed paragraphs are rewritten; every other byte of the
   * original package is preserved), falling back to a full repack for
   * structural edits — the same two-tier path the editor's save uses.
   */
  async toBuffer(): Promise<ArrayBuffer> {
    const document = this.toDocument();
    const selective = await this.trySelectiveSave(document);
    if (selective) {
      return selective;
    }
    return repackDocx({ ...document, originalBuffer: this.originalBuffer });
  }

  /**
   * Attempt the selective patch, treating a throw the same as a decline. Odd
   * source XML can make the paragraph diff throw rather than return `null`; a
   * full repack is the correct lossless fallback in both cases, so the fallback
   * is the graceful handling — no separate error surface is needed here.
   */
  private async trySelectiveSave(document: Document): Promise<ArrayBuffer | null> {
    try {
      return await attemptSelectiveSave(document, this.originalBuffer, {
        changedParaIds: getChangedParagraphIds(this.state),
        structuralChange: hasStructuralChanges(this.state),
        hasUntrackedChanges: hasUntrackedChanges(this.state),
      });
    } catch {
      return null;
    }
  }

  /**
   * Count every tracked change an accept-all / reject-all sweep resolves: the
   * inline insertion / deletion groups {@link getChanges} enumerates, plus
   * paragraph-level changes (`pPrMark`, `_propertyChanges`) that live on
   * paragraph attrs rather than inline marks.
   */
  private countTrackedChanges(): number {
    let count = this.getChanges().length;
    this.state.doc.descendants((node) => {
      if (node.type.name !== "paragraph") {
        return undefined;
      }
      const pPrMark = node.attrs["pPrMark"];
      if (pPrMark !== null && pPrMark !== undefined) {
        count += 1;
      }
      const propertyChanges = node.attrs["_propertyChanges"];
      if (Array.isArray(propertyChanges) && propertyChanges.length > 0) {
        count += 1;
      }
      return false;
    });
    return count;
  }

  /** Map each snapshot block's start position to its stable id. */
  private blockStartIds(): Map<number, string> {
    const starts = new Map<number, string>();
    for (const anchor of Object.values(this.snapshot().anchors)) {
      starts.set(anchor.from, anchor.id);
    }
    return starts;
  }

  /** Map each anchored comment id to its anchored text and containing block id. */
  private commentAnchors(): Map<number, { text: string; blockId: string | null }> {
    const commentType = this.state.schema.marks["comment"];
    const anchors = new Map<number, { text: string; blockId: string | null }>();
    if (!commentType) {
      return anchors;
    }
    const blockStarts = this.blockStartIds();
    let currentBlockId: string | null = null;

    this.state.doc.descendants((node, pos) => {
      if (node.isTextblock) {
        currentBlockId = blockStarts.get(pos) ?? null;
        return true;
      }
      if (!node.isInline || node.text === undefined) {
        return undefined;
      }
      const text = node.text;
      for (const mark of node.marks) {
        if (mark.type !== commentType || typeof mark.attrs["commentId"] !== "number") {
          continue;
        }
        const commentId = mark.attrs["commentId"];
        const existing = anchors.get(commentId);
        if (!existing) {
          anchors.set(commentId, { text, blockId: currentBlockId });
          continue;
        }
        existing.text += text;
        existing.blockId ??= currentBlockId;
      }
      return undefined;
    });

    return anchors;
  }

  /**
   * Drive a ProseMirror command against the reviewer's headless state via the
   * same `{ state, dispatch }` seam {@link applyOperations} uses, retaining the
   * resulting state for {@link toBuffer}.
   */
  private runCommand(command: Command): boolean {
    const view = {
      state: this.state,
      dispatch: (transaction: Transaction) => {
        view.state = view.state.apply(transaction);
      },
    };
    const handled = command(view.state, view.dispatch);
    this.state = view.state;
    return handled;
  }
}

/** Options for {@link applyFolioAIEditsToBuffer}. */
export type ApplyFolioAIEditsToBufferOptions = {
  /** Default author for tracked changes and comments. (default: `"AI"`) */
  author?: string;
  /** `"tracked-changes"` (default) produces ins/del redlines; `"direct"` edits in place. */
  mode?: FolioAIEditApplyMode;
  /**
   * The snapshot the `operations`' block ids were built against. Omit to
   * snapshot the freshly parsed buffer. Pass it when the ids were derived in a
   * separate process (e.g. server-side `@stll/folio-core/server` block ids) so
   * anchor resolution matches exactly.
   */
  snapshot?: FolioAIEditSnapshot;
};

/** Result of {@link applyFolioAIEditsToBuffer}. */
export type ApplyFolioAIEditsToBufferResult = FolioAIEditApplyResult & {
  /** The reviewed `.docx` as a new buffer. */
  buffer: ArrayBuffer;
};

/**
 * One-shot headless review: parse `buffer`, apply `operations`, and return the
 * reviewed `.docx` buffer alongside the applied / skipped breakdown. Convenience
 * wrapper over {@link FolioDocxReviewer} for callers that already hold the
 * operations to run.
 */
export const applyFolioAIEditsToBuffer = async (
  buffer: ArrayBuffer,
  operations: FolioAIEditOperation[],
  options: ApplyFolioAIEditsToBufferOptions = {},
): Promise<ApplyFolioAIEditsToBufferResult> => {
  const reviewer = await FolioDocxReviewer.fromBuffer(buffer, {
    ...(options.author !== undefined && { author: options.author }),
  });
  const { applied, skipped } = reviewer.applyOperations(operations, {
    ...(options.mode !== undefined && { mode: options.mode }),
    ...(options.snapshot !== undefined && { snapshot: options.snapshot }),
  });
  const reviewed = await reviewer.toBuffer();
  return { buffer: reviewed, applied, skipped };
};
