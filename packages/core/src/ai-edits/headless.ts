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
 * Scope: main, header, and footer blocks. Footnote and endnote bodies build on
 * the note serialization path separately.
 */

import { TaggedError } from "better-result";
import { EditorState } from "prosemirror-state";
import type { Command } from "prosemirror-state";
import type { Plugin } from "prosemirror-state";
import type { Transaction } from "prosemirror-state";

import { createReply } from "../docx/replyToComment";
import { attemptSelectiveSave } from "../docx/selectiveSave";
import { getHeaderFooterText } from "../docx/headerFooterParser";
import {
  getEndnoteText,
  getFootnoteText,
  isSeparatorEndnote,
  isSeparatorFootnote,
} from "../docx/footnoteParser";
import { parseDocx } from "../docx/parser";
import { repackDocx } from "../docx/rezip";
import {
  acceptAIEditRevision,
  acceptAllChanges,
  rejectAIEditRevision,
  rejectAllChanges,
} from "../prosemirror/commands/comments";
import { proseDocToBlocks, updateDocumentContent } from "../prosemirror/conversion/fromProseDoc";
import { headerFooterToProseDoc, toProseDoc } from "../prosemirror/conversion/toProseDoc";
import {
  getChangedParagraphIds,
  hasStructuralChanges,
  hasUntrackedChanges,
  ignoreTrackedChanges,
} from "../prosemirror/extensions/features/ParagraphChangeTrackerExtension";
import { schema, singletonManager } from "../prosemirror/schema";
import type { Comment } from "../types/content";
import type { Document, HeaderFooter } from "../types/document";
import { deterministicHexId } from "../utils/hexId";
import {
  applyFolioDocumentOperations,
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperationBatch,
  type FolioDocumentOperationResult,
  type FolioDocumentOperationUndoHandle,
  type FolioDocumentOperationUndoResult,
} from "../document-operations";
import { buildAnnotatedBlockText } from "./clean-text";
import {
  getCommentAnchorsFromDoc,
  getTrackedChangesFromDoc,
  type FolioReviewChange,
  type FolioReviewChangeKind,
} from "./read";
import { createFolioAIEditSnapshot, normalizeFolioAIBlockText } from "./snapshot";
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
let undoHandleCursor = Date.now();

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

type FolioDocumentOperationUndoEntry = {
  undoHandle: FolioDocumentOperationUndoHandle;
  story: FolioEditableDocumentStoryHandle;
  beforeState: EditorState;
  afterState: EditorState;
  createdCommentsLengthBefore: number;
  createdCommentsLengthAfter: number;
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
  /** Password for Agile-encrypted .docx files (Office 2010+). */
  password?: string | undefined;
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

/** Options for {@link FolioDocxReviewer.applyDocumentOperations}. */
export type FolioApplyDocumentOperationsOptions = Omit<FolioApplyOperationsOptions, "mode">;

/** Options for {@link FolioDocxReviewer.getContentAsText}. */
export type FolioGetContentAsTextOptions = {
  /**
   * Render tracked changes and comment anchors inline as `<ins>` / `<del>` /
   * `<comment>` tags instead of the default flattened, post-tracked-changes
   * text. (default: `false`)
   */
  annotated?: boolean;
};

export type FolioDocumentStoryHandle =
  | { type: "main" }
  | { type: "header" | "footer"; relationshipId: string }
  | { type: "footnote" | "endnote"; noteId: number };

export type FolioEditableDocumentStoryHandle = Exclude<
  FolioDocumentStoryHandle,
  { type: "footnote" | "endnote" }
>;

export type FolioDocumentStory = {
  handle: FolioDocumentStoryHandle;
  text: string;
};

export class FolioDocumentStoryNotFoundError extends TaggedError(
  "FolioDocumentStoryNotFoundError",
)<{
  message: string;
  story: FolioEditableDocumentStoryHandle;
}>() {}

export type FolioApplyDocumentOperationsToStoryOptions = FolioApplyDocumentOperationsOptions & {
  story: FolioEditableDocumentStoryHandle;
  batch: FolioDocumentOperationBatch;
};

type FolioHeaderFooterStoryHandle = Extract<
  FolioEditableDocumentStoryHandle,
  { type: "header" | "footer" }
>;

type FolioHeaderFooterStoryState = {
  handle: FolioHeaderFooterStoryHandle;
  initialState: EditorState;
  state: EditorState;
};

type ApplyDocumentOperationsInternalOptions = {
  story: FolioEditableDocumentStoryHandle;
  batch: FolioDocumentOperationBatch;
  snapshot?: FolioAIEditSnapshot;
  createUndoEntry: boolean;
};

const MAIN_STORY = Object.freeze({ type: "main" } as const);

const headerFooterStoryKey = ({ type, relationshipId }: FolioHeaderFooterStoryHandle): string =>
  `${type}:${relationshipId}`;

// The change shape and its pure reader now live in `./read` so a live editor
// can produce the same `FolioReviewChange[]` from its own doc; the reviewer
// delegates to `getTrackedChangesFromDoc` and re-exports the types here so its
// public surface is unchanged.
export type { FolioReviewChange, FolioReviewChangeKind } from "./read";

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

/** Input for {@link FolioDocxReviewer.replyTo}. */
export type FolioReviewReplyInput = {
  /** Reply body text. */
  text: string;
  /** Reply author; defaults to the reviewer's author. */
  author?: string;
  initials?: string;
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
const formatBlockLine = (block: FolioAIBlock, text: string): string => {
  const label = `[${block.id}]`;
  if (block.kind === "heading") {
    return `${label} (h${headingLevel(block)}) ${text}`;
  }
  if (block.kind === "listItem") {
    return `${label} ${block.displayLabel ?? "•"} ${text}`;
  }
  return `${label} ${text}`;
};

const formatBlockForLLM = (block: FolioAIBlock): string => formatBlockLine(block, block.text);

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
  private readonly headerFooterStoryStates = new Map<string, FolioHeaderFooterStoryState>();
  private readonly createdComments: Comment[] = [];
  private readonly documentOperationUndoEntries: FolioDocumentOperationUndoEntry[] = [];
  /**
   * Resolved-state overrides recorded by {@link resolveComment}, keyed by
   * comment id. Applied on read ({@link getComments}) and on write
   * ({@link toDocument}) rather than mutating the parsed `Comment` objects in
   * place, matching how the parser/serializer treat `Comment` as immutable
   * (`commentParser.ts` replaces array slots via spread; the serializer
   * types its inputs `readonly Comment[]`).
   */
  private readonly resolvedOverrides = new Map<number, boolean>();

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
      password: options.password,
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
      originalBuffer: baseDocument.originalBuffer ?? buffer,
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

  /** Snapshot one editable story into stable, operation-ready blocks. */
  snapshotStory(story: FolioEditableDocumentStoryHandle): FolioAIEditSnapshot | null {
    const state = this.getEditableStoryState(story);
    return state ? createFolioAIEditSnapshot(state.doc) : null;
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
    const { applied, skipped } = this.applyDocumentOperationsInternal({
      story: MAIN_STORY,
      batch: {
        version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
        operations,
        mode: options.mode ?? "tracked-changes",
      },
      ...(options.snapshot !== undefined && { snapshot: options.snapshot }),
      createUndoEntry: false,
    });
    return { applied, skipped };
  }

  /**
   * Apply a versioned document-operation batch against the current state.
   * This is the contract entry point for serialized callers; the legacy
   * {@link applyOperations} method delegates here so both APIs keep identical
   * edit semantics.
   */
  applyDocumentOperations(
    batch: FolioDocumentOperationBatch,
    options: FolioApplyDocumentOperationsOptions = {},
  ): FolioDocumentOperationResult {
    return this.applyDocumentOperationsInternal({
      story: MAIN_STORY,
      batch,
      ...(options.snapshot !== undefined && { snapshot: options.snapshot }),
      createUndoEntry: true,
    });
  }

  /** Apply a versioned operation batch to the main story, a header, or a footer. */
  applyDocumentOperationsToStory({
    story,
    batch,
    snapshot,
  }: FolioApplyDocumentOperationsToStoryOptions): FolioDocumentOperationResult {
    return this.applyDocumentOperationsInternal({
      story,
      batch,
      ...(snapshot !== undefined && { snapshot }),
      createUndoEntry: true,
    });
  }

  private applyDocumentOperationsInternal({
    story,
    batch,
    snapshot,
    createUndoEntry,
  }: ApplyDocumentOperationsInternalOptions): FolioDocumentOperationResult {
    const beforeState = this.requireEditableStoryState(story);
    const createdCommentsLengthBefore = this.createdComments.length;
    const view = {
      state: beforeState,
      dispatch: (transaction: Transaction) => {
        view.state = view.state.apply(transaction);
      },
    };

    const result = applyFolioDocumentOperations({
      view,
      snapshot: snapshot ?? createFolioAIEditSnapshot(beforeState.doc),
      batch,
      story: story.type === "main" ? "main" : story,
      author: this.author,
      createCommentId: (text) => {
        const comment = createReviewerComment(text, this.author);
        this.createdComments.push(comment);
        return comment.id;
      },
      ...(createUndoEntry && {
        createUndoHandle: () => ({
          type: "documentOperationUndo",
          id: `headless-${String(undoHandleCursor++)}`,
        }),
      }),
    });

    this.setEditableStoryState(story, view.state);
    if (result.undoHandle !== null) {
      this.documentOperationUndoEntries.push({
        undoHandle: result.undoHandle,
        story,
        beforeState,
        afterState: view.state,
        createdCommentsLengthBefore,
        createdCommentsLengthAfter: this.createdComments.length,
      });
    }
    return result;
  }

  /** Undo the latest unchanged document-operation batch and its created comments. */
  undoDocumentOperations(
    undoHandle: FolioDocumentOperationUndoHandle,
  ): FolioDocumentOperationUndoResult {
    const entryIndex = this.documentOperationUndoEntries.findIndex(
      (entry) => entry.undoHandle.type === undoHandle.type && entry.undoHandle.id === undoHandle.id,
    );
    if (entryIndex === -1) {
      return { status: "rejected", undoHandle, reason: "unknownHandle" };
    }
    if (entryIndex !== this.documentOperationUndoEntries.length - 1) {
      return { status: "rejected", undoHandle, reason: "notLatest" };
    }

    const entry = this.documentOperationUndoEntries.at(-1);
    if (!entry) {
      return { status: "rejected", undoHandle, reason: "unknownHandle" };
    }
    if (
      this.requireEditableStoryState(entry.story) !== entry.afterState ||
      this.createdComments.length !== entry.createdCommentsLengthAfter
    ) {
      return { status: "rejected", undoHandle, reason: "documentChanged" };
    }

    this.setEditableStoryState(entry.story, entry.beforeState);
    this.createdComments.length = entry.createdCommentsLengthBefore;
    this.documentOperationUndoEntries.pop();
    return { status: "undone", undoHandle };
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
   *
   * With `{ annotated: true }` each block's text is rendered redline-aware:
   * tracked insertions/deletions and comment anchors appear inline as
   * `<ins>` / `<del>` / `<comment>` tags (see {@link buildAnnotatedBlockText})
   * for prompt-embedding parity with a live editor's redline view. The default
   * (clean) output flattens tracked changes and is unchanged.
   */
  getContentAsText(options: FolioGetContentAsTextOptions = {}): string {
    if (!options.annotated) {
      return this.getContent().map(formatBlockForLLM).join("\n");
    }
    const snapshot = this.snapshot();
    const startById = new Map<string, number>();
    for (const anchor of Object.values(snapshot.anchors)) {
      startById.set(anchor.id, anchor.from);
    }
    return snapshot.blocks
      .map((block) => {
        const from = startById.get(block.id);
        const node = from === undefined ? null : this.state.doc.nodeAt(from);
        const text = node ? buildAnnotatedBlockText(node) : block.text;
        return formatBlockLine(block, text);
      })
      .join("\n");
  }

  /**
   * Header / footer and footnote / endnote text as labeled, LLM-ready lines,
   * one per non-empty part: `[header default] …`, `[footer default] …`,
   * `[footnote #N] …`, `[endnote #N] …`. Header and footer lines reflect
   * in-memory edits; note bodies reflect the parsed source. Empty parts and
   * separator notes are omitted.
   */
  getNotesAsText(): string {
    const pkg = this.baseDocument.package;
    const lines: string[] = [];

    const pushHeaderFooter = (
      map: Map<string, HeaderFooter> | undefined,
      label: "header" | "footer",
    ) => {
      if (!map) {
        return;
      }
      for (const [relationshipId, hf] of map) {
        const handle = { type: label, relationshipId } as const;
        const text = this.getHeaderFooterStoryText(handle, hf);
        if (text.length > 0) {
          lines.push(`[${label} ${hf.hdrFtrType}] ${text}`);
        }
      }
    };

    pushHeaderFooter(pkg.headers, "header");
    pushHeaderFooter(pkg.footers, "footer");

    for (const footnote of pkg.footnotes ?? []) {
      if (isSeparatorFootnote(footnote)) {
        continue;
      }
      const text = normalizeFolioAIBlockText(getFootnoteText(footnote));
      if (text.length > 0) {
        lines.push(`[footnote #${footnote.id}] ${text}`);
      }
    }
    for (const endnote of pkg.endnotes ?? []) {
      if (isSeparatorEndnote(endnote)) {
        continue;
      }
      const text = normalizeFolioAIBlockText(getEndnoteText(endnote));
      if (text.length > 0) {
        lines.push(`[endnote #${endnote.id}] ${text}`);
      }
    }

    return lines.join("\n");
  }

  /** Discover every readable document story through a typed, serializable handle. */
  listStories(): FolioDocumentStory[] {
    const pkg = this.baseDocument.package;
    const stories: FolioDocumentStory[] = [
      { handle: { type: "main" }, text: this.getContentAsText() },
    ];
    for (const [relationshipId, header] of pkg.headers ?? []) {
      const handle = { type: "header", relationshipId } as const;
      stories.push({
        handle,
        text: this.getHeaderFooterStoryText(handle, header),
      });
    }
    for (const [relationshipId, footer] of pkg.footers ?? []) {
      const handle = { type: "footer", relationshipId } as const;
      stories.push({
        handle,
        text: this.getHeaderFooterStoryText(handle, footer),
      });
    }
    for (const footnote of pkg.footnotes ?? []) {
      if (!isSeparatorFootnote(footnote)) {
        stories.push({
          handle: { type: "footnote", noteId: footnote.id },
          text: normalizeFolioAIBlockText(getFootnoteText(footnote)),
        });
      }
    }
    for (const endnote of pkg.endnotes ?? []) {
      if (!isSeparatorEndnote(endnote)) {
        stories.push({
          handle: { type: "endnote", noteId: endnote.id },
          text: normalizeFolioAIBlockText(getEndnoteText(endnote)),
        });
      }
    }
    return stories;
  }

  /** Read one discovered story; returns null when its handle is no longer present. */
  readStory(handle: FolioDocumentStoryHandle): FolioDocumentStory | null {
    const stories = this.listStories();
    if (handle.type === "main") {
      return stories.at(0) ?? null;
    }
    if (handle.type === "header" || handle.type === "footer") {
      return (
        stories.find(
          ({ handle: candidate }) =>
            candidate.type === handle.type && candidate.relationshipId === handle.relationshipId,
        ) ?? null
      );
    }
    return (
      stories.find(
        ({ handle: candidate }) =>
          (candidate.type === "footnote" || candidate.type === "endnote") &&
          candidate.type === handle.type &&
          candidate.noteId === handle.noteId,
      ) ?? null
    );
  }

  /**
   * The tracked changes (insertions and deletions) present in the body, read
   * from the same `insertion` / `deletion` marks the editor renders. Each
   * carries the revision id {@link acceptChange} / {@link rejectChange} resolve
   * against. Runs of one revision within a block fold into a single entry.
   */
  getChanges(filter?: FolioReviewChangeFilter): FolioReviewChange[] {
    const changes = getTrackedChangesFromDoc(this.state.doc);
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
    const definitions = this.withResolvedOverrides([
      ...(this.baseDocument.package.document.comments ?? []),
      ...this.createdComments,
    ]);
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
   * Add a reply to a comment thread. Pass a {@link FolioReviewComment} from
   * {@link getComments} or its id; the reply threads under that comment's root
   * (Word threads are flat). On {@link toBuffer} the reply is written as a real
   * Word reply — linked via `commentsExtended.xml` and given its own
   * `commentRange` markers + reference anchored on the parent's range. Returns
   * the created reply, or `null` when the target comment is absent.
   */
  replyTo(
    target: FolioReviewComment | number,
    input: FolioReviewReplyInput,
  ): FolioReviewCommentReply | null {
    const parentId = typeof target === "number" ? target : target.id;
    const existing = [
      ...(this.baseDocument.package.document.comments ?? []),
      ...this.createdComments,
    ];
    const reply = createReply(existing, parentId, {
      author: input.author ?? this.author,
      text: input.text,
      ...(input.initials !== undefined ? { initials: input.initials } : {}),
    });
    if (!reply) {
      return null;
    }
    this.createdComments.push(reply);
    return { id: reply.id, author: reply.author, date: reply.date ?? null, text: input.text };
  }

  /**
   * Mark a comment thread resolved, or reopen a previously resolved one. Pass
   * the id from {@link getComments} / {@link FolioReviewComment.id}. Applies
   * only to the target comment id, not cascaded to its replies: Word keys the
   * resolved marker off the comment's own `w15:commentEx` entry
   * (`commentSerializer.ts`'s `buildCommentExtendedEntries` reads only
   * `comment.done` for the id it is building an entry for), so resolving the
   * thread root is sufficient and reply entries need no `done` flag of their
   * own. On {@link toBuffer} the state is written to `commentsExtended.xml`
   * (`w15:done`) through the same channel `replyTo`-created comments use.
   * Returns `false` when no comment with that id exists.
   */
  resolveComment(commentId: string, options: { resolved?: boolean } = {}): boolean {
    const resolved = options.resolved ?? true;
    const existing = [
      ...(this.baseDocument.package.document.comments ?? []),
      ...this.createdComments,
    ];
    const target = existing.find((comment) => String(comment.id) === commentId);
    if (!target) {
      return false;
    }
    this.resolvedOverrides.set(target.id, resolved);
    return true;
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
    this.mergeEditedHeaderFooterStories(document);
    if (this.createdComments.length > 0 || this.resolvedOverrides.size > 0) {
      document.package.document.comments = this.withResolvedOverrides([
        ...(document.package.document.comments ?? []),
        ...this.createdComments,
      ]);
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

  private getEditableStoryState(story: FolioEditableDocumentStoryHandle): EditorState | null {
    if (story.type === "main") {
      return this.state;
    }
    const key = headerFooterStoryKey(story);
    const existing = this.headerFooterStoryStates.get(key);
    if (existing) {
      return existing.state;
    }
    const source = this.getHeaderFooterStory(story);
    if (!source) {
      return null;
    }
    const state = ensureDeterministicParaIdsInState(
      EditorState.create({
        schema,
        doc: headerFooterToProseDoc(source.content, {
          ...(this.baseDocument.package.styles !== undefined && {
            styles: this.baseDocument.package.styles,
          }),
          ...(this.baseDocument.package.theme !== undefined && {
            theme: this.baseDocument.package.theme,
          }),
        }),
        plugins: singletonManager.getPlugins(),
      }),
    );
    this.headerFooterStoryStates.set(key, { handle: story, initialState: state, state });
    return state;
  }

  private requireEditableStoryState(story: FolioEditableDocumentStoryHandle): EditorState {
    const state = this.getEditableStoryState(story);
    if (state) {
      return state;
    }
    throw new FolioDocumentStoryNotFoundError({
      message: `Document story ${JSON.stringify(story)} was not found.`,
      story,
    });
  }

  private setEditableStoryState(story: FolioEditableDocumentStoryHandle, state: EditorState): void {
    if (story.type === "main") {
      this.state = state;
      return;
    }
    const entry = this.headerFooterStoryStates.get(headerFooterStoryKey(story));
    if (!entry) {
      throw new FolioDocumentStoryNotFoundError({
        message: `Document story ${JSON.stringify(story)} was not found.`,
        story,
      });
    }
    entry.state = state;
  }

  private getHeaderFooterStory(story: FolioHeaderFooterStoryHandle): HeaderFooter | undefined {
    const stories =
      story.type === "header"
        ? this.baseDocument.package.headers
        : this.baseDocument.package.footers;
    return stories?.get(story.relationshipId);
  }

  private getHeaderFooterStoryText(
    story: FolioHeaderFooterStoryHandle,
    source: HeaderFooter,
  ): string {
    const state = this.headerFooterStoryStates.get(headerFooterStoryKey(story))?.state;
    return normalizeFolioAIBlockText(state?.doc.textContent ?? getHeaderFooterText(source));
  }

  private mergeEditedHeaderFooterStories(document: Document): void {
    let headers: Map<string, HeaderFooter> | undefined;
    let footers: Map<string, HeaderFooter> | undefined;
    for (const entry of this.headerFooterStoryStates.values()) {
      if (entry.state === entry.initialState) {
        continue;
      }
      const source = this.getHeaderFooterStory(entry.handle);
      if (!source) {
        continue;
      }
      const edited = { ...source, content: proseDocToBlocks(entry.state.doc) };
      if (entry.handle.type === "header") {
        headers ??= new Map(document.package.headers);
        headers.set(entry.handle.relationshipId, edited);
        continue;
      }
      footers ??= new Map(document.package.footers);
      footers.set(entry.handle.relationshipId, edited);
    }
    if (headers) {
      document.package.headers = headers;
    }
    if (footers) {
      document.package.footers = footers;
    }
  }

  /**
   * Count every tracked change an accept-all / reject-all sweep resolves: the
   * inline insertion / deletion groups {@link getChanges} enumerates, plus the
   * property-change records living on node attrs rather than inline marks —
   * paragraph-level (`pPrMark`, `_propertyChanges`, the inline sectPr's
   * `propertyChanges`) and table-level (`tblPrChange` / `trPrChange` /
   * `tcPrChange`).
   */
  private countTrackedChanges(): number {
    const hasEntries = (value: unknown): boolean => Array.isArray(value) && value.length > 0;
    let count = this.getChanges().length;
    this.state.doc.descendants((node) => {
      const typeName = node.type.name;
      if (typeName === "table" && hasEntries(node.attrs["tblPrChange"])) {
        count += 1;
      }
      if (typeName === "tableRow" && hasEntries(node.attrs["trPrChange"])) {
        count += 1;
      }
      if (
        (typeName === "tableCell" || typeName === "tableHeader") &&
        hasEntries(node.attrs["tcPrChange"])
      ) {
        count += 1;
      }
      if (typeName !== "paragraph") {
        return undefined;
      }
      const pPrMark = node.attrs["pPrMark"];
      if (pPrMark !== null && pPrMark !== undefined) {
        count += 1;
      }
      if (hasEntries(node.attrs["_propertyChanges"])) {
        count += 1;
      }
      const sectionProperties = node.attrs["_sectionProperties"] as
        | { propertyChanges?: unknown }
        | null
        | undefined;
      if (hasEntries(sectionProperties?.propertyChanges)) {
        count += 1;
      }
      return false;
    });
    return count;
  }

  /** Apply any {@link resolveComment} overrides recorded for these comments. */
  private withResolvedOverrides(comments: readonly Comment[]): Comment[] {
    if (this.resolvedOverrides.size === 0) {
      return [...comments];
    }
    return comments.map((comment) => {
      const override = this.resolvedOverrides.get(comment.id);
      return override === undefined ? comment : { ...comment, done: override };
    });
  }

  /** Map each anchored comment id to its anchored text and containing block id. */
  private commentAnchors(): Map<number, { text: string; blockId: string | null }> {
    const anchors = new Map<number, { text: string; blockId: string | null }>();
    for (const anchor of getCommentAnchorsFromDoc(this.state.doc)) {
      anchors.set(anchor.commentId, { text: anchor.quote, blockId: anchor.blockId });
    }
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
