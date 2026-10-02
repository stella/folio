import { panic, Result, TaggedError } from "better-result";
import type { Node as PMNode } from "prosemirror-model";
import { TextSelection, type EditorState, type Transaction } from "prosemirror-state";
import {
  applyDocumentOps,
  DOCUMENT_OP_TYPES,
  normalizeForOps,
  OP_STORIES,
  validateOpsDocument,
  findStoryBody,
  documentStories,
  sameStory,
  type OpStory,
  type DocumentOp,
  type TextPosition,
  type TouchedBlocks,
} from "@stll/docx-core/ops";
import { hasIllegalXmlCharacters } from "@stll/docx-core";

import { splitsGraphemeCluster, splitsSurrogatePair } from "../ai-edits/character-boundaries";
import {
  cloneDocumentWithParagraphPropertySources,
  copyDocumentParagraphPropertySourceContract,
  copyParagraphPropertySource,
} from "../docx/paragraphPropertySource";
import { markPackageChange } from "../prosemirror/extensions/features/ParagraphChangeTrackerExtension";
import {
  toProseDoc,
  headerFooterToProseDoc,
  footnoteToProseDoc,
} from "../prosemirror/conversion/toProseDoc";
import type { Document, Paragraph, StyleDefinitions, TextFormatting } from "../types/document";

export class CanonicalSessionError extends TaggedError("CanonicalSessionError")<{
  message: string;
}> {}

const refuse = (message: string) => Result.err(new CanonicalSessionError({ message }));

const operationChangesPackage = (op: DocumentOp): boolean => {
  switch (op.type) {
    case DOCUMENT_OP_TYPES.CREATE_HEADER_FOOTER:
    case DOCUMENT_OP_TYPES.REMOVE_HEADER_FOOTER:
    case DOCUMENT_OP_TYPES.ADD_NOTE:
    case DOCUMENT_OP_TYPES.REMOVE_NOTE:
    case DOCUMENT_OP_TYPES.SET_SECTION_PROPS:
    case DOCUMENT_OP_TYPES.RESTORE_STORY_PARTS:
      return true;
    default: {
      if ("story" in op) return !sameStory(op.story, OP_STORIES.MAIN);
      if ("at" in op) return !sameStory(op.at.story, OP_STORIES.MAIN);
      if ("from" in op) return !sameStory(op.from.story, OP_STORIES.MAIN);
      const unreachable: never = op;
      return unreachable;
    }
  }
};

const operationChangesBodyProjection = (op: DocumentOp): boolean => {
  switch (op.type) {
    case DOCUMENT_OP_TYPES.CREATE_HEADER_FOOTER:
    case DOCUMENT_OP_TYPES.REMOVE_HEADER_FOOTER:
    case DOCUMENT_OP_TYPES.ADD_NOTE:
    case DOCUMENT_OP_TYPES.REMOVE_NOTE:
    case DOCUMENT_OP_TYPES.SET_SECTION_PROPS:
    case DOCUMENT_OP_TYPES.RESTORE_STORY_PARTS:
      return true;
    default: {
      if ("story" in op) return sameStory(op.story, OP_STORIES.MAIN);
      if ("at" in op) return sameStory(op.at.story, OP_STORIES.MAIN);
      if ("from" in op) return sameStory(op.from.story, OP_STORIES.MAIN);
      const unreachable: never = op;
      return unreachable;
    }
  }
};

export type CanonicalSelection = { anchor: TextPosition; head: TextPosition };
export type CanonicalOrigin = "input" | "undo" | "redo";
export const CANONICAL_PROJECTION_META = "folioCanonicalProjection";

type ParagraphAddress = {
  blockId: string;
  start: number;
  text: string;
  boundaries: readonly (readonly number[])[];
  node: PMNode;
  source: Paragraph;
};

type CanonicalProjectionOptions = {
  doc: PMNode;
  paragraphs: readonly ParagraphAddress[];
  story: OpStory;
};

/** Paragraph gaps map logical UTF-16 units and zero-width note marks to PM positions. */
class CanonicalProjection {
  readonly doc: PMNode;
  readonly story: OpStory;
  private readonly paragraphs: readonly ParagraphAddress[];

  constructor({ doc, paragraphs, story }: CanonicalProjectionOptions) {
    this.story = story;
    this.doc = doc;
    this.paragraphs = paragraphs;
  }

  addressAt(position: number): Result<TextPosition, CanonicalSessionError> {
    if (!Number.isInteger(position)) return refuse("The input position is not an integer.");
    const paragraph = this.paragraphs.find(
      ({ start, node }) => position >= start && position <= start + node.content.size,
    );
    if (paragraph === undefined) return refuse("The input is outside a plain paragraph.");
    const relative = position - paragraph.start;
    const offset = paragraph.boundaries.findIndex((gaps) => gaps.includes(relative));
    if (offset < 0) return refuse("The input splits a note reference.");
    if (splitsSurrogatePair(paragraph.text, offset)) {
      return refuse("The input would split a surrogate pair.");
    }
    const gaps = paragraph.boundaries.at(offset) ?? panic("Missing canonical boundary");
    return Result.ok({
      story: this.story,
      blockId: paragraph.blockId,
      offset,
      ...(gaps.length > 1 ? { zeroWidthBefore: gaps.indexOf(relative) } : {}),
    });
  }

  positionAt(address: TextPosition): Result<number, CanonicalSessionError> {
    const paragraph = this.paragraphs.find(({ blockId }) => blockId === address.blockId);
    if (
      !sameStory(address.story, this.story) ||
      paragraph === undefined ||
      !Number.isInteger(address.offset) ||
      address.offset < 0 ||
      address.offset > paragraph.text.length ||
      splitsSurrogatePair(paragraph.text, address.offset)
    ) {
      return refuse("The canonical selection is outside a plain paragraph.");
    }
    const gaps = paragraph.boundaries.at(address.offset) ?? panic("Missing canonical boundary");
    const gapIndex = address.zeroWidthBefore ?? gaps.length - 1;
    if (!Number.isInteger(gapIndex) || gapIndex < 0 || gapIndex >= gaps.length)
      return refuse("The canonical selection is outside a paragraph gap.");
    return Result.ok(paragraph.start + (gaps.at(gapIndex) ?? panic("Missing canonical gap")));
  }

  selectionAt(state: EditorState): Result<CanonicalSelection, CanonicalSessionError> {
    if (!(state.selection instanceof TextSelection)) {
      return refuse("Canonical input requires a text selection.");
    }
    const anchor = this.addressAt(state.selection.anchor);
    if (anchor.isErr()) return anchor;
    const head = this.addressAt(state.selection.head);
    if (head.isErr()) return head;
    return Result.ok({ anchor: anchor.value, head: head.value });
  }

  paragraph(blockId: string): ParagraphAddress | undefined {
    return this.paragraphs.find((paragraph) => paragraph.blockId === blockId);
  }
}

type ChangedStoriesOptions = { document: Document; previous: Document };
/** Immutable operations retain content identities for every untouched story. */
const changedStories = ({ document, previous }: ChangedStoriesOptions): OpStory[] => {
  const stories: OpStory[] = [];
  if (document.package.document.content !== previous.package.document.content)
    stories.push(OP_STORIES.MAIN);
  for (const kind of ["header", "footer"] as const) {
    const parts = kind === "header" ? document.package.headers : document.package.footers;
    const prior = kind === "header" ? previous.package.headers : previous.package.footers;
    if (parts === prior) continue;
    for (const [rId, body] of parts ?? []) {
      if (body.content !== prior?.get(rId)?.content) stories.push({ kind, rId });
    }
  }
  for (const kind of ["footnote", "endnote"] as const) {
    const notes = kind === "footnote" ? document.package.footnotes : document.package.endnotes;
    const prior = kind === "footnote" ? previous.package.footnotes : previous.package.endnotes;
    if (notes === prior) continue;
    const priorById = new Map(prior?.map((note) => [note.id, note]));
    for (const note of notes ?? []) {
      const original = priorById.get(note.id);
      if (note.content !== original?.content || note.noteType !== original?.noteType)
        stories.push({ kind, id: note.id });
    }
  }
  return stories;
};

/** Carry private paragraph source captures across immutable story operations. */
type PreservePropertySourcesOptions = {
  target: Document;
  source: Document;
  stories?: readonly OpStory[];
};
const preservePropertySources = ({
  target,
  source,
  stories = documentStories(source),
}: PreservePropertySourcesOptions): void => {
  for (const story of stories) {
    const originals = findStoryBody(source, story)?.content ?? [];
    const derived = findStoryBody(target, story)?.content ?? [];
    if (originals === derived) continue;
    const paragraphs = new Map(
      derived.flatMap((block) =>
        block.type === "paragraph" ? [[block.paraId, block] as const] : [],
      ),
    );
    for (const original of originals) {
      if (original.type !== "paragraph") continue;
      const next = paragraphs.get(original.paraId);
      if (next && next !== original) copyParagraphPropertySource(next, original);
    }
  }
  copyDocumentParagraphPropertySourceContract(target, source);
};

type AuthoredInputFormattingOptions = {
  paragraph: Paragraph;
  offset: number;
  affinity: "before" | "after";
};

/** PM caret input inherits left; range replacement inherits the first replaced unit. */
const authoredInputFormatting = ({
  paragraph,
  offset,
  affinity,
}: AuthoredInputFormattingOptions): TextFormatting => {
  const unit = affinity === "before" && offset > 0 ? offset - 1 : offset;
  let end = 0;
  for (const run of paragraph.content) {
    if (run.type !== "run") panic("Canonical input encountered a non-text run.");
    for (const child of run.content) {
      if (child.type === "noteMarker") continue;
      end += child.type === "text" ? child.text.length : 1;
    }
    if (unit < end) return run.formatting ?? {};
  }
  if (end !== 0) panic("Canonical input could not find the authored insertion formatting.");
  return {};
};

const supportsSeed = (document: Document, stories = documentStories(document)): boolean => {
  if ((document.package.document.comments?.length ?? 0) > 0) return false;
  return stories.every((story) => {
    const body = findStoryBody(document, story);
    if (story !== OP_STORIES.MAIN && (story.kind === "footnote" || story.kind === "endnote")) {
      const notes =
        story.kind === "footnote" ? document.package.footnotes : document.package.endnotes;
      const note = notes?.find(({ id }) => id === story.id);
      if (note?.noteType !== undefined && note.noteType !== "normal") return true;
    }
    const content = body?.content;
    return (
      content !== undefined &&
      content.length > 0 &&
      content.every(
        (paragraph) =>
          paragraph.type === "paragraph" &&
          paragraph.pPrMark === undefined &&
          paragraph.reviewCarrier === undefined &&
          (paragraph.propertyChanges?.length ?? 0) === 0 &&
          paragraph.content.every(
            (run) =>
              run.type === "run" &&
              (run.propertyChanges?.length ?? 0) === 0 &&
              run.content.every(
                (child) =>
                  child.type === "noteMarker" ||
                  child.type === "footnoteRef" ||
                  child.type === "endnoteRef" ||
                  (child.type === "text" &&
                    !hasIllegalXmlCharacters(child.text) &&
                    !/[\t\r\n]/u.test(child.text)),
              ),
          ),
      )
    );
  });
};

type ProjectStoryOptions = {
  document: Document;
  styles: StyleDefinitions | null | undefined;
  story?: OpStory;
};

const project = ({
  document,
  styles,
  story = OP_STORIES.MAIN,
}: ProjectStoryOptions): Result<CanonicalProjection, CanonicalSessionError> => {
  const body = findStoryBody(document, story);
  if (!body) return refuse("The canonical story no longer exists.");
  const options = {
    ...(styles == null ? {} : { styles }),
    ...(document.package.theme === undefined ? {} : { theme: document.package.theme }),
  };
  const converted = Result.try({
    try: () => {
      if (story === OP_STORIES.MAIN) return toProseDoc(document, options);
      if (story.kind === "header" || story.kind === "footer")
        return headerFooterToProseDoc(body.content, options);
      return footnoteToProseDoc(body.content, options);
    },
    catch: (cause) =>
      new CanonicalSessionError({
        message: `Canonical projection failed: ${cause instanceof Error ? cause.message : String(cause)}`,
      }),
  });
  if (converted.isErr()) return converted;
  const paragraphs: ParagraphAddress[] = [];
  let failure: CanonicalSessionError | undefined;
  converted.value.forEach((node, offset, index) => {
    const source = body.content.at(index);
    if (source?.type !== "paragraph" || source.paraId === undefined) {
      failure = new CanonicalSessionError({
        message: "The projection changed paragraph structure.",
      });
      return;
    }
    let text = "";
    let renderedText = "";
    let renderedSize = 0;
    const boundaries: number[][] = [[0]];
    for (const run of source.content) {
      if (run.type !== "run") return;
      for (const child of run.content) {
        if (child.type === "text") {
          for (const unit of child.text.split("")) {
            text += unit;
            renderedText += unit;
            renderedSize += 1;
            boundaries.push([renderedSize]);
          }
        } else if (child.type === "footnoteRef" || child.type === "endnoteRef") {
          text += "\uFFFC";
          renderedText += String(child.id);
          renderedSize += String(child.id).length;
          boundaries.push([renderedSize]);
        } else if (child.type === "noteMarker") {
          renderedSize += 1;
          const gaps = boundaries.at(-1) ?? panic("Missing canonical note gap");
          gaps.push(renderedSize);
        }
      }
    }
    if (
      node.type.name !== "paragraph" ||
      node.attrs["paraId"] !== source.paraId ||
      node.textContent !== renderedText ||
      node.content.size !== renderedSize
    ) {
      failure = new CanonicalSessionError({
        message: "The paragraph cannot be projected as plain text.",
      });
      return;
    }
    paragraphs.push({ blockId: source.paraId, start: offset + 1, text, boundaries, node, source });
  });
  if (failure !== undefined) return Result.err(failure);
  if (paragraphs.length !== body.content.length) {
    return refuse("The projection changed paragraph structure.");
  }
  return Result.ok(new CanonicalProjection({ doc: converted.value, paragraphs, story }));
};

/** Input intent, never projection-step adjacency, determines the journal partition. */
export type CanonicalInputSemantic =
  | "typing"
  | "deleteBackward"
  | "deleteForward"
  | "paste"
  | "composition"
  | "replacement"
  | "structure";

const UNDO_GROUP_WINDOW_MS = 500;

const samePosition = (left: TextPosition, right: TextPosition): boolean =>
  left.story === right.story &&
  left.blockId === right.blockId &&
  left.offset === right.offset &&
  (left.zeroWidthBefore ?? 0) === (right.zeroWidthBefore ?? 0);

const sameSelection = (left: CanonicalSelection, right: CanonicalSelection): boolean =>
  samePosition(left.anchor, right.anchor) && samePosition(left.head, right.head);

type AppliedJournalEntry = {
  type: "applied";
  ops: readonly DocumentOp[];
  inverse: readonly DocumentOp[];
  preSelection: CanonicalSelection;
  postSelection: CanonicalSelection;
  version: number;
  origin: "input";
  semantic: CanonicalInputSemantic;
  grouping: "run" | "isolated";
  time: number;
  boundary: number;
};

type AppliedJournalGroup = {
  entries: AppliedJournalEntry[];
  undo: { type: "entries" } | { type: "replayed"; inverse: readonly DocumentOp[] };
  preSelection: CanonicalSelection;
  postSelection: CanonicalSelection;
};

const continuesGroup = (previous: AppliedJournalEntry, next: AppliedJournalEntry): boolean => {
  switch (next.semantic) {
    case "typing":
    case "deleteBackward":
    case "deleteForward":
      return (
        previous.grouping === "run" &&
        next.grouping === "run" &&
        previous.semantic === next.semantic &&
        previous.boundary === next.boundary &&
        next.time >= previous.time &&
        next.time - previous.time <= UNDO_GROUP_WINDOW_MS &&
        sameSelection(previous.postSelection, next.preSelection) &&
        samePosition(next.preSelection.anchor, next.preSelection.head)
      );
    case "paste":
    case "composition":
    case "replacement":
    case "structure":
      return false;
    default: {
      const unreachable: never = next.semantic;
      return panic(`Unknown canonical input semantic ${unreachable}`);
    }
  }
};

type UndoneJournalEntry = {
  type: "undone";
  group: AppliedJournalGroup;
  redoOps: readonly DocumentOp[];
};

export type CanonicalCommit = {
  transaction: Transaction;
  document: Document;
  projection: CanonicalProjection;
  bodyProjection: CanonicalProjection;
  touched: TouchedBlocks;
  version: number;
  origin: CanonicalOrigin;
  /** Publish only after every projection/plugin transaction has been accepted. */
  publish: () => Result<void, CanonicalSessionError>;
};

type ProjectionAuthorization = { session: CanonicalSession; version: number; doc: PMNode };
const authorizedProjections = new WeakMap<Transaction, ProjectionAuthorization>();

/** Metadata is descriptive; transaction identity and session version authorize mutations. */
export const isCanonicalProjectionTransaction = (
  transaction: Transaction,
  session: CanonicalSession,
): boolean => {
  const authorization = authorizedProjections.get(transaction);
  return (
    authorization?.session === session &&
    authorization.version === session.version &&
    transaction.doc.eq(authorization.doc)
  );
};

type PublishCanonicalProjectionOptions = {
  state: EditorState;
  commit: CanonicalCommit;
  session: CanonicalSession;
};

/** The model publishes only after PM and every plugin accept its exact projection. */
export const publishCanonicalProjection = ({
  state,
  commit,
  session,
}: PublishCanonicalProjectionOptions) => {
  if (!isCanonicalProjectionTransaction(commit.transaction, session))
    return refuse("The staged canonical projection is unauthorized.");
  const applied = Result.try({
    try: () => state.applyTransaction(commit.transaction),
    catch: (cause) =>
      new CanonicalSessionError({
        message: `Canonical plugin staging failed: ${cause instanceof Error ? cause.message : String(cause)}`,
      }),
  });
  if (applied.isErr()) return applied;
  const staged = applied.value;
  if (
    !staged.transactions.includes(commit.transaction) ||
    !staged.state.doc.eq(commit.projection.doc) ||
    staged.transactions.some(
      (transaction) =>
        transaction.docChanged && !isCanonicalProjectionTransaction(transaction, session),
    )
  )
    return refuse("A plugin refused or changed the canonical projection.");
  const published = commit.publish();
  if (published.isErr()) return published;
  return Result.ok(staged);
};

type StageOptions = {
  state: EditorState;
  ops: readonly DocumentOp[];
  selection: CanonicalSelection;
  origin: CanonicalOrigin;
  story?: OpStory;
  previousProjection?: CanonicalProjection;
  onPublish: (inverse: readonly DocumentOp[], version: number) => void;
};

type CanonicalSessionSeedOptions = {
  document: Document;
  projection: CanonicalProjection;
  styles: StyleDefinitions | null | undefined;
};

type CanonicalReplaceTextInput = {
  from: number;
  to: number;
  text: string;
  story?: OpStory;
  semantic?: CanonicalInputSemantic;
  time?: number;
};

type CachedStoryProjection = {
  content: NonNullable<ReturnType<typeof findStoryBody>>["content"];
  theme: Document["package"]["theme"];
  projection: CanonicalProjection;
};

const storyProjectionKey = (story: Exclude<OpStory, typeof OP_STORIES.MAIN>): string => {
  switch (story.kind) {
    case "header":
    case "footer":
      return `${story.kind}:${story.rId}`;
    case "footnote":
    case "endnote":
      return `${story.kind}:${story.id}`;
    default: {
      const unreachable: never = story;
      return panic(`Unknown canonical story ${unreachable}`);
    }
  }
};

/** Immutable model authority with a journal staged independently of the PM view. */
class CanonicalSession {
  private currentDocument: Document;
  private currentProjection: CanonicalProjection;
  private currentVersion = 0;
  private readonly storyProjections = new Map<string, CachedStoryProjection>();
  private currentSelection: CanonicalSelection | null = null;
  private readonly applied: AppliedJournalGroup[] = [];
  private groupingBoundary = 0;
  private lifecycle: { type: "committed" } | { type: "composing" } = { type: "committed" };
  private readonly undone: UndoneJournalEntry[] = [];
  private readonly styles: StyleDefinitions | null | undefined;

  constructor({ document, projection, styles }: CanonicalSessionSeedOptions) {
    this.currentDocument = document;
    this.currentProjection = projection;
    this.styles = styles == null ? styles : structuredClone(styles);
  }

  get document(): Document {
    if (this.isComposing)
      throw new CanonicalSessionError({
        message: "Composition must finish before taking a snapshot.",
      });
    return this.currentDocument;
  }
  get projection(): CanonicalProjection {
    return this.currentProjection;
  }
  get selection(): CanonicalSelection | null {
    return this.currentSelection;
  }
  get version(): number {
    return this.currentVersion;
  }
  get canUndo(): boolean {
    return this.applied.length > 0;
  }
  get canRedo(): boolean {
    return this.undone.length > 0;
  }

  get isComposing(): boolean {
    return this.lifecycle.type === "composing";
  }

  breakUndoGroup(): void {
    this.groupingBoundary += 1;
  }

  beginComposition(): Result<void, CanonicalSessionError> {
    if (this.isComposing) return refuse("A canonical composition is already pending.");
    this.lifecycle = { type: "composing" };
    this.breakUndoGroup();
    return Result.ok(undefined);
  }

  endComposition(): void {
    this.lifecycle = { type: "committed" };
    this.breakUndoGroup();
  }

  private checkState(
    state: EditorState,
    projection = this.projection,
  ): Result<void, CanonicalSessionError> {
    if (this.isComposing) return refuse("Composition must finish before committing another edit.");
    if (!state.doc.eq(projection.doc)) {
      return refuse("The input uses a stale canonical projection.");
    }
    return Result.ok(undefined);
  }

  prepareReplace(
    state: EditorState,
    {
      from,
      to,
      text,
      story = OP_STORIES.MAIN,
      semantic = "replacement",
      time = Date.now(),
    }: CanonicalReplaceTextInput,
  ): Result<CanonicalCommit, CanonicalSessionError> {
    const projected = this.projectStory(story);
    if (projected.isErr()) return projected;
    const projection = projected.value;
    const checked = this.checkState(state, projection);
    if (checked.isErr()) return checked;
    if (!Number.isFinite(time)) return refuse("The input time is invalid.");
    if (state.storedMarks !== null && state.storedMarks.length > 0) {
      return refuse("Stored formatting is not supported in the canonical session.");
    }
    if (from > to || hasIllegalXmlCharacters(text) || /[\t\r\n]/u.test(text)) {
      return refuse("Canonical input accepts a well-formed same-paragraph text replacement.");
    }
    if (from === to && text.length === 0) return refuse("The input makes no text change.");
    const start = projection.addressAt(from);
    if (start.isErr()) return start;
    const end = projection.addressAt(to);
    if (end.isErr()) return end;
    if (start.value.blockId !== end.value.blockId) {
      return refuse("Cross-paragraph replacement is not supported in the canonical session.");
    }
    const preSelection = projection.selectionAt(state);
    if (preSelection.isErr()) return preSelection;
    const ops: DocumentOp[] = [];
    if (from !== to) {
      ops.push({ type: DOCUMENT_OP_TYPES.DELETE_RANGE, from: start.value, to: end.value });
    }
    if (text.length > 0) {
      const paragraph = projection.paragraph(start.value.blockId);
      if (paragraph === undefined) panic("Canonical input lost its addressed paragraph.");
      const runProps = authoredInputFormatting({
        paragraph: paragraph.source,
        offset: start.value.offset,
        affinity: from === to ? "before" : "after",
      });
      ops.push({ type: DOCUMENT_OP_TYPES.INSERT_TEXT, at: start.value, text, runProps });
    }
    const caret = {
      ...start.value,
      offset: start.value.offset + text.length,
      ...(text.length > 0 && start.value.zeroWidthBefore !== undefined
        ? { zeroWidthBefore: 0 }
        : {}),
    };
    const postSelection = { anchor: caret, head: caret };
    const isCaret = state.selection.empty;
    const run =
      isCaret &&
      ((semantic === "typing" && from === to && from === state.selection.head) ||
        (semantic === "deleteBackward" && text.length === 0 && to === state.selection.head) ||
        (semantic === "deleteForward" && text.length === 0 && from === state.selection.head));
    return this.stage({
      state,
      story,
      previousProjection: projection,
      ops,
      selection: postSelection,
      origin: "input",
      onPublish: (inverse, version) => {
        const entry = {
          type: "applied",
          ops,
          inverse,
          preSelection: preSelection.value,
          postSelection,
          version,
          origin: "input",
          semantic,
          grouping: run ? "run" : "isolated",
          time,
          boundary: this.groupingBoundary,
        } as const satisfies AppliedJournalEntry;
        const group = this.applied.at(-1);
        const previous = group?.entries.at(-1);
        if (group !== undefined && previous !== undefined && continuesGroup(previous, entry)) {
          group.entries.push(entry);
          group.postSelection = postSelection;
        } else {
          this.applied.push({
            entries: [entry],
            undo: { type: "entries" },
            preSelection: preSelection.value,
            postSelection,
          });
        }
        this.undone.length = 0;
      },
    });
  }

  projectStory(story: OpStory): Result<CanonicalProjection, CanonicalSessionError> {
    if (story === OP_STORIES.MAIN) return Result.ok(this.projection);
    const key = storyProjectionKey(story);
    const body = findStoryBody(this.currentDocument, story);
    if (body === undefined) {
      this.storyProjections.delete(key);
      return refuse("The canonical story no longer exists.");
    }
    const cached = this.storyProjections.get(key);
    const theme = this.currentDocument.package.theme;
    if (cached?.content === body.content && cached.theme === theme)
      return Result.ok(cached.projection);
    const projected = project({ document: this.currentDocument, styles: this.styles, story });
    if (projected.isOk())
      this.storyProjections.set(key, { content: body.content, theme, projection: projected.value });
    return projected;
  }

  prepareOperations(
    state: EditorState,
    ops: readonly DocumentOp[],
  ): Result<CanonicalCommit, CanonicalSessionError> {
    if (ops.length === 0) return refuse("The operation batch is empty.");
    const selected = this.projection.selectionAt(state);
    if (selected.isErr()) return selected;
    return this.stage({
      state,
      ops,
      selection: selected.value,
      origin: "input",
      onPublish: (inverse, version) => {
        this.applied.push({
          entries: [
            {
              type: "applied",
              ops,
              inverse,
              preSelection: selected.value,
              postSelection: selected.value,
              version,
              origin: "input",
              semantic: "structure",
              grouping: "isolated",
              time: Date.now(),
              boundary: this.groupingBoundary,
            },
          ],
          undo: { type: "entries" },
          preSelection: selected.value,
          postSelection: selected.value,
        });
        this.undone.length = 0;
      },
    });
  }

  prepareUndo(
    state: EditorState,
    story: OpStory = OP_STORIES.MAIN,
  ): Result<CanonicalCommit, CanonicalSessionError> {
    const group = this.applied.at(-1);
    if (group === undefined) return refuse("There is no canonical edit to undo.");
    return this.stage({
      state,
      story,
      ops:
        group.undo.type === "replayed"
          ? group.undo.inverse
          : group.entries.toReversed().flatMap((entry) => entry.inverse),
      selection: group.preSelection,
      origin: "undo",
      onPublish: (redoOps) => {
        this.applied.pop();
        this.undone.push({ type: "undone", group, redoOps });
        this.breakUndoGroup();
      },
    });
  }

  prepareRedo(
    state: EditorState,
    story: OpStory = OP_STORIES.MAIN,
  ): Result<CanonicalCommit, CanonicalSessionError> {
    const undone = this.undone.at(-1);
    if (undone === undefined) return refuse("There is no canonical edit to redo.");
    return this.stage({
      state,
      story,
      ops: undone.redoOps,
      selection: undone.group.postSelection,
      origin: "redo",
      onPublish: (inverse) => {
        this.undone.pop();
        this.applied.push({
          entries: undone.group.entries,
          undo: { type: "replayed", inverse },
          preSelection: undone.group.preSelection,
          postSelection: undone.group.postSelection,
        });
        this.breakUndoGroup();
      },
    });
  }

  private stage({
    state,
    ops,
    selection,
    origin,
    onPublish,
    story = OP_STORIES.MAIN,
    previousProjection,
  }: StageOptions): Result<CanonicalCommit, CanonicalSessionError> {
    const previous =
      previousProjection === undefined ? this.projectStory(story) : Result.ok(previousProjection);
    if (previous.isErr()) return previous;
    const checked = this.checkState(state, previous.value);
    if (checked.isErr()) return checked;
    const applied = applyDocumentOps(this.currentDocument, ops);
    if (applied.isErr()) return refuse(applied.error.message);
    const changed = changedStories({
      document: applied.value.document,
      previous: this.currentDocument,
    });
    if (!supportsSeed(applied.value.document, changed))
      return refuse("The operations produce unsupported canonical story content.");
    preservePropertySources({
      target: applied.value.document,
      source: this.currentDocument,
      stories: changed,
    });
    const projectedStory = findStoryBody(applied.value.document, story) ? story : OP_STORIES.MAIN;
    const projected = project({
      document: applied.value.document,
      styles: this.styles,
      story: projectedStory,
    });
    if (projected.isErr()) return projected;
    const bodyProjection = (() => {
      if (sameStory(projectedStory, OP_STORIES.MAIN)) return projected;
      if (ops.some(operationChangesBodyProjection))
        return project({ document: applied.value.document, styles: this.styles });
      return Result.ok(this.currentProjection);
    })();
    if (bodyProjection.isErr()) return bodyProjection;
    const transaction = state.tr;
    if (!transaction.doc.eq(projected.value.doc))
      transaction.replaceWith(0, transaction.doc.content.size, projected.value.doc.content);
    if (sameStory(selection.anchor.story, story) && sameStory(selection.head.story, story)) {
      const anchor = projected.value.positionAt(selection.anchor);
      const head = projected.value.positionAt(selection.head);
      if (anchor.isOk() && head.isOk())
        transaction.setSelection(TextSelection.create(transaction.doc, anchor.value, head.value));
      else
        transaction.setSelection(
          TextSelection.near(
            transaction.doc.resolve(Math.min(state.selection.anchor, transaction.doc.content.size)),
          ),
        );
    }
    transaction.setMeta(CANONICAL_PROJECTION_META, {
      type: "canonical",
      origin,
      version: this.version + 1,
    });
    if (ops.some(operationChangesPackage)) markPackageChange(transaction);
    transaction.setMeta("addToHistory", false);
    const baseVersion = this.version;
    const baseBoundary = this.groupingBoundary;
    authorizedProjections.set(transaction, {
      session: this,
      version: baseVersion,
      doc: projected.value.doc,
    });
    return Result.ok({
      transaction,
      document: applied.value.document,
      projection: projected.value,
      bodyProjection: bodyProjection.value,
      touched: applied.value.touched,
      version: baseVersion + 1,
      origin,
      publish: () => {
        if (this.isComposing || this.groupingBoundary !== baseBoundary)
          return refuse("The staged canonical commit crossed a composition or selection boundary.");
        if (this.version !== baseVersion) return refuse("The staged canonical commit is stale.");
        if (!isCanonicalProjectionTransaction(transaction, this)) {
          return refuse("The staged canonical projection was changed after preparation.");
        }
        if (projectedStory !== OP_STORIES.MAIN) {
          const body = findStoryBody(applied.value.document, projectedStory);
          if (body === undefined) panic("A staged story projection lost its source.");
          this.storyProjections.set(storyProjectionKey(projectedStory), {
            content: body.content,
            theme: applied.value.document.package.theme,
            projection: projected.value,
          });
        }
        this.currentSelection = selection;
        this.currentDocument = applied.value.document;
        this.currentProjection = bodyProjection.value;
        this.currentVersion = baseVersion + 1;
        onPublish(applied.value.inverse, this.currentVersion);
        return Result.ok(undefined);
      },
    });
  }
}

export type { CanonicalProjection, CanonicalSession };

export const createCanonicalSession = (
  document: Document,
  styles?: StyleDefinitions | null,
): Result<CanonicalSession, CanonicalSessionError> => {
  if (!supportsSeed(document)) {
    return refuse(
      "Canonical sessions currently require plain paragraphs and note references without revisions.",
    );
  }
  const owned = cloneDocumentWithParagraphPropertySources(document);
  const normalized = normalizeForOps(owned);
  preservePropertySources({ target: normalized, source: owned });
  const validated = validateOpsDocument(normalized);
  if (validated.isErr()) return refuse(validated.error.message);
  const projected = project({ document: normalized, styles });
  if (projected.isErr()) return projected;
  return Result.ok(
    new CanonicalSession({ document: normalized, projection: projected.value, styles }),
  );
};

/** Native character deletion stays in one paragraph and consumes a whole grapheme. */
export const deletionRange = (
  state: EditorState,
  direction: "backward" | "forward",
): Result<{ from: number; to: number }, CanonicalSessionError> => {
  const { selection } = state;
  if (!(selection instanceof TextSelection) || !selection.$from.sameParent(selection.$to)) {
    return refuse("Canonical deletion requires a same-paragraph text selection.");
  }
  if (!selection.empty) return Result.ok({ from: selection.from, to: selection.to });
  const parent = selection.$from.parent;
  // Invisible atoms still occupy PM positions; retain that space for grapheme boundaries.
  const text = parent.textBetween(0, parent.content.size, "", "\uFFFC");
  const offset = selection.$from.parentOffset;
  const adjacent =
    direction === "backward" ? parent.childBefore(offset).node : parent.childAfter(offset).node;
  if (adjacent?.type.name === "noteMarker")
    return refuse("Automatic note marks cannot be deleted as characters.");
  if (splitsSurrogatePair(text, offset))
    return refuse("The deletion would split a surrogate pair.");
  if (splitsGraphemeCluster(text, offset)) {
    return refuse("Character deletion requires a caret at a grapheme boundary.");
  }
  if (
    (direction === "backward" && offset === 0) ||
    (direction === "forward" && offset === text.length)
  ) {
    return refuse("Paragraph joins are not supported in the canonical session.");
  }
  let boundary = offset + (direction === "backward" ? -1 : 1);
  while (splitsGraphemeCluster(text, boundary)) boundary += direction === "backward" ? -1 : 1;
  const position = selection.$from.start() + boundary;
  return Result.ok(
    direction === "backward"
      ? { from: position, to: selection.to }
      : { from: selection.from, to: position },
  );
};
