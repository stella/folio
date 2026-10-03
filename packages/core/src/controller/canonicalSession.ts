import { panic, Result, TaggedError } from "better-result";
import type { Node as PMNode } from "prosemirror-model";
import { TextSelection, type EditorState, type Transaction } from "prosemirror-state";
import {
  applyDocumentOps,
  DOCUMENT_OP_TYPES,
  normalizeForOps,
  OP_STORIES,
  validateOpsDocument,
  compileEditorIntent,
  createEditorIntentIdAllocator,
  editorParagraphGroups,
  type EditorIntent,
  type EditorIntentMode,
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
import type { Document, Paragraph, StyleDefinitions } from "../types/document";

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
    case DOCUMENT_OP_TYPES.CREATE_NUMBERING_INSTANCE:
    case DOCUMENT_OP_TYPES.DELETE_NUMBERING_INSTANCE:
    case DOCUMENT_OP_TYPES.SET_SECTION_ENDPOINT:
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
    case DOCUMENT_OP_TYPES.CREATE_NUMBERING_INSTANCE:
    case DOCUMENT_OP_TYPES.DELETE_NUMBERING_INSTANCE:
    case DOCUMENT_OP_TYPES.SET_SECTION_ENDPOINT:
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
export type CanonicalSessionMode = { type: "editing" } | { type: "suggesting"; author: string };
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

  inputAddressAt(position: number): Result<TextPosition, CanonicalSessionError> {
    const address = this.addressAt(position);
    if (address.isErr()) return address;
    const paragraph = this.paragraph(address.value.blockId);
    if (paragraph === undefined) panic("Input lost its paragraph.");
    let offset = address.value.offset;
    paragraph.node.forEach((node, start) => {
      if (
        node.marks.some((mark) => mark.type.name === "deletion") &&
        offset > start &&
        offset < start + node.nodeSize
      )
        offset = start + node.nodeSize;
    });
    return Result.ok({ ...address.value, offset });
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

const supportsTextContent = (content: Paragraph["content"]): boolean =>
  content.every((item) => {
    switch (item.type) {
      case "run":
        return item.content.every(
          (child) =>
            child.type === "noteMarker" ||
            child.type === "footnoteRef" ||
            child.type === "endnoteRef" ||
            (child.type === "text" &&
              !hasIllegalXmlCharacters(child.text) &&
              !/[\t\r\n]/u.test(child.text)),
        );
      case "insertion":
      case "deletion":
        return supportsTextContent(item.content);
      default:
        return false;
    }
  });

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
          paragraph.reviewCarrier === undefined &&
          supportsTextContent(paragraph.content),
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
    const appendContent = (content: Paragraph["content"]): void => {
      for (const run of content) {
        if (run.type === "insertion" || run.type === "deletion") {
          appendContent(run.content);
          continue;
        }
        if (run.type !== "run") panic("Canonical projection encountered unsupported content");
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
    };
    appendContent(source.content);
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
  sameStory(left.story, right.story) &&
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

type JournalInputOptions = Pick<StageOptions, "state" | "ops" | "story"> & {
  preSelection: CanonicalSelection;
  postSelection: CanonicalSelection;
  semantic?: CanonicalInputSemantic;
  time?: number;
  grouping?: "run" | "isolated";
};

type IntentInputOptions = Pick<JournalInputOptions, "semantic" | "time" | "grouping"> & {
  intent:
    | Exclude<EditorIntent, { type: "splitParagraph" }>
    | Omit<Extract<EditorIntent, { type: "splitParagraph" }>, "newBlockId">;
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
  private readonly allocateIntentIds = createEditorIntentIdAllocator();
  private mode: CanonicalSessionMode = { type: "editing" };
  private readonly sourceOwners = new Map<string, Paragraph>();
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
    this.rememberSources(document);
  }

  setMode(mode: CanonicalSessionMode): void {
    if (
      this.mode.type !== mode.type ||
      (this.mode.type === "suggesting" &&
        mode.type === "suggesting" &&
        this.mode.author !== mode.author)
    )
      this.breakUndoGroup();
    this.mode =
      mode.type === "editing" ? { type: "editing" } : { type: "suggesting", author: mode.author };
  }

  private rememberSources(document: Document): void {
    for (const block of document.package.document.content) {
      if (block.type === "paragraph" && block.paraId !== undefined)
        this.sourceOwners.set(block.paraId, block);
    }
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
      return refuse("Canonical input accepts a well-formed text replacement.");
    }
    if (from === to && text.length === 0) return refuse("The input makes no text change.");
    const start = projection.inputAddressAt(from);
    if (start.isErr()) return start;
    const end = projection.inputAddressAt(to);
    if (end.isErr()) return end;
    const run =
      state.selection.empty &&
      ((semantic === "typing" && from === to && from === state.selection.head) ||
        (semantic === "deleteBackward" && text.length === 0 && to === state.selection.head) ||
        (semantic === "deleteForward" && text.length === 0 && from === state.selection.head));
    return this.prepareIntent(state, {
      intent: {
        type: "replaceText",
        from: start.value,
        to: end.value,
        text,
      },
      semantic,
      time,
      grouping: run ? "run" : "isolated",
    });
  }

  prepareSplit(state: EditorState): Result<CanonicalCommit, CanonicalSessionError> {
    const checked = this.checkState(state);
    if (checked.isErr()) return checked;
    const at = this.projection.inputAddressAt(state.selection.from);
    if (at.isErr()) return at;
    const to = this.projection.inputAddressAt(state.selection.to);
    if (to.isErr()) return to;
    return this.prepareIntent(state, {
      intent: {
        type: "splitParagraph",
        at: at.value,
        to: to.value,
      },
    });
  }

  prepareJoin(
    state: EditorState,
    direction: "backward" | "forward",
  ): Result<CanonicalCommit, CanonicalSessionError> {
    const checked = this.checkState(state);
    if (checked.isErr()) return checked;
    if (!state.selection.empty) return refuse("Join requires a collapsed text selection.");
    const at = this.projection.inputAddressAt(state.selection.from);
    if (at.isErr()) return at;
    const groups = editorParagraphGroups(this.document, at.value.story);
    const index = groups.findIndex(({ paragraphs }) =>
      paragraphs.some((paragraph) => paragraph.paraId === at.value.blockId),
    );
    if (index < 0 || (direction === "backward" && index === 0))
      return refuse("There is no adjacent paragraph to join.");
    const firstGroup = groups.at(direction === "backward" ? index - 1 : index);
    const secondGroup = groups.at(direction === "backward" ? index : index + 1);
    const first = firstGroup?.paragraphs.at(-1);
    const second = secondGroup?.paragraphs.at(0);
    if (!first?.paraId || !second?.paraId) return refuse("There is no adjacent paragraph to join.");
    if (!isCanonicalJoinBoundary(state, direction))
      return refuse("Join requires a caret at a paragraph boundary.");
    return this.prepareIntent(state, {
      intent: {
        type: "joinParagraphs",
        story: at.value.story,
        blockId: first.paraId,
        nextBlockId: second.paraId,
      },
    });
  }

  private prepareIntent(
    state: EditorState,
    { intent, ...journal }: IntentInputOptions,
  ): Result<CanonicalCommit, CanonicalSessionError> {
    const story = (() => {
      switch (intent.type) {
        case "replaceText":
        case "formatRun":
        case "insertAtom":
          return intent.from.story;
        case "splitParagraph":
        case "formatParagraph":
          return intent.at.story;
        case "joinParagraphs":
          return intent.story;
        case "setList": {
          const first = intent.items.at(0);
          if (first === undefined) panic("A list input must address a paragraph.");
          return first.at.story;
        }
        default: {
          const unreachable: never = intent;
          return panic(`Unknown canonical intent ${unreachable}`);
        }
      }
    })();
    const projected = this.projectStory(story);
    if (projected.isErr()) return projected;
    const checked = this.checkState(state, projected.value);
    if (checked.isErr()) return checked;
    const preSelection = projected.value.selectionAt(state);
    if (preSelection.isErr()) return preSelection;
    const ids = this.allocateIntentIds(this.document, intent);
    const mode =
      this.mode.type === "editing"
        ? ({ type: "editing", newIds: ids.newIds } as const satisfies EditorIntentMode)
        : ({
            type: "suggesting",
            revision: {
              id: ids.revisionId,
              author: this.mode.author,
              date: new Date().toISOString(),
            },
            newIds: ids.newIds,
          } as const satisfies EditorIntentMode);
    const compiled = compileEditorIntent(this.document, {
      intent: intent.type === "splitParagraph" ? { ...intent, newBlockId: ids.newBlockId } : intent,
      mode,
    });
    if (compiled.isErr()) return refuse(compiled.error.message);
    return this.prepareJournalled({
      state,
      ops: compiled.value.ops,
      story,
      ...journal,
      preSelection: preSelection.value,
      postSelection: {
        anchor: compiled.value.selection,
        head: compiled.value.selection,
      },
    });
  }

  prepareResolve(
    state: EditorState,
    {
      revisionIds,
      resolution,
    }: { revisionIds: readonly number[]; resolution: "accept" | "reject" },
  ): Result<CanonicalCommit, CanonicalSessionError> {
    const checked = this.checkState(state);
    if (checked.isErr()) return checked;
    const preSelection = this.projection.selectionAt(state);
    if (preSelection.isErr()) return preSelection;
    const ops = [
      {
        type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
        story: OP_STORIES.MAIN,
        revisionIds,
        decision: resolution,
      },
    ] as const;
    // Resolution can retire the selected paragraph; choose a valid result anchor before staging.
    const applied = applyDocumentOps(this.document, ops);
    if (applied.isErr()) return refuse(applied.error.message);
    const projected = project({ document: applied.value.document, styles: this.styles });
    if (projected.isErr()) return projected;
    const surviving = projected.value.paragraph(preSelection.value.head.blockId);
    const first = applied.value.document.package.document.content.at(0);
    if (first?.type !== "paragraph" || first.paraId === undefined)
      return refuse("Resolution left no paragraph.");
    let offset = surviving ? Math.min(preSelection.value.head.offset, surviving.text.length) : 0;
    if (surviving && splitsSurrogatePair(surviving.text, offset)) offset -= 1;
    const caret = { story: OP_STORIES.MAIN, blockId: surviving?.blockId ?? first.paraId, offset };
    return this.prepareJournalled({
      state,
      ops,
      preSelection: preSelection.value,
      postSelection: { anchor: caret, head: caret },
    });
  }

  private prepareJournalled({
    state,
    ops,
    preSelection,
    postSelection,
    semantic = "structure",
    time = Date.now(),
    grouping = "isolated",
    story = preSelection.anchor.story,
  }: JournalInputOptions): Result<CanonicalCommit, CanonicalSessionError> {
    return this.stage({
      state,
      story,
      ops,
      selection: postSelection,
      origin: "input",
      onPublish: (inverse, version) => {
        const entry = {
          type: "applied",
          ops,
          inverse,
          preSelection,
          postSelection,
          version,
          origin: "input",
          semantic,
          grouping,
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
            preSelection,
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
    const stagedSources = new Map(this.sourceOwners);
    for (const op of ops) {
      if (op.type !== DOCUMENT_OP_TYPES.SPLIT_BLOCK) continue;
      const source = stagedSources.get(op.at.blockId);
      if (source && !stagedSources.has(op.newBlockId)) stagedSources.set(op.newBlockId, source);
    }
    for (const block of applied.value.document.package.document.content) {
      if (block.type !== "paragraph" || block.paraId === undefined) continue;
      const source = stagedSources.get(block.paraId);
      if (source && block !== source) copyParagraphPropertySource(block, source);
    }
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
    const structureChanged =
      applied.value.touched.inserted.length > 0 || applied.value.touched.removed.length > 0;
    if (structureChanged) {
      transaction.replaceWith(0, transaction.doc.content.size, projected.value.doc.content);
    } else {
      const modifiedParagraphs = applied.value.touched.modified
        .flatMap((blockId) => {
          const paragraph = previous.value.paragraph(blockId);
          // A journal entry can modify another story while this view stays active.
          return paragraph === undefined ? [] : [paragraph];
        })
        .sort((left, right) => right.start - left.start);
      for (const oldParagraph of modifiedParagraphs) {
        const nextParagraph = projected.value.paragraph(oldParagraph.blockId);
        if (nextParagraph === undefined)
          return refuse("The operation changed paragraph structure.");
        transaction.replaceWith(
          oldParagraph.start - 1,
          oldParagraph.start - 1 + oldParagraph.node.nodeSize,
          nextParagraph.node,
        );
      }
    }
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
        this.rememberSources(applied.value.document);
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

const hasDeletedParagraphMark = (node: PMNode): boolean => {
  const mark: unknown = node.attrs["pPrMark"];
  return (
    typeof mark === "object" &&
    mark !== null &&
    "kind" in mark &&
    (mark.kind === "del" || mark.kind === "moveFrom")
  );
};

/** Visible graphemes concatenate across retained deletions and deleted paragraph marks. */
const visibleDeletionContext = (state: EditorState) => {
  const { selection } = state;
  const index = selection.$from.index(0);
  let first = index;
  while (first > 0 && hasDeletedParagraphMark(state.doc.child(first - 1))) first -= 1;
  let last = index;
  while (last + 1 < state.doc.childCount && hasDeletedParagraphMark(state.doc.child(last)))
    last += 1;
  let text = "";
  const physicalGaps: number[] = [];
  let visibleOffset = 0;
  state.doc.forEach((paragraph, paragraphOffset, paragraphIndex) => {
    if (paragraphIndex < first || paragraphIndex > last) return;
    paragraph.forEach((node, offset) => {
      if (node.marks.some((mark) => mark.type.name === "deletion")) return;
      const value = node.isText ? (node.text ?? "") : "\uFFFC";
      for (let unit = 0; unit < value.length; unit += 1) {
        const position = paragraphOffset + 1 + offset + unit;
        physicalGaps[text.length] = position;
        text += value.charAt(unit);
        physicalGaps.push(position + (node.isText ? 1 : node.nodeSize));
        if (position < selection.from) visibleOffset += 1;
      }
    });
  });
  return { text, physicalGaps, visibleOffset };
};

export const isCanonicalJoinBoundary = (
  state: EditorState,
  direction: "backward" | "forward",
): boolean => {
  if (!(state.selection instanceof TextSelection) || !state.selection.empty) return false;
  const { text, visibleOffset } = visibleDeletionContext(state);
  return visibleOffset === (direction === "backward" ? 0 : text.length);
};

/** Native character deletion consumes one whole visible grapheme. */
export const deletionRange = (
  state: EditorState,
  direction: "backward" | "forward",
): Result<{ from: number; to: number }, CanonicalSessionError> => {
  const { selection } = state;
  if (!(selection instanceof TextSelection))
    return refuse("Canonical deletion requires a text selection.");
  if (!selection.empty) return Result.ok({ from: selection.from, to: selection.to });
  const parent = selection.$from.parent;
  const offset = selection.$from.parentOffset;
  const adjacent =
    direction === "backward" ? parent.childBefore(offset).node : parent.childAfter(offset).node;
  if (adjacent?.type.name === "noteMarker")
    return refuse("Automatic note marks cannot be deleted as characters.");
  const { text, physicalGaps, visibleOffset } = visibleDeletionContext(state);
  if (splitsSurrogatePair(text, visibleOffset))
    return refuse("The deletion would split a surrogate pair.");
  if (splitsGraphemeCluster(text, visibleOffset))
    return refuse("Character deletion requires a caret at a grapheme boundary.");
  if (
    (direction === "backward" && visibleOffset === 0) ||
    (direction === "forward" && visibleOffset === text.length)
  )
    return refuse("There is no visible character to delete in this paragraph.");
  let boundary = visibleOffset + (direction === "backward" ? -1 : 1);
  while (splitsGraphemeCluster(text, boundary)) boundary += direction === "backward" ? -1 : 1;
  const position = physicalGaps.at(boundary);
  if (position === undefined) panic("Canonical deletion lost a grapheme boundary.");
  return Result.ok(
    direction === "backward"
      ? { from: position, to: selection.to }
      : { from: selection.from, to: position },
  );
};
