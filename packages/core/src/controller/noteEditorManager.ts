/** Framework-neutral lifecycle owner for editable footnote and endnote stories. */
import { EditorState } from "prosemirror-state";
import { type DirectEditorProps, EditorView } from "prosemirror-view";

import { OP_STORIES } from "@stll/docx-core/ops";

import { CanonicalSessionError } from "../canonical/canonicalSession";
import {
  isSeparatorEndnote,
  isSeparatorFootnote,
} from "../docx/footnoteParser";
import {
  createNativeNavigationPlugin,
  createNativeSelectionRefreshPlugin,
} from "../prosemirror/nativeSelection";
import { noteReferenceTransactionIssue } from "../prosemirror/noteReferenceOccurrences";
import { createDocumentNumberingPlugin } from "../prosemirror/plugins/documentNumbering";
import { createDocumentStyleContextPlugin } from "../prosemirror/plugins/documentStyles";
import {
  CANONICAL_GAP,
  type CanonicalRefusalGround,
} from "../types/canonicalCapabilities";
import type {
  Document,
  Endnote,
  Footnote,
  StyleDefinitions,
} from "../types/document";
import {
  type DeepReadonly,
  readModel,
  type ReadonlyDocument,
} from "../types/documentModel";
import type { NoteStoryKey } from "../types/editor-story";
import {
  createCanonicalStoryEditor,
  refreshCanonicalStoryView,
} from "./canonicalStoryEditor";
import type { HiddenEditorApi } from "./hiddenEditorApi";

export type { NoteStoryKey } from "../types/editor-story";
export type NoteStoryKind = NoteStoryKey["kind"];

export type NoteEditorTransaction = NoteStoryKey & {
  docChanged: boolean;
  selectionChanged: boolean;
  view: EditorView;
};

export type NoteEditorManagerDeps = {
  createView?:
    | ((mountNode: HTMLElement, props: DirectEditorProps) => EditorView)
    | undefined;
  getCanonicalApi?: (() => HiddenEditorApi | null) | undefined;
  /**
   * The canonical session's own document, read without a copy (see
   * `HiddenEditorManager.readCanonicalDocument`); null outside a session.
   */
  readCanonicalDocument: () => ReadonlyDocument | null;
  onSessionRefusal?:
    | ((reason: string, ground: CanonicalRefusalGround, error?: Error) => void)
    | undefined;
  getHost: () => HTMLElement | null;
  getStyles: () => DeepReadonly<StyleDefinitions> | null | undefined;
  onTransaction?: ((transaction: NoteEditorTransaction) => void) | undefined;
};

const createEditorView = (
  mountNode: HTMLElement,
  props: DirectEditorProps,
): EditorView => new EditorView(mountNode, props);

export type NoteEditorManager = {
  activate: (story: NoteStoryKey | null) => EditorView | null;
  destroy: () => void;
  getActive: () => NoteStoryKey | null;
  getView: (story: NoteStoryKey) => EditorView | null;
  listStories: () => NoteStoryKey[];
  /**
   * The session's document with every story's edits. It is read-only: a
   * writer copies it first (`cloneDocumentWithParagraphPropertySources`).
   */
  snapshotDocument: () => ReadonlyDocument;
  sync: () => void;
};

type NoteStory = Footnote | Endnote;

type MountedView = {
  mountNode: HTMLElement;
  note: NoteStoryKey;
  view: EditorView;
};

const storyMapKey = ({ kind, noteId }: NoteStoryKey): string =>
  `${kind}:${String(noteId)}`;

const isNormalFootnote = (note: Footnote): boolean =>
  !isSeparatorFootnote(note);
const isNormalEndnote = (note: Endnote): boolean => !isSeparatorEndnote(note);

export const enumerateDocumentNoteStories = (
  document: Document | null,
): NoteStoryKey[] => [
  ...(document?.package.footnotes ?? [])
    .filter(isNormalFootnote)
    .map(({ id }) => ({ kind: "footnote" as const, noteId: id })),
  ...(document?.package.endnotes ?? [])
    .filter(isNormalEndnote)
    .map(({ id }) => ({ kind: "endnote" as const, noteId: id })),
];

const resolveNote = (
  document: Document | null,
  story: NoteStoryKey,
): NoteStory | null => {
  const notes =
    story.kind === "footnote"
      ? document?.package.footnotes
      : document?.package.endnotes;
  return notes?.find(({ id }) => id === story.noteId) ?? null;
};

export const createNoteEditorManager = (
  deps: NoteEditorManagerDeps,
): NoteEditorManager => {
  const mounted = new Map<string, MountedView>();
  let active: NoteStoryKey | null = null;

  const updateVisibility = (): void => {
    const activeKey = active ? storyMapKey(active) : null;
    for (const [key, story] of mounted) {
      story.mountNode.hidden = key !== activeKey;
    }
  };

  const destroyMounted = (story: MountedView): void => {
    story.view.destroy();
    story.mountNode.remove();
  };

  const destroy = (): void => {
    for (const story of mounted.values()) destroyMounted(story);
    mounted.clear();
    active = null;
  };

  const sync = (): void => {
    const host = deps.getHost();
    if (!host) return;
    // Pending IME input owns its view until the shared session commits.
    if (!deps.getCanonicalApi?.()?.getCanonicalStoryProjection(OP_STORIES.MAIN))
      return;
    const owned = deps.readCanonicalDocument();
    const document = owned && readModel(owned);
    const styles = deps.getStyles();
    const numbering = document?.package.numbering;
    const wanted = new Map(
      enumerateDocumentNoteStories(document).map(
        (story) => [storyMapKey(story), story] as const,
      ),
    );
    const activeKey = active ? storyMapKey(active) : null;

    for (const [key, story] of mounted) {
      if (wanted.has(key)) continue;
      destroyMounted(story);
      mounted.delete(key);
    }

    for (const [key, storyKey] of wanted) {
      if (!mounted.has(key) && key !== activeKey) continue;
      if (!resolveNote(document, storyKey)) continue;
      const existing = mounted.get(key);
      if (existing) {
        if (existing.mountNode.parentElement !== host)
          host.append(existing.mountNode);
        refreshCanonicalStoryView({
          view: existing.view,
          story: { kind: storyKey.kind, id: storyKey.noteId },
          api: deps.getCanonicalApi?.(),
        });
        continue;
      }

      const canonical = createCanonicalStoryEditor({
        story: { kind: storyKey.kind, id: storyKey.noteId },
        getView: () => view,
        getApi: () => deps.getCanonicalApi?.() ?? null,
        enabled: () => true,
        onRefusal: deps.onSessionRefusal,
        onSelectionChange: () =>
          deps.onTransaction?.({
            ...storyKey,
            view,
            docChanged: false,
            selectionChanged: true,
          }),
      });
      const canonicalProjection = deps
        .getCanonicalApi?.()
        ?.getCanonicalStoryProjection({
          kind: storyKey.kind,
          id: storyKey.noteId,
        });
      if (!canonicalProjection) continue;
      const mountNode = host.ownerDocument.createElement("div");
      mountNode.dataset["noteKind"] = storyKey.kind;
      mountNode.dataset["noteId"] = String(storyKey.noteId);
      host.append(mountNode);
      const view = (deps.createView ?? createEditorView)(mountNode, {
        ...canonical.props,
        state: EditorState.create({
          doc: canonicalProjection,
          plugins: [
            createDocumentNumberingPlugin(numbering),
            createDocumentStyleContextPlugin(styles),
            createNativeSelectionRefreshPlugin(),
            createNativeNavigationPlugin(),
          ],
        }),
        dispatchTransaction(transaction) {
          if (canonical.dispatch(transaction)) return;
          const nextState = view.state.apply(transaction);
          const docChanged = !nextState.doc.eq(view.state.doc);
          view.updateState(nextState);
          const noteIssue = noteReferenceTransactionIssue(transaction);
          if (noteIssue)
            deps.onSessionRefusal?.(
              noteIssue.message,
              { gap: CANONICAL_GAP.dispatch },
              noteIssue,
            );
          deps.onTransaction?.({
            ...storyKey,
            docChanged,
            selectionChanged: transaction.selectionSet,
            view,
          });
        },
      });
      mounted.set(key, { mountNode, note: storyKey, view });
    }
    if (active && !wanted.has(storyMapKey(active))) active = null;
    updateVisibility();
  };

  return {
    activate: (story) => {
      active = story;
      sync();
      if (!story) return null;
      const view = mounted.get(storyMapKey(story))?.view ?? null;
      if (view) return view;
      active = null;
      updateVisibility();
      return null;
    },
    destroy,
    getActive: () => active,
    getView: (story) => mounted.get(storyMapKey(story))?.view ?? null,
    listStories: () => {
      const owned = deps.readCanonicalDocument();
      return enumerateDocumentNoteStories(owned && readModel(owned));
    },
    snapshotDocument: () => {
      const canonical = deps.getCanonicalApi?.()?.getCanonicalDocument();
      if (canonical) return canonical;
      throw new CanonicalSessionError({
        ground: { gap: CANONICAL_GAP.authorityRouting },
        reason: "refused",
        message: "Canonical story snapshot is unavailable.",
      });
    },
    sync,
  };
};
