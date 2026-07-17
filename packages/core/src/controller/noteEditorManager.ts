/** Framework-neutral lifecycle owner for editable footnote and endnote stories. */

import { EditorState } from "prosemirror-state";
import type { EditorState as EditorStateT } from "prosemirror-state";
import type { Plugin } from "prosemirror-state";
import type { Node as PMNode } from "prosemirror-model";
import { EditorView } from "prosemirror-view";

import { isSeparatorEndnote, isSeparatorFootnote } from "../docx/footnoteParser";
import { proseDocToBlocks } from "../prosemirror/conversion/fromProseDoc";
import { footnoteToProseDoc } from "../prosemirror/conversion/toProseDoc";
import { ExtensionManager } from "../prosemirror/extensions/ExtensionManager";
import { ensureBaseDirectionInState } from "../prosemirror/extensions/features/AutoBidiDetectionExtension";
import { ensureParaIdsInState } from "../prosemirror/extensions/features/ParaIdAllocatorExtension";
import { createStarterKit } from "../prosemirror/extensions/StarterKit";
import { createDocumentStylesPlugin } from "../prosemirror/plugins/documentStyles";
import { schema } from "../prosemirror/schema";
import type {
  BlockContent,
  Document,
  Endnote,
  Footnote,
  StyleDefinitions,
  Theme,
} from "../types/document";

export type NoteStoryKind = "footnote" | "endnote";

export type NoteStoryKey = {
  kind: NoteStoryKind;
  noteId: number;
};

export type NoteEditorTransaction = NoteStoryKey & {
  docChanged: boolean;
  selectionChanged: boolean;
  view: EditorView;
};

export type NoteEditorManagerDeps = {
  getDocument: () => Document | null;
  getHost: () => HTMLElement | null;
  getPlugins?: (() => Plugin[]) | undefined;
  getStyles: () => StyleDefinitions | null | undefined;
  getTheme: () => Theme | null | undefined;
  onTransaction?: ((transaction: NoteEditorTransaction) => void) | undefined;
};

export type NoteEditorManager = {
  activate: (story: NoteStoryKey | null) => EditorView | null;
  destroy: () => void;
  getActive: () => NoteStoryKey | null;
  getView: (story: NoteStoryKey) => EditorView | null;
  listStories: () => NoteStoryKey[];
  snapshotDocument: (document: Document) => Document;
  sync: () => void;
};

type NoteStory = Footnote | Endnote;
const EMPTY_PLUGINS: Plugin[] = [];

type MountedView = {
  appliedContent: BlockContent[];
  appliedNote: NoteStory;
  appliedPlugins: Plugin[];
  appliedProseDocument: PMNode;
  appliedStyles: StyleDefinitions | null | undefined;
  appliedTheme: Theme | null | undefined;
  dirty: boolean;
  manager: ExtensionManager;
  mountNode: HTMLElement;
  note: NoteStoryKey;
  view: EditorView;
};

const storyMapKey = ({ kind, noteId }: NoteStoryKey): string => `${kind}:${String(noteId)}`;

const isNormalFootnote = (note: Footnote): boolean => !isSeparatorFootnote(note);
const isNormalEndnote = (note: Endnote): boolean => !isSeparatorEndnote(note);

export const enumerateDocumentNoteStories = (document: Document | null): NoteStoryKey[] => [
  ...(document?.package.footnotes ?? [])
    .filter(isNormalFootnote)
    .map(({ id }) => ({ kind: "footnote" as const, noteId: id })),
  ...(document?.package.endnotes ?? [])
    .filter(isNormalEndnote)
    .map(({ id }) => ({ kind: "endnote" as const, noteId: id })),
];

const resolveNote = (document: Document | null, story: NoteStoryKey): NoteStory | null => {
  const notes =
    story.kind === "footnote" ? document?.package.footnotes : document?.package.endnotes;
  return notes?.find(({ id }) => id === story.noteId) ?? null;
};

const noteToProseDocument = (
  note: NoteStory,
  styles: StyleDefinitions | null | undefined,
  theme: Theme | null | undefined,
) => {
  const options: { styles?: StyleDefinitions; theme?: Theme | null } = {};
  if (styles) options.styles = styles;
  if (theme !== undefined) options.theme = theme;
  return footnoteToProseDoc(note.content, options);
};

const buildInitialState = (
  document: PMNode,
  styles: StyleDefinitions | null | undefined,
  manager: ExtensionManager,
  externalPlugins: Plugin[],
): EditorStateT => {
  return ensureBaseDirectionInState(
    ensureParaIdsInState(
      EditorState.create({
        doc: document,
        schema,
        plugins: [...manager.getPlugins(), ...externalPlugins, createDocumentStylesPlugin(styles)],
      }),
    ),
  );
};

export const createNoteEditorManager = (deps: NoteEditorManagerDeps): NoteEditorManager => {
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
    story.manager.destroy();
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
    const document = deps.getDocument();
    const styles = deps.getStyles();
    const theme = deps.getTheme();
    const externalPlugins = deps.getPlugins?.() ?? EMPTY_PLUGINS;
    const wanted = new Map(
      enumerateDocumentNoteStories(document).map((story) => [storyMapKey(story), story] as const),
    );
    const activeKey = active ? storyMapKey(active) : null;

    for (const [key, story] of mounted) {
      if (wanted.has(key)) continue;
      destroyMounted(story);
      mounted.delete(key);
    }

    for (const [key, storyKey] of wanted) {
      if (!mounted.has(key) && key !== activeKey) continue;
      const note = resolveNote(document, storyKey);
      if (!note) continue;
      const existing = mounted.get(key);
      if (existing) {
        if (existing.mountNode.parentElement !== host) host.append(existing.mountNode);
        const nextProseDocument = noteToProseDocument(note, styles, theme);
        const contextIsCurrent =
          existing.appliedStyles === styles &&
          existing.appliedTheme === theme &&
          existing.appliedPlugins === externalPlugins;
        const referencesAreCurrent =
          existing.appliedNote === note && existing.appliedContent === note.content;
        const contentIsCurrent =
          referencesAreCurrent || existing.appliedProseDocument.eq(nextProseDocument);
        if (contentIsCurrent && contextIsCurrent) {
          existing.appliedNote = note;
          existing.appliedContent = note.content;
          existing.appliedProseDocument = nextProseDocument;
          continue;
        }
        existing.view.updateState(
          buildInitialState(nextProseDocument, styles, existing.manager, externalPlugins),
        );
        existing.appliedNote = note;
        existing.appliedContent = note.content;
        existing.appliedStyles = styles;
        existing.appliedTheme = theme;
        existing.appliedPlugins = externalPlugins;
        existing.appliedProseDocument = nextProseDocument;
        existing.dirty = false;
        continue;
      }

      const manager = new ExtensionManager(createStarterKit());
      manager.buildSchema();
      manager.initializeRuntime();
      const mountNode = host.ownerDocument.createElement("div");
      mountNode.dataset["noteKind"] = storyKey.kind;
      mountNode.dataset["noteId"] = String(storyKey.noteId);
      host.append(mountNode);
      const proseDocument = noteToProseDocument(note, styles, theme);
      const view = new EditorView(mountNode, {
        state: buildInitialState(proseDocument, styles, manager, externalPlugins),
        dispatchTransaction(transaction) {
          view.updateState(view.state.apply(transaction));
          const mountedStory = mounted.get(key);
          if (mountedStory && transaction.docChanged) mountedStory.dirty = true;
          deps.onTransaction?.({
            ...storyKey,
            docChanged: transaction.docChanged,
            selectionChanged: transaction.selectionSet,
            view,
          });
        },
      });
      mounted.set(key, {
        appliedContent: note.content,
        appliedNote: note,
        appliedPlugins: externalPlugins,
        appliedProseDocument: proseDocument,
        appliedStyles: styles,
        appliedTheme: theme,
        dirty: false,
        manager,
        mountNode,
        note: storyKey,
        view,
      });
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
    listStories: () => enumerateDocumentNoteStories(deps.getDocument()),
    snapshotDocument: (document) => {
      let footnotes = document.package.footnotes;
      let endnotes = document.package.endnotes;
      let footnotesChanged = false;
      let endnotesChanged = false;
      for (const story of mounted.values()) {
        if (!story.dirty) continue;
        if (story.note.kind === "footnote") {
          const index = footnotes?.findIndex(({ id }) => id === story.note.noteId) ?? -1;
          const current = index === -1 ? null : footnotes?.at(index);
          if (!current) continue;
          if (!footnotesChanged) {
            footnotes = [...(footnotes ?? [])];
            footnotesChanged = true;
          }
          const updated: Footnote = { ...current, content: proseDocToBlocks(story.view.state.doc) };
          footnotes[index] = updated;
          story.appliedNote = updated;
          story.appliedContent = updated.content;
          story.appliedProseDocument = noteToProseDocument(
            updated,
            story.appliedStyles,
            story.appliedTheme,
          );
          continue;
        }
        const index = endnotes?.findIndex(({ id }) => id === story.note.noteId) ?? -1;
        const current = index === -1 ? null : endnotes?.at(index);
        if (!current) continue;
        if (!endnotesChanged) {
          endnotes = [...(endnotes ?? [])];
          endnotesChanged = true;
        }
        const updated: Endnote = { ...current, content: proseDocToBlocks(story.view.state.doc) };
        endnotes[index] = updated;
        story.appliedNote = updated;
        story.appliedContent = updated.content;
        story.appliedProseDocument = noteToProseDocument(
          updated,
          story.appliedStyles,
          story.appliedTheme,
        );
      }
      if (!footnotesChanged && !endnotesChanged) return document;
      return {
        ...document,
        package: {
          ...document.package,
          ...(footnotesChanged && footnotes ? { footnotes } : {}),
          ...(endnotesChanged && endnotes ? { endnotes } : {}),
        },
      };
    },
    sync,
  };
};
