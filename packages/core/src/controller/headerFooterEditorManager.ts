import { EditorState } from "prosemirror-state";
/**
 * Framework-neutral lifecycle owner for persistent header/footer EditorViews.
 *
 * Each distinct relationship id owns one off-screen ProseMirror view. Painted
 * header/footer instances that share that relationship therefore share live
 * editor state, and adapters only decide when to synchronize or destroy the
 * manager.
 */
import { type DirectEditorProps, EditorView } from "prosemirror-view";

import { OP_STORIES } from "@stll/docx-core/ops";

import { CanonicalSessionError } from "../canonical/canonicalSession";
import {
  createNativeNavigationPlugin,
  createNativeSelectionRefreshPlugin,
} from "../prosemirror/nativeSelection";
import { noteReferenceTransactionIssue } from "../prosemirror/noteReferenceOccurrences";
import { createDocumentNumberingPlugin } from "../prosemirror/plugins/documentNumbering";
import { createDocumentStyleContextPlugin } from "../prosemirror/plugins/documentStyles";
import {
  CANONICAL_CONDITION,
  CANONICAL_GAP,
  type CanonicalRefusalGround,
} from "../types/canonicalCapabilities";
import type {
  Document,
  HeaderFooter,
  StyleDefinitions,
} from "../types/document";
import {
  type DeepReadonly,
  readModel,
  type ReadonlyDocument,
} from "../types/documentModel";
import {
  createCanonicalStoryEditor,
  refreshCanonicalStoryView,
} from "./canonicalStoryEditor";
import type { HiddenEditorApi } from "./hiddenEditorApi";

export type HeaderFooterPartKind = "header" | "footer";

export type HeaderFooterPartKey = {
  kind: HeaderFooterPartKind;
  rId: string;
};

export type HeaderFooterEditorTransaction = {
  docChanged: boolean;
  kind: HeaderFooterPartKind;
  rId: string;
  selectionChanged: boolean;
  view: EditorView;
};

export type HeaderFooterEditorManagerDeps = {
  createView?:
    | ((mountNode: HTMLElement, props: DirectEditorProps) => EditorView)
    | undefined;
  getReadOnly?: (() => boolean) | undefined;
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
  onTransaction?:
    ((transaction: HeaderFooterEditorTransaction) => void) | undefined;
};

const createEditorView = (
  mountNode: HTMLElement,
  props: DirectEditorProps,
): EditorView => new EditorView(mountNode, props);

export type HeaderFooterEditorManager = {
  destroy: () => void;
  getView: (rId: string) => EditorView | null;
  listSlots: () => HeaderFooterPartKey[];
  /**
   * The session's document with every story's edits. It is read-only: a
   * writer copies it first (`cloneDocumentWithParagraphPropertySources`).
   */
  snapshotDocument: () => ReadonlyDocument;
  sync: () => void;
};

export const enumerateHeaderFooterParts = ({
  headers,
  footers,
}: {
  headers: Map<string, HeaderFooter> | undefined;
  footers: Map<string, HeaderFooter> | undefined;
}): HeaderFooterPartKey[] => {
  const parts: HeaderFooterPartKey[] = [];
  if (headers) {
    for (const rId of headers.keys()) {
      parts.push({ kind: "header", rId });
    }
  }
  if (footers) {
    for (const rId of footers.keys()) {
      if (!headers?.has(rId)) {
        parts.push({ kind: "footer", rId });
      }
    }
  }
  return parts;
};

export const enumerateDocumentHeaderFooterParts = (
  document: Document | null,
): HeaderFooterPartKey[] => {
  if (!document?.package) {
    return [];
  }
  return enumerateHeaderFooterParts({
    headers: document.package.headers,
    footers: document.package.footers,
  });
};

const resolveHeaderFooter = (
  document: Document | null,
  { kind, rId }: HeaderFooterPartKey,
): HeaderFooter | null => {
  const parts =
    kind === "header" ? document?.package.headers : document?.package.footers;
  return parts?.get(rId) ?? null;
};

type MountedView = {
  kind: HeaderFooterPartKind;
  mountNode: HTMLElement;
  rId: string;
  view: EditorView;
};

export const createHeaderFooterEditorManager = (
  deps: HeaderFooterEditorManagerDeps,
): HeaderFooterEditorManager => {
  const mounted = new Map<string, MountedView>();

  const destroyMounted = (part: MountedView): void => {
    part.view.destroy();
    part.mountNode.remove();
  };

  const destroy = (): void => {
    for (const part of mounted.values()) {
      destroyMounted(part);
    }
    mounted.clear();
  };

  const sync = (): void => {
    const host = deps.getHost();
    if (!host) {
      return;
    }

    const getReadOnly = deps.getReadOnly;
    if (getReadOnly) {
      const readOnly = getReadOnly();
      for (const part of mounted.values()) {
        const expectedEditable = !readOnly;
        if (part.view.editable === expectedEditable) {
          continue;
        }
        part.view.setProps({
          editable: () => !getReadOnly(),
        });
      }
    }

    // Pending IME input owns its view until the shared session commits.
    if (!deps.getCanonicalApi?.()?.getCanonicalStoryProjection(OP_STORIES.MAIN))
      return;

    const owned = deps.readCanonicalDocument();
    const document = owned && readModel(owned);
    const styles = deps.getStyles();
    const numbering = document?.package.numbering;
    const wanted = new Map(
      enumerateDocumentHeaderFooterParts(document).map(
        (part) => [part.rId, part] as const,
      ),
    );

    for (const [rId, part] of mounted) {
      if (wanted.has(rId)) {
        continue;
      }
      destroyMounted(part);
      mounted.delete(rId);
    }

    for (const part of wanted.values()) {
      if (!resolveHeaderFooter(document, part)) {
        continue;
      }

      const existing = mounted.get(part.rId);
      if (existing) {
        if (existing.mountNode.parentElement !== host) {
          host.append(existing.mountNode);
        }
        refreshCanonicalStoryView({
          view: existing.view,
          story: { kind: part.kind, rId: part.rId },
          api: deps.getCanonicalApi?.(),
        });
        continue;
      }

      const canonical = createCanonicalStoryEditor({
        story: { kind: part.kind, rId: part.rId },
        getView: () => view,
        getApi: () => deps.getCanonicalApi?.() ?? null,
        enabled: () => true,
        onRefusal: deps.onSessionRefusal,
        onSelectionChange: () =>
          deps.onTransaction?.({
            kind: part.kind,
            rId: part.rId,
            view,
            docChanged: false,
            selectionChanged: true,
          }),
      });
      const canonicalProjection = deps
        .getCanonicalApi?.()
        ?.getCanonicalStoryProjection({ kind: part.kind, rId: part.rId });
      if (!canonicalProjection) continue;

      const mountNode = host.ownerDocument.createElement("div");
      mountNode.dataset["hfRId"] = part.rId;
      mountNode.dataset["hfKind"] = part.kind;
      host.append(mountNode);
      const view = (deps.createView ?? createEditorView)(mountNode, {
        ...canonical.props,
        editable: () => !(deps.getReadOnly?.() ?? false),
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
          if ((deps.getReadOnly?.() ?? false) && transaction.docChanged) return;
          if (canonical.dispatch(transaction)) return;
          const nextState = view.state.apply(transaction);
          const docChanged = !nextState.doc.eq(view.state.doc);
          const noteIssue = noteReferenceTransactionIssue(transaction);
          if (noteIssue)
            deps.onSessionRefusal?.(
              noteIssue.message,
              { gap: CANONICAL_GAP.dispatch },
              noteIssue,
            );
          view.updateState(nextState);
          deps.onTransaction?.({
            docChanged,
            kind: part.kind,
            rId: part.rId,
            selectionChanged: transaction.selectionSet,
            view,
          });
        },
      });

      mounted.set(part.rId, {
        kind: part.kind,
        mountNode,
        rId: part.rId,
        view,
      });
    }
  };

  return {
    destroy,
    getView: (rId) => mounted.get(rId)?.view ?? null,
    listSlots: () =>
      [...mounted.values()].map(({ kind, rId }) => ({
        kind,
        rId,
      })),
    snapshotDocument: () => {
      const canonical = deps.getCanonicalApi?.()?.getCanonicalDocument();
      if (canonical) return canonical;
      throw new CanonicalSessionError({
        ground: { condition: CANONICAL_CONDITION.notReady },
        reason: "refused",
        message: "Canonical story snapshot is unavailable.",
      });
    },
    sync,
  };
};
